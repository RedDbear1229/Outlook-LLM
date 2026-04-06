using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using MailPrioritizer.Models;
using MailPrioritizer.Utils;

namespace MailPrioritizer.Services
{
    /// <summary>
    /// R-05: 분석 결과 로컬 인덱스 (SQLite).
    /// %AppData%/MailPrioritizer/index.db 에 저장.
    /// 통계 조회와 CSV 내보내기를 COM 순회 없이 DB 쿼리로 처리한다.
    /// </summary>
    public class IndexDatabase : IDisposable
    {
        private static readonly string DbPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MailPrioritizer", "index.db");

        private SQLiteConnection _connection;
        private readonly object _lock = new object();

        public IndexDatabase()
        {
            Open();
            EnsureSchema();
            Logger.Info("IndexDatabase: opened at " + DbPath);
        }

        // ──────────────────────────────────────────────────────────────
        // 초기화
        // ──────────────────────────────────────────────────────────────

        private void Open()
        {
            string dir = Path.GetDirectoryName(DbPath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            _connection = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;");
            _connection.Open();
        }

        private void EnsureSchema()
        {
            const string ddl = @"
CREATE TABLE IF NOT EXISTS mails (
    entry_id     TEXT PRIMARY KEY,
    subject      TEXT NOT NULL DEFAULT '',
    sender       TEXT NOT NULL DEFAULT '',
    priority     TEXT NOT NULL DEFAULT 'normal',
    summary      TEXT NOT NULL DEFAULT '',
    reason       TEXT NOT NULL DEFAULT '',
    model_name   TEXT NOT NULL DEFAULT '',
    analyzed_at  TEXT NOT NULL DEFAULT '',
    is_rule_based INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_priority ON mails(priority);
CREATE INDEX IF NOT EXISTS idx_analyzed_at ON mails(analyzed_at);";

            lock (_lock)
            using (var cmd = new SQLiteCommand(ddl, _connection))
                cmd.ExecuteNonQuery();
        }

        // ──────────────────────────────────────────────────────────────
        // 쓰기
        // ──────────────────────────────────────────────────────────────

        /// <summary>
        /// 분석 결과를 인덱스에 저장한다 (INSERT OR REPLACE).
        /// ApplyAnalysisToMail 직후 호출.
        /// </summary>
        public void Upsert(string entryId, string subject, string sender, MailAnalysis analysis)
        {
            if (string.IsNullOrEmpty(entryId)) return;

            const string sql = @"
INSERT OR REPLACE INTO mails
    (entry_id, subject, sender, priority, summary, reason, model_name, analyzed_at, is_rule_based)
VALUES
    (@entryId, @subject, @sender, @priority, @summary, @reason, @modelName, @analyzedAt, @isRuleBased)";

            lock (_lock)
            {
                try
                {
                    using (var cmd = new SQLiteCommand(sql, _connection))
                    {
                        cmd.Parameters.AddWithValue("@entryId",    entryId);
                        cmd.Parameters.AddWithValue("@subject",    subject ?? "");
                        cmd.Parameters.AddWithValue("@sender",     sender ?? "");
                        cmd.Parameters.AddWithValue("@priority",   analysis.Priority.ToString().ToLowerInvariant());
                        cmd.Parameters.AddWithValue("@summary",    analysis.Summary ?? "");
                        cmd.Parameters.AddWithValue("@reason",     analysis.PriorityReason ?? "");
                        cmd.Parameters.AddWithValue("@modelName",  analysis.ModelName ?? "");
                        cmd.Parameters.AddWithValue("@analyzedAt", analysis.AnalyzedAt.ToString("o"));
                        cmd.Parameters.AddWithValue("@isRuleBased", analysis.IsRuleBased ? 1 : 0);
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn("IndexDatabase.Upsert failed for entryId=" + entryId + ": " + ex.Message);
                }
            }
        }

        /// <summary>메일 삭제 또는 재분석 플래그 초기화 시 인덱스에서 제거한다.</summary>
        public void Remove(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) return;
            lock (_lock)
            {
                try
                {
                    using (var cmd = new SQLiteCommand("DELETE FROM mails WHERE entry_id = @id", _connection))
                    {
                        cmd.Parameters.AddWithValue("@id", entryId);
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn("IndexDatabase.Remove failed: " + ex.Message);
                }
            }
        }

        /// <summary>전체 재분석 시 인덱스를 비운다.</summary>
        public void Clear()
        {
            lock (_lock)
            {
                try
                {
                    using (var cmd = new SQLiteCommand("DELETE FROM mails", _connection))
                        cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Logger.Warn("IndexDatabase.Clear failed: " + ex.Message);
                }
            }
        }

        // ──────────────────────────────────────────────────────────────
        // 읽기 — 통계
        // ──────────────────────────────────────────────────────────────

        /// <summary>
        /// DB 기반 통계 조회. COM 순회 없이 SQL COUNT/GROUP BY 사용.
        /// inboxTotal: Outlook COM에서 직접 가져온 전체 메일 수.
        /// </summary>
        public MailProcessor.InboxStats GetStats(int inboxTotal)
        {
            var stats = new MailProcessor.InboxStats { Total = inboxTotal };

            const string sql = @"
SELECT priority, COUNT(*) as cnt
FROM mails
GROUP BY priority";

            lock (_lock)
            {
                try
                {
                    using (var cmd = new SQLiteCommand(sql, _connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string p = reader.GetString(0);
                            int cnt  = reader.GetInt32(1);
                            stats.Analyzed += cnt;
                            switch (p)
                            {
                                case "urgent": stats.Urgent += cnt; break;
                                case "high":   stats.High   += cnt; break;
                                case "normal": stats.Normal += cnt; break;
                                case "low":    stats.Low    += cnt; break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn("IndexDatabase.GetStats failed: " + ex.Message);
                }
            }

            return stats;
        }

        // ──────────────────────────────────────────────────────────────
        // 읽기 — 내보내기
        // ──────────────────────────────────────────────────────────────

        /// <summary>분석된 전체 메일을 CSV 행 목록으로 반환한다.</summary>
        public List<MailIndexRow> GetAllAnalyzed()
        {
            var rows = new List<MailIndexRow>();

            const string sql = @"
SELECT entry_id, subject, sender, priority, summary, reason, analyzed_at
FROM mails
ORDER BY analyzed_at DESC";

            lock (_lock)
            {
                try
                {
                    using (var cmd = new SQLiteCommand(sql, _connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            rows.Add(new MailIndexRow
                            {
                                EntryId    = reader.GetString(0),
                                Subject    = reader.GetString(1),
                                Sender     = reader.GetString(2),
                                Priority   = reader.GetString(3),
                                Summary    = reader.GetString(4),
                                Reason     = reader.GetString(5),
                                AnalyzedAt = reader.GetString(6)
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn("IndexDatabase.GetAllAnalyzed failed: " + ex.Message);
                }
            }

            return rows;
        }

        /// <summary>인덱스에 존재하는 EntryID 집합을 반환한다 (배치 수집 최적화용).</summary>
        public HashSet<string> GetAllEntryIds()
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            lock (_lock)
            {
                try
                {
                    using (var cmd = new SQLiteCommand("SELECT entry_id FROM mails", _connection))
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            ids.Add(reader.GetString(0));
                }
                catch (Exception ex)
                {
                    Logger.Warn("IndexDatabase.GetAllEntryIds failed: " + ex.Message);
                }
            }

            return ids;
        }

        // ──────────────────────────────────────────────────────────────
        // 정리
        // ──────────────────────────────────────────────────────────────

        public void Dispose()
        {
            if (_connection != null)
            {
                try { _connection.Close(); _connection.Dispose(); }
                catch { }
                _connection = null;
            }
        }

        // ──────────────────────────────────────────────────────────────
        // DTO
        // ──────────────────────────────────────────────────────────────

        public class MailIndexRow
        {
            public string EntryId    { get; set; }
            public string Subject    { get; set; }
            public string Sender     { get; set; }
            public string Priority   { get; set; }
            public string Summary    { get; set; }
            public string Reason     { get; set; }
            public string AnalyzedAt { get; set; }
        }
    }
}
