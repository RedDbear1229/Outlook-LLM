using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MailPrioritizer.Utils;
using Newtonsoft.Json;

namespace MailPrioritizer.Services
{
    /// <summary>
    /// LLM 분석 실패 메일의 EntryID를 보관하고 다음 시작 시 자동 재시도한다.
    /// 큐는 %AppData%/MailPrioritizer/retry_queue.json 에 영속 저장된다.
    /// </summary>
    public class RetryQueue
    {
        private static readonly string QueuePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MailPrioritizer", "retry_queue.json");

        private List<RetryItem> _items;
        private static readonly object _lock = new object();

        public const int MaxRetries = 3;

        public class RetryItem
        {
            public string EntryId { get; set; }
            public DateTime FailedAt { get; set; }
            public int RetryCount { get; set; }
        }

        public RetryQueue()
        {
            _items = LoadFromDisk();
            Logger.Info("RetryQueue: loaded " + _items.Count + " pending items");
        }

        public int Count
        {
            get { lock (_lock) return _items.Count; }
        }

        /// <summary>실패한 메일을 큐에 추가한다. 이미 존재하면 무시.</summary>
        public void Enqueue(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) return;
            lock (_lock)
            {
                if (_items.Any(x => x.EntryId == entryId)) return;
                _items.Add(new RetryItem
                {
                    EntryId = entryId,
                    FailedAt = DateTime.Now,
                    RetryCount = 0
                });
                SaveToDisk();
            }
            Logger.Info("RetryQueue: enqueued entryId=" + entryId + " (total=" + Count + ")");
        }

        /// <summary>성공적으로 분석된 메일을 큐에서 제거한다.</summary>
        public void Remove(string entryId)
        {
            lock (_lock)
            {
                int removed = _items.RemoveAll(x => x.EntryId == entryId);
                if (removed > 0) SaveToDisk();
            }
        }

        /// <summary>현재 큐의 모든 항목을 반환한다 (복사본).</summary>
        public List<RetryItem> GetAll()
        {
            lock (_lock) return new List<RetryItem>(_items);
        }

        /// <summary>재시도 횟수를 1 증가시키고, 최대 횟수 초과 시 큐에서 제거한다.</summary>
        public void IncrementRetry(string entryId)
        {
            lock (_lock)
            {
                var item = _items.FirstOrDefault(x => x.EntryId == entryId);
                if (item == null) return;
                item.RetryCount++;
                if (item.RetryCount >= MaxRetries)
                {
                    _items.Remove(item);
                    Logger.Warn("RetryQueue: max retries exceeded, dropping entryId=" + entryId);
                }
                SaveToDisk();
            }
        }

        private List<RetryItem> LoadFromDisk()
        {
            try
            {
                if (!File.Exists(QueuePath)) return new List<RetryItem>();
                string json = File.ReadAllText(QueuePath);
                return JsonConvert.DeserializeObject<List<RetryItem>>(json) ?? new List<RetryItem>();
            }
            catch (Exception ex)
            {
                Logger.Warn("RetryQueue.LoadFromDisk failed: " + ex.Message);
                return new List<RetryItem>();
            }
        }

        private void SaveToDisk()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_items, Formatting.Indented);
                File.WriteAllText(QueuePath, json);
            }
            catch (Exception ex)
            {
                Logger.Warn("RetryQueue.SaveToDisk failed: " + ex.Message);
            }
        }
    }
}
