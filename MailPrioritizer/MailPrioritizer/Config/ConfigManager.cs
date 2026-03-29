using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using MailPrioritizer.Models;
using MailPrioritizer.Utils;
using Newtonsoft.Json;

namespace MailPrioritizer.Config
{
    /// <summary>
    /// 설정 로드/저장 및 API 토큰 DPAPI 암호화 관리.
    /// 설정 파일: %AppData%\MailPrioritizer\config.json
    /// </summary>
    public class ConfigManager
    {
        private static readonly string ConfigDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         "MailPrioritizer");

        private static readonly string ConfigPath = Path.Combine(ConfigDir, "config.json");

        /// <summary>설정 파일에서 로드. 없으면 기본값 생성.</summary>
        public AppConfig Load()
        {
            try
            {
                EnsureDirectory();

                if (!File.Exists(ConfigPath))
                {
                    var defaults = new AppConfig();
                    Save(defaults, plainToken: "");
                    return defaults;
                }

                string json = File.ReadAllText(ConfigPath, Encoding.UTF8);
                var config = JsonConvert.DeserializeObject<AppConfig>(json) ?? new AppConfig();

                // 암호화된 토큰 복호화 (디스크에서 읽은 값 → 메모리 평문)
                if (!string.IsNullOrEmpty(config.Llm.ApiToken))
                {
                    config.Llm.ApiToken = DecryptToken(config.Llm.ApiToken);
                }

                Logger.Info("Config loaded: endpoint=" + config.Llm.Endpoint
                            + ", model=" + config.Llm.ModelName);
                return config;
            }
            catch (Exception ex)
            {
                Logger.Error("ConfigManager.Load failed, returning defaults", ex);
                return new AppConfig();
            }
        }

        /// <summary>설정 저장. plainToken이 null이 아니면 암호화하여 저장.</summary>
        public void Save(AppConfig config, string plainToken = null)
        {
            try
            {
                EnsureDirectory();

                // 저장 전 토큰 암호화
                string tokenToStore = "";
                if (plainToken != null)
                {
                    tokenToStore = string.IsNullOrEmpty(plainToken)
                        ? ""
                        : EncryptToken(plainToken);
                }
                else if (!string.IsNullOrEmpty(config.Llm.ApiToken))
                {
                    // ApiToken은 메모리 내 항상 평문 → 암호화하여 저장
                    tokenToStore = EncryptToken(config.Llm.ApiToken);
                }

                // 저장용 복사본 생성 (원본 훼손 방지)
                var saveConfig = DeepClone(config);
                saveConfig.Llm.ApiToken = tokenToStore;

                string json = JsonConvert.SerializeObject(saveConfig, Formatting.Indented);
                File.WriteAllText(ConfigPath, json, Encoding.UTF8);
                Logger.Info("Config saved");
            }
            catch (Exception ex)
            {
                Logger.Error("ConfigManager.Save failed", ex);
                throw new InvalidOperationException("설정 저장 실패: " + ex.Message, ex);
            }
        }

        private void EnsureDirectory()
        {
            Directory.CreateDirectory(ConfigDir);
        }

        private static string EncryptToken(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return "";
            byte[] data = Encoding.UTF8.GetBytes(plainText);
            byte[] encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(encrypted);
        }

        private static string DecryptToken(string encryptedBase64)
        {
            if (string.IsNullOrEmpty(encryptedBase64)) return "";
            try
            {
                byte[] encrypted = Convert.FromBase64String(encryptedBase64);
                byte[] decrypted = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(decrypted);
            }
            catch (Exception ex)
            {
                Logger.Error("DecryptToken failed, returning empty token", ex);
                return "";
            }
        }

        private static AppConfig DeepClone(AppConfig config)
        {
            string json = JsonConvert.SerializeObject(config);
            return JsonConvert.DeserializeObject<AppConfig>(json);
        }
    }
}
