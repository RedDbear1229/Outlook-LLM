using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Microsoft.Win32;

namespace MailPrioritizer.Installer
{
    /// <summary>
    /// HKCU 레지스트리 등록, manifest 생성, Resiliency 복구 로직.
    /// 관리자 권한 불필요 (HKCU 범위만 사용).
    /// </summary>
    internal static class RegistrationHelper
    {
        private const string AddInSubKey      = @"Software\Microsoft\Office\16.0\Outlook\Addins\MailPrioritizer";
        private const string ResiliencySubKey = @"Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems";
        private const string AddinName        = "MailPrioritizer";

        public enum Status
        {
            NotRegistered,
            RegisteredOk,
            LoadBehaviorDisabled,
            ManifestMissing
        }

        // ── 상태 조회 ────────────────────────────────────────────────────────
        public static Status GetStatus(out string details)
        {
            details = string.Empty;
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(AddInSubKey))
            {
                if (key == null)
                {
                    details = "레지스트리 키 없음";
                    return Status.NotRegistered;
                }

                int lb = Convert.ToInt32(key.GetValue("LoadBehavior", 3));
                if (lb == 0)
                {
                    details = "LoadBehavior = 0 (Outlook에 의해 비활성됨)";
                    return Status.LoadBehaviorDisabled;
                }

                string manifestVal = (key.GetValue("Manifest") as string) ?? string.Empty;
                string manifestPath = ExtractManifestPath(manifestVal);
                if (!string.IsNullOrEmpty(manifestPath) && !File.Exists(manifestPath))
                {
                    details = "Manifest 파일 없음: " + manifestPath;
                    return Status.ManifestMissing;
                }

                details = "LoadBehavior=" + lb;
                return Status.RegisteredOk;
            }
        }

        // ── 설치 ─────────────────────────────────────────────────────────────
        public static void Install(string dllDirectory, Action<string> log)
        {
            string manifestPath = Path.Combine(dllDirectory, "MailPrioritizer.dll.manifest");

            log("Manifest 파일 생성 중...");
            WriteManifest(manifestPath);
            log("  → " + manifestPath);

            log("레지스트리 키 작성 중...");
            WriteRegistry(manifestPath);
            log("  → HKCU\\" + AddInSubKey);

            log("설치 완료.");
        }

        // ── 제거 ─────────────────────────────────────────────────────────────
        public static void Uninstall(Action<string> log)
        {
            log("레지스트리 키 삭제 중...");
            Registry.CurrentUser.DeleteSubKeyTree(AddInSubKey, throwOnMissingSubKey: false);
            log("제거 완료.");
        }

        // ── 복구 ─────────────────────────────────────────────────────────────
        public static void Repair(string dllDirectory, Action<string> log)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(AddInSubKey, writable: true))
            {
                if (key == null)
                {
                    log("레지스트리 키 없음 → 새로 설치합니다.");
                    Install(dllDirectory, log);
                    return;
                }

                // LoadBehavior 복원
                int lb = Convert.ToInt32(key.GetValue("LoadBehavior", 3));
                if (lb != 3)
                {
                    key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    log("LoadBehavior → 3 으로 복원");
                }
                else
                {
                    log("LoadBehavior = 3 (정상)");
                }

                // Manifest 재생성 (없거나 경로가 깨진 경우)
                string manifestVal  = (key.GetValue("Manifest") as string) ?? string.Empty;
                string manifestPath = ExtractManifestPath(manifestVal);
                if (string.IsNullOrEmpty(manifestPath) || !File.Exists(manifestPath))
                {
                    string newPath = Path.Combine(dllDirectory, "MailPrioritizer.dll.manifest");
                    log("Manifest 파일 재생성 중...");
                    WriteManifest(newPath);
                    WriteRegistry(newPath);
                    log("  → " + newPath);
                }
                else
                {
                    log("Manifest 파일 정상: " + manifestPath);
                }
            }

            // Resiliency 비활성 목록 정리
            ClearResiliencyEntries(log);
            log("복구 완료.");
        }

        // ── Resiliency 정리 ───────────────────────────────────────────────────
        public static void ClearResiliencyEntries(Action<string> log)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(ResiliencySubKey, writable: true))
            {
                if (key == null) return;
                foreach (string name in key.GetValueNames())
                {
                    try
                    {
                        byte[] bytes = key.GetValue(name) as byte[];
                        if (bytes == null) continue;
                        string text = Encoding.Unicode.GetString(bytes);
                        if (text.IndexOf(AddinName, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            key.DeleteValue(name);
                            log("Resiliency 비활성 항목 제거: " + name);
                        }
                    }
                    catch { /* 개별 항목 오류는 무시 */ }
                }
            }
        }

        // ── DLL 디렉터리 자동 탐색 ───────────────────────────────────────────
        /// <summary>
        /// MailPrioritizer.dll이 있는 디렉터리를 자동으로 찾습니다.
        /// 1순위: 이 EXE와 같은 폴더 (배포 환경)
        /// 2순위: 솔루션 트리의 bin\Debug, bin\Release (개발 환경)
        /// </summary>
        public static string FindDllDirectory()
        {
            const string target = "MailPrioritizer.dll";

            // 1) EXE 위치와 동일 폴더 (publish/ 배포 환경)
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            if (File.Exists(Path.Combine(exeDir, target)))
                return exeDir;

            // 2) 솔루션 트리 탐색 (개발 환경)
            // Installer EXE: solution\MailPrioritizer.Installer\bin\{Config}\
            // Main DLL:      solution\MailPrioritizer\bin\{Config}\
            string[] relPaths = {
                @"..\..\..\MailPrioritizer\bin\Debug",
                @"..\..\..\MailPrioritizer\bin\Release",
            };
            foreach (string rel in relPaths)
            {
                string full = Path.GetFullPath(Path.Combine(exeDir, rel));
                if (File.Exists(Path.Combine(full, target)))
                    return full;
            }

            return null;
        }

        // ── Outlook 프로세스 확인 ─────────────────────────────────────────────
        public static bool IsOutlookRunning()
        {
            return Process.GetProcessesByName("OUTLOOK").Length > 0;
        }

        // ── 내부 헬퍼 ────────────────────────────────────────────────────────
        private static void WriteRegistry(string manifestPath)
        {
            string manifestUri = "file:///" + manifestPath.Replace('\\', '/') + "|vstolocal";
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(AddInSubKey))
            {
                key.SetValue("Description",  "메일 분석 및 우선순위 분류 Add-in");
                key.SetValue("FriendlyName", "MailPrioritizer");
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                key.SetValue("Manifest",     manifestUri);
            }
        }

        private static string ExtractManifestPath(string manifestValue)
        {
            if (string.IsNullOrEmpty(manifestValue)) return null;
            return manifestValue
                .Replace("file:///", string.Empty)
                .Replace("|vstolocal", string.Empty)
                .Trim()
                .Replace('/', '\\');
        }

        /// <summary>
        /// VSTO 런타임이 DLL을 로드하기 위한 최소 manifest XML을 생성합니다.
        /// |vstolocal 플래그 사용 시 hash 검증이 생략되므로 DLL hash 불필요.
        /// </summary>
        private static void WriteManifest(string path)
        {
            const string xml =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<asmv1:assembly manifestVersion=""1.0""
  xmlns:asmv1=""urn:schemas-microsoft-com:asm.v1""
  xmlns=""urn:schemas-microsoft-com:asm.v2""
  xmlns:asmv2=""urn:schemas-microsoft-com:asm.v2""
  xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
  xmlns:co.v1=""urn:schemas-microsoft-com:clickonce.v1""
  xsi:schemaLocation=""urn:schemas-microsoft-com:asm.v1 assembly.adaptive.xsd"">
  <asmv1:assemblyIdentity name=""MailPrioritizer.dll"" version=""1.0.0.0""
    publicKeyToken=""0000000000000000"" language=""neutral""
    processorArchitecture=""msil"" type=""win32"" />
  <description xmlns=""urn:schemas-microsoft-com:asm.v1"">MailPrioritizer</description>
  <application />
  <entryPoint>
    <co.v1:customHostSpecified />
  </entryPoint>
  <trustInfo>
    <security>
      <applicationRequestMinimum>
        <PermissionSet Unrestricted=""true"" ID=""Custom"" SameSite=""site"" />
        <defaultAssemblyRequest permissionSetReference=""Custom"" />
      </applicationRequestMinimum>
      <requestedPrivileges xmlns=""urn:schemas-microsoft-com:asm.v3"">
        <requestedExecutionLevel level=""asInvoker"" uiAccess=""false"" />
      </requestedPrivileges>
    </security>
  </trustInfo>
  <vstav3:addIn xmlns:vstav3=""urn:schemas-microsoft-com:vsta.v3"">
    <vstav3:entryPointsCollection>
      <vstav3:entryPoints>
        <vstav3:entryPoint class=""MailPrioritizer.ThisAddIn"">
          <assemblyIdentity name=""MailPrioritizer"" version=""1.0.0.0""
            language=""neutral"" processorArchitecture=""msil"" />
        </vstav3:entryPoint>
      </vstav3:entryPoints>
    </vstav3:entryPointsCollection>
    <vstav3:update enabled=""false"" />
    <vstav3:application>
      <vstov4:customizations xmlns:vstov4=""urn:schemas-microsoft-com:vsto.v4"">
        <vstov4:customization>
          <vstov4:appAddIn application=""Outlook"" loadBehavior=""3"" keyName=""MailPrioritizer"">
            <vstov4:friendlyName>MailPrioritizer</vstov4:friendlyName>
            <vstov4:description>메일 분석 및 우선순위 분류 Add-in</vstov4:description>
            <vstov4.1:ribbonTypes xmlns:vstov4.1=""urn:schemas-microsoft-com:vsto.v4.1"">
              <vstov4.1:ribbonType>MailPrioritizer.Ribbon.MailRibbon</vstov4.1:ribbonType>
            </vstov4.1:ribbonTypes>
          </vstov4:appAddIn>
        </vstov4:customization>
      </vstov4:customizations>
    </vstav3:application>
  </vstav3:addIn>
</asmv1:assembly>";
            File.WriteAllText(path, xml, Encoding.UTF8);
        }
    }
}
