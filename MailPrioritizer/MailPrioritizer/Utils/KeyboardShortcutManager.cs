using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using MailPrioritizer.Models;

namespace MailPrioritizer.Utils
{
    /// <summary>
    /// R-08: 전역 키보드 단축키 관리자.
    /// RegisterHotKey Win32 API + 메시지 전용 NativeWindow로 구현.
    /// 단축키: Ctrl+Shift+M (선택 메일 분석), Ctrl+Shift+1~4 (우선순위 즉시 변경)
    /// </summary>
    internal sealed class KeyboardShortcutManager : NativeWindow, IDisposable
    {
        private const int WM_HOTKEY   = 0x0312;
        private const int MOD_CONTROL = 0x0002;
        private const int MOD_SHIFT   = 0x0004;

        private const int ID_ANALYZE_SELECTED = 9001;
        private const int ID_PRIORITY_URGENT  = 9002;
        private const int ID_PRIORITY_HIGH    = 9003;
        private const int ID_PRIORITY_NORMAL  = 9004;
        private const int ID_PRIORITY_LOW     = 9005;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private readonly SynchronizationContext _syncContext;
        private bool _disposed;

        public KeyboardShortcutManager(SynchronizationContext syncContext)
        {
            _syncContext = syncContext;

            // 메시지 전용 창 생성 (화면에 표시 안 됨)
            var cp = new CreateParams();
            cp.Parent = new IntPtr(-3); // HWND_MESSAGE
            CreateHandle(cp);

            Register(ID_ANALYZE_SELECTED, MOD_CONTROL | MOD_SHIFT, (int)Keys.M);
            Register(ID_PRIORITY_URGENT,  MOD_CONTROL | MOD_SHIFT, (int)Keys.D1);
            Register(ID_PRIORITY_HIGH,    MOD_CONTROL | MOD_SHIFT, (int)Keys.D2);
            Register(ID_PRIORITY_NORMAL,  MOD_CONTROL | MOD_SHIFT, (int)Keys.D3);
            Register(ID_PRIORITY_LOW,     MOD_CONTROL | MOD_SHIFT, (int)Keys.D4);
        }

        private void Register(int id, int modifiers, int vk)
        {
            if (!RegisterHotKey(Handle, id, modifiers, vk))
                Logger.Warn(string.Format(
                    "KeyboardShortcutManager: RegisterHotKey failed for id={0} vk={1} (another app may own it)",
                    id, vk));
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                int id = m.WParam.ToInt32();
                _syncContext.Post(_ => HandleHotKey(id), null);
            }
            base.WndProc(ref m);
        }

        private void HandleHotKey(int id)
        {
            try
            {
                switch (id)
                {
                    case ID_ANALYZE_SELECTED:
                        Globals.ThisAddIn.TriggerAnalyzeSelected();
                        break;
                    case ID_PRIORITY_URGENT:
                        Globals.ThisAddIn.TriggerChangePriority(Priority.Urgent);
                        break;
                    case ID_PRIORITY_HIGH:
                        Globals.ThisAddIn.TriggerChangePriority(Priority.High);
                        break;
                    case ID_PRIORITY_NORMAL:
                        Globals.ThisAddIn.TriggerChangePriority(Priority.Normal);
                        break;
                    case ID_PRIORITY_LOW:
                        Globals.ThisAddIn.TriggerChangePriority(Priority.Low);
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("KeyboardShortcutManager.HandleHotKey: error id=" + id, ex);
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            if (Handle != IntPtr.Zero)
            {
                UnregisterHotKey(Handle, ID_ANALYZE_SELECTED);
                UnregisterHotKey(Handle, ID_PRIORITY_URGENT);
                UnregisterHotKey(Handle, ID_PRIORITY_HIGH);
                UnregisterHotKey(Handle, ID_PRIORITY_NORMAL);
                UnregisterHotKey(Handle, ID_PRIORITY_LOW);
                DestroyHandle();
            }
        }
    }
}
