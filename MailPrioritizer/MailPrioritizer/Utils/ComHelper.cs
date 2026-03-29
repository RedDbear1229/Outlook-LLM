using System;
using System.Runtime.InteropServices;

namespace MailPrioritizer.Utils
{
    /// <summary>
    /// Outlook COM 객체 해제 유틸리티.
    /// VSTO에서 COM 객체를 해제하지 않으면 Outlook이 종료되지 않거나 메모리 누수가 발생한다.
    /// 모든 Outlook COM 접근은 try/finally 블록에서 이 클래스를 사용해야 한다.
    /// </summary>
    public static class ComHelper
    {
        /// <summary>COM 객체를 안전하게 해제한다. null이면 무시.</summary>
        public static void Release(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch (Exception)
                {
                    // 이미 해제된 객체 무시
                }
            }
        }

        /// <summary>
        /// 여러 COM 객체를 역순으로 해제한다 (자식 → 부모 순서).
        /// 예: ReleaseAll(prop, props, folder, folders)
        /// </summary>
        public static void ReleaseAll(params object[] comObjects)
        {
            if (comObjects == null) return;
            for (int i = comObjects.Length - 1; i >= 0; i--)
            {
                Release(comObjects[i]);
            }
        }
    }
}
