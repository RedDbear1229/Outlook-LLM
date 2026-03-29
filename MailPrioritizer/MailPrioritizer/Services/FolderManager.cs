using System;
using MailPrioritizer.Models;
using MailPrioritizer.Utils;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailPrioritizer.Services
{
    /// <summary>
    /// Outlook 우선순위 폴더 생성/조회 및 메일 이동 관리.
    /// 기본 저장소(Default Store)의 받은편지함 아래에 우선순위 하위 폴더를 생성한다.
    /// 모든 COM 객체는 ComHelper로 명시적 해제.
    /// </summary>
    public class FolderManager
    {
        private readonly Outlook.Application _app;
        private AppConfig _config;

        public FolderManager(Outlook.Application app, AppConfig config)
        {
            _app = app;
            _config = config;
        }

        public void ReloadConfig(AppConfig newConfig)
        {
            _config = newConfig;
        }

        /// <summary>대상 저장소의 받은편지함을 반환한다. 호출자가 ComHelper.Release()로 해제.</summary>
        public Outlook.MAPIFolder GetTargetInbox()
        {
            Outlook.NameSpace session = null;
            Outlook.Store store = null;
            try
            {
                session = _app.Session;
                string storeId = _config.Processing.TargetStoreId;
                if (string.IsNullOrEmpty(storeId))
                    return session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                store = session.GetStoreFromID(storeId);
                return store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            }
            catch (Exception ex)
            {
                Logger.Warn("GetTargetInbox: store not found, falling back to default. " + ex.Message);
                // session이 이미 있으면 재사용, 없으면 다시 획득
                if (session == null) session = _app.Session;
                return session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            }
            finally
            {
                ComHelper.ReleaseAll(store, session);
            }
        }

        /// <summary>
        /// 우선순위에 해당하는 폴더를 반환한다. 없으면 생성.
        /// 반환된 MAPIFolder는 호출자가 ComHelper.Release()로 해제해야 한다.
        /// </summary>
        public Outlook.MAPIFolder GetOrCreatePriorityFolder(Priority priority)
        {
            Outlook.MAPIFolder inbox = null;
            Outlook.Folders inboxFolders = null;
            Outlook.MAPIFolder parentFolder = null;
            Outlook.Folders parentFolders = null;

            try
            {
                inbox = GetTargetInbox();
                inboxFolders = inbox.Folders;

                // 상위 폴더 (접두사) 찾기/생성
                parentFolder = FindOrCreateFolder(inboxFolders, _config.Classification.FolderPrefix);
                parentFolders = parentFolder.Folders;

                // 우선순위 하위 폴더
                string folderName = _config.Classification.GetFolderName(priority);
                Outlook.MAPIFolder targetFolder = FindOrCreateFolder(parentFolders, folderName);

                // targetFolder는 호출자가 해제
                return targetFolder;
            }
            finally
            {
                // targetFolder 제외하고 모두 해제
                ComHelper.ReleaseAll(parentFolders, parentFolder, inboxFolders, inbox);
            }
        }

        /// <summary>메일을 우선순위 폴더로 이동한다.</summary>
        public void MoveToFolder(Outlook.MailItem mail, Priority priority)
        {
            Logger.Info("MoveToFolder: moving mail to " + priority + " folder");
            Outlook.MAPIFolder targetFolder = null;
            try
            {
                targetFolder = GetOrCreatePriorityFolder(priority);
                mail.Move(targetFolder);
            }
            finally
            {
                ComHelper.Release(targetFolder);
            }
        }

        /// <summary>
        /// 폴더를 이름으로 찾아 반환. 없으면 생성.
        /// 반환된 MAPIFolder는 호출자가 해제해야 한다.
        /// </summary>
        private static Outlook.MAPIFolder FindOrCreateFolder(
            Outlook.Folders folders, string name)
        {
            // 이름으로 찾기 (인덱서는 없으면 COMException 발생)
            try
            {
                return folders[name];
            }
            catch (Exception)
            {
                Logger.Info("FindOrCreateFolder: creating new folder \"" + name + "\"");
                return folders.Add(name, Outlook.OlDefaultFolders.olFolderInbox);
            }
        }

    }
}
