using System;
using Microsoft.Win32; // For 'Registry'
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;

namespace wei_outlook_add_in
{
    class Util
    {
        private const string key = @"HKEY_CURRENT_USER\Software\wei-outlook-add-in";
        private const string LogFileName = @"wei-outlook-add-in.log";

        internal static bool GetFlag(string name)
        {
            bool value = false;
            try
            {
                value = (int)Registry.GetValue(key, name, 0) != 0;
            }
            catch
            {
            }
            return value;
        }

        internal static void SetFlag(string name, bool value)
        {
            Registry.SetValue(key, name, value ? 1 : 0);
        }

        internal static string GetValue(string name)
        {
            string value = "";
            try
            {
                value = (string)Registry.GetValue(key, name, "");
            }
            catch
            {
            }
            return value;
        }

        internal static void SetValue(string name, string value)
        {
            Registry.SetValue(key, name, value);
        }

        internal static bool IsMailItem(object objectToInspect)
        {
            bool isMailItem = false;

            try
            {
                if (String.Equals(
                    (string)objectToInspect.GetType().InvokeMember(
                        "MessageClass",
                        BindingFlags.GetProperty,
                        null,
                        objectToInspect,
                        null),
                    "IPM.Note",
                    StringComparison.Ordinal) == true)
                {
                    isMailItem = true;
                }
            }
            catch (Exception)
            {
            }

            return isMailItem;
        }

        internal static bool IsFolderDefaultOutbox(Outlook.Explorer explorer, Outlook.Folder folder)
        {
            Outlook.Folder outbox = explorer.Application.GetNamespace("MAPI").GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderOutbox) as Outlook.Folder;
            bool isOutbox = folder.EntryID == outbox.EntryID;
            return isOutbox;
        }

        internal static bool IsCurrentFolderDefaultOutbox(Outlook.Explorer explorer)
        {
            Outlook.Folder folder = explorer.CurrentFolder as Outlook.Folder;
            bool result = IsFolderDefaultOutbox(explorer, folder);
            return result;
        }

        internal static void Log(
            string logMessage,
            bool outputTimestamp = true,
            bool newLineBeforeData = false,
            bool newLineAfterData = true,
            string logFileName = LogFileName)
        {
            string basePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string fullPath = basePath + @"\" + logFileName;

            using (StreamWriter w = File.AppendText(fullPath))
            {
                if (outputTimestamp == true)
                {
                    w.Write(
                        "\r\n{0} {1} : ",
                        DateTime.Now.ToLongDateString(),
                        DateTime.Now.ToLongTimeString()
                        );
                }
                if (newLineBeforeData == true)
                {
                    w.Write("\r\n");
                }
                w.Write("{0}", logMessage);
                if (newLineAfterData == true)
                {
                    w.Write("\r\n");
                }
            }
        }

        internal static void BackupEmail(Outlook.Application app, String mailItemEntryId, bool autoFlag)
        {
            Outlook.Folder backupFolder = null;
            try
            {
                Outlook.MailItem mailItem = app.Session.GetItemFromID(mailItemEntryId);
                mailItem.UnRead = false;
                mailItem.Categories = "";
                mailItem.Save();

                if (autoFlag == true)
                {
                    if (IsMailItemHasFlaggedPreviousMail(app, mailItemEntryId) == true)
                    {
                        mailItem.FlagIcon = Outlook.OlFlagIcon.olRedFlagIcon;
                        mailItem.Save();
                    }
                }

                string backupPathName = Util.GetValue("AttachmentsSavingFolder");

                try
                {
                    AddAttachmentLinkToBodyEnd(app, mailItemEntryId, backupPathName);
                    SaveAttachment(app, mailItemEntryId, backupPathName);
                }
                catch (DirectoryNotFoundException)
                {
                    MessageBox.Show(
                        "Save email attachments failed, please check if the folder is exist.",
                        "Save attachments failed");
                }
                DeleteAttachment(app, mailItemEntryId);

                backupFolder = GetBackupFolder(mailItem);
                BackupEmailToFolder(app, mailItemEntryId, backupFolder);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.ToString());
                MessageBox.Show(
                    e.ToString(),
                    e.ToString() + "Backup email failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1);
            }
        }

        internal static void BackupEmailToFolder(Outlook.Application app, String mailItemEntryId, Outlook.Folder folder)
        {
            Outlook.Folder mailCurrentFolder = null;

            try
            {
                Outlook.MailItem mailItem = app.Session.GetItemFromID(mailItemEntryId);
                mailCurrentFolder = mailItem.Parent as Outlook.Folder;
                Debug.Assert(mailCurrentFolder != null);

                if (mailCurrentFolder.EntryID != folder.EntryID)
                {
                    mailItem = app.Session.GetItemFromID(mailItemEntryId);
                    mailItem.Move(folder);
                }
                else
                {
                    // the mail has already been in the target folder we want, so don't move, just close it.
                    mailItem.Close(Outlook.OlInspectorClose.olSave);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                MessageBox.Show(ex.ToString());
            }
        }

        internal static Outlook.Folder GetBackupFolder(Outlook.MailItem mail)
        {
            string autoBackupPstPathNamePattern = GetValue("AutoBackupPstPathNamePattern");
            string autoBackupOutlookDataFileNamePattern = GetValue("AutoBackupOutlookDataFileNamePattern");
            string autoBackupOutlookFolderNamePattern = GetValue("AutoBackupOutlookFolderNamePattern");

            string mailReceivedTime_Year = mail.ReceivedTime.Year.ToString();
            string mailReceivedTime_Month = mail.ReceivedTime.Month.ToString();

            string autoBackupPstPathName = Regex.Replace(autoBackupPstPathNamePattern, @"\{mailReceivedTime_Year\}", mailReceivedTime_Year, RegexOptions.Compiled | RegexOptions.Singleline);
            autoBackupPstPathName = Regex.Replace(autoBackupPstPathName, @"\{mailReceivedTime_Month\}", mailReceivedTime_Month, RegexOptions.Compiled | RegexOptions.Singleline);

            string autoBackupOutlookDataFileName = Regex.Replace(autoBackupOutlookDataFileNamePattern, @"\{mailReceivedTime_Year\}", mailReceivedTime_Year, RegexOptions.Compiled | RegexOptions.Singleline);
            autoBackupOutlookDataFileName = Regex.Replace(autoBackupOutlookDataFileName, @"\{mailReceivedTime_Month\}", mailReceivedTime_Month, RegexOptions.Compiled | RegexOptions.Singleline);

            string autoBackupOutlookFolderName = Regex.Replace(autoBackupOutlookFolderNamePattern, @"\{mailReceivedTime_Year\}", mailReceivedTime_Year, RegexOptions.Compiled | RegexOptions.Singleline);
            autoBackupOutlookFolderName = Regex.Replace(autoBackupOutlookFolderName, @"\{mailReceivedTime_Month\}", mailReceivedTime_Month, RegexOptions.Compiled | RegexOptions.Singleline);

            string autoBackupOutlookFolderPathName = @"\\" + autoBackupOutlookDataFileName + @"\" + autoBackupOutlookFolderName;

            Outlook.Folder targetFolder = GetFolder(mail.Application, autoBackupOutlookFolderPathName);

            if (targetFolder == null)
            {
                Outlook.Store store = GetStore(mail.Application, autoBackupPstPathName, autoBackupOutlookDataFileName, true);
                targetFolder = GetFolder(mail.Application, autoBackupOutlookFolderPathName, true);
            }

            return targetFolder;
        }

        private static bool IsMailItemHasFlaggedPreviousMail(Outlook.Application app, String mailItemEntryId)
        {
            bool answer = false;

            // This example uses only MailItem. Other item types such as MeetingItem and PostItem can participate in the conversation.
            Outlook.MailItem mailItem = app.Session.GetItemFromID(mailItemEntryId);
            if (mailItem is Outlook.MailItem)
            {
                // Determine the store of the mail item.
                Outlook.Folder folder = mailItem.Parent as Outlook.Folder;
                Outlook.Store store = folder.Store;

                if (store.IsConversationEnabled == true)
                {
                    // Obtain a Coversation object.
                    Outlook.Conversation conv = mailItem.GetConversation();
                    // Check for null Conversation
                    Outlook.Table table = conv.GetTable();

                    Debug.WriteLine("Conversation Items Count: " + table.GetRowCount().ToString());
                    Debug.WriteLine("Conversation Items from Table:");

                    table.Columns.Add("FlagRequest");
                    while (!table.EndOfTable)
                    {
                        Outlook.Row nextRow = table.GetNextRow();

                        string msg = nextRow["Subject"] + ", " + nextRow["LastModificationTime"] + ", Flag: " + nextRow["FlagRequest"];
                        Debug.WriteLine(msg);

                        if ((nextRow["FlagRequest"] != "") && (nextRow["FlagRequest"] != null))
                        {
                            answer = true;
                            break;
                        }
                    }
                }
            }
            return answer;
        }

        private static Outlook.Folder AddFolder(dynamic parent, string folderName)
        {
            Outlook.Folders folders = parent.Folders;
            Outlook.Folder newFolder = folders.Add(folderName) as Outlook.Folder;
            return newFolder;
        }

        private static Outlook.Folder GetFolder(Outlook.Application application, string folderPath, bool createIfNotExist = false)
        {
            Outlook.Folder returnFolder = null;

            try
            {
                // Remove leading "\" characters.
                folderPath = folderPath.TrimStart("\\".ToCharArray());

                // Split the folder path into individual folder names.
                string[] folders = folderPath.Split("\\".ToCharArray());

                // Retrieve a reference to the root folder.
                returnFolder = application.Session.Folders[folders[0]] as Outlook.Folder;

                if (returnFolder == null)
                {
                    if (createIfNotExist == true)
                    {
                        returnFolder = AddFolder(application.Session, folders[0]);
                    }
                }

                // If the root folder exists, look in subfolders
                if (returnFolder != null)
                {
                    Outlook.Folders subFolders = null;
                    string folderName;

                    // Look through folder names, skipping the first
                    // folder, which you already retrieved
                    for (int i = 1; i < folders.Length; ++i)
                    {
                        folderName = folders[i];
                        subFolders = returnFolder.Folders;

                        Outlook.Folder childFolder = null;
                        try
                        {
                            childFolder = subFolders[folderName] as Outlook.Folder;
                        }
                        catch
                        {
                            Debug.Assert(childFolder == null);
                            if (createIfNotExist == true)
                            {
                                childFolder = AddFolder(returnFolder, folderName);
                            }
                        }

                        returnFolder = childFolder;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                // Do nothing at all -- just return a null reference.
                returnFolder = null;
            }

            return returnFolder;
        }

        private static Outlook.Store AddStore(Outlook.Application application, string pstPath, string displayName)
        {
            application.Session.AddStoreEx(pstPath, Outlook.OlStoreType.olStoreDefault);

            // Trick to set the 'display name' of a newly created store
            Outlook.Folder lastFolder = application.GetNamespace("MAPI").Folders.GetLast() as Outlook.Folder;
            lastFolder.Name = displayName;

            Outlook.Store store = GetStore(application, pstPath);

            // [TODO] set 'Display reminders and tasks from this folder in the To-Do Bar' checkbox
            Outlook.Folder folder = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderToDo) as Outlook.Folder;

            return store;
        }

        internal static Outlook.Store GetStore(Outlook.Application application, string pstPath, string displayName = "", bool createIfNotExist = false)
        {
            Outlook.Stores stores = application.Session.Stores;
            foreach (Outlook.Store store in stores)
            {
                if (store.FilePath == pstPath)
                {
                    return store;
                }
            }

            if (createIfNotExist == true)
            {
                return AddStore(application, pstPath, displayName);
            }
            else
            {
                return null;
            }
        }

        private static void RemoveStore(Outlook.Application application, string filePath)
        {
            Outlook.Stores stores = application.Session.Stores;
            foreach (Outlook.Store store in stores)
            {
                if (store.FilePath == filePath)
                {
                    Outlook.Folder folder = store.GetRootFolder() as Outlook.Folder;
                    application.Session.RemoveStore(folder);
                }
            }
        }

        internal static string GetFirstAccountSmtpAddress(Outlook.Application application)
        {
            Outlook.Accounts accounts = null;
            string smtpAddress = "";

            try
            {
                accounts = application.Session.Accounts;

                foreach (Outlook.Account account in accounts)
                {
                    smtpAddress = account.SmtpAddress;
                }
            }
            catch
            {
            }

            return smtpAddress;
        }

        internal static string GetSMTPAddress(Outlook.AddressEntry entry)
        {
            Debug.Assert(entry != null);

            string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            // Now we have an AddressEntry representing the Sender
            if (entry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeAgentAddressEntry ||
                entry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
            {
                // Use the ExchangeUser object PrimarySMTPAddress
                Outlook.ExchangeUser exchUser = entry.GetExchangeUser();
                if (exchUser != null)
                {
                    return exchUser.PrimarySmtpAddress;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return entry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
            }
        }

        internal static string GetSenderSMTPAddress(Outlook.MailItem mail)
        {
            Debug.Assert(mail != null);
            if (mail.SenderEmailType == "EX")
            {
                Outlook.AddressEntry sender = mail.Sender;
                if (sender != null)
                {
                    return GetSMTPAddress(sender);
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return mail.SenderEmailAddress;
            }
        }

        private static bool IsInlineImage(Outlook.MailItem mail, Outlook.Attachment attachment)
        {
            const string PR_ATTACH_METHOD = "http://schemas.microsoft.com/mapi/proptag/0x37050003";
            const string PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003";

            Debug.Assert(attachment != null);

            if (mail.BodyFormat == Outlook.OlBodyFormat.olFormatPlain)
            {
                // if this is a plain text email, every attachment is a non-inline attachment
                return false;
            }
            else if (mail.BodyFormat == Outlook.OlBodyFormat.olFormatRichText)
            {
                // if the body format is RTF, the non-embedded attachment would be of the PR_ATTACH_METHOD property is not 6 (ATTACH_OLD)
                if ((int)attachment.PropertyAccessor.GetProperty(PR_ATTACH_METHOD) != 6)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else if (mail.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
            {
                // if the body format is HTML, the non-embedded attachment would be of the PR_ATTACH_FLAGS property is not 4 (ATT_MHTML_REF)
                if ((int)attachment.PropertyAccessor.GetProperty(PR_ATTACH_FLAGS) != 4)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                Debug.Assert(false);
                return true;
            }
        }

        internal static void AddAttachmentLinkToBodyEnd(Outlook.Application app, String mailItemEntryId, string storagePathName)
        {
            if (GetFlag("AddAttachmentsLinkToEmailEndWhenDelete") == true)
            {
                Outlook.MailItem mail = app.Session.GetItemFromID(mailItemEntryId);
                if (mail.Attachments.Count > 0)
                {
                    if (Directory.Exists(storagePathName) == false)
                    {
                        throw new DirectoryNotFoundException();
                    }

                    string attachmentLinks = "";

                    for (int i = 1; i <= mail.Attachments.Count; ++i)
                    {
                        if (IsInlineImage(mail, mail.Attachments[i]) == false)
                        {
                            if (attachmentLinks == "")
                            {
                                attachmentLinks += @"<div>======== Saved attachment links ========<br>";
                            }

                            string fileName = "";
                            try
                            {
                                fileName = mail.Attachments[i].FileName;
                            }
                            catch
                            {
                                // can not get the file name, then ignore this attachment.
                            }

                            if (fileName != "")
                            {
                                attachmentLinks +=
                                    "<a href=\"file:///" +
                                    storagePathName + (storagePathName.EndsWith(@"\") ? "" : @"\") +
                                    mail.Attachments[i].FileName + "\">" +
                                    storagePathName + (storagePathName.EndsWith(@"\") ? "" : @"\") +
                                    mail.Attachments[i].FileName + @"</a><br>";
                            }
                        }
                    }

                    if (attachmentLinks != "")
                    {
                        attachmentLinks += @"==============================</div>";

                        mail = app.Session.GetItemFromID(mailItemEntryId);
                        mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                        mail.HTMLBody = Regex.Replace(mail.HTMLBody,
                            @"</body>",
                            @"<br>" + attachmentLinks + @"</body>",
                            RegexOptions.Compiled | RegexOptions.IgnoreCase);
                        mail.Save();
                    }
                }
            }
        }

        internal static void DeleteAttachment(Outlook.Application app, String mailItemEntryId)
        {
            if (GetFlag("DeleteAttachmentsWhenBackup") == true)
            {
                Outlook.MailItem mailItem = app.Session.GetItemFromID(mailItemEntryId);
                int count = mailItem.Attachments.Count;
                int keepCount = 0;

                for (int i = 0; i < count; ++i)
                {
                    if (IsInlineImage(mailItem, mailItem.Attachments[1 + keepCount]) == false)
                    {
                        string fileName = "";

                        try
                        {
                            fileName = mailItem.Attachments[1 + keepCount].FileName;
                        }
                        catch
                        {
                            // can not get the fileename, then ignore this attachment.
                        }

                        if (fileName != "")
                        {
                            mailItem.Attachments[1 + keepCount].Delete();
                            mailItem.Save();
                        }
                        else
                        {
                            ++keepCount;
                        }
                    }
                    else
                    {
                        ++keepCount;
                    }
                }
            }
        }

        internal static void SaveAttachment(Outlook.Application app, String mailItemEntryId, string storagePathName)
        {
            if (GetFlag("SaveAttachmentsToLocalFolderWhenBackup") == true)
            {
                Outlook.MailItem mailItem = app.Session.GetItemFromID(mailItemEntryId);
                if (mailItem.Attachments.Count > 0)
                {
                    if (Directory.Exists(storagePathName) == false)
                    {
                        throw new DirectoryNotFoundException();
                    }

                    for (int i = 1; i <= mailItem.Attachments.Count; ++i)
                    {
                        if (IsInlineImage(mailItem, mailItem.Attachments[i]) == false)
                        {
                            string fileName = "";

                            try
                            {
                                fileName = mailItem.Attachments[i].FileName;
                            }
                            catch
                            {
                                // can not get the filename, then ignore this attachment.
                            }

                            if (fileName != "")
                            {
                                mailItem.Attachments[i].SaveAsFile(storagePathName + @"\" + mailItem.Attachments[i].FileName);
                            }
                        }
                    }
                }
            }
        }

        internal static Microsoft.Office.Interop.Word.Document GetWordEditor(Outlook.Inspector inspector)
        {
            // Check that the email editor is Word editor
            // Although "always" is a Word editor in Outlook 2013, it's best done perform this check
            if (inspector.EditorType == Outlook.OlEditorType.olEditorWord && inspector.IsWordMail())
            {
                return inspector.WordEditor;
            }
            else
            {
                return null;
            }
        }

        [Conditional("DEBUG")]
        internal static void ShowMailInfo(Outlook.MailItem mail)
        {
            Debug.Assert(mail != null);

            Util.Log(
                "MessageClass: " + mail.MessageClass + "\r\n" +
                "Class: " + mail.Class + "\r\n" +
                "Subject: " + mail.Subject + "\r\n" +
                "SenderName: " + mail.SenderName + "\r\n" +
                "SenderEmailAddress: " + mail.SenderEmailAddress + "\r\n" +
                "TO: " + mail.To + "\r\n" +
                "CC: " + mail.CC + "\r\n" +
                "BCC: " + mail.BCC + "\r\n" +
                "Category: " + mail.Categories + "\r\n" +
                "Company: " + mail.Companies + "\r\n" +
                "Creation Time: " + mail.CreationTime + "\r\n" +
                "DeferredDeliveryTime: " + mail.DeferredDeliveryTime + "\r\n" +
                "LastModificationTime: " + mail.LastModificationTime + "\r\n" +
                "ReceivedTime: " + mail.ReceivedTime + "\r\n" +
                "SentOn: " + mail.SentOn + "\r\n" +
                "DownloadState: " + mail.DownloadState + "\r\n" +
                "FlagRequest: " + mail.FlagRequest + "\r\n" +
                "ReceivedByName: " + mail.ReceivedByName + "\r\n" +
                "Saved: " + mail.Saved + "\r\n" +
                "Namespace - CurrentProfileName: " + mail.Session.CurrentProfileName + "\r\n" +
                "Namespace - Type: " + mail.Session.Type,
                newLineBeforeData: true);
        }

        [Conditional("DEBUG")]
        internal static void ShowFolderInfo(Outlook.Folder folder)
        {
            if (folder != null)
            {
                Util.Log(
                    "Name: " + folder.Name + "\r\n" +
                    "Folder Path: " + folder.FolderPath + "\r\n" +
                    "Default MessageClass: " + folder.DefaultMessageClass + "\r\n" +
                    "Current View: " + folder.CurrentView.Name + "\r\n" +
                    "Description: " + folder.Description + "\r\n" +
                    "IsSharePointFolder: " + folder.IsSharePointFolder + "\r\n" +
                    "Store (Display Name): " + folder.Store.DisplayName + "\r\n" +
                    "StoreID: " + folder.StoreID + "\r\n" +
                    "Namespace - CurrentProfileName: " + folder.Session.CurrentProfileName + "\r\n" +
                    "Namespace - CurrentUser: " + folder.Session.CurrentUser + "\r\n" +
                    "Namespace - ExchangeConnectionMode: " + folder.Session.ExchangeConnectionMode,
                    newLineBeforeData: true);

                Outlook.PropertyAccessor propertyAccessor = folder.PropertyAccessor;
                // get 'PR_FOLDER_TYPE' property value, 2 means 'search folder'
                // search folder does not have 'folders' property
                if (propertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x36010003") != 2)
                {
                    foreach (Outlook.Folder item in folder.Folders)
                    {
                        ShowFolderInfo(item);
                    }
                }
            }
        }

        [Conditional("DEBUG")]
        internal static void ShowAllStoreInfo(Outlook.Application application)
        {
            Outlook.Stores stores = application.Session.Stores;
            foreach (Outlook.Store store in stores)
            {
                if (store.IsDataFileStore == true)
                {
                    Util.Log(
                        "Store (DisplayName): " + store.DisplayName + "\r\n" +
                        "File Path: " + store.FilePath + "\r\n" +
                        "Root Folder: " + store.GetRootFolder().FolderPath + "\r\n" +
                        "ExchangeStoreType: " + store.ExchangeStoreType.ToString() + "\r\n" +
                        "IsCachedExchange: " + store.IsCachedExchange + "\r\n" +
                        "IsDataFileStore: " + store.IsDataFileStore + "\r\n" +
                        "IsInstantSearchEnabled: " + store.IsInstantSearchEnabled + "\r\n" +
                        "IsOpen: " + store.IsOpen + "\r\n" +
                        "StoreID: " + store.StoreID + "\r\n" +
                        "Namespace - CurrentProfileName: " + store.Session.CurrentProfileName + "\r\n" +
                        "Namespace - CurrentUser: " + store.Session.CurrentUser + "\r\n" +
                        "Namespace - ExchangeConnectionMode: " + store.Session.ExchangeConnectionMode,
                        newLineBeforeData: true);

                    ShowFolderInfo(store.GetRootFolder() as Outlook.Folder);
                }
            }
        }

        [Conditional("DEBUG")]
        internal static void ShowAllFolderInfo(Outlook.Application application)
        {
            Outlook.Folders folders = application.Session.Folders;
            foreach (Outlook.Folder folder in folders)
            {
                ShowFolderInfo(folder);
            }
        }

        [Conditional("DEBUG")]
        internal static void ShowAllAccountInfo(Outlook.Application application)
        {
            Outlook.Accounts accounts = null;

            try
            {
                // The Namespace Object (Session) has a collection of accounts.
                accounts = application.Session.Accounts;

                // Loop over all accounts and print detail account information.
                // All properties of the Account object are read-only.
                foreach (Outlook.Account account in accounts)
                {
                    Log(
                        // The DisplayName property represents the friendly name of the account.
                        "DisplayName: " + account.DisplayName + "\r\n" +
                        // The UserName property provides an account-based context to determine identity.
                        "UserName: " + account.UserName + "\r\n" +
                        // The SmtpAddress property provides the SMTP address for the account.
                        "SmtpAddress: " + account.SmtpAddress + "\r\n" +
                        // The AccountType property indicates the type of the account.
                        "AccountType: ",
                        newLineBeforeData: true, newLineAfterData: false);

                    switch (account.AccountType)
                    {
                        case Outlook.OlAccountType.olExchange:
                            Log("Exchange", outputTimestamp: false, newLineBeforeData: false, newLineAfterData: true);
                            break;

                        case Outlook.OlAccountType.olHttp:
                            Log("Http", outputTimestamp: false, newLineBeforeData: false, newLineAfterData: true);
                            break;

                        case Outlook.OlAccountType.olImap:
                            Log("Imap", outputTimestamp: false, newLineBeforeData: false, newLineAfterData: true);
                            break;

                        case Outlook.OlAccountType.olOtherAccount:
                            Log("Other", outputTimestamp: false, newLineBeforeData: false, newLineAfterData: true);
                            break;

                        case Outlook.OlAccountType.olPop3:
                            Log("Pop3", outputTimestamp: false, newLineBeforeData: false, newLineAfterData: true);
                            break;
                    }
                }
            }
            catch
            {
            }
        }
    }
}
