using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    class BackupEmailUtil {
        private static Outlook.Store GetStore(Outlook.Application app, string pstPath) {
            Outlook.Stores stores = app.Session.Stores;
            foreach (Outlook.Store store in stores) {
                if (store.FilePath == pstPath) {
                    return store;
                }
            }
            return null;
        }

        private static Outlook.Store AddStore(Outlook.Application app, string pstPath, string displayName) {
            app.Session.AddStoreEx(pstPath, Outlook.OlStoreType.olStoreDefault);

            Outlook.Folder lastFolder = app.GetNamespace("MAPI").Folders.GetLast() as Outlook.Folder;
            lastFolder.Name = displayName;

            Outlook.Store store = GetStore(app, pstPath);
            return store;
            // [TODO] set 'Display reminders and tasks from this folder in the To-Do Bar' checkbox
        }

        private static Outlook.Folder AddFolderInternal(Outlook.Folders folders, string folderName) {
            Outlook.Folder newFolder = folders.Add(folderName) as Outlook.Folder;
            return newFolder;
        }

        private static Outlook.Folder AddFolder(Outlook.NameSpace ns, string folderName) {
            return AddFolderInternal(ns.Folders, folderName);
        }

        private static Outlook.Folder AddFolder(Outlook.Folder folder, string folderName) {
            return AddFolderInternal(folder.Folders, folderName);
        }

        private static Outlook.Folder GetFolder(Outlook.Application application, string folderPath) {
            folderPath = folderPath.TrimStart(@"\".ToCharArray()); // Remove leading "\" characters.
            string[] folders = folderPath.Split(@"\".ToCharArray()); // Split the folder path into individual folder names.

            Outlook.Folder returnFolder = null;
            try {
                returnFolder = application.Session.Folders[folders[0]] as Outlook.Folder; // Retrieve a reference to the root folder.
            } catch (COMException) {
                returnFolder = AddFolder(application.Session, folders[0]);
            }

            // If the root folder exists, look in subfolders
            if (returnFolder != null) {
                // Look through folder names, skipping the first folder, which you already retrieved
                for (int i = 1; i < folders.Length; ++i) {
                    string folderName = folders[i];
                    Outlook.Folder childFolder = null;
                    try {
                        childFolder = returnFolder.Folders[folderName] as Outlook.Folder;
                    } catch {
                        Debug.Assert(childFolder == null);
                        childFolder = AddFolder(returnFolder, folderName);
                    }
                    returnFolder = childFolder;
                }
            }

            return returnFolder;
        }

        private static Outlook.Folder GetBackupFolder(Outlook.MailItem mailItem) {
            string mailReceivedYear = mailItem.ReceivedTime.Year.ToString();
            string mailReceivedMonth = mailItem.ReceivedTime.Month.ToString();

            string backupName = @"" + mailReceivedYear + @"_" + mailReceivedMonth;

            string backupStoreFileName = Config.EmailBackupPath + @"\" + backupName + @".pst";
            Outlook.Store store = GetStore(mailItem.Application, backupStoreFileName);
            if (store == null) {
                AddStore(mailItem.Application, backupStoreFileName, backupName);
            }

            string backupFolderName = @"\\" + backupName;
            Outlook.Folder backupFolder = GetFolder(mailItem.Application, backupFolderName);

            return backupFolder;
        }

        private static bool AttachmentIsInlineImage(Outlook.MailItem mailItem, Outlook.Attachment attachment) {
            Debug.Assert(attachment != null);

            const string PR_ATTACH_METHOD = "http://schemas.microsoft.com/mapi/proptag/0x37050003";
            const string PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003";

            if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatPlain) {
                // if this is a plain text email, every attachment is a non-inline attachment
                return false;
            } else if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatRichText) {
                // if the body format is RTF, the non-embedded attachment would be of the PR_ATTACH_METHOD property is not 6 (ATTACH_OLD)
                return (int)attachment.PropertyAccessor.GetProperty(PR_ATTACH_METHOD) == 6;
            } else if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML) {
                // if the body format is HTML, the non-embedded attachment would be of the PR_ATTACH_FLAGS property is not 4 (ATT_MHTML_REF)
                return (int)attachment.PropertyAccessor.GetProperty(PR_ATTACH_FLAGS) == 4;
            } else {
                Debug.Assert(false);
                return false;
            }
        }

        private static void AddAttachmentLinkToBodyEnd(Outlook.MailItem mailItem) {
            if (mailItem.Attachments.Count > 0) {
                string attachmentLinks = "";

                for (int i = 1; i <= mailItem.Attachments.Count; ++i) {
                    if (AttachmentIsInlineImage(mailItem, mailItem.Attachments[i]) == false) {
                        if (attachmentLinks == "") {
                            attachmentLinks += @"<div>======== Saved attachment links ========<br>";
                        }

                        string fileName = "";
                        try {
                            fileName = mailItem.Attachments[i].FileName;
                        } catch {
                            // can not get the file name, then ignore this attachment.
                        }

                        if (fileName != "") {
                            string location =
                                Config.AttachmentBackupPath + (Config.AttachmentBackupPath.EndsWith(@"\") ? "" : @"\") +
                                mailItem.Attachments[i].FileName;
                            string uri = Regex.Replace(location, @"\\", @"/", RegexOptions.IgnoreCase);
                            attachmentLinks += "<a href=\"file:///" + uri + "\">" + uri + @"</a><br>";
                        }
                    }
                }

                if (attachmentLinks != "") {
                    attachmentLinks += @"========================================</div>";

                    mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                    mailItem.HTMLBody = Regex.Replace(mailItem.HTMLBody,
                        @"</body>",
                        @"<br>" + attachmentLinks + @"</body>",
                        RegexOptions.IgnoreCase);
                    mailItem.Save();
                }
            }
        }

        private static void SaveAttachment(Outlook.MailItem mailItem) {
            if (mailItem.Attachments.Count > 0) {
                for (int i = 1; i <= mailItem.Attachments.Count; ++i) {
                    if (AttachmentIsInlineImage(mailItem, mailItem.Attachments[i]) == false) {
                        string fileName = "";
                        try {
                            fileName = mailItem.Attachments[i].FileName;
                        } catch {
                            // can not get the filename, then ignore this attachment.
                        }

                        if (fileName != "") {
                            string location =
                                Config.AttachmentBackupPath + (Config.AttachmentBackupPath.EndsWith(@"\") ? "" : @"\") +
                                mailItem.Attachments[i].FileName;
                            mailItem.Attachments[i].SaveAsFile(location);
                        }
                    }
                }
            }
        }

        private static void DeleteAttachment(Outlook.MailItem mailItem) {
            int count = mailItem.Attachments.Count;
            int keepCount = 0;
            for (int i = 0; i < count; ++i) {
                if (AttachmentIsInlineImage(mailItem, mailItem.Attachments[1 + keepCount]) == false) {
                    string fileName = "";
                    try {
                        fileName = mailItem.Attachments[1 + keepCount].FileName;
                    } catch {
                        // can not get the fileename, then ignore this attachment.
                    }

                    if (fileName != "") {
                        mailItem.Attachments[1 + keepCount].Delete();
                        mailItem.Save();
                    } else {
                        ++keepCount;
                    }
                } else {
                    ++keepCount;
                }
            }
        }

        internal static void BackupEmail(Outlook.MailItem mailItem) {
            Outlook.Folder backupFolder = GetBackupFolder(mailItem);
            Outlook.MAPIFolder parentFolder = mailItem.Parent as Outlook.MAPIFolder;

            if (parentFolder.FolderPath == backupFolder.FolderPath) {
                mailItem.Close(Outlook.OlInspectorClose.olSave);
            } else {
                mailItem.UnRead = false;
                mailItem.Save();

                AddAttachmentLinkToBodyEnd(mailItem);
                SaveAttachment(mailItem);
                DeleteAttachment(mailItem);

                mailItem.Move(backupFolder);
            }
        }
    }
}
