using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    class EmailFlagUtil {
        private static void ClearMailItemFlagIfItsInSentItemFolder(Outlook.MailItem mailItem) {
            Outlook.Folder parentFolder = mailItem.Parent as Outlook.Folder;
            Debug.Assert(parentFolder != null);
            if (parentFolder.FolderPath.EndsWith("\\Sent Items")) {
                mailItem.ClearTaskFlag();
                mailItem.Save();
            }
        }

        private static bool IsMailItemOrAnyItsChildrenFlagged(Outlook.MailItem mailItem, Outlook.Conversation conv) {
            bool answer = false;

            if ((mailItem.FlagRequest != null) && (mailItem.FlagRequest != "")) {
                answer = true;
            }

            Outlook.SimpleItems items = conv.GetChildren(mailItem);
            if (items.Count > 0) {
                foreach (object myItem in items) {
                    if (myItem is Outlook.MailItem) {
                        bool result = IsMailItemOrAnyItsChildrenFlagged(myItem as Outlook.MailItem, conv);
                        if (result == true) {
                            answer = true;
                        }
                    }
                }
            }

            ClearMailItemFlagIfItsInSentItemFolder(mailItem);
            return answer;
        }

        private static bool IsMailItemHasFlaggedPreviousMail(Outlook.MailItem mailItem) {
            bool answer = false;

            if (mailItem is Outlook.MailItem) {
                Outlook.Folder folder = mailItem.Parent as Outlook.Folder;
                Debug.Assert(folder != null);
                Outlook.Store store = folder.Store;

                if (store.IsConversationEnabled == true) {
                    Outlook.Conversation conv = mailItem.GetConversation();
                    if (conv != null) {
                        Outlook.SimpleItems simpleItems = conv.GetRootItems();
                        foreach (object item in simpleItems) {
                            if (item is Outlook.MailItem) {
                                bool result = IsMailItemOrAnyItsChildrenFlagged(item as Outlook.MailItem, conv);
                                if (result == true) {
                                    answer = true;
                                }
                            }
                        }
                    }
                }
            }

            return answer;
        }

        internal static void FlagEmail(Outlook.MailItem mailItem) {
            if (IsMailItemHasFlaggedPreviousMail(mailItem) == true) {
                mailItem.MarkAsTask(Outlook.OlMarkInterval.olMarkNoDate);
                mailItem.Save();
            }
        }
    }
}
