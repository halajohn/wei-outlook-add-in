using System;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    class Util {
        internal static Microsoft.Office.Interop.Word.Document GetWordEditor(Outlook.Inspector inspector) {
            // https://docs.microsoft.com/zh-tw/office/vba/api/outlook.inspector.wordeditor
            // The WordEditor property is only valid if the IsWordMail method returns True and the EditorType property is olEditorWord.
            if (inspector.IsWordMail() && inspector.EditorType == Outlook.OlEditorType.olEditorWord) {
                return inspector.WordEditor as Microsoft.Office.Interop.Word.Document;
            } else {
                return null;
            }
        }

        internal static bool IsMailItem(object obj) {
            return (string)obj.GetType().InvokeMember("MessageClass", BindingFlags.GetProperty, null, obj, null) == "IPM.Note";
        }

        internal static string OutlookVersion() {
            string version = Globals.ThisAddIn.Application.Version;
            if (version.StartsWith("16")) {
                return "2016";
            } else {
                return "unknown";
            }
        }

        private static string GetSMTPAddress(Outlook.AddressEntry entry) {
            Debug.Assert(entry != null);

            if (entry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeAgentAddressEntry ||
                entry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) {
                Outlook.ExchangeUser exchUser = entry.GetExchangeUser();
                if (exchUser != null) {
                    return exchUser.PrimarySmtpAddress;
                } else {
                    return null;
                }
            } else {
                string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                return entry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
            }
        }

        internal static string GetSenderSMTPAddress(Outlook.MailItem mailItem) {
            Debug.Assert(mailItem != null);
            if (mailItem.SenderEmailType == "EX") {
                Outlook.AddressEntry sender = mailItem.Sender;
                if (sender != null) {
                    return GetSMTPAddress(sender);
                } else {
                    return null;
                }
            } else {
                return mailItem.SenderEmailAddress;
            }
        }

        // ex: High priority ==> btnHighPriority
        internal static string FromLabelToId(string label) {
            string titleCaseLabel = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(label.ToLower());
            string[] tokens = titleCaseLabel.Split(new string[] { " " }, StringSplitOptions.None);
            return @"btn" + string.Concat(tokens);
        }

        // ex: TodayILiveInTheUSAWithSimon ==> Today I Live In The USA With Simon
        private static string SplitCamelCase(string source) {
            Regex r = new Regex(@"
                (?<=[A-Z])(?=[A-Z][a-z]) |
                (?<=[^A-Z])(?=[A-Z]) |
                (?<=[A-Za-z])(?=[^A-Za-z])", RegexOptions.IgnorePatternWhitespace);
            return r.Replace(source, " ");
        }

        // ex: btnHighPriority ==> High priority
        internal static string FromIdToLabel(string id) {
            id = id.Substring(3);
            id = SplitCamelCase(id);
            string[] tokens = id.Split(new string[] { " " }, StringSplitOptions.None);
            for (int i = 1; i < tokens.Length; ++i) {
                tokens[i] = tokens[i].ToLower();
            }
            return string.Concat(tokens);
        }
    }
}
