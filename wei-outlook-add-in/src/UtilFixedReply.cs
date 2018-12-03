using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    class FixedReplyUtil {
        internal class Data {
            public string label;
            public string text;
        }

        internal static void AddText(Outlook.MailItem mailItem, string text) {
            Debug.Assert(mailItem != null);

            StringBuilder stringBuilder = new StringBuilder(@"$1");
            stringBuilder.Append(text);

            mailItem.HTMLBody = Regex.Replace(mailItem.HTMLBody,
                @"(<body[^>]*>)",
                stringBuilder.ToString(),
                RegexOptions.IgnoreCase);
        }
    }
}
