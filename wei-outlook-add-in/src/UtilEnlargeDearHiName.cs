using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    class EnlargeDearHiNameUtil {
        internal static void PerformEnlarge(Outlook.MailItem mailItem) {
            mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            mailItem.HTMLBody = Regex.Replace(
                mailItem.HTMLBody,
                @"(?<=>(?:&nbsp;|\s)*(?:Dear|dear|Hi|hi|HI|hI))(?<Space>(?:&nbsp;|\s|,)+)(?<Name>[^,\s]+?.*?)(?<Last><o:p>|<br>)",
                delegate (Match match) {
                    string v = match.Groups["Space"].Value + "<b><u><span style='font-size:22.0pt'>" + match.Groups["Name"].Value + "</span></u></b>" + match.Groups["Last"].Value;
                    return v;
                });
            mailItem.Save();
        }
    }
}
