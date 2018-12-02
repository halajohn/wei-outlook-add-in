using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    class FilterEmailUtil {
        internal static void FilterOutUnwantedEmail(Outlook.MailItem mailItem) {
            string header = mailItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E") as string;
            if (header.Contains("X-Mailer: nodemailer")) {
                mailItem.Categories = "No need to popup new mail alarm";
                mailItem.Save();
            }
        }
    }
}
