using System.Diagnostics;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    class AutoBccUtil {
        internal static void AddBcc(Outlook.MailItem mailItem, ref bool Cancel) {
            Debug.Assert(mailItem != null);

            if (Config.EnableAutoBcc == true) {
                if (mailItem != null) {
                    // TODO: if strBcc == "", the following line will raise an exception, why?
                    Outlook.Recipient objRecip = mailItem.Recipients.Add(Config.AutoBccEmailAddress);
                    objRecip.Type = (int)Outlook.OlMailRecipientType.olBCC;

                    if (objRecip.Resolve() == false) {
                        DialogResult result =
                            MessageBox.Show("Could not resolve the Bcc recipient. Do you still want to send the message ?",
                                            "Could Not Resolve Bcc Recipient",
                                            MessageBoxButtons.YesNo,
                                            MessageBoxIcon.Question,
                                            MessageBoxDefaultButton.Button2);
                        if (result == DialogResult.No) {
                            Cancel = true;
                        }
                    }
                }
            }
        }
    }
}
