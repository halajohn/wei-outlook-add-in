using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook; // TODO: why I can not using Microsoft.Office.Interop here?

namespace wei_outlook_add_in
{
    // 'ComVisible(true)' is very important, otherwise, Outlook COM can not find this 'PropertyPage'.
    [ComVisible(true)]
    public class PropertyPage_WeiOutlookAddIn : UserControl, Outlook.PropertyPage
    {
        public void Apply()
        {
            // Save stuff to registry here
        }

        public bool Dirty
        {
            get; private set;
        }

        public void GetPageInfo(ref string HelpFile, ref int HelpContext)
        {
            // Can safely ignore this method, and do nothing
        }

        [DispId(-518)]
        public string PageCaption
        {
            get
            {
                return "Wei Outlook Add-in Options";
            }
        }
    }
}
