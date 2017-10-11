using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in
{
    public partial class ThisAddIn
    {
        public static ribbon1 ribbonExtensibilityObject;

        // Holds a reference to the Application.Inspectors collection.
        // Required to get notifications for NewInspector events.
        //
        // Have to remember this to prevent .Net's GC to clean the NewInspector event handler
        private Outlook.Inspectors inspectors;

        // A dictionary that holds a reference to the inspectors handled by the add-in.
        private Dictionary<Guid, InspectorWrapper> _wrappedInspectors;

        private Outlook.Application application;
        private Outlook.Explorer explorer;
        private Redemption.SafeExplorer sExplorer;

        private Outlook.Items items;

        // Startup method is called when Outlook loads the add-in
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _wrappedInspectors = new Dictionary<Guid, InspectorWrapper>();
            inspectors = Globals.ThisAddIn.Application.Inspectors;
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(WrapInspector);

            // Also handle existing Inspectors
            // (for example, double-clicking a .msg file).
            foreach (Outlook.Inspector inspector in inspectors)
            {
                WrapInspector(inspector);
            }

            // keep the following references to keep event handlers alive through GC
            application = Globals.ThisAddIn.Application;
            explorer = application.ActiveExplorer();
            sExplorer = new Redemption.SafeExplorer();
            items = Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items;

            {
                // Enable 'property page' only on Outlook 2010 and later
                int majorVersion = int.Parse(this.Application.Version.Split('.')[0]);
                bool hasBackstage = (majorVersion >= 14);
                if (!hasBackstage)
                {
                    application.OptionsPagesAdd += new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_AddOptionsPages);
                }
            }

            // folder switch event
            explorer.FolderSwitch += new Outlook.ExplorerEvents_10_FolderSwitchEventHandler(Explorer_SwitchFolder);

            // selection change event
            explorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);

            // an event for sending an email
            application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);

            // an event for receiving an email
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Inbox_ItemAdd);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785

            // Clean up.
            _wrappedInspectors.Clear();
            inspectors.NewInspector -= new Outlook.InspectorsEvents_NewInspectorEventHandler(WrapInspector);
            inspectors = null;

            explorer.FolderSwitch -= new Outlook.ExplorerEvents_10_FolderSwitchEventHandler(Explorer_SwitchFolder);
            explorer.SelectionChange -= new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);
            explorer = null;

            sExplorer = null;

            application.OptionsPagesAdd -= new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_AddOptionsPages);
            application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
            application = null;

            items.ItemAdd -= new Outlook.ItemsEvents_ItemAddEventHandler(Inbox_ItemAdd);
            items = null;

            ribbonExtensibilityObject = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void Explorer_SelectionChange()
        {
            if (application.ActiveExplorer().Selection.Count > 0)
            {
                dynamic msg = application.ActiveExplorer().Selection[1];
                application.ActiveExplorer().RemoveFromSelection(msg);
                application.ActiveExplorer().AddToSelection(msg);
                sExplorer.Item = application.ActiveExplorer();
                dynamic wdDoc = sExplorer.ReadingPane.WordEditor;
                if (wdDoc != null)
                {
                    wdDoc.Windows(1).View.Zoom.Percentage = 150;
                }
            }

            ribbonExtensibilityObject.InvalidateGroupReply();
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            if (Util.GetFlag("EnableAutoBcc") == true)
            {
                string strBcc = Util.GetValue("AutoBccEmailAddr");

                if (strBcc == "")
                {
                    MessageBox.Show("Automatically auto BCC has been selected, but no email address has been provided",
                                    "Auto BCC fialed",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error,
                                    MessageBoxDefaultButton.Button1);
                    return;
                }

                Outlook.MailItem mailItem = Item as Outlook.MailItem;
                if (mailItem != null)
                {
                    // TODO: if strBcc == "", the following line will raise an exceptino, why?
                    Outlook.Recipient objRecip = mailItem.Recipients.Add(strBcc);
                    objRecip.Type = (int)Outlook.OlMailRecipientType.olBCC;

                    if (objRecip.Resolve() == false)
                    {
                        DialogResult result =
                            MessageBox.Show("Could not resolve the Bcc recipient. Do you want still to send the message ?",
                                            "Could Not Resolve Bcc Recipient",
                                            MessageBoxButtons.YesNo,
                                            MessageBoxIcon.Question,
                                            MessageBoxDefaultButton.Button2);

                        if (result == DialogResult.No)
                        {
                            Cancel = true;
                        }
                    }
                }
            }
        }

        private void Application_AddOptionsPages(Outlook.PropertyPages pages)
        {
            pages.Add(new PropertyPage_WeiOutlookAddIn(), "seems to be ignored...");
        }

        private void Inbox_ItemAdd(object item)
        {
            Outlook.MailItem mail = item as Outlook.MailItem;
            if (mail != null)
            {
                if (mail.MessageClass == "IPM.Note" &&
                    Util.GetFlag("EnableAutoBackupEmailFromMe") == true &&
                    Util.GetSenderSMTPAddress(mail) == Util.GetValue("AutoBackupMyEmailAddr"))
                {
                    if (mail.Recipients.Count == 1)
                    {
                        if (Util.GetSMTPAddress(mail.Recipients[1].AddressEntry) ==
                            Util.GetValue("AutoBackupMyEmailAddr"))
                        {
                            // send only to myself, do not backup this email.
                        }
                        else
                        {
                            Util.BackupEmail(application, mail.EntryID, false);
                        }
                    }
                    else
                    {
                        Util.BackupEmail(application, mail.EntryID, false);
                    }
                }
            }
        }

        private void Explorer_SwitchFolder()
        {
            ribbonExtensibilityObject.InvalidateGroupReply();
        }

        // Wrap an Inspector, if required, and store it in memory to get events of the wrapped Inspector.
        // <param name="inspector">The Outlook Inspector instance.</param>
        void WrapInspector(Outlook.Inspector inspector)
        {
            InspectorWrapper wrapper = InspectorWrapper.GetWrapperFor(inspector);
            if (wrapper != null)
            {
                // Register the Closed event.
                wrapper.Closed += new InspectorWrapperClosedEventHandler(wrapper_Closed);
                // Remember the inspector in memory.
                _wrappedInspectors[wrapper.Id] = wrapper;
            }
        }

        // Method is called when an inspector has been closed.
        // Removes reference from memory.
        // <param name="id">The unique id of the closed inspector</param>
        void wrapper_Closed(Guid id)
        {
            _wrappedInspectors.Remove(id);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbonExtensibilityObject = new ribbon1();
            return ribbonExtensibilityObject;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
