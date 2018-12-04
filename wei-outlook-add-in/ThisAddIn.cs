using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    public partial class ThisAddIn {
        private Outlook.Inspectors Inspectors;
        private Dictionary<Guid, InspectorWrapper> WrappedInspectors;
        private Outlook.Explorer Explorer;
        private Redemption.SafeExplorer sExplorer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e) {
            Config.ReadFromFile();

            Inspectors = Globals.ThisAddIn.Application.Inspectors;
            WrappedInspectors = new Dictionary<Guid, InspectorWrapper>();

            Inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(WrapInspector);
            foreach (Outlook.Inspector inspector in Inspectors) {
                WrapInspector(inspector);
            }

            Explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            Explorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(ExplorerSelectionChange);

            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(ApplicationItemSend);
            Application.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(ApplicationNewMailEx);

            sExplorer = new Redemption.SafeExplorer();

            CategoryUtil.UpdateCategories(Application);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void WrapInspector(Outlook.Inspector inspector) {
            InspectorWrapper wrapper = InspectorWrapper.GetWrapperFor(inspector);
            if (wrapper != null) {
                wrapper.Closed += new InspectorWrapperClosedEventHandler(WrapperClosed);
                WrappedInspectors[wrapper.Id] = wrapper;
            }
        }

        private void WrapperClosed(Guid id) {
            WrappedInspectors.Remove(id);
        }

        // The implementation here is the best I can do.
        // drawback:
        // 1) click an email in a folder at the first time would not zoom it.
        // 2) it's best to disable "show as conversations" of all folders
        private void ExplorerSelectionChange_() {
            if (Explorer.Selection.Count > 0) {
                Outlook.MailItem mailItem = Explorer.Selection[1] as Outlook.MailItem;

                if (mailItem != null) {
                    Microsoft.Office.Interop.Word.Document wdDoc = null;
                    if (Util.OutlookVersion() == "2016") {
                        var previewPane = Explorer.GetType().InvokeMember("PreviewPane", BindingFlags.GetProperty, null, Explorer, null);
                        try {
                            wdDoc = (Microsoft.Office.Interop.Word.Document)previewPane.GetType().InvokeMember("WordEditor", BindingFlags.GetProperty, null, previewPane, null);
                        } catch (TargetInvocationException ex) {
                            Debug.Print(ex.ToString());
                        }
                    } else if (Util.OutlookVersion() == "2013") {
                        sExplorer.Item = Explorer;
                        wdDoc = sExplorer.ReadingPane.WordEditor;
                    } else {
                        Debug.Assert(false);
                        wdDoc = null;
                    }

                    if (wdDoc != null) {
                        Microsoft.Office.Interop.Word.Zoom zoom = wdDoc.Windows[1].View.Zoom;
                        zoom.Percentage = Config.Zoom;
                    }
                }
            }
        }

        private void ExplorerSelectionChange() {
            ExplorerSelectionChange_();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // Although NewMailEx event is triggered before outlook rule processing, these 2 run asynchronized,
        // so the rule processing doesn't guarentee to run "after" the completion of NewMailEx handler.
        // And NewMailEx event will not be triggered for every new mail if a lot of new mails coming in a short period of time.
        //
        // so the most reliable way to process every new mail is to add an outlook email rule to "run a script" for
        // every new mail.
        //
        // put the following function into outlook vba editor, under "TheOutlookSession", and create an email rule
        // to "run a script" this one.
        //
        // Enable "run a script" in Outlook 2013:
        // Create a DWORD "EnableUnsafeClientMailRules" under
        // HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Security, and set to 1
        //
        // Public Sub XXX(Item As Outlook.MailItem)
        //     Header = Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
        //     Pos = InStr(Header, "X-Mailer: nodemailer")
        //     If Pos Then
        //         Item.Categories = "No need to popup new mail alarm"
        //         Item.Save
        //     End If
        // End Sub
        private void ApplicationNewMailEx(string EntryIDCollection) {
            Outlook.NameSpace nameSpace = Application.GetNamespace("MAPI");
            string[] entryIds = EntryIDCollection.Split(',');
            for (int i = 0; i < entryIds.Length; ++i) {
                Outlook.MailItem mailItem = nameSpace.GetItemFromID(entryIds[i]) as Outlook.MailItem;
                if (mailItem != null) {
                    FilterEmailUtil.FilterOutUnwantedEmail(mailItem);
                    if (Config.AutoBackupEmailFromMe == true && Util.GetSenderSMTPAddress(mailItem) == Config.MyEmailAddress) {
                        BackupEmailUtil.MarkEmailReadAndClearAllCategories(mailItem);
                        EmailFlagUtil.FlagEmail(mailItem);
                        BackupEmailUtil.BackupEmail(mailItem);
                    }
                }
            }
        }

        private void ApplicationItemSend(object Item, ref bool Cancel) {
            if (Item is Outlook.MailItem) {
                AutoBccUtil.AddBcc(Item as Outlook.MailItem, ref Cancel);
                if (Cancel == false) {
                    EnlargeDearHiNameUtil.PerformEnlarge(Item as Outlook.MailItem);
                    ConvertChineseUtil.ConvertEmailChineseAccordingToRecipientDept(Item as Outlook.MailItem);
                }
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            return new Ribbon1();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
