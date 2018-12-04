using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using wei_outlook_add_in.Properties;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility {
        private Office.IRibbonUI ribbon;

        public Ribbon1() {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("wei_outlook_add_in.src.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
        }

        private bool ControlIsInInspector(Office.IRibbonControl control, ref Outlook.Inspector inspector) {
            inspector = control.Context as Outlook.Inspector;
            return inspector != null;
        }

        private bool ControlIsInExplorer(Office.IRibbonControl control, ref Outlook.Explorer explorer) {
            explorer = control.Context as Outlook.Explorer;
            return explorer != null;
        }

        public Bitmap GetBackupEmailImage(Office.IRibbonControl control) {
            return Resources.Backup_email;
        }

        public void BtnBackupEmail_Click(Office.IRibbonControl control) {
            Outlook.MailItem mailItem = null;
            Outlook.Explorer explorer = null;
            Outlook.Inspector inspector = null;

            if (ControlIsInInspector(control, ref inspector) == true) {
                mailItem = inspector.CurrentItem as Outlook.MailItem;
                if (mailItem != null) {
                    if (BackupEmailUtil.IsEmailAlreadyInBackupFolder(mailItem)) {
                        mailItem.Close(Outlook.OlInspectorClose.olSave);
                    } else {
                        BackupEmailUtil.MarkEmailUnreadAndClearAllCategories(mailItem);
                        EmailFlagUtil.FlagEmail(mailItem);
                        BackupEmailUtil.BackupEmail(mailItem);
                    }
                }
            } else if (ControlIsInExplorer(control, ref explorer) == true) {
                try {
                    // I have to wrap 'explorer.selction' into a try block,
                    // becasue outlook will raise an exception on this line when the first page is 'Outlook Today'
                    Outlook.Selection selection = explorer.Selection;
                    foreach (var selected in selection) {
                        mailItem = selected as Outlook.MailItem;
                        if (mailItem != null) {
                            if (BackupEmailUtil.IsEmailAlreadyInBackupFolder(mailItem) == false) {
                                EmailFlagUtil.FlagEmail(mailItem);
                            }
                        }
                    }
                    foreach (var selected in selection) {
                        mailItem = selected as Outlook.MailItem;
                        if (mailItem != null) {
                            if (BackupEmailUtil.IsEmailAlreadyInBackupFolder(mailItem) == false) {
                                BackupEmailUtil.MarkEmailUnreadAndClearAllCategories(mailItem);
                                BackupEmailUtil.BackupEmail(mailItem);
                            }
                        }
                    }
                } catch (COMException ex) {
                    Debug.WriteLine(ex.ToString());
                }
            }
        }

        public string GetDynamicMenuCategoryContent(Office.IRibbonControl control) {
            StringBuilder stringBuilder = new StringBuilder(@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">");

            foreach (CategoryUtil.Data category in Config.Categories) {
                string id = Util.FromLabelToId(category.label);

                stringBuilder
                    .Append(@"<button")
                    .Append(@" id=""").Append(id).Append(@"""")
                    .Append(@" label=""").Append(category.label).Append(@"""")
                    .Append(@" tag=""").Append(category.label).Append(@"""")
                    .Append(@" onAction=""OnCategoriesAction""")
                    .Append(@" getImage=""getCategoriesImage""/>");
            }
            stringBuilder.Append(@"</menu>");

            return stringBuilder.ToString();
        }

        public Bitmap GetDynamicMenuCategoryImage(Office.IRibbonControl control) {
            return Resources.Categories;
        }

        public void OnCategoriesAction(Office.IRibbonControl control) {
            Outlook.Inspector inspector = null;
            if (ControlIsInInspector(control, ref inspector) == true) {
                Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
                if (mailItem != null) {
                    mailItem.Categories = "";
                    mailItem.Categories = control.Tag;
                    mailItem.Save();
                }
                return;
            }

            Outlook.Explorer explorer = null;
            if (ControlIsInExplorer(control, ref explorer) == true) {
                Outlook.Selection selection = explorer.Selection;
                foreach (var selected in selection) {
                    Outlook.MailItem mailItem = selected as Outlook.MailItem;
                    if (mailItem != null) {
                        mailItem.Categories = "";
                        mailItem.Categories = control.Tag;
                        mailItem.Save();
                    }
                }
                return;
            }
        }

        public Bitmap GetCategoriesImage(Office.IRibbonControl control) {
            string userProfileFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string imageFilename = userProfileFolder + @"\wei-outlook-add-in\" + control.Tag + @".png";
            return new Bitmap(imageFilename);
        }

        public string GetDynamicMenuFixedReplyContent(Office.IRibbonControl control) {
            StringBuilder stringBuilder = new StringBuilder(@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">");

            foreach (FixedReplyUtil.Data fixedReply in Config.FixedReplies) {
                string id = Util.FromLabelToId(fixedReply.label);

                string text = Regex.Replace(fixedReply.text, @"<", @"&lt;", RegexOptions.IgnoreCase);
                text = Regex.Replace(text, @">", @"&gt;", RegexOptions.IgnoreCase);
                text = Regex.Replace(text, @"""", @"&quot;", RegexOptions.IgnoreCase);

                stringBuilder
                    .Append(@"<button")
                    .Append(@" id=""").Append(id).Append(@"""")
                    .Append(@" label=""").Append(fixedReply.label).Append(@"""")
                    .Append(@" tag=""").Append(text).Append(@"""")
                    .Append(@" onAction=""OnFixedRepliesAction""")
                    .Append(@" getImage=""getFixedRepliesImage""/>");
            }
            stringBuilder.Append(@"</menu>");

            return stringBuilder.ToString();
        }

        public Bitmap GetDynamicMenuFixedReplyImage(Office.IRibbonControl control) {
            return Resources.Fixed_reply;
        }

        public void OnFixedRepliesAction(Office.IRibbonControl control) {
            Outlook.Inspector inspector = null;
            if (ControlIsInInspector(control, ref inspector) == true) {
                Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
                if ((mailItem != null) && (mailItem.Sent == false) /* compose mode */) {
                    FixedReplyUtil.AddText(mailItem, control.Tag);
                    mailItem.Save();
                }
                return;
            }

            Outlook.Explorer explorer = null;
            if (ControlIsInExplorer(control, ref explorer) == true) {
                Outlook.Selection selection = explorer.Selection;
                foreach (var selected in selection) {
                    Outlook.MailItem mailItem = selected as Outlook.MailItem;
                    if (mailItem != null) {
                        if (mailItem.Application.ActiveExplorer().ActiveInlineResponse != null) {
                            Outlook.MailItem inlineMailItem = mailItem.Application.ActiveExplorer().ActiveInlineResponse as Outlook.MailItem;
                            if (inlineMailItem != null) {
                                FixedReplyUtil.AddText(inlineMailItem, control.Tag);
                                inlineMailItem.Save();
                            }
                        }
                    }
                }
                return;
            }
        }

        public Bitmap GetFixedRepliesImage(Office.IRibbonControl control) {
            string userProfileFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string imageFilename = userProfileFolder + @"\wei-outlook-add-in\" + Util.FromIdToLabel(control.Id) + @".png";
            return new Bitmap(imageFilename);
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
