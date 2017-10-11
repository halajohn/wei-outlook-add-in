// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Drawing;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in
{
    [ComVisible(true)]
    public class ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public ribbon1()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("wei_outlook_add_in.src.ribbon1.xml");
        }

        // Ribbon Callbacks
        // Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void btnBackupEmail_Click(Office.IRibbonControl control)
        {
            Outlook.MailItem mail = null;
            Outlook.Explorer explorer = null;
            Outlook.Inspector inspector = null;

            if (ControlIsInInspector(control, ref inspector) == true)
            {
                mail = inspector.CurrentItem as Outlook.MailItem;
                if (mail != null)
                {
                    Util.BackupEmail(mail.Application, mail.EntryID, true);
                }
                return;
            }

            if (ControlIsInExplorer(control, ref explorer) == true)
            {
                Outlook.Selection selection = null;
                try
                {
                    // TODO: remove this 'Today' check, it should be disabled in Today page.

                    // I have to wrap 'explorer.selction' into a try block,
                    // becasue outlook will raise an exception on this line when the first page is 'Outlook Today'
                    selection = explorer.Selection;
                    foreach (dynamic selected in selection)
                    {
                        mail = selected as Outlook.MailItem;
                        if (mail != null)
                        {
                            Util.BackupEmail(mail.Application, mail.EntryID, true);
                        }
                    }
                }
                catch (COMException ex)
                {
                    Debug.WriteLine(ex.ToString());
                }
                return;
            }
        }

        public enum PriorityType
        {
            HIGH,
            NORMAL,
            PROJECT_1,
            PROJECT_2,
        };

        public void btnHighPriorityEmail_Click(Office.IRibbonControl control)
        {
            btnPriorityEmail_Click(control, PriorityType.HIGH);
        }

        public void btnNormalPriorityEmail_Click(Office.IRibbonControl control)
        {
            btnPriorityEmail_Click(control, PriorityType.NORMAL);
        }

        public void btnProject1Email_Click(Office.IRibbonControl control)
        {
            btnPriorityEmail_Click(control, PriorityType.PROJECT_1);
        }

        public void btnProject2Email_Click(Office.IRibbonControl control)
        {
            btnPriorityEmail_Click(control, PriorityType.PROJECT_2);
        }

        public void btnPriorityEmail_Click(Office.IRibbonControl control, PriorityType priorityType)
        {
            Outlook.MailItem mail = null;
            Outlook.Explorer explorer = null;
            Outlook.Inspector inspector = null;

            if (ControlIsInInspector(control, ref inspector) == true)
            {
                mail = inspector.CurrentItem as Outlook.MailItem;
                if (mail != null)
                {
                    switch (priorityType)
                    {
                        case PriorityType.HIGH:
                            mail.Categories = "";
                            mail.Categories = "1. 等別人回覆 (急)";
                            break;

                        case PriorityType.NORMAL:
                            mail.Categories = "";
                            mail.Categories = "2. 等別人回覆 (不急)";
                            break;

                        case PriorityType.PROJECT_1:
                            mail.Categories = "";
                            mail.Categories = "3. Project 1";
                            break;

                        case PriorityType.PROJECT_2:
                            mail.Categories = "";
                            mail.Categories = "4. Project 2";
                            break;
                    }
                    mail.Save();
                }
                return;
            }

            if (ControlIsInExplorer(control, ref explorer) == true)
            {
                Outlook.Selection selection = null;
                try
                {
                    // TODO: remove this 'Today' check, if should be disabled in Today page.

                    // I have to wrap 'explorer.selection' into a try block,
                    // because Outlook will raise an exception on this line when the first page is 'Outlook Today'
                    selection = explorer.Selection;
                    foreach (dynamic selected in selection)
                    {
                        mail = selected as Outlook.MailItem;
                        if (mail != null)
                        {
                            switch (priorityType)
                            {
                                case PriorityType.HIGH:
                                    mail.Categories = "";
                                    mail.Categories = "1. 等別人回覆 (急)";
                                    break;

                                case PriorityType.NORMAL:
                                    mail.Categories = "";
                                    mail.Categories = "2. 等別人回覆 (不急)";
                                    break;

                                case PriorityType.PROJECT_1:
                                    mail.Categories = "";
                                    mail.Categories = "3. Project 1";
                                    break;

                                case PriorityType.PROJECT_2:
                                    mail.Categories = "";
                                    mail.Categories = "4. Project 2";
                                    break;
                            }
                            mail.Save();
                        }
                    }
                }
                catch (COMException ex)
                {
                    Debug.WriteLine(ex.ToString());
                }
                return;
            }
        }

        public void GroupReply_Click(Office.IRibbonControl control, bool pressed)
        {
            Outlook.MailItem mail = null;
            int index = int.Parse(control.Tag);

            Outlook.Explorer explorer = null;
            Outlook.Inspector inspector = null;

            if (ControlIsInInspector(control, ref inspector) == true)
            {
                mail = inspector.CurrentItem as Outlook.MailItem;
                if (mail != null)
                {
                    var action = mail.Actions[index];
                    action.Enabled = !pressed;

                    mail.Save();

                    InvalidateControl("btnNoReplyAll");
                    InvalidateControl("btnNoReply");
                    InvalidateControl("btnForward");
                }
                return;
            }

            if (ControlIsInExplorer(control, ref explorer) == true)
            {
                Outlook.Selection selection = null;
                try
                {
                    // I have to wrap 'explorer.Selection' into a try block,
                    // because Outlook will raise an exception on this line when the first page is 'Outlook Today'
                    selection = explorer.Selection;
                    if (selection.Count == 1)
                    {
                        mail = selection[1] as Outlook.MailItem;
                        if (mail != null)
                        {
                            var action = mail.Actions[index];
                            action.Enabled = !pressed;

                            mail.Save();

                            InvalidateControl("btnNoReplyAll");
                            InvalidateControl("btnNoReply");
                            InvalidateControl("btnForward");
                        }
                    }
                }
                catch (COMException /* ex */)
                {
                }
                return;
            }
        }

        public bool GroupReply_IsPressed(Office.IRibbonControl control)
        {
            Outlook.MailItem item = null;
            bool pressed = false;
            int index = int.Parse(control.Tag);

            Outlook.Explorer explorer = null;
            Outlook.Inspector inspector = null;

            if (ControlIsInInspector(control, ref inspector) == true)
            {
                item = inspector.CurrentItem as Outlook.MailItem;
                if (item != null)
                {
                    var action = item.Actions[index];
                    pressed = !action.Enabled;
                }
                return pressed;
            }

            if (ControlIsInExplorer(control, ref explorer) == true)
            {
                if (Util.IsCurrentFolderDefaultOutbox(explorer) == true)
                {
                    return false;
                }

                Outlook.Selection selection = null;
                try
                {
                    selection = explorer.Selection;
                    if (selection.Count == 1)
                    {
                        item = selection[1] as Outlook.MailItem;
                        if (item != null)
                        {
                            var action = item.Actions[index];
                            pressed = !action.Enabled;
                        }
                    }
                }
                catch (COMException /* ex */)
                {
                }
                return pressed;
            }

            return pressed;
        }

        public Bitmap GroupReply_GetImage(Office.IRibbonControl control)
        {
            switch (control.Tag)
            {
                case "1":
                    return GetImage("disable_mail_reply", 64);
                case "2":
                    return GetImage("disable_mail_reply_all", 64);
                case "3":
                    return GetImage("disable_mail_forward", 64);

                default:
                    Debug.Assert(false);
                    return null as Bitmap;
            }
        }

        public bool GroupReply_IsEnabled(Office.IRibbonControl control)
        {
            Outlook.Explorer explorer = null;
            Outlook.Inspector inspector = null;

            if (ControlIsInExplorer(control, ref explorer) == false)
            {
                // should be in an inspector

                ControlIsInInspector(control, ref inspector);
                Debug.Assert(inspector != null);

                Outlook.MailItem mail = inspector.CurrentItem as Outlook.MailItem;
                if (mail != null)
                {
                    Outlook.Folder folder = mail.Parent as Outlook.Folder;
                    if (folder != null)
                    {
                        if ((Util.IsFolderDefaultOutbox(inspector.Application.ActiveExplorer(), folder) == true) ||
                            (folder.Store.ExchangeStoreType == Outlook.OlExchangeStoreType.olNotExchange))
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }

            Outlook.Selection selection = null;

            // I have to wrap 'explorer.selction' into a try block,
            // becasue outlook will raise an exception on this line when the first page is 'Outlook Today'
            try
            {
                selection = explorer.Selection;
                if (selection.Count != 1)
                {
                    return false;
                }
                else
                {
                    Outlook.MailItem mail = selection[1] as Outlook.MailItem;

                    if (mail != null)
                    {
                        Outlook.Folder folder = mail.Parent as Outlook.Folder;
                        Debug.Assert(folder != null);

                        switch (control.Tag)
                        {
                            case "1":
                            case "2":
                            case "3":
                                if ((Util.IsCurrentFolderDefaultOutbox(explorer) == true) ||
                                    (folder.Store.ExchangeStoreType == Outlook.OlExchangeStoreType.olNotExchange))
                                {
                                    return false;
                                }
                                else
                                {
                                    return true;
                                }

                            default:
                                Debug.Assert(false);
                                return false;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            return false;
        }

        public Bitmap Button_GetImage(Office.IRibbonControl control)
        {
            switch (control.Tag)
            {
                case "BackupEmail":
                    return GetImage("backup_email", 64);

                default:
                    Debug.Assert(false);
                    return null as Bitmap;
            }
        }

        public Bitmap GroupPriority_GetImage(Office.IRibbonControl control)
        {
            switch (control.Tag)
            {
                case "HighPriorityEmail":
                    return GetImage("high_priority_email", 64);

                case "NormalPriorityEmail":
                    return GetImage("normal_priority_email", 64);

                case "Project1Email":
                    return GetImage("project_1", 64);

                case "Project2Email":
                    return GetImage("project_2", 64);

                default:
                    Debug.Assert(false);
                    return null as Bitmap;
            }
        }

        public bool Flag_GetPressed(Office.IRibbonControl control)
        {
            return Util.GetFlag(control.Tag);
        }

        public void Flag_Action(Office.IRibbonControl control, bool pressed)
        {
            Util.SetFlag(control.Tag, pressed);

            if (control.Tag == "AutoDetectCurrentAccountEmailAddr" && pressed == true)
            {
                // use the 1st account info
                string smtpAddress = Util.GetFirstAccountSmtpAddress(((Outlook.Explorer)control.Context).Application);
                Util.SetValue("AutoBccEmailAddr", smtpAddress);
                InvalidateControl("editBoxAutoBccEmailAddr");
            }
            else if (control.Tag == "AutoDetectMyEmailAddr" && pressed == true)
            {
                // use the 1st account info
                string smtpAddress = Util.GetFirstAccountSmtpAddress(((Outlook.Explorer)control.Context).Application);
                Util.SetValue("AutoBackupMyEmailAddr", smtpAddress);
                InvalidateControl("editBoxAutoBackupMyEmailAddr");
            }
        }

        public string EditBox_GetText(Office.IRibbonControl control)
        {
            if (control.Tag == "AutoBccEmailAddr")
            {
                if (Util.GetFlag("AutoDetectCurrentAccountEmailAddr") == true)
                {
                    // use the 1st account info
                    return Util.GetFirstAccountSmtpAddress(((Outlook.Explorer)control.Context).Application);
                }
                else
                {
                    return Util.GetValue(control.Tag);
                }
            }
            else if (control.Tag == "AutoBackupMyEmailAddr")
            {
                if (Util.GetFlag("AutoDetectMyEmailAddr") == true)
                {
                    // use the 1st account info
                    return Util.GetFirstAccountSmtpAddress(((Outlook.Explorer)control.Context).Application);
                }
                else
                {
                    return Util.GetValue(control.Tag);
                }
            }
            else
            {
                return Util.GetValue(control.Tag);
            }
        }

        public void EditBox_OnChange(Office.IRibbonControl control, string text)
        {
            Util.SetValue(control.Tag, text);
        }

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private bool ControlIsInExplorer(Office.IRibbonControl control, ref Outlook.Explorer explorer)
        {
            explorer = control.Context as Outlook.Explorer;
            if (explorer == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool ControlIsInExplorer(Office.IRibbonControl control)
        {
            Outlook.Explorer explorer = null;
            return ControlIsInExplorer(control, ref explorer);
        }

        private bool ControlIsInInspector(Office.IRibbonControl control, ref Outlook.Inspector inspector)
        {
            inspector = control.Context as Outlook.Inspector;
            if (inspector == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool ControlIsInInspector(Office.IRibbonControl control)
        {
            Outlook.Inspector inspector = null;
            return ControlIsInInspector(control, ref inspector);
        }

        public void InvalidateControl(string ControlID)
        {
            if (ribbon != null)
            {
                ribbon.InvalidateControl(ControlID);
            }
        }

        public void InvalidateGroupReply()
        {
            if (ribbon != null)
            {
                ribbon.InvalidateControl("btnNoReplyAll");
                ribbon.InvalidateControl("btnNoReply");
                ribbon.InvalidateControl("btnNoForward");
            }
        }

        private static Bitmap GetImage(string imageName, int targetSize)
        {
            int screenSize = GetScaledSize(targetSize);
            var bitmap = Properties.Resources.ResourceManager.GetObject(imageName + "_" + targetSize + "x" + targetSize) as Bitmap;
            if (targetSize < screenSize)
            {
                var scaledBitmap = new Bitmap(screenSize, screenSize);
                using (var graphics = Graphics.FromImage(scaledBitmap))
                {
                    int offset = (screenSize - targetSize) / 2;
                    graphics.DrawImage(bitmap, new Rectangle(offset, offset, targetSize, targetSize));
                }
                bitmap.Dispose();
                bitmap = scaledBitmap;
            }
            return bitmap;
        }

        private static int GetScaledSize(int requestedSize)
        {
            return requestedSize * GetDPI() / 96;
        }

        private static int GetDPI()
        {
            return GetDeviceCaps(88 /*LOGPIXELSX*/ );
        }

        private static int GetDeviceCaps(int index)
        {
            IntPtr hdc = GetDC(IntPtr.Zero);
            int val = GetDeviceCaps(hdc, index);
            ReleaseDC(IntPtr.Zero, hdc);
            return val;
        }

        #endregion

        #region Imports

        [DllImport("gdi32.dll")]
        private extern static int GetDeviceCaps(IntPtr hdc, int index);

        [DllImport("user32.dll")]
        private extern static int ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("user32.dll")]
        private extern static IntPtr GetDC(IntPtr hwnd);

        #endregion
    }
}
