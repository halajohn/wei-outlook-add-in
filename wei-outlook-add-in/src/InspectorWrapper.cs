using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace wei_outlook_add_in {
    internal delegate void InspectorWrapperClosedEventHandler(Guid id);

    internal class InspectorWrapper {
        public event InspectorWrapperClosedEventHandler Closed;

        public Guid Id { get; private set; }
        public Outlook.Inspector Inspector { get; private set; }

        public InspectorWrapper(Outlook.Inspector inspector) {
            Id = Guid.NewGuid();
            Inspector = inspector;

            ((Outlook.InspectorEvents_10_Event)Inspector).Close += new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate += new Outlook.InspectorEvents_10_ActivateEventHandler(Activate);
        }

        private void Inspector_Close_() {
            ((Outlook.InspectorEvents_10_Event)Inspector).Close -= new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate -= new Outlook.InspectorEvents_10_ActivateEventHandler(Activate);
            Closed?.Invoke(Id);
        }

        private void Inspector_Close() {
            Inspector_Close_();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        protected virtual void Activate() {
            System.Windows.Forms.Application.DoEvents();

            Microsoft.Office.Interop.Word.Document wdDoc = Util.GetWordEditor(Inspector);
            if (wdDoc != null) {
                wdDoc.Windows[1].View.Zoom.Percentage = Config.Zoom;
            }
        }

        public static InspectorWrapper GetWrapperFor(Outlook.Inspector inspector) {
            if (inspector.CurrentItem is Outlook.MailItem) {
                return new InspectorWrapper(inspector);
            } else {
                return null;
            }
        }
    }
}
