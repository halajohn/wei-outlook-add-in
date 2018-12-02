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

        private void Inspector_Close() {
            ((Outlook.InspectorEvents_10_Event)Inspector).Close -= new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate -= new Outlook.InspectorEvents_10_ActivateEventHandler(Activate);
            Closed?.Invoke(Id);
        }

        protected virtual void Activate() {
            Microsoft.Office.Interop.Word.Document wdDoc = Util.GetWordEditor(Inspector);
            if (wdDoc != null) {
                wdDoc.Windows[1].View.Zoom.Percentage = Config.Zoom;
            }
        }

        public static InspectorWrapper GetWrapperFor(Outlook.Inspector inspector) {
            if (true == Util.IsMailItem(inspector.CurrentItem as object)) {
                return new InspectorWrapper(inspector);
            } else {
                return null;
            }
        }
    }
}
