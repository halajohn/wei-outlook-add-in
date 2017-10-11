using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Reflection;

namespace wei_outlook_add_in
{
    // Event handler used to correctly clean up resources.
    // <param name="id">The unique id of the Inspector instance.</param>
    internal delegate void InspectorWrapperClosedEventHandler(Guid id);

    // The base class for all inspector wrappers.
    internal abstract class InspectorWrapper
    {
        // Event notification for the InspectorWrapper.Closed event.
        // This event is raised when an inspector has been closed.
        public event InspectorWrapperClosedEventHandler Closed;

        // The unique ID that identifies the inspector window.
        public Guid Id { get; private set; }

        // The Outlook Inspector instance.
        public Outlook.Inspector Inspector { get; private set; }

        // .ctor
        // <param name="inspector">The Outlook Inspector instance that should be handled.</param>
        public InspectorWrapper(Outlook.Inspector inspector)
        {
            Id = Guid.NewGuid();
            Inspector = inspector;

            // Register Inspector events here
            ((Outlook.InspectorEvents_10_Event)Inspector).Close += new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate += new Outlook.InspectorEvents_10_ActivateEventHandler(Activate);
            ((Outlook.InspectorEvents_10_Event)Inspector).Deactivate += new Outlook.InspectorEvents_10_DeactivateEventHandler(Deactivate);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMaximize += new Outlook.InspectorEvents_10_BeforeMaximizeEventHandler(BeforeMaximize);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMinimize += new Outlook.InspectorEvents_10_BeforeMinimizeEventHandler(BeforeMinimize);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMove += new Outlook.InspectorEvents_10_BeforeMoveEventHandler(BeforeMove);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeSize += new Outlook.InspectorEvents_10_BeforeSizeEventHandler(BeforeSize);
            ((Outlook.InspectorEvents_10_Event)Inspector).PageChange += new Outlook.InspectorEvents_10_PageChangeEventHandler(PageChange);

            // Initialize is called to give the derived wrappers.
            Initialize();
        }

        // Event handler for the Inspector Close event.
        private void Inspector_Close()
        {
            // Call the Close method - the derived classes can implement cleanup code
            // by overriding the Close method.
            Close();

            // Unregister Inspector events.
            ((Outlook.InspectorEvents_10_Event)Inspector).Close -= new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate -= new Outlook.InspectorEvents_10_ActivateEventHandler(Activate);
            ((Outlook.InspectorEvents_10_Event)Inspector).Deactivate -= new Outlook.InspectorEvents_10_DeactivateEventHandler(Deactivate);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMaximize -= new Outlook.InspectorEvents_10_BeforeMaximizeEventHandler(BeforeMaximize);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMinimize -= new Outlook.InspectorEvents_10_BeforeMinimizeEventHandler(BeforeMinimize);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMove -= new Outlook.InspectorEvents_10_BeforeMoveEventHandler(BeforeMove);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeSize -= new Outlook.InspectorEvents_10_BeforeSizeEventHandler(BeforeSize);
            ((Outlook.InspectorEvents_10_Event)Inspector).PageChange -= new Outlook.InspectorEvents_10_PageChangeEventHandler(PageChange);

            // Clean up resources and do a GC.Collect().
            Inspector = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // Raise the Close event.
            if (Closed != null)
            {
                Closed(Id);
            }
        }

        protected virtual void Initialize() { }

        // Method is called when another page of the inspector has been selected.
        // <param name="ActivePageName">The active page name by reference.</param>
        protected virtual void PageChange(ref string ActivePageName) { }

        // Method is called before the inspector is resized.
        // <param name="Cancel">To prevent resizing, set Cancel to true.</param>
        protected virtual void BeforeSize(ref bool Cancel) { }

        // Method is called before the inspector is moved around.
        // <param name="Cancel">To prevent moving, set Cancel to true.</param>
        protected virtual void BeforeMove(ref bool Cancel) { }

        // Method is called before the inspector is minimized.
        // <param name="Cancel">To prevent minimizing, set Cancel to true.</param>
        protected virtual void BeforeMinimize(ref bool Cancel) { }

        // Method is called before the inspector is maximized.
        // <param name="Cancel">To prevent maximizing, set Cancel to true.</param>
        protected virtual void BeforeMaximize(ref bool Cancel) { }

        // Method is called when the inspector is deactivated.
        protected virtual void Deactivate() { }

        // Method is called when the inspector is activated.
        protected virtual void Activate()
        {
            dynamic wdDoc = Util.GetWordEditor(Inspector);
            if (wdDoc != null)
            {
                wdDoc.Windows(1).View.Zoom.Percentage = 150;
            }
        }

        // Derived classes can do a cleanup by overriding this method.
        protected virtual void Close() { }

        // This factory method returns a specific InspectorWrapper or null if not handled.
        // <param name=”inspector”>The Outlook Inspector instance.</param>
        // Returns the specific wrapper or null.
        public static InspectorWrapper GetWrapperFor(Outlook.Inspector inspector)
        {
            // Retrieve the message class by using late binding.
            string messageClass = inspector.CurrentItem.GetType().InvokeMember("MessageClass", BindingFlags.GetProperty, null, inspector.CurrentItem, null);

            // Depending on the message class, you can instantiate a
            // different wrapper explicitly for a given message class by
            // using a switch statement.
            switch (messageClass)
            {
                // case "IPM.Contact": return new ContactItemWrapper(inspector);
                // case "IPM.Journal": return new ContactItemWrapper(inspector);
                case "IPM.Note": return new MailItemWrapper(inspector);
                // case "IPM.Post": return new PostItemWrapper(inspector);
                // case "IPM.Task": return new TaskItemWrapper(inspector);
            }

            // Or, check if the message class begins with a specific fragment.
            if (messageClass.StartsWith("IPM.Contact.X4U"))
            {
                // return new X4UContactItemWrapper(inspector);
            }

            // Or, check the interface type of the item.
            if (inspector.CurrentItem is Outlook.AppointmentItem)
            {
                // return new AppointmentItemWrapper(inspector);
            }

            // No wrapper is found.
            return null;
        }
    }
}
