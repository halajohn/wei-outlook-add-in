using Outlook = Microsoft.Office.Interop.Outlook;
using System;

namespace wei_outlook_add_in
{
    internal class MailItemWrapper : InspectorWrapper
    {
        // .ctor
        // <param name="inspector">The Outlook Inspector instance that is to be handled.</param>
        public MailItemWrapper(Outlook.Inspector inspector)
            : base(inspector)
        {
        }

        // The Object instance behind the Inspector, which is the current item.
        public Outlook.MailItem Item { get; private set; }

        // Method is called when the wrapper has been initialized.
        protected override void Initialize()
        {
            // Get the item of the current Inspector.
            Item = (Outlook.MailItem)Inspector.CurrentItem;

            // Register for the item events.
            Item.Open += new Outlook.ItemEvents_10_OpenEventHandler(Item_Open);
            Item.Write += new Outlook.ItemEvents_10_WriteEventHandler(Item_Write);
        }

        // This method is called when the item is saved.
        // <param name="Cancel">When set to true, the save operation is cancelled.</param>
        void Item_Write(ref bool Cancel)
        {
            //TODO: Implement something 
        }

        // This method is called when the item is visible and the UI is initialized.
        // <param name="Cancel">When you set this property to true, the Inspector is closed.</param>
        void Item_Open(ref bool Cancel)
        {
            // 'Sent == false' means a compose inspector
            if (Item.Sent == false)
            {
                var action = Item.Actions[2];
                bool wanted_action_value = !(Util.GetFlag("DisableReplyAllForNewMsg"));

                // If the ReplyAll's value (Action[2].Enabled) meet what user wants,
                // do not set the same value again, otherwise, Outlook would think the user has changed the property,
                // and request user to save that email item. (Even if the new value equals to the old value)
                if (action.Enabled != wanted_action_value)
                {
                    action.Enabled = wanted_action_value;

                    // TODO: In Outlook 2013 MSDN:
                    // "If a mail item is an inline reply, calling Save on that mail item may fail
                    // and result in unexpected behavior.
                    Item.Save();
                }
            }

            if (ThisAddIn.ribbonExtensibilityObject != null)
            {
                ThisAddIn.ribbonExtensibilityObject.InvalidateGroupReply();
            }
        }

        // The Close method is called when the inspector has been closed.
        // The UI is gone, cannot access it here.
        protected override void Close()
        {
            // Unregister events.
            Item.Write -= new Outlook.ItemEvents_10_WriteEventHandler(Item_Write);
            Item.Open -= new Outlook.ItemEvents_10_OpenEventHandler(Item_Open);

            // Set item to null to keep a reference in memory of the garbage collector.
            Item = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
