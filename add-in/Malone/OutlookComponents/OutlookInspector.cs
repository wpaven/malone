using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;
using System.Windows.Forms;

namespace Malone
{

    /// <summary>
    /// This class tracks the state of an Outlook Inspector window for your
    /// add-in and ensures that what happens in this window is handled correctly.
    /// </summary>
    class OutlookInspector
    {
        #region Instance Variables


        private Outlook.Inspector m_Window;             // wrapped window object
        // Use these instance variables to handle item-level events
        private Outlook.MailItem m_Mail;                // wrapped MailItem
        private Outlook.AppointmentItem m_Appointment;  // wrapped AppointmentItem
        private Outlook.ContactItem m_Contact;          // wrapped ContactItem
        private Outlook.ContactItem m_Task;             // wrapped TaskItem
        // Define other class-level item instance variables as needed
        public Tools.CustomTaskPane taskPane;
        #endregion

        #region Events

        public event EventHandler Close;
        public event EventHandler<InvalidateEventArgs> InvalidateControl;

        #endregion

        #region Constructor

        /// <summary>
        /// Create a new instance of the tracking class for a particular 
        /// inspector and custom task pane.
        /// </summary>
        /// <param name="inspector">A new inspector window to track</param>
        ///<remarks></remarks>
        public OutlookInspector(Outlook.Inspector inspector)
        {
            m_Window = inspector;

            // Hookup the close event
            ((Outlook.InspectorEvents_Event)inspector).Close +=
                new Outlook.InspectorEvents_CloseEventHandler(
                OutlookInspectorWindow_Close);

            taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(
                new UserControl1(), "My task pane", m_Window);
            taskPane.Visible = true;
//  taskPane.VisibleChanged += new EventHandler(TaskPane_VisibleChanged);

            // Hookup item-level events as needed
            // For example, the following code hooks up PropertyChange
            // event for a ContactItem
            //OutlookItem olItem = new OutlookItem(inspector.CurrentItem);
            //if(olItem.Class==Outlook.OlObjectClass.olContact)
            //{
            //    m_Contact = olItem.InnerObject as Outlook.ContactItem;
            //    m_Contact.PropertyChange +=
            //        new Outlook.ItemEvents_10_PropertyChangeEventHandler(
            //        m_Contact_PropertyChange);
            //}

        }
        #endregion

        #region Event Handlers
        void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            //Globals.ThisAddIn.CustomTaskPanes[this.taskPane].Visible = true;
            MessageBox.Show("Toggle");
            
            if(taskPane.Visible)
            {
                taskPane.Visible = false;
            }
            else
            {
                taskPane.Visible = true;
            }
           //Globals.ThisAddIn
           // Globals.Ribbons[m_Window]   Ribbon1.toggleButton1.Checked =
              //  taskPane.Visible;
         
           // Globals.Ribbons[m_Window].Ribbon1.toggleButton1.Checked =
            //    taskPane.Visible;
        }

        /// <summary>
        /// Event Handler for the inspector close event.
        /// </summary>
        private void OutlookInspectorWindow_Close()
        {
            // Unhook events from any item-level instance variables
            //m_Contact.PropertyChange -= 
            //    Outlook.ItemEvents_10_PropertyChangeEventHandler(
            //    m_Contact_PropertyChange);

            // Unhook events from the window

            if (taskPane != null)
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
            }

            taskPane = null;

            ((Outlook.InspectorEvents_Event)m_Window).Close -=
                new Outlook.InspectorEvents_CloseEventHandler(
                OutlookInspectorWindow_Close);

            // Raise the OutlookInspector close event
            if (Close != null)
            {
                Close(this, EventArgs.Empty);
            }

            // Unhook any item-level instance variables
            //m_Contact = null;
            m_Window = null;
        }


        //void  m_Contact_PropertyChange(string Name)
        //{
        //    // Implement PropertyChange here
        //}
        #endregion

        #region Methods
        private void RaiseInvalidateControl(string controlID)
        {
            if (InvalidateControl != null)
                InvalidateControl(this, new InvalidateEventArgs(controlID));
        }
        #endregion

        #region Properties

        /// <summary>
        /// The actual Outlook inspector window wrapped by this instance
        /// </summary>
        internal Outlook.Inspector Window
        {
            get { return m_Window; }
        }

        #endregion

        #region Helper Class
        public class InvalidateEventArgs : EventArgs
        {
            private string m_ControlID;

            public InvalidateEventArgs(string controlID)
            {
                m_ControlID = controlID;
            }

            public string ControlID
            {
                get { return m_ControlID; }
            }
        }
        #endregion
    }
}
