using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;
using CardaxFTEIAPILib;
using Scripting;

namespace FTCAPI_Centrak
{
    public partial class Form1 : Form, IFTMiddleware2, IFTTrigger2
    {
        //Create the event input API. 
        private IFTExternalEventInput2 iAPI;
        //Used for connecting IFTMiddleware2 and IFTTrigger2 manually.
        private int cookieMiddle;
        private int cookieTrigger;
        private IConnectionPoint icpMiddle;
        private IConnectionPoint icpTrigger;
        //Script file.
        private FileSystemObject fsoScriptFile;
        private int callCounter;
        private DateTime callStart;
        private TextStream tsScriptFile;
        public Form1()
        {
            InitializeComponent();
        }
        [MTAThread]
        private void Form1_Load(object sender, EventArgs e)
        {
            InitialConnection();
            label1.Text = "Disconnect";
            label1.ForeColor = System.Drawing.Color.Red;
        }

        private void InitialConnection()
        {
            try
            {
                //Create the API component as a CardaxFTEIAPIClass object.
                iAPI = new CardaxFTEIAPIClass();

                //cboRegionCode.SelectedIndex = 0;
                //optLogEvent_CheckedChanged(optLogEvent, new EventArgs());
                //fsoScriptFile = new FileSystemObjectClass();

                //Connect to the IFTMiddleware2 interface (i.e. a non-default outbound interface on the API).
                IConnectionPointContainer icpc = (IConnectionPointContainer)iAPI;
                Guid guidMiddle = new Guid("ED5CDD41-4393-4b88-BF30-F250905D57C8");
                icpc.FindConnectionPoint(ref guidMiddle, out icpMiddle);
                icpMiddle.Advise(this, out cookieMiddle);

                //Connect to the IFTTrigger2 interface (i.e. a non-default outbound interface on the API).
                Guid guidTrigger = new Guid("C6BDE0BA-B4B1-495e-BE28-0D68A5C9AFFA");
                icpc.FindConnectionPoint(ref guidTrigger, out icpTrigger);
                icpTrigger.Advise(this, out cookieTrigger);
            }
            catch (Exception exception)
            {
                String sText = "Error while loading: " + exception.Message;

                MessageBox.Show(sText, "API error", MessageBoxButtons.OK);
            }
        }

        public void notifyItemDeregistered(string sSystemID, string sItemID)
        {
           // txtLog.Text = txtLog.Text + "notifyItemDeregistered. sSystemID: " + sSystemID + ", sItemID: " + sItemID + "\r\n";
        }
        public void notifyItemRegistered(string sSystemID, string sItemID, string sConfig)
        {
            //txtLog.Text = txtLog.Text + "notifyItemRegistered. sSystemID: " + sSystemID + ", sItemID: " + sItemID + ", sConfig: " + sConfig + "\r\n";
            if (label1.InvokeRequired)
            {
                label1.Invoke(new MethodInvoker(delegate { label1.Text = "Connect "+ sSystemID + " " + sItemID; label1.ForeColor = System.Drawing.Color.Green; }));
            }
            //if (name == "MyName")
            //{
            //    // do whatever
            //}
            //label1.Text = "Connect";
            //label1.ForeColor = System.Drawing.Color.Green;
        }
        public void notifySystemDeregistered(string sSystemID)
        {
           // txtLog.Text = txtLog.Text + "notifySystemDeregistered. sSystemID: " + sSystemID + "\r\n";
        }
        public void notifySystemRegistered(string sSystemID, string sTypeID, string sConfig)
        {
           // txtLog.Text = txtLog.Text + "notifySystemRegistered. sSystemID: " + sSystemID + ", sTypeID: " + sTypeID + ", sConfig: " + sConfig + "\r\n";
        }
        public void notifyAlarmAcknowledged(string sSystemID, int lEventID)
        {
           // txtLog.Text = txtLog.Text + "notifyAlarmAcknowledged. sSystemID: " + sSystemID + ", EventID: " + Convert.ToString(lEventID) + "\r\n";
        }

        /*
         * The following methods represent the implementaion of IFTTrigger2 interface.
         */
        public void triggerEvent(string sSystemID, string sItem, string sTrigger)
        {
            //txtTrigger.Text = txtTrigger.Text + "triggerEvent. sSystemID: " + sSystemID + ", sItem: " + sItem + ", sTrigger: " + sTrigger + "\r\n";
        }
        public void triggerOutput(string sSystemID, string sItem, int lState, string sTrigger)
        {
            //txtTrigger.Text = txtTrigger.Text + "triggerOutput. sSystemID: " + sSystemID + ", sItem: " + sItem + ", lState: " + Convert.ToString(lState) + ", sTrigger: " + sTrigger + "\r\n";
        }

        private void btnOn_Click(object sender, EventArgs e)
        {
            DateTime dNow = DateTime.Now;
            DateTime dNowUTC = dNow.ToUniversalTime();
            iAPI.notifyStatus("ES01","ESI01", 1, 0, 0, "Alarm Trigger");
            iAPI.logEvent(1, 1, dNowUTC, 1, "ES01", "ESI01", "Alarm On", "Alarm Trigger");
        }

        private void btnOff_Click(object sender, EventArgs e)
        {
            iAPI.notifyStatus("ES01", "ESI01", 0, 0, 0, "Resume Normal");
            iAPI.notifyRestore(1, "ES01", "ESI01");
        }


    }
}
