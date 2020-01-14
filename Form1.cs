using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO.Ports;
using Awesomium.Core;
using System.Reflection;
using System.Threading;
using SHDocVw;

namespace WebKiosk
{
    public partial class Form1 : Form
    {
        static SerialPort port;
        SHDocVw.InternetExplorer ie; // = new SHDocVw.InternetExplorer();

        private bool documentLoaded;
        private bool documentPrinted;

        private string data = "";
        
        public Form1()
        {
            InitializeComponent();
            ie = new SHDocVw.InternetExplorer();
            ie.DocumentComplete += ie_DocumentComplete;
            ie.PrintTemplateTeardown += ie_PrintTemplateTeardown;

            webC.DocumentReady += WebViewOnDocumentReady;

            if (Properties.Settings.Default.HideButton) button1.Hide();
            port = new SerialPort(); 
            port.DataReceived += new SerialDataReceivedEventHandler(sp_DataReceived);
            port.PortName = Properties.Settings.Default.COMPort;
            try
            {
            port.Open();
            }
            catch (System.IO.IOException e)
            {
                Console.WriteLine("Error reading port {0}. Message = {1}", port.PortName, e.Message);
                MessageBox.Show("Error reading port " + port.PortName, "My Application", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                button1.Show();
                try
                {
                    port.Close();
                }
                catch (System.IO.IOException ec)
                { }
            }
        }
//  http://stackoverflow.com/questions/13280727/can-i-call-application-methods-from-javascript-in-awesomium
        private void WebViewOnDocumentReady(object sender, UrlEventArgs urlEventArgs)
        {
            webC.DocumentReady -= WebViewOnDocumentReady;
            JSObject jsobject = webC.CreateGlobalJavascriptObject("jsobject");
            jsobject.Bind("PrintPage", false, JSHandler);
            jsobject.Bind("getprintenable", true, JSHandler);
            //Console.WriteLine("Printing set {0}", Properties.Settings.Default.PrintEnable);
            //webC.ExecuteJavascriptWithResult("SetPrintEnable11('" + Properties.Settings.Default.PrintEnable + "')");

        }

        private void ie_DocumentComplete(object pDisp, ref object URL)
        {
            documentLoaded = true;
        }

        private void ie_PrintTemplateTeardown(object pDisp)
        {
            documentPrinted = true;
        }

        public void Printie(string htmlFilename)
        {
            documentLoaded = false;
            documentPrinted = false;

            object missing = Missing.Value;

            ie.Navigate(htmlFilename, ref missing, ref missing, ref missing, ref missing);
            while (!documentLoaded && ie.QueryStatusWB(OLECMDID.OLECMDID_PRINT) != OLECMDF.OLECMDF_ENABLED)
                Thread.Sleep(100);

            ie.ExecWB(OLECMDID.OLECMDID_PRINT, OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, ref missing, ref missing);
            while (!documentPrinted)
                Thread.Sleep(100);

            System.Threading.Thread.Sleep(1000);
            webC.ExecuteJavascriptWithResult("endprint()");
        }

        private void JSHandler(object sender, JavascriptMethodEventArgs args)
        {
            switch (args.MethodName)
            {
                case "PrintPage":
                    System.Diagnostics.Debug.Print("Printing file {0}", args.Arguments[0]);
                    Printie(args.Arguments[0]);
                    break;
                case "getprintenable":
                    args.Result = Properties.Settings.Default.PrintEnable;
                    break;
            }

            //if (args.MustReturnValue)
            //{
            //    Console.WriteLine("Got method call with return request");
            //    args.Result = "Returning " + args.Arguments[0];
            //}
            //else
            //{
            //    Console.WriteLine("Got method call with no return request");
            //}
        }
 
        private delegate void SetText(string text);

        private void sp_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
             Thread.Sleep(500);
             data = port.ReadExisting();
             webC.Invoke(new SetText(si_DataReceived), new object[] { data });
        }

        private string getProximity(string data)
        {
            if (data.Substring(8, 2) != "\r\n") return "error";  
            string code = data.Substring(0, 8);
            int num = Convert.ToInt32(code,16);
            if (num == 0) return "error"; 
            int prx = num;
            if ( Properties.Settings.Default.Use26 ) prx = num >> 1;
            prx = prx & 65535;
            int grp = num >> 16;
            if (Properties.Settings.Default.Use26) grp = num >> 17; 
            grp = grp & 255;
            prx = prx + grp * 100000;
            string ret = Convert.ToString(prx);
            return ret;
        }
        
        private void si_DataReceived(string data)
        {
            if (webC.IsDocumentReady)
            { 
                webC.ExecuteJavascriptWithResult("app.helper.changeProximityCode('" + getProximity(data) + "')");
                webC.Update();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            webC.LoadHTML("<!DOCTYPE html><html><head><style>body{height:1024px;width:1280px;background-color:#fec552}#s{margin-left:432px;margin-top:150px;width:150px}.s{background-color:#ED860D;float:left;height:32px;margin-left:17px;width:32px;-webkit-animation-name:bounce_s;-webkit-animation-duration:2s;-webkit-animation-iteration-count:infinite;-webkit-animation-direction:normal;-webkit-border-radius:21px}#s1{-webkit-animation-delay:.4s}#s2{-webkit-animation-delay:.9s}#s3{-webkit-animation-delay:1.17s}@-webkit-keyframes bounce_s{50%{background-color:#FFF}}</style><body><div id=s><div id=s1 class=s></div><div id=s2 class=s></div><div id=s3 class=s></div></div>");
            webC.Source = new Uri(Properties.Settings.Default.BaseURL);
            webC.ShowContextMenu += Html5_ShowContextMenu;
        }

        void Html5_ShowContextMenu(object sender, Awesomium.Core.ContextMenuEventArgs e)
        {
            e.Handled = true;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            webC.Left = 0;
            webC.Top = 0;
            webC.Width = Width;
            webC.Height = Height; 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ie.Quit();
            port.Close();
            Close();
        }

    }
}
