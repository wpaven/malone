using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;
using System.Runtime.InteropServices;

namespace Malone
{
    [ComVisible(true)]
    public partial class UserControl1 : UserControl
    {
// private AddinConfiguration ac = AddinConfiguration.GetInstance();
        private string webUrl = "";

        public UserControl1()
        {
            InitializeComponent();
// webUrl = ac.getWebURL();

            webBrowser1.AllowWebBrowserDrop = false;
            webBrowser1.IsWebBrowserContextMenuEnabled = false;
            webBrowser1.WebBrowserShortcutsEnabled = false;
            webBrowser1.ObjectForScripting = this;
            //webBrowser1.ScriptErrorsSuppressed = true;
            //webBrowser1.Navigate(webUrl);
            webBrowser1.Navigate("http://localhost:8023/outlookSamples");
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        public void testMail(string myHtml)
        {
            //Outlook.Inspector ins = Globals.ThisAddIn.Application.ActiveInspector();
            //Outlook.MailItem mi = (Outlook.MailItem)ins;


            Outlook.Application myApplication = Globals.ThisAddIn.Application;
            Outlook.Explorer myActiveExplorer = (Outlook.Explorer)myApplication.ActiveExplorer();
            //Outlook.Inspector myInspector = (Outlook.Inspector)myApplication.ActiveInspector();
            Outlook.Selection selection = myActiveExplorer.Selection;
            Outlook.Inspector myInspector = (Outlook.Inspector)myApplication.ActiveInspector();


            if (selection.Count == 1 && selection[1] is Outlook.MailItem)
            {
                /* try
                 {
                     MessageBox.Show("IN IF");
                     //Object selObject = this.ActiveExplorer().Selection[1];

                     Outlook.MailItem mail = (Outlook.MailItem)selection[1];
                     //mail.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                     //mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                     mail.Body = "HELLO THERE!!";
                     //mail.HTMLBody = "<HTML><BODY>Enter the message text here. </BODY></HTML>";
                     mail.Save();
                    
                 }
                 catch (Exception e)
                 {
                     MessageBox.Show("ERROR :" + e.Message);
                 }
                 */
                try
                {
                    Outlook.MailItem mi = (Outlook.MailItem)myInspector.CurrentItem;
                    //mi.Body = "HELLO THERE AGAIN";
                    mi.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                    mi.HTMLBody = myHtml; //"<HTML><BODY>Enter the message text here. </BODY></HTML>";
                    mi.Save();
                }
                catch (Exception e)
                {
                    MessageBox.Show("ERROR 2 :" + e.Message);
                }
            }

        }

        public static void downloadFile(string url, string sourcefile, string user, string pwd)
        {
            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                Client.DownloadFile(url, sourcefile);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public void addAttachment(string path, string title, string url, string user, string pwd)
        {//C:\Users\paven\AppData\Local\Temp


            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            //remove these two when we remove if
            Outlook.Explorer myActiveExplorer = (Outlook.Explorer)myApplication.ActiveExplorer();
            Outlook.Selection selection = myActiveExplorer.Selection;



            Outlook.Inspector myInspector = (Outlook.Inspector)myApplication.ActiveInspector();

            //remove if, check activeIns other way
            if (selection.Count == 1 && selection[1] is Outlook.MailItem)
            {


                try
                {   //object source = @"C:\Users\paven\AppData\Local\Temp\TEST.txt";
                    object source = null;
                    object missing = Type.Missing;
                    Outlook.MailItem mi = (Outlook.MailItem)myInspector.CurrentItem;
                    //mi.Body = "HELLO THERE AGAIN";
                    //mi.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                    //mi.HTMLBody = myHtml; //"<HTML><BODY>Enter the message text here. </BODY></HTML>";
                    string tmpdoc = "";
                    try
                    {
                        tmpdoc = path + title;
                        // MessageBox.Show("TEMPDOC" + tmpdoc);
                        source = tmpdoc;
                        downloadFile(url, tmpdoc, user, pwd);


                    }
                    catch (Exception e)
                    {
                        string errorMsg = e.Message;
                        MessageBox.Show("error: " + errorMsg);
                    }

                    mi.Attachments.Add(source, Outlook.OlAttachmentType.olByValue, missing, missing);
                    mi.Save();
                }
                catch (Exception e)
                {
                    MessageBox.Show("ERROR 3 :" + e.Message);
                }
            }
        }
    }
}
