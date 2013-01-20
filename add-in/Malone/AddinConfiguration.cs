using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;

namespace Malone
{
    class AddinConfiguration
    {
        bool debugMsg = false;
        private static AddinConfiguration instance;

        private string webUrl = "";
        private string uploadXqy = "";

        private string isPaneEnabled = "";
        private bool paneEnabled = false;

        private AddinConfiguration()
        {
            initializeConfig();
        }

        public static AddinConfiguration GetInstance()
        {

            if (instance == null)
            {
                instance = new AddinConfiguration();
            }

            return instance;
        }

        private void initializeConfig()
        {

            RegistryKey regKey1 = Registry.CurrentUser;
            regKey1 = regKey1.OpenSubKey(@"MaloneAddinConfiguration\Malone");

            if (regKey1 == null)
            {
                if (debugMsg)
                    MessageBox.Show("KEY IS  NULL");

            }
            else
            {
                if (debugMsg)
                    MessageBox.Show("KEY IS: " + regKey1.GetValue("URL"));

                webUrl = (string)regKey1.GetValue("URL");
                isPaneEnabled = (string)regKey1.GetValue("MTPEnabled");

                //only used user/auth for button, can pass into function from js
                //user = (string)regKey1.GetValue("User");
                //auth = (string)regKey1.GetValue("Auth");
                uploadXqy = (string)regKey1.GetValue("UploadXqy");

                if (isPaneEnabled.ToUpper().Equals("TRUE"))
                {
                    paneEnabled = true;
                }

            }

        }

        public string getWebURL()
        {
            return webUrl;
        }


        public string getUploadXqy()
        {
            return uploadXqy;
        }

        public bool getPaneEnabled()
        {
            return paneEnabled;
        }
    }
}
