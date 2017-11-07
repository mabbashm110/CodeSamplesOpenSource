using System;
using System.Linq;
using System.Runtime.InteropServices;
using MetroFramework.Controls;
using System.Net;
using System.IO;

namespace EmailMarketing.Library.StandardLibraries
{
    public static class NotificationDefaults
    {
        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);
        public static readonly string NotificationFolder = AppDomain.CurrentDomain.BaseDirectory + "Notify/";
        public static readonly string SupportURL = "<<removed>>";
        public static readonly string NotificationsURL = "<<removed>>";
        public static readonly string DeveloperURL = "http://www.freewindowsapps.com";
        public static readonly string NoInternetNotification = "It appears that you are not connected to the internet.\nCheck your internet connection and try again.";
        public static string FileLoc { get; set; }

        public static bool CheckInternet()
        {
            int Desc;
            return InternetGetConnectedState(out Desc, 0);
        }

        private static void CreateNotificationsDirectory()
        {
            if (!Directory.Exists(NotificationFolder))
            {
                Directory.CreateDirectory(NotificationFolder);
            }
        }

        /// <summary>
        /// Setting Tab Preferences for tabMetro and mTabPage Name
        /// </summary>
        /// <param name="ad">Notification<T>"/></param>
        /// <param name="mTabPage">Metro Tab Page Name</param>
        /// <param name="tabMetro">Metro Tab Name</param>
        /// <param name="browser">Web Browser Name for Help</param>
        /// <param name="pbname">Picture Box Name for Advert Image</param>
        public static void SetTabPreferences(out Notification ad, MetroTabPage mTabPage = null, MetroTabControl tabMetro = null)
        {
            bool internet = CheckInternet();
            ad = new Notification();
            if (mTabPage.Name.Contains("Home"))
            {
                tabMetro.SelectedTab = mTabPage;
                ad = null;
            }
            else if (mTabPage.Name == "tbNotifications")
            {
                string message = string.Empty;
                ad = ProcessNotificationsFile(out message);
            }
        }

        public static Notification ProcessNotificationsFile(out string message)
        {
            Notification ad = new Notification();
            CreateNotificationsDirectory();
            if (CheckInternet() == true)
            {
                WebClient webclient = new WebClient();
                webclient.DownloadFile(NotificationsURL, @NotificationFolder + "/offers.txt");
                webclient.Dispose();
            }
            try
            {
                string[] marketingLines = File.ReadAllLines(@NotificationFolder + "/offers.txt");
                foreach (string line in marketingLines)
                {
                    ad.ProductName = marketingLines[0];
                    ad.ProductDescription = marketingLines[1];
                    ad.URLText = marketingLines[2];
                    ad.ImageLocation = marketingLines[3];
                    ad.URLTagClick = marketingLines[4];
                }
                message = string.Empty;
                return ad;
            }
            catch (Exception)
            {
                message = NoInternetNotification;
                return ad;
            }

        }
    }
}
