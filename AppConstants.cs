using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailMarketing.Library
{
    public class AppConstants
    {
        public static readonly string DirectoryPath = AppDomain.CurrentDomain.BaseDirectory;
        public static readonly string CampaignFolderName = AppDomain.CurrentDomain.BaseDirectory + "Campaigns/";
        public static readonly string GroupFolderName = AppDomain.CurrentDomain.BaseDirectory + "UserGroups/";
        public static readonly string ReportsFolderName = AppDomain.CurrentDomain.BaseDirectory + "Reports/";
        public static readonly string ImportDemoCSVFileLocation = AppDomain.CurrentDomain.BaseDirectory + "import.csv";
    }
}
