using EmailMarketing.Library;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using MetroFramework;

namespace EmailMarketing.Desktop
{
    public partial class frmUnsubscribe : MetroForm
    {
        public frmUnsubscribe()
        {
            InitializeComponent();
            PopulateUserGroups();
        }

        private void btnUnsubUsers_Click(object sender, EventArgs e)
        {
            int totalUnsubscribed = 0;
            if (cboUnsubPickGroup.Text == "All Groups")
            {
                Helper.UnsubscribeFromAllGroups(AppConstants.GroupFolderName, txtUnsubEmails.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None).ToList(), out totalUnsubscribed);
            }
            else
            {
                string fileName = AppConstants.GroupFolderName + cboUnsubPickGroup.Text + ".csv";

                if (fileName != null)
                {
                    Helper.UnsubscribeFromSingleGroup(fileName, txtUnsubEmails.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None).ToList(), out totalUnsubscribed);
                }
            }
            
            MetroMessageBox.Show(this, "User unsubscribed");
        }

        void PopulateUserGroups()
        {
            if (Directory.Exists(AppConstants.GroupFolderName))
            {
                cboUnsubPickGroup.Items.Clear();
                cboUnsubPickGroup.Items.Add("All Groups");

                foreach (var file in Directory.GetFiles(AppConstants.GroupFolderName))
                {
                    cboUnsubPickGroup.Items.Add(Path.GetFileNameWithoutExtension(file));
                }
            }
        }
    }
}
