using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework;
using EmailMarketing.Library;
using System.IO;
using MetroFramework.Forms;
using System.Net;
using System.Diagnostics;
using System.Runtime;
using System.Runtime.InteropServices;
using EmailMarketing.Library.StandardLibraries;

namespace EmailMarketing.Desktop
{
    public partial class frmEMM : MetroForm
    {
        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);
        string NotificationFolder = AppDomain.CurrentDomain.BaseDirectory + "Reports/";
        AppConstants appPaths = new AppConstants();
        private BackgroundWorker bw = new BackgroundWorker();
        List<string> attachmentFileNames = new List<string>();
        List<CPerson> persons = new List<CPerson>();
        List<CPerson> sendMailPersons = new List<CPerson>();
        CCampaign currentCampaign = null;
        Notification notify;

        public frmEMM()
        {
            InitializeComponent();

            bw.WorkerReportsProgress = true;
            bw.WorkerSupportsCancellation = true;

            bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);

            dgvCMCampaign.AutoGenerateColumns = false;
            dgvUSUsers.AutoGenerateColumns = false;

            NotificationDefaults.CheckInternet();
            notify = new Notification();
            //Load default page
            NotificationDefaults.SetTabPreferences(out notify, tbHomeUsersGroups, tabMetro);
            //checkHTMLorPlain();
        }

        /// <summary>
        /// Check if there is a valid internet connection
        /// </summary>
        /// <returns></returns>
        bool CheckInternet()
        {
            int Desc;
            return InternetGetConnectedState(out Desc, 0);
        }

        private void frmEMM_Load(object sender, EventArgs e)
        {
            CheckDirectoryStatus();
            PopulateUserGroups();
            PopulateEmailTemplates();
        }

        /// <summary>
        /// Check if Directories have been created in the first launch.
        /// </summary>
        private void CheckDirectoryStatus()
        {
            if (Properties.Settings.Default.Directories == false)
            {
                try
                {
                    Helper.CreateDirectories();
                    Properties.Settings.Default.Directories = true;
                    Properties.Settings.Default.Save();
                    lblAppStatus.ForeColor = Color.Black;
                    lblAppStatus.Text = "Directories created successfully.";
                }
                catch (Exception)
                {
                    lblAppStatus.ForeColor = Color.Red;
                    lblAppStatus.Text = "There was an error creating directories. Make sure you have the right permissions or contact support.";
                }
            }
        }

        void PopulateUserGroups()
        {
            if (Directory.Exists(AppConstants.GroupFolderName))
            {
                cboCMUserGroup.Items.Clear();
                cboUSPickGroup.Items.Clear();

                foreach (var file in Directory.GetFiles(AppConstants.GroupFolderName))
                {
                    cboCMUserGroup.Items.Add(Path.GetFileNameWithoutExtension(file));
                    cboUSPickGroup.Items.Add(Path.GetFileNameWithoutExtension(file));
                }
            }
        }

        /// <summary>
        /// UI Populate Grid based on group selection from the combo box (User Groups Tab)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboUSPickGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filePath = AppConstants.GroupFolderName + cboUSPickGroup.Text + ".csv";
            if (File.Exists(filePath))
            {
                int currentTImportPerFile = 0;
                int currentTRemovedDuplicatesPerFile = 0;

                persons = new List<CPerson>();
                persons = Helper.GetPersonFromFile(filePath, persons, out currentTImportPerFile, out currentTRemovedDuplicatesPerFile);
                dgvUSUsers.DataSource = persons;
            }
        }

        /// <summary>
        /// RECHECK FUNCTION
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboUSPickGroup_TextChanged(object sender, EventArgs e)
        {
            if (Directory.GetFiles(AppConstants.GroupFolderName).Count(x => String.Equals(Path.GetFileNameWithoutExtension(x), cboUSPickGroup.Text, StringComparison.OrdinalIgnoreCase)) == 0)
            {
                persons = new List<CPerson>();
                dgvUSUsers.DataSource = persons;
            }
            else
            {
                cboUSPickGroup.SelectedItem = cboUSPickGroup.Text;
            }
        }

        /// <summary>
        /// Import CSV files to the Application and display in the DataGridView of Subscription Management
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUSImportUsers_Click(object sender, EventArgs e)
        {
            if (btnUSImportUsers.Text == "Import Users" && !String.IsNullOrEmpty(txtUSRename.Text))
            {
                DialogResult csvDialog = MetroMessageBox.Show(this, "To import contacts, your CSV file should be in the right format." +
                    "\nFormat should be Email Address, Name of the Person (Optional), Telephone (Optional)" +
                    "\nYou can import the test file, IMPORT.CSV to try before you import your file" + 
                    "\nClick OK to open the csv file format before you import." +
                    "\nCancel to continue importing.", "EMM - Before you Import CSV Files", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (csvDialog == DialogResult.OK)
                {
                    Process.Start(AppConstants.ImportDemoCSVFileLocation);
                }
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "CSV Files (.csv) | *.csv";
                dialog.Multiselect = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    int totalImported = 0;
                    int removedDuplicates = 0;

                    foreach (string filepath in dialog.FileNames)
                    {
                        int currentTImportPerFile = 0;
                        int currentTRemovedDuplicatesPerFile = 0;
                        persons = Helper.GetPersonFromFile(filepath, persons, out currentTImportPerFile, out currentTRemovedDuplicatesPerFile);
                        totalImported += currentTImportPerFile;
                        removedDuplicates += currentTRemovedDuplicatesPerFile;
                    }

                    lblAppStatus.Text = "STATUS: Users Imported: " + totalImported.ToString() + " Duplicates Removed: " + removedDuplicates.ToString();
                    dgvUSUsers.DataSource = persons;
                    btnUSImportUsers.Text = "Save Group";
                }
            }
            else if (btnUSImportUsers.Text == "Save Group")
            {
                if (!String.IsNullOrEmpty(txtUSRename.Text))
                {
                    string textContent = "";
                    string filePath = Path.Combine(AppConstants.GroupFolderName, txtUSRename.Text + ".csv");
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }

                    foreach (var item in persons)
                    {
                        textContent = textContent + item.EmailAddress + "," + item.ContactPerson + "," + item.Tel + "," + item.Status + "\r\n";
                    }

                    if (!Directory.Exists(AppConstants.GroupFolderName))
                    {
                        Helper.CreateDirectories();
                    }

                    File.WriteAllText(filePath, textContent);
                    PopulateUserGroups();
                    lblAppStatus.Text = "User group saved successfully";
                }
            }
            else
            {
                MetroMessageBox.Show(this, "Kindly provide the name of the User Group to import", "EMM User Group Name Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Rename a user group
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUSRenameGrp_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(cboUSPickGroup.Text) && (!String.IsNullOrEmpty(txtUSRename.Text)))
            {
                string oldFilePath = AppConstants.GroupFolderName + cboUSPickGroup.Text + ".csv";
                string newFilePath = AppConstants.GroupFolderName + txtUSRename.Text + ".csv";
                File.Move(oldFilePath, newFilePath);

                lblAppStatus.Text = "User group renamed successfully";
            }
            else
            {
                lblAppStatus.Text = "Pick a user group before renaming";
            }
        }

        /// <summary>
        /// UI Rename User Group
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtUSRename_TextChanged(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtUSRename.Text) && (!String.IsNullOrEmpty(cboUSPickGroup.Text)))
            {
                btnUSRenameGrp.Visible = true;
            }
            else
            {
                //MetroMessageBox.Show(this, "Either you have not selected the group\nor specified the new name for the group you are trying to rename", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// UI Load Unsubscribe Form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUSUnsubscribe_Click(object sender, EventArgs e)
        {
            frmUnsubscribe frm = new frmUnsubscribe();
            frm.ShowDialog();

        }

        /// <summary>
        /// UI HTML and Plain Text
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleHTMLPlain_CheckedChanged(object sender, EventArgs e)
        {
            checkHTMLorPlain();
        }

        /// <summary>
        /// Check if the Campaign is in HTML or Plain Text
        /// </summary>
        void checkHTMLorPlain()
        {
            if (toggleHTMLPlain.Checked)
            {
                lblHTMLPlain.Text = "HTML Content Enabled";
            }
            else
            {
                lblHTMLPlain.Text = "Plain Text Content";
            }
        }

        /// <summary>
        /// UI Populate Email Templates for use
        /// </summary>
        void PopulateEmailTemplates()
        {
            //List<string> emailTemplates = new List<string>();
            if (!Directory.Exists(AppConstants.CampaignFolderName))
            {
                Helper.CreateDirectories();
            }

            if (Directory.Exists(AppConstants.CampaignFolderName))
            {
                foreach (var file in Directory.GetFiles(AppConstants.CampaignFolderName))
                {
                    cboCMPickCampaign.Items.Add(Path.GetFileNameWithoutExtension(file));
                    //emailTemplates.Add(Path.GetFileNameWithoutExtension(file));
                }
            }
        }

        /// <summary>
        /// UI Save Email Template
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnETSaveTemplate_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtETCampaignName.Text) && (!String.IsNullOrEmpty(txtETSubject.Text) && (!String.IsNullOrEmpty(txtETEmailContent.Text))))
            {
                string textContent = txtETSubject.Text + "*()" + txtETEmailContent.Text + "*()" + toggleHTMLPlain.Checked.ToString();
                string message = Helper.CreateEmailTemplate(textContent, txtETCampaignName.Text, AppConstants.CampaignFolderName);
                lblAppStatus.Text = message;
                PopulateEmailTemplates();
            }
            else
            {
                MetroMessageBox.Show(this, "Kindly provide the details to save the email template", "EMM Email Template", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// UI Delete Email Template
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnETDeleteTemplate_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// UI Add User Groups to send an email campaign
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCMAddGroup_Click(object sender, EventArgs e)
        {
            string filepath = AppConstants.GroupFolderName + cboCMUserGroup.Text + ".csv";
            if (File.Exists(filepath))
            {
                int currentTotalImportPerFile = 0;
                int currentTotalDupPerFile = 0;
                var currentPersons = Helper.GetPersonFromFile(filepath, persons, out currentTotalImportPerFile, out currentTotalDupPerFile).Where(x => x.Status != EmailStatus.Unsubscribed).ToList();

                //Add Mail Send Pending Status
                currentPersons = currentPersons.Select(x => { x.Status = EmailStatus.Pending; return x; }).ToList();
                sendMailPersons.AddRange(currentPersons);
                var sendingList = new BindingList<CPerson>(sendMailPersons);
                dgvCMCampaign.DataSource = sendingList;
            }
        }

        /// <summary>
        /// Clear all added users from the email campaign
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCMClearCampaign_Click(object sender, EventArgs e)
        {
            //persons.Clear();
            //dgvCMCampaign.DataSource = null;
            dgvCMCampaign.Rows.Clear();
            dgvCMCampaign.Refresh();
        }

        /// <summary>
        /// Manage Controls when the worker runs (Disable when working, re-enable when cancelled or stopped)
        /// </summary>
        /// <param name="status"></param>
        void ManageControls(bool status)
        {
            txtCMAPIKey.Enabled = status;
            txtCMFooter.Enabled = status;
            txtCMfromEmail.Enabled = status;
            txtCMFromName.Enabled = status;
            cboCMPickCampaign.Enabled = status;
            cboCMUserGroup.Enabled = status;
            dgvCMCampaign.Enabled = status;
            btnCMAddGroup.Enabled = status;
            btnCMClearCampaign.Enabled = status;
            btnCMAttachUpload.Enabled = attachmentFileNames.Count > 0 ? status : false;
            btnCMAttachRemove.Enabled = status;
            lblCMTotalAttachments.Enabled = status;
            btnCMStartResume.Enabled = status;
            btnCMCancelAbort.Enabled = !status;     //Should be enabled to cancel the campaign send
        }

        /// <summary>
        /// Multi-threading worker completion
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ManageControls(true);
            if (e.Result != null && e.Result.ToString() == "-1")
            {
                lblAppStatus.Text = "Sending your campaign now.";
            }
            else if ((e.Cancelled == true))
            {
                lblAppStatus.Text = "User cancelled the campaign.";
            }

            else if (!(e.Error == null))
            {
                lblAppStatus.Text = "Possible issues in sending your email campaign,\nplease verify that the CSV import was as per the guidelines specified.";
            }
            else
            {
                lblAppStatus.Text = "Campaign Sent Successfully.";
            }
        }

        /// <summary>
        /// Multi-threading worker progress status update
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            sendMailPersons[e.ProgressPercentage].Status = (EmailStatus)e.UserState;
            lblAppStatus.Text = "Sending " + (e.ProgressPercentage + 1).ToString() + " out of " + (sendMailPersons.Count).ToString();
            dgvCMCampaign.Refresh();
        }

        /// <summary>
        /// Multi-threading worker process
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            var index = sendMailPersons.FindIndex(a => a.Status == EmailStatus.Pending);
            if (index == -1)
            {
                e.Result = "-1";
                return;
            }
            for (int i = index; i < sendMailPersons.Count; i++)
            {
                if ((bw.CancellationPending == true))
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    try
                    {
                        bw.ReportProgress(i, EmailStatus.Processing);
                        Helper.SendGridSendMail(sendMailPersons[i].ContactPerson, sendMailPersons[i].EmailAddress, txtCMFromName.Text,
                            txtCMfromEmail.Text, currentCampaign.Subject, currentCampaign.isHTML, currentCampaign.EmailContents,
                            attachmentFileNames, txtCMFooter.Text, txtCMAPIKey.Text);

                        System.Threading.Thread.Sleep(2000);
                        bw.ReportProgress(i, EmailStatus.Sent);
                    }
                    catch (Exception ex)
                    {
                        bw.ReportProgress(i, EmailStatus.Error);
                    }
                }
            }
        }

        private void btnCMStartResume_Click(object sender, EventArgs e)
        {
            if (File.Exists(Path.Combine(AppConstants.CampaignFolderName, cboCMPickCampaign.Text + ".txt")))
            {
                string textContent = File.ReadAllText(Path.Combine(AppConstants.CampaignFolderName, cboCMPickCampaign.Text + ".txt"));
                //You need an array of strings with the split option. You cannot pass a single string
                //Debug to check
                var values = textContent.Split(new string[] { "*()" }, StringSplitOptions.None).ToList();
                currentCampaign = new CCampaign
                {
                    Subject = values[0],
                    EmailContents = values[1],
                    isHTML = Convert.ToBoolean(values[2])
                };

                bw.RunWorkerAsync();
            }
        }

        private void tabMetro_Selected(object sender, TabControlEventArgs e)
        {
            ProcessTabs();
        }

        /// <summary>
        /// Go to the trackable advertising link
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lnkProductLink_Click(object sender, EventArgs e)
        {
            Process.Start(lnkProductLink.Tag + "");
        }

        /// <summary>
        /// Manage tabs with internet requirements
        /// </summary>
        private void ProcessTabs()
        {
            bool internet = NotificationDefaults.CheckInternet();
            if (tabMetro.SelectedTab == tbNotifications)
            {
                NotificationDefaults.SetTabPreferences(out notify, tbNotifications, tabMetro);
                lblProductName.Text = notify.ProductName;
                lblProductDesc.Text = notify.ProductDescription;
                lnkProductLink.Text = notify.URLText;
                lnkProductLink.Tag = notify.URLTagClick;
                pbProductImage.SizeMode = PictureBoxSizeMode.StretchImage;
                pbProductImage.ImageLocation = notify.ImageLocation;
            }
            else if (tabMetro.SelectedTab == tbHelp)
            {
                if (internet == true)
                {
                    webHelp.Navigate(NotificationDefaults.SupportURL);
                }
                else
                {
                    MetroMessageBox.Show(this, NotificationDefaults.NoInternetNotification, "EMM Internet Unavailable", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void webHelp_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            lblAppStatus.Text = "Connecting to support website, please wait..";
        }

        private void webHelp_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            lblAppStatus.Text = "Support site connected";
        }

        private void btnCMNewCampaign_Click(object sender, EventArgs e)
        {
            dgvCMCampaign.Rows.Clear();
            cboCMPickCampaign.Text = string.Empty;
            txtCMFooter.Text = string.Empty;
            txtCMfromEmail.Text = string.Empty;
            txtCMFromName.Text = string.Empty;
        }
    }
}
