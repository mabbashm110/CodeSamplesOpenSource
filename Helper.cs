using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SendGrid;
using System.Net.Mail;
using System.IO;

namespace EmailMarketing.Library
{
    public static class Helper
    {
        /// <summary>
        /// Create Base Directories for Campaigns, UserGroups and Reports
        /// </summary>
        public static void CreateDirectories()
        {
            if (!Directory.Exists(AppConstants.CampaignFolderName))
            {
                Directory.CreateDirectory(AppConstants.CampaignFolderName);
            }

            if (!Directory.Exists(AppConstants.GroupFolderName))
            {
                Directory.CreateDirectory(AppConstants.GroupFolderName);
            }

            if (!Directory.Exists(AppConstants.ReportsFolderName))
            {
                Directory.CreateDirectory(AppConstants.ReportsFolderName);
            }
        }

        /// <summary>
        /// Import CSV File and Create an internally readable CSV Group
        /// </summary>
        /// <param name="fileName">Location of the CSV file</param>
        /// <param name="persons">List of people being imported</param>
        /// <param name="totalImported">Number of records imported</param>
        /// <param name="removedDuplicates">Number of duplicate emails automatically removed</param>
        /// <returns></returns>
        public static List<CPerson> GetPersonFromFile(string fileName, List<CPerson> persons, out int totalImported, out int removedDuplicates)
        {
            totalImported = 0;
            removedDuplicates = 0;
            //statusMessage = "";

            using (StreamReader sr = new StreamReader(fileName))
            {
                string currentLine = string.Empty;

                while ((currentLine = sr.ReadLine()) != null)
                {
                    var values = currentLine.Split(',');

                    //Check duplicate emails here
                    if (!persons.Any(x => string.Equals(x.EmailAddress, values[0], StringComparison.OrdinalIgnoreCase)))
                    {
                        persons.Add(new CPerson
                        {
                            EmailAddress = values[0],
                            ContactPerson = values.Count() > 1 ? values[1] : string.Empty,
                            Tel = values.Count() > 2 ? values[2] : string.Empty,
                            Status = values.Count() > 3 ? (EmailStatus)Enum.Parse(typeof(EmailStatus), values[3]) : EmailStatus.Subscribed,
                            //SubscriptionDate = values.Count() > 4 ? Convert.ToDateTime(values[4]) : DateTime.Now
                        });
                        totalImported++;
                    }
                    else
                    {
                        removedDuplicates++;
                    }
                }
                //statusMessage = "Import successful";
                return persons;
            }
        }

        /// <summary>
        /// Unsubscribe Users from all groups
        /// </summary>
        /// <param name="folderName">Location of the User Group Folders</param>
        /// <param name="emailAddresses">Array of email addresses to be unsubscribed</param>
        /// <param name="totalUnsubscribed">Number of users unsubscribed</param>
        /// <param name="message">Status result of the unsubscribe function</param>
        public static void UnsubscribeFromAllGroups(string folderName, List<string> emailAddresses, out int totalUnsubscribed)
        {
            totalUnsubscribed = 0;
            //message = "";

            try
            {
                int filesCount = Directory.GetFiles(folderName).Length;
                string[] filesinFolder = Directory.GetFiles(folderName);

                if (filesCount > 0 && filesinFolder != null)
                {
                    for (int i = 0; i < filesCount; i++)
                    {
                        string currentFile = filesinFolder[i];
                        var persons = File.ReadLines(currentFile).Select(ParsePersonEmailFromFile);
                        string textContent = string.Empty;

                        foreach (var person in persons)
                        {
                            if (emailAddresses.Any(x => x == person.EmailAddress))
                            {
                                textContent += person.EmailAddress + "," + person.ContactPerson + "," + person.Tel + "," + EmailStatus.Unsubscribed + "," + person.SubscriptionDate.ToString() + "\r\n";
                                totalUnsubscribed++;
                                //message = "The user(s) were successfully unsubscribed.";
                            }
                            else
                            {
                                textContent += person.EmailAddress + "," + person.ContactPerson + "," + person.Tel + "," + person.Status + "," + person.SubscriptionDate.ToString() + "\r\n";
                            }
                        }
                        File.WriteAllText(currentFile, textContent);
                    }
                }
            }
            catch (Exception)
            {
                //message = "There was an error in unsubscribing the user." + 
            }
        }

        /// <summary>
        /// Unsubscribe from a single file
        /// </summary>
        /// <param name="folderName">Location of the User Group Folders</param>
        /// <param name="emailAddresses">Array of email addresses to be unsubscribed</param>
        /// <param name="totalUnsubscribed">Number of users unsubscribed</param>
        /// <param name="message">Status result of the unsubscribe function</param>
        public static void UnsubscribeFromSingleGroup(string fileName, List<string> emailAddresses, out int totalUnsubscribed)
        {
            totalUnsubscribed = 0;
            //message = "";
            try
            {
                var persons = File.ReadLines(fileName).Select(ParsePersonEmailFromFile);
                string textContent = string.Empty;

                foreach (var person in persons)
                {
                    if (emailAddresses.Any(x => x == person.EmailAddress))
                    {
                        textContent += person.EmailAddress + "," + person.ContactPerson + "," + person.Tel + "," + EmailStatus.Unsubscribed + "," + person.SubscriptionDate.ToString() + "\r\n";
                        totalUnsubscribed++;
                        //message = "The user(s) were successfully unsubscribed.";
                    }
                    else
                    {
                        textContent += person.EmailAddress + "," + person.ContactPerson + "," + person.Tel + "," + person.Status + "," + person.SubscriptionDate.ToString() + "\r\n";
                    }
                }
                File.WriteAllText(fileName, textContent);
            }
            catch (Exception)
            {
                //message = "There was an error in unsubscribing the user." + 
            }
        }

        /// <summary>
        /// Function to store objects as Persons and retreive information as Persons object
        /// </summary>
        /// <param name="line">Reading each line in a csv for a match</param>
        /// <returns></returns>
        private static CPerson ParsePersonEmailFromFile(string line)
        {
            string[] values = line.Split(',');
            return new CPerson
            {
                EmailAddress = values[0],
                ContactPerson = values[1],
                Status = (EmailStatus)Enum.Parse(typeof(EmailStatus), values[3]),
                //SubscriptionDate = values.Count() > 4 ? Convert.ToDateTime(values[4]) : (DateTime?)null
            };
        }

        /// <summary>
        /// Creates an Email Template
        /// </summary>
        /// <param name="textContent">Template Content with HTML or Plain Text</param>
        /// <param name="campaignName">Campaign Name</param>
        /// <param name="campaignFolderName">Campaign Folder Location</param>
        /// <returns>Message upon successful completion</returns>
        public static string CreateEmailTemplate(string textContent, string campaignName, string campaignFolderName)
        {
            if (!Directory.Exists(campaignFolderName))
            {
                CreateDirectories();
            }

            File.WriteAllText(Path.Combine(campaignFolderName, campaignName + ".txt"), textContent);
            string message = "STATUS: Email Template created successfully.";
            return message;
        }

        /// <summary>
        /// Sending email through SendGrid System (Can be eventually moved out to a different class)
        /// </summary>
        /// <param name="recipientName">Customer Name</param>
        /// <param name="recipientEmail">Customer Email</param>
        /// <param name="fromName">Senders Name</param>
        /// <param name="fromEmail">Senders Email</param>
        /// <param name="emailSubject">Email Subject</param>
        /// <param name="HTMLorPlain">Email in Plain or HTML context</param>
        /// <param name="emailContents">Email Contents</param>
        /// <param name="attachmentFileNames">Attachments to the email</param>
        /// <param name="footer">Email Footer</param>
        /// <param name="apiKey">SendGrid's API Key</param>
        public static void SendGridSendMail(string recipientName, string recipientEmail, string fromName, string fromEmail, string emailSubject, bool HTMLorPlain, string emailContents, List<string> attachmentFileNames, string footer, string apiKey)
        {
            SendGridMessage sendGridMessage = new SendGridMessage();
            sendGridMessage.From = new MailAddress(fromEmail, fromName);

            List<MailAddress> toAddresses = new List<MailAddress>();
            if (!string.IsNullOrEmpty(recipientName))
            {
                toAddresses.Add(new MailAddress(recipientEmail, recipientName));
            }
            else
            {
                toAddresses.Add(new MailAddress(recipientEmail, recipientName));
            }
            sendGridMessage.To = toAddresses.ToArray();

            sendGridMessage.Subject = emailSubject;
            if (HTMLorPlain == true)
            {
                sendGridMessage.Html = emailContents;
            }
            else
            {
                sendGridMessage.Text = emailContents;
            }

            if (attachmentFileNames.Count > 0)
            {
                foreach (string file in attachmentFileNames)
                {
                    Stream stream = new FileStream(file, FileMode.Open);
                    sendGridMessage.AddAttachment(stream, Path.GetFileName(file));
                    stream.Close();
                }
            }

            if (!string.IsNullOrEmpty(footer))
            {
                sendGridMessage.EnableFooter(footer);
            }

            sendGridMessage.EnableClickTracking(true);

            Web web = new Web(apiKey);
            web.DeliverAsync(sendGridMessage);

        }
    }
}
