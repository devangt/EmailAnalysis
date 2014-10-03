using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace EmailAnalysis
{
    public class EmailAnalyser
    {
        private const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        public List<Dictionary<string, object>> GetEmailData()
        {
            var application = GetApplicationObject();

            Microsoft.Office.Interop.Outlook.ExchangeUser currentUser = application.Session.CurrentUser.AddressEntry.GetExchangeUser();

            var currentUserEmail = currentUser.PrimarySmtpAddress;

            var @namespace = application.GetNamespace("MAPI");

            var inbox = @namespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

            var explorers = application.Explorers;

            var explorer = explorers.Add(inbox, Microsoft.Office.Interop.Outlook.OlFolderDisplayMode.olFolderDisplayNormal);

            explorer.Activate();

            var emailData = new List<Dictionary<string, object>>();

            AnalyseEmails(currentUserEmail, emailData, inbox);

            return emailData;
        }

        private void AnalyseEmails(string currentUserEmail, List<Dictionary<string, object>> emailData, Microsoft.Office.Interop.Outlook.MAPIFolder folder)
        {
            var testFolderName = ConfigurationManager.AppSettings["TestFolderName"];

            var folderName = folder.Name;

            Console.WriteLine(folderName);

            foreach (dynamic item in folder.Items)
            {
                Microsoft.Office.Interop.Outlook.MailItem mailItem;

                try
                {
                    mailItem = item;
                }
                catch
                {
                    continue;
                }

                try
                {
                    var subject = mailItem.Subject;

                    Console.WriteLine("{0} - {1}", folderName, subject);

                    var data = new Dictionary<string, object>();

                    // test
                    data.Add("TestFolder", folderName == testFolderName);

                    // attachments
                    data.Add("HasAttachments", mailItem.Attachments.Count > 0);

                    // Direct

                    var recipients = GetEmails(mailItem.Recipients);
                    var sentDirect = recipients.Equals(currentUserEmail, StringComparison.OrdinalIgnoreCase);
                    data.Add("SentDirect", sentDirect);

                    // contains numbers / a time
                    data.Add("MayContainATime", MayContainATime(subject));

                    // Replies, forwards
                    data.Add("IsREorFW", IsREorFW(subject));

                    // time of day
                    var recieved = mailItem.ReceivedTime;
                    data.Add("ReceivedDayOfWeek", recieved.DayOfWeek);
                    data.Add("ReceivedHour", recieved.Hour);

                    // Subject word count
                    data.Add("SubjectWordCount", WordCount(subject));
                    // Body word count
                    data.Add("BodyWordCount", WordCount(mailItem.Body));

                    string senderEmail = mailItem.SenderEmailAddress;

                    // sender domain
                    string domain = EmailDomain(senderEmail);
                    data.Add("SenderDomain", domain);

                    // has cc
                    data.Add("HasCC", !String.IsNullOrWhiteSpace(mailItem.CC));

                    // Importance
                    data.Add("Importance", (int)mailItem.Importance);

                    // Format
                    data.Add("BodyFormat", (int)mailItem.BodyFormat);

                    data.Add("SpecialCharacterCount", SpecialCharacterCount(subject));
                    data.Add("NumberCount", NumberCount(subject));

                    // folderName, recieved and subject - for debugging purposes
                    data.Add("Recieved", recieved);
                    data.Add("Subject", mailItem.Subject);
                    data.Add("FolderName", folderName);

                    emailData.Add(data);
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.ToString());
                }
            }

            foreach (Microsoft.Office.Interop.Outlook.MAPIFolder subFolder in folder.Folders)
            {
                AnalyseEmails(currentUserEmail, emailData, subFolder);
            }
        }

        private int SpecialCharacterCount(string subject)
        {
            if (String.IsNullOrWhiteSpace(subject))
            {
                return 0;
            }

            MatchCollection collection = Regex.Matches(subject, @"[^0-9A-Za-z_ :-]");
            return collection.Count;
        }

        private int NumberCount(string subject)
        {
            if (String.IsNullOrWhiteSpace(subject))
            {
                return 0;
            }

            MatchCollection collection = Regex.Matches(subject, @"[0-9]");
            return collection.Count;
        }

        private string GetEmails(dynamic addresses)
        {
            var emails = new List<string>();
            foreach (dynamic address in addresses)
            {
                var email = Email(address);
                emails.Add(email);
            }

            return String.Join(",", emails);
        }

        private dynamic Email(dynamic address)
        {
            var email = address.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
            return email;
        }

        private bool MayContainATime(string subject)
        {
            if (String.IsNullOrWhiteSpace(subject))
            {
                return false;
            }

            var mayContainATimeFormat1 = Regex.IsMatch(subject, @"\b(?:0?[0-9]|1[0-9]|2[0-3]):[0-5][0-9]\b");

            var mayContainATimeFormat2 = Regex.IsMatch(subject, @"\b(0?[1-9]|1[012])(:[0-5]\d)? ?[APap][mM]\b", RegexOptions.IgnoreCase);

            return mayContainATimeFormat1 || mayContainATimeFormat2;
        }

        private bool IsREorFW(string subject)
        {
            if (String.IsNullOrWhiteSpace(subject))
            {
                return false;
            }

            var isREorFW = Regex.IsMatch(subject, @"(\b(re)\b:|\b(fd)\b:|\b(fwd)\b:)", RegexOptions.IgnoreCase);

            return isREorFW;
        }

        private string EmailDomain(string senderEmailAddress)
        {
            senderEmailAddress = senderEmailAddress ?? String.Empty;
            var @index = senderEmailAddress.IndexOf('@');
            return index == -1 ? String.Empty : senderEmailAddress.Substring(index + 1).ToLowerInvariant();
        }

        private int WordCount(string s)
        {
            if (String.IsNullOrWhiteSpace(s))
            {
                return 0;
            }

            MatchCollection collection = Regex.Matches(s, @"[\S]+");
            return collection.Count;
        }

        private Microsoft.Office.Interop.Outlook.Application GetApplicationObject()
        {
            // http://msdn.microsoft.com/en-us/library/office/ff462097(v=office.15).aspx
            Microsoft.Office.Interop.Outlook.Application application = null;

            // Check whether there is an Outlook process running.
            if (Process.GetProcessesByName("OUTLOOK").Any())
            {
                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                application = Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
            }
            else
            {
                // If not, create a new instance of Outlook and log on to the default profile.
                application = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon(String.Empty, String.Empty, Missing.Value, Missing.Value);
                nameSpace = null;
            }

            // Return the Outlook Application object.
            return application;
        }

        public static void OutputDataset(List<Dictionary<string, object>> emailData, string filename = "EmailDataset.csv")
        {
            var sb = new StringBuilder();

            sb.AppendLine(String.Join(",", emailData[0].Keys));

            foreach (var data in emailData)
            {
                sb.AppendLine(String.Join(",", data.Values.Select(ToCsvString)));
            }

            File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), filename),
                sb.ToString());
        }

        private static string ToCsvString(object item)
        {
            if (item == null)
            {
                return String.Empty;
            }

            return EscapeCsv(item.ToString());
        }

        private static string EscapeCsv(string s)
        {
            if (s == null)
            {
                return null;
            }

            var escapedValue = new StringBuilder();
            bool applyQuotes = false;

            foreach (var c in s)
            {
                switch (c)
                {
                    case '\n':
                        escapedValue.Append("\\n");
                        applyQuotes = true;
                        break;
                    case '\t':
                        escapedValue.Append("\\t");
                        applyQuotes = true;
                        break;
                    case ',':
                        escapedValue.Append(",");
                        applyQuotes = true;
                        break;
                    case ' ':
                        escapedValue.Append(" ");
                        applyQuotes = true;
                        break;
                    case '"':
                        escapedValue.Append("\\\"");
                        applyQuotes = true;
                        break;
                    default: escapedValue.Append(c);
                        break;
                }
            }

            if (applyQuotes)
            {
                return @"""" + escapedValue + @"""";
            }

            return escapedValue.ToString();
        }
    }
}