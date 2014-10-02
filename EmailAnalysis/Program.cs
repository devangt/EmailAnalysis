using System;
using System.Configuration;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailAnalysis
{
    class Program
    {
        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        static void Main(string[] args)
        {
            var application = GetApplicationObject();

            Outlook.ExchangeUser currentUser = application.Session.CurrentUser.AddressEntry.GetExchangeUser();

            var currentUserEmail = currentUser.PrimarySmtpAddress;

            var @namespace = application.GetNamespace("MAPI");

            var inbox = @namespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            var explorers = application.Explorers;

            var explorer = explorers.Add(inbox, Outlook.OlFolderDisplayMode.olFolderDisplayNormal);

            explorer.Activate();

            var emailData = new List<Dictionary<string, object>>();

            AnalyseEmails(currentUserEmail, emailData, inbox);

            OutputDataset(emailData);

            Console.WriteLine("DONE");
            Console.ReadLine();
        }

        private static void OutputDataset(List<Dictionary<string, object>> emailData)
        {
            var sb = new StringBuilder();

            sb.AppendLine(string.Join(",", emailData[0].Keys));

            foreach (var data in emailData)
            {
                sb.AppendLine(string.Join(",", data.Values.Select(ToCsvString)));
            }

            File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "EmailDataset.csv"),
                sb.ToString());
        }

        private static string ToCsvString(object item)
        {
            if (item == null)
            {
                return string.Empty;
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

        private static void AnalyseEmails(string currentUserEmail, List<Dictionary<string, object>> emailData, Outlook.MAPIFolder folder)
        {
            var testFolderName = ConfigurationManager.AppSettings["TestFolderName"];

            var folderName = folder.Name;

            Console.WriteLine(folderName);

            foreach (dynamic item in folder.Items)
            {
                Outlook.MailItem mailItem;

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
                    data.Add("HasCC", !string.IsNullOrWhiteSpace(mailItem.CC));

                    // Importance
                    data.Add("Importance", (int)mailItem.Importance);

                    // Format
                    data.Add("BodyFormat", (int)mailItem.BodyFormat);

                    data.Add("SpecialCharacterCount", SpecialCharacterCount(subject));
                    data.Add("NumberCount", NumberCount(subject));

                    // subject - for debugging purposes
                    data.Add("Subject", mailItem.Subject);

                    emailData.Add(data);
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.ToString());
                }
            }

            foreach (Outlook.MAPIFolder subFolder in folder.Folders)
            {
                AnalyseEmails(currentUserEmail, emailData, subFolder);
            }
        }

        private static int SpecialCharacterCount(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
            {
                return 0;
            }

            MatchCollection collection = Regex.Matches(subject, @"[^0-9A-Za-z_ :-]");
            return collection.Count;
        }

        private static int NumberCount(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
            {
                return 0;
            }

            MatchCollection collection = Regex.Matches(subject, @"[0-9]");
            return collection.Count;
        }

        private static string GetEmails(dynamic addresses)
        {
            var emails = new List<string>();
            foreach (dynamic address in addresses)
            {
                var email = Email(address);
                emails.Add(email);
            }

            return string.Join(",", emails);
        }

        private static dynamic Email(dynamic address)
        {
            var email = address.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
            return email;
        }

        private static bool MayContainATime(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
            {
                return false;
            }

            var mayContainATimeFormat1 = Regex.IsMatch(subject, @"\b(?:0?[0-9]|1[0-9]|2[0-3]):[0-5][0-9]\b");

            var mayContainATimeFormat2 = Regex.IsMatch(subject, @"\b(0?[1-9]|1[012])(:[0-5]\d)? ?[APap][mM]\b", RegexOptions.IgnoreCase);

            return mayContainATimeFormat1 || mayContainATimeFormat2;
        }

        private static bool IsREorFW(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
            {
                return false;
            }

            var isREorFW = Regex.IsMatch(subject, @"(\b(re)\b:|\b(fd)\b:|\b(fwd)\b:)", RegexOptions.IgnoreCase);

            return isREorFW;
        }

        private static string EmailDomain(string senderEmailAddress)
        {
            senderEmailAddress = senderEmailAddress ?? string.Empty;
            var @index = senderEmailAddress.IndexOf('@');
            return @index == -1 ? string.Empty : senderEmailAddress.Substring(@index + 1).ToLowerInvariant();
        }

        private static int WordCount(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                return 0;
            }

            MatchCollection collection = Regex.Matches(s, @"[\S]+");
            return collection.Count;
        }

        private static Outlook.Application GetApplicationObject()
        {
            // http://msdn.microsoft.com/en-us/library/office/ff462097(v=office.15).aspx
            Outlook.Application application = null;

            // Check whether there is an Outlook process running.
            if (Process.GetProcessesByName("OUTLOOK").Any())
            {
                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            else
            {
                // If not, create a new instance of Outlook and log on to the default profile.
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon(string.Empty, string.Empty, Missing.Value, Missing.Value);
                nameSpace = null;
            }

            // Return the Outlook Application object.
            return application;
        }
    }
}
