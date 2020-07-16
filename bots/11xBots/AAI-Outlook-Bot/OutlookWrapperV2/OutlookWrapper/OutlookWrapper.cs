/*
Copyright 2020 Automation Anywhere, Inc.
This software is licensed under Automation Anywhere, Inc. Free Source Code License.

This Automation Anywhere Bot is issued under the AAI Free Source Code License.
You may obtain a copy of License at 
https://github.com/AutomationAnywhere/AAI-Botstore-Open-Source-Bots/blob/master/LICENSE
 */

using System;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
[assembly: DefaultDllImportSearchPaths(DllImportSearchPath.System32)]

namespace OutlookWrapper
{
    public class OutlookWrapper
    {
        private Outlook.Application application;
        private Outlook.NameSpace nameSpace;
        private Outlook.Account account;
        private Outlook.MAPIFolder folder;
        private Outlook.MailItem mail;
        private Outlook.AppointmentItem appointment;
        private Outlook.MeetingItem meeting;
        private Outlook.OlDefaultFolders defaultFolderSaved;
        private string folderPathSaved = string.Empty;
        private string profileNameSaved = string.Empty;
        private string profilePasswordSaved = string.Empty;
        private string accountNameSaved = string.Empty;
        private bool isOutlookVisible = false;

        private const string RETURNVALUE = "Success";
        private const string ERROR = "[ERROR]:";


        private string PR_SMTP_ADDRESS;
        public void SetPR_SMTP_ADDRESS(string pr_smtp_address)
        {
            PR_SMTP_ADDRESS = pr_smtp_address;
        }


        private string PR_SENT_REPRESENTING_ENTRYID;
        public void SetPR_SENT_REPRESENTING_ENTRYID(string pr_sent_smtp_address)
        {
            PR_SENT_REPRESENTING_ENTRYID = pr_sent_smtp_address;

        }

        private const int SW_MAXIMIZE = 3;
        private const int SW_MINIMIZE = 6;
        private enum REPLYTYPE
        {
              Reply
            , ReplyAll
        }
        private enum COLLECTIONTYPE
        {
              ToRecipients
            , CcRecipients
            , BccRecipients
            , ReplyRecipients
            , Attachments
        }

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


        #region Outlook Actions

        public string LaunchOutlook()
        {
            try
            {
                // Check whether there is an Outlook process running.
                if (Process.GetProcessesByName("OUTLOOK").Count() == 0)
                {
                    // If not, create a new instance of Outlook and log on to the default profile.
                    var outlookApp = new Process();
                    outlookApp.StartInfo = new ProcessStartInfo("OUTLOOK.EXE");
                    //  outlookApp.StartInfo.Verb = "runas";    // Only required if Outlook is to be lauched with Administrator rights. 
                    outlookApp.Start();
                    // Following steps are needed to ensure that the Outlook Application instance is added
                    // to Running Object Table (ROT), so that "Marshal.GetActiveObject" won't fail. 
                    Thread.Sleep(20000);
                    outlookApp.WaitForInputIdle();
                    ShowWindow(outlookApp.MainWindowHandle, SW_MINIMIZE);       // Minimize Window to force addition to ROT
                    Thread.Sleep(3000);
                    ShowWindow(outlookApp.MainWindowHandle, SW_MAXIMIZE);       // Restore Window
                    Thread.Sleep(2000);

                    isOutlookVisible = true;
                }

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                ReleaseComObject(application);
                ReleaseComObject(nameSpace);
                ReleaseComObject(account);
                ReleaseComObject(folder);
                ReleaseComObject(mail);

                return ERROR + ex.ToString();
            }
        }

        public string CloseOutlook()
        {
            try
            {
                application.Quit();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
            finally
            {
                ReleaseComObject(application);
                ReleaseComObject(nameSpace);
                ReleaseComObject(account);
                ReleaseComObject(folder);
                ReleaseComObject(mail);
            }

            return RETURNVALUE;
        }

        private void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }

        private void Reconnect()
        {
            // This Sleep/Delay is necessary for objects/processes to end completely before reconnecting.
            Thread.Sleep(5000);

            // NOTE: ORDER/SEQUNCE OF FOLLOWING COMMANDS MATTER. KEEP AS IS. 
            if (isOutlookVisible)
                LaunchOutlook();

            // Profile
            SelectProfile(profileNameSaved, profilePasswordSaved);

            // Account
            if (!string.IsNullOrWhiteSpace(accountNameSaved))
                SelectAccount(accountNameSaved);

            // Folder
            if (Enum.IsDefined(typeof(Outlook.OlDefaultFolders), defaultFolderSaved))
                SelectDefaultFolder(defaultFolderSaved);
            else if (!string.IsNullOrWhiteSpace(folderPathSaved))
                SelectFolderByPath(folderPathSaved);
            else
                SelectInbox();
        }

        public string SelectProfile()
        {
            return SelectProfile(string.Empty, string.Empty);   // Login with Default Profile or Existing Session
        }

        private string SelectProfile(string ProfileName, string ProfilePassword)
        {
            try
            {
                // Check whether there is an Outlook process running.
                if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
                {
                    // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                    application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                    nameSpace = application.GetNamespace("MAPI");
                    isOutlookVisible = true;
                }
                else
                {
                    // If not, create a new instance of Outlook and log on to the provided profile.
                    application = new Outlook.Application();
                    nameSpace = application.GetNamespace("MAPI");
                    isOutlookVisible = false;
                }

                nameSpace.Logon(ProfileName, ProfilePassword, false, true);

                profileNameSaved = ProfileName;
                profilePasswordSaved = ProfilePassword;
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                ReleaseComObject(application);
                ReleaseComObject(nameSpace);
                ReleaseComObject(account);
                ReleaseComObject(folder);
                ReleaseComObject(mail);

                return ERROR + ex.ToString();
            }
        }

        public string SelectAccount()
        {
            try
            {
                account = application.Session.Accounts[1];

                accountNameSaved = account.DisplayName;
                return RETURNVALUE;
                
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SelectAccount(string AccountName)
        {
            try
            {
                ReleaseComObject(account);
                Console.WriteLine("Acc Object :" + account);
                foreach (Outlook.Account acct in application.Session.Accounts)
                {
                    Console.WriteLine(acct.AccountType + "\t" + acct.DisplayName + "");
                    //Console.WriteLine(acc.ToString());

                    if (acct.DisplayName.Trim().Equals(AccountName))
                    {
                        ReleaseComObject(account);
                        account = acct;
                        Console.WriteLine( " Inside If :\t" + account.DisplayName + "");
                        break;
                    }
                }
                if (account == null) return ERROR + "Invalid account name!";

                accountNameSaved = account.DisplayName;
                Console.WriteLine(accountNameSaved);
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SelectInbox()
        {
            try
            {
                return SelectDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SelectCalendar()
        {
            try
            {
                var result = SelectDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                if (result != RETURNVALUE)
                    return result; 
                
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        private string SelectDefaultFolder(Outlook.OlDefaultFolders DefaultFolder)
        {
            try
            {
                folder = account.Session.GetDefaultFolder(DefaultFolder);
                Console.WriteLine("\nDescription : " + folder.Description + "\nStoreId : " + folder.StoreID + " \n Folder :" + folder.ToString());
                defaultFolderSaved = DefaultFolder;     // Only one of these two should be set at a time 
                folderPathSaved = string.Empty;         // for Reconnect() to work correctly
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SelectFolderByPath(string FolderPath)
        {
            Outlook.Folder fldr;
            string backslash = @"\";
            try
            {
                if (FolderPath.StartsWith(@"\\"))
                {
                    FolderPath = FolderPath.Remove(0, 2);
                }
                String[] folders = FolderPath.Split(backslash.ToCharArray());
                fldr = application.Session.Folders[folders[0]] as Outlook.Folder;
                if (fldr == null)
                    return ERROR + "Folder path not found!";

                for (int i = 1; i <= folders.GetUpperBound(0); i++)
                {
                    Outlook.Folders subFolders = fldr.Folders;
                    fldr = subFolders[folders[i]] as Outlook.Folder;
                    if (fldr == null)
                        return ERROR + "Folder path not found!";
                }

                folder = fldr;
                folderPathSaved = FolderPath;           // Only one of these two should be set at a time 
                defaultFolderSaved = 0;                 // for Reconnect() to work correctly

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SelectMailItem(string ID)
        {
            try
            {
                mail = (Outlook.MailItem)(nameSpace).GetItemFromID(ID);
                if (mail == null)
                    return ERROR + "Invalid Mail ID.";

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SelectAppointment(string ID)
        {
            try
            {
                var index = ID.IndexOf("||");

                if (index < 0)
                {
                    appointment = (Outlook.AppointmentItem)(nameSpace).GetItemFromID(ID);
                }
                else
                {
                    var startDate = ID.Substring(index + 2);
                    ID = ID.Remove(index);

                    appointment = (Outlook.AppointmentItem)(nameSpace).GetItemFromID(ID);
                    var recurPattern = appointment.GetRecurrencePattern();

                    try
                    {
                        var specificAppt = recurPattern.GetOccurrence(DateTime.Parse(startDate));
                        appointment = specificAppt;
                    }
                    catch (Exception ex)
                    {
                        // Move on. 
                    }

                    ReleaseComObject(recurPattern);
                }

                if (appointment == null)
                    return ERROR + "Invalid Appointment ID.";
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SelectMeetingRequest(string ID)
        {
            try
            {
                meeting = (Outlook.MeetingItem)(nameSpace).GetItemFromID(ID);
                if (meeting == null)
                    return ERROR + "Invalid Meeting Request ID.";
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetAccountInformation()
        {
            try
            {
                // The Namespace Object (Session) has a collection of accounts.
                Outlook.Accounts accounts = application.Session.Accounts;

                // Concatenate a message with information about all accounts.
                StringBuilder builder = new StringBuilder();

                // Loop over all accounts and print detail account information.
                // All properties of the Account object are read-only.
                foreach (Outlook.Account account in accounts)
                {

                    // The DisplayName property represents the friendly name of the account.
                    builder.AppendFormat("DisplayName: {0}\n", account.DisplayName);

                    // The UserName property provides an account-based context to determine identity.
                    builder.AppendFormat("UserName: {0}\n", account.UserName);

                    // The SmtpAddress property provides the SMTP address for the account.
                    builder.AppendFormat("SmtpAddress: {0}\n", account.SmtpAddress);

                    // The AccountType property indicates the type of the account.
                    builder.Append("AccountType: ");
                    switch (account.AccountType)
                    {

                        case Outlook.OlAccountType.olExchange:
                            builder.AppendLine("Exchange");
                            break;

                        case Outlook.OlAccountType.olHttp:
                            builder.AppendLine("Http");
                            break;

                        case Outlook.OlAccountType.olImap:
                            builder.AppendLine("Imap");
                            break;

                        case Outlook.OlAccountType.olOtherAccount:
                            builder.AppendLine("Other");
                            break;

                        case Outlook.OlAccountType.olPop3:
                            builder.AppendLine("Pop3");
                            break;
                    }

                    builder.AppendLine();
                }

                // Display the account information.
                return builder.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string[] GetFoldersArray()
        {
            List<string> Folders = new List<string>();

            try
            {
                EnumerateFolders(application.Session.DefaultStore.GetRootFolder() as Outlook.Folder, Folders);
                return Folders.ToArray();
            }
            catch (Exception ex)
            {
                Folders.Add(ERROR + ex.ToString());
                return Folders.ToArray();
            }
        }

        private void EnumerateFolders(Outlook.Folder folder, List<string> list)
        {
            try
            {
                Outlook.Folders childFolders = folder.Folders;
                if (childFolders.Count > 0)
                {
                    foreach (Outlook.Folder childFolder in childFolders)
                    {
                        // Write the folder path.
                        list.Add(childFolder.FolderPath);
                        // Call EnumerateFolders using childFolder.
                        EnumerateFolders(childFolder, list);
                    }
                }
            }
            catch (Exception ex)
            {
                list.Add(ex.Message);
            }
        }

        public string[] GetSharedFoldersArray()
        {
            List<string> Folders = new List<string>();

            try
            {
                Outlook.Stores stores = application.Session.Stores;

                for (int i = 1; i <= stores.Count; i++)
                {
                    Outlook.Store store = stores[i];
                    if (store.ExchangeStoreType == Microsoft.Office.Interop.Outlook.OlExchangeStoreType.olExchangeMailbox ||
                        store.ExchangeStoreType == Microsoft.Office.Interop.Outlook.OlExchangeStoreType.olAdditionalExchangeMailbox)
                    {
                        try
                        {
                            var rootFolder = store.GetRootFolder() as Outlook.Folder;
                            if (rootFolder != null)
                                EnumerateFolders(rootFolder, Folders);
                        }
                        catch   // Mostly due to lack of Permissions to read the folder(s). 
                        {
                            continue;
                        }
                    }
                }

                return Folders.ToArray();
            }
            catch (Exception ex)
            {
                Folders.Add(ERROR + ex.ToString());
                return Folders.ToArray();
            }
        }

        #endregion

        #region Get Mails Collection/Array 

        public string[] GetAllMailIDsArray(string NumberOfItems)
        {
            return PrepareGetMailIDsArrayResponse(NumberOfItems, string.Empty);
        }

        public string[] GetUnReadMailIDsArray(string NumberOfItems)
        {
            return PrepareGetMailIDsArrayResponse(NumberOfItems, "[Unread]=true");
        }

        public string[] GetReadMailIDsArray(string NumberOfItems)
        {
            return PrepareGetMailIDsArrayResponse(NumberOfItems, "[Unread]=false");
        }

        public string[] GetMailIDsArrayByFilter(string Filter)
        {
            return PrepareGetMailIDsArrayResponse(Int32.MaxValue.ToString(), Filter);
        }

        public string[] GetMailIDsArrayByDateRange(string StartDate, string EndDate)
        {
            List<string> IDs = new List<string>();

            try
            {
                var parsedStartDate = DateTime.Parse(StartDate);
                var parsedEndDate = DateTime.Parse(EndDate);
                string Filter = string.Format("[ReceivedTime] >= '{0} 12:00 AM' And [ReceivedTime] <= '{1} 11:59 PM'",
                    parsedStartDate.ToString("MM/dd/yyyy"), parsedEndDate.ToString("MM/dd/yyyy"));

                return PrepareGetMailIDsArrayResponse(Int32.MaxValue.ToString(), Filter);
            }
            catch (Exception ex)
            {
                IDs.Add(ERROR + "StartDate and/or EndDate not in valid Date Time format.");
                return IDs.ToArray();
            }
        }

        private string[] PrepareGetMailIDsArrayResponse(string NumberOfItems, string Filter)
        {
            List<string> IDs = new List<string>();

            try
            {
                Outlook.Items items = null;
                Console.WriteLine("\nDescription : " + folder.Description + "\nStoreId : " + folder.StoreID);
                if (string.IsNullOrWhiteSpace(Filter))
                    items = folder.Items;
                else
                    items = folder.Items.Restrict(Filter);

                items.Sort("[ReceivedTime]", false);

                int numberOfItemsReturned = 0;
                for (int i = items.Count; (numberOfItemsReturned < Convert.ToInt32(NumberOfItems)) && (i > 0); i--)
                {
                    if (i <= items.Count && i > 0)
                    {
                        mail = items[i] as Outlook.MailItem;
                        if (mail != null)
                        {
                            Console.WriteLine(mail.Subject.ToString());
                            IDs.Add(mail.EntryID);
                            numberOfItemsReturned++;
                        }
                    }
                }

                return IDs.ToArray();
            }

            catch (Exception ex)
            {
                IDs.Add(ERROR + ex.ToString());
                return IDs.ToArray();
            }
        }

        #endregion

        #region Get MailItem Data

        public string GetSenderSMTPAddress()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                if (mail.SenderEmailType == "EX")
                {
                    Outlook.AddressEntry sender = mail.Sender;
                    if (sender != null)
                    {
                        //Now we have an AddressEntry representing the Sender
                        if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry
                            || sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                        {
                            //Use the ExchangeUser object PrimarySMTPAddress
                            Outlook.ExchangeUser exchUser = sender.GetExchangeUser();
                            if (exchUser != null)
                            {
                                return exchUser.PrimarySmtpAddress;
                            }
                            else
                            {
                                return null;
                            }
                        }
                        else
                        {
                            return sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                        }
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return mail.SenderEmailAddress;
                }
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string[] GetSMTPAddressForCCRecipientsArray()
        {
            List<string> Addresses = new List<string>();

            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                foreach (Outlook.Recipient recip in mail.Recipients)
                {
                    if (recip.Type == (int)(Outlook.OlMailRecipientType.olCC))
                    {
                        Outlook.PropertyAccessor pa = recip.PropertyAccessor;
                        string smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                        Addresses.Add(smtpAddress);
                    }
                }
                return Addresses.ToArray();
            }
            catch (Exception ex)
            {
                Addresses.Add(ERROR + ex.ToString());
                return Addresses.ToArray();
            }
        }

        public string[] GetSMTPAddressForToRecipientsArray()
        {
            List<string> Addresses = new List<string>();

            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                foreach (Outlook.Recipient recip in mail.Recipients)
                {
                    if (recip.Type == (int)(Outlook.OlMailRecipientType.olTo))
                    {
                        Outlook.PropertyAccessor pa = recip.PropertyAccessor;
                        string smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                        Addresses.Add(smtpAddress);
                    }
                }
                return Addresses.ToArray();
            }
            catch (Exception ex)
            {
                Addresses.Add(ERROR + ex.ToString());
                return Addresses.ToArray();
            }
        }

        public string GetSMTPAddressForRecipientsString()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                return mail.To.ToString(); 
            }
            catch (Exception ex)
            {           
                return ERROR + ex.ToString();
            }
        }

        public string GetSMTPAddressCCRecipientsString()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                return mail.CC;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetImportance()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }


                return mail.Importance.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetVotingResponse()
        {
            try
            {
                if (mail == null)
                {
                    return ERROR;
                }

                return mail.VotingResponse.ToString();

            }
            catch (Exception)
            {
                return "No Voting Mails Found";
            }
        }

        public string GetReadReceiptRequested()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                return mail.ReadReceiptRequested.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetReceivedTime()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                return mail.ReceivedTime.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetDownloadState()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                return mail.DownloadState.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetNumberOfAttachments()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }
                return mail.Attachments.Count.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string[] GetListOfAttachments()
        {
            List<string> IDs = new List<string>();

            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                if (mail.Attachments.Count > 0)
                {
                    Outlook.Attachments attachments = mail.Attachments;
                    foreach (Outlook.Attachment attachment in attachments)
                    {
                        IDs.Add(attachment.FileName);
                    }
                }

                return IDs.ToArray();
            }
            catch (Exception ex)
            {
                IDs.Add(ERROR + ex.ToString());
                return IDs.ToArray();
            }
        }

        public string GetBody()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }
                return mail.Body;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetBodyHTML()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }
                return mail.HTMLBody;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetIsBodyHTML()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }
                return (mail.BodyFormat == Outlook.OlBodyFormat.olFormatHTML).ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetIsBodyRTF()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }
                return (mail.BodyFormat == Outlook.OlBodyFormat.olFormatRichText).ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetSubject()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }
                return mail.Subject;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetSentOnDate()

        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }
                return mail.SentOn.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetIsUnReadStatus()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }
                return mail.UnRead.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetFlagStatus()
        {
            try
            {
                if (mail == null)
                {
                    throw new ArgumentNullException();
                }

                if (mail.FlagStatus == Outlook.OlFlagStatus.olFlagComplete)
                    return "Complete";
                else if (mail.FlagStatus == Outlook.OlFlagStatus.olFlagMarked)
                    return "Marked";
                else
                    return string.Empty;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        #endregion

        #region Mail Actions

        public string MarkAsRead(string ID)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                mail.UnRead = false;
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string MarkAsUnRead(string ID)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                mail.UnRead = true;
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }


        public string DownloadAttachments(string ID, string DownloadAttachmentPath)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                if (!Directory.Exists(DownloadAttachmentPath)) return ERROR + "Download file path for attachments does not exist.";

                Outlook.Attachments attachments = mail.Attachments;
                foreach (Outlook.Attachment attachment in attachments)
                {
                    attachment.SaveAsFile(Path.Combine(DownloadAttachmentPath, attachment.FileName));
                }
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string DisplayMail(string ID)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                mail.Display(); 
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string CloseMail(string ID)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                mail.Close(Outlook.OlInspectorClose.olDiscard);     // THIS KILLS ALL OUTLOOK OBJECTS. THEREFORE RECONNECT() METHOD BELOW IS REQUIRED. 
                ReleaseComObject(mail);
                Reconnect();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string DeleteMail(string ID)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                mail.Delete();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string MoveMail(string ID, string FolderPath)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                Outlook.Folder fldr;
                string backslash = @"\";
                
                if (FolderPath.StartsWith(@"\\"))
                {
                    FolderPath = FolderPath.Remove(0, 2);
                }
                String[] folders = FolderPath.Split(backslash.ToCharArray());
                fldr = application.Session.Folders[folders[0]] as Outlook.Folder;
                if (fldr == null)
                    return ERROR + "Folder path not found!";

                for (int i = 1; i <= folders.GetUpperBound(0); i++)
                {
                    Outlook.Folders subFolders = fldr.Folders;
                    fldr = subFolders[folders[i]] as Outlook.Folder;
                    if (fldr == null)
                        return ERROR + "Folder path not found!";
                }

                mail.Move(fldr);
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SaveMailAsPDF(string ID, string FilePath)
        {
            try
            {
                var wordDocPath = Path.ChangeExtension(FilePath, "mht");

                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                // Save Mail as .mht file
                mail.SaveAs(wordDocPath, Outlook.OlSaveAsType.olMHTML);

                // Create a new instance of Word 
                var wordApplication = new Word.Application();
                Thread.Sleep(1000);

                // Open .mht file as Word Document and Export it to PDF format
                var wordDoc = wordApplication.Documents.Open(wordDocPath) as Word.Document;
                wordDoc.ExportAsFixedFormat(FilePath, Word.WdExportFormat.wdExportFormatPDF);
                wordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                wordApplication.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);

                ReleaseComObject(wordDoc);
                ReleaseComObject(wordApplication);

                // Delete the .mht file
                File.Delete(wordDocPath);

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SaveMailAsMSG(string ID, string FilePath)
        {
            try
            {
                var savePath = Path.ChangeExtension(FilePath, "msg");

                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                // Save Mail as .msg file
                mail.SaveAs(savePath, Outlook.OlSaveAsType.olMSG);

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SaveMailAsTXT(string ID, string FilePath)
        {
            try
            {
                var savePath = Path.ChangeExtension(FilePath, "txt");

                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                // Save Mail as .txt file
                mail.SaveAs(savePath, Outlook.OlSaveAsType.olTXT);

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SaveMailAsHTML(string ID, string FilePath)
        {
            try
            {
                var savePath = Path.ChangeExtension(FilePath, "html");

                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                // Save Mail as .html file
                mail.SaveAs(savePath, Outlook.OlSaveAsType.olHTML);

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string ClearMailFlag(string ID)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                mail.ClearTaskFlag();
                mail.Save();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string MarkMailFlagComplete(string ID)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                mail.FlagStatus = Outlook.OlFlagStatus.olFlagComplete;
                mail.Save();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string MarkMailFlag(string ID)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                mail.FlagStatus = Outlook.OlFlagStatus.olFlagMarked;
                mail.Save();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SendMail(string ToRecipients, string CcRecipients, string BccRecipients, string Subject, string Attachments, string Body, string IsBodyHTML, string ReplyRecipients)
        {
            try
            {
                Outlook.MailItem newMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

                PopulateCollectionForMail(newMail, COLLECTIONTYPE.ToRecipients, ToRecipients);

                PopulateCollectionForMail(newMail, COLLECTIONTYPE.CcRecipients, CcRecipients);

                PopulateCollectionForMail(newMail, COLLECTIONTYPE.BccRecipients, BccRecipients);

                PopulateCollectionForMail(newMail, COLLECTIONTYPE.ReplyRecipients, ReplyRecipients);

                newMail.Recipients.ResolveAll();

                PopulateCollectionForMail(newMail, COLLECTIONTYPE.Attachments, Attachments);

                // Subject
                newMail.Subject = Subject;

                // Body
                if (IsBodyHTML.Trim().ToUpper() == "TRUE" || IsBodyHTML.Trim() == "1")
                {
                    newMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                    newMail.HTMLBody = Body;
                }
                else
                {
                    newMail.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                    newMail.Body = Body;
                }

                newMail.SendUsingAccount = account;
                newMail.Send();

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string ReplyMail(string ID, string CcRecipients, string BccRecipients, string OptionalSubjectOverwrite, string Attachments, string Body, string IsBodyOverwrite, string ReplyRecipients)
        {
            return RespondToMail(ID, CcRecipients, BccRecipients, OptionalSubjectOverwrite, Attachments, Body, IsBodyOverwrite, ReplyRecipients, REPLYTYPE.Reply);
        }

        public string ReplyAllMail(string ID, string CcRecipients, string BccRecipients, string OptionalSubjectOverwrite, string Attachments, string Body, string IsBodyOverwrite, string ReplyRecipients)
        {
            return RespondToMail(ID, CcRecipients, BccRecipients, OptionalSubjectOverwrite, Attachments, Body, IsBodyOverwrite, ReplyRecipients, REPLYTYPE.ReplyAll);
        }

        private string RespondToMail(string ID, string CcRecipients, string BccRecipients, string OptionalSubjectOverwrite, string Attachments, string Body, string IsBodyOverwrite, string ReplyRecipients, REPLYTYPE ReplyType)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                Outlook.MailItem replyMail = null;

                if (ReplyType == REPLYTYPE.Reply)
                    replyMail = mail.Reply() as Outlook.MailItem;
                else if (ReplyType == REPLYTYPE.ReplyAll)
                    replyMail = mail.ReplyAll() as Outlook.MailItem;

                PopulateCollectionForMail(replyMail, COLLECTIONTYPE.CcRecipients, CcRecipients);

                PopulateCollectionForMail(replyMail, COLLECTIONTYPE.BccRecipients, BccRecipients);

                PopulateCollectionForMail(replyMail, COLLECTIONTYPE.ReplyRecipients, ReplyRecipients);

                replyMail.Recipients.ResolveAll();

                PopulateCollectionForMail(replyMail, COLLECTIONTYPE.Attachments, Attachments);

                // Subject
                if (!string.IsNullOrWhiteSpace(OptionalSubjectOverwrite)) replyMail.Subject = OptionalSubjectOverwrite;

                SetBody(replyMail, Body, IsBodyOverwrite);

                replyMail.SendUsingAccount = account;
                replyMail.Send();

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string ForwardMail(string ID, string ToRecipients, string CcRecipients, string BccRecipients, string OptionalSubjectOverwrite, string Attachments, string Body, string IsBodyOverwrite, string ReplyRecipients)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                Outlook.MailItem forwardMail = mail.Forward() as Outlook.MailItem;

                PopulateCollectionForMail(forwardMail, COLLECTIONTYPE.ToRecipients, ToRecipients);

                PopulateCollectionForMail(forwardMail, COLLECTIONTYPE.CcRecipients, CcRecipients);

                PopulateCollectionForMail(forwardMail, COLLECTIONTYPE.BccRecipients, BccRecipients);

                PopulateCollectionForMail(forwardMail, COLLECTIONTYPE.ReplyRecipients, ReplyRecipients);

                forwardMail.Recipients.ResolveAll();

                PopulateCollectionForMail(forwardMail, COLLECTIONTYPE.Attachments, Attachments);

                // Subject
                if (!string.IsNullOrWhiteSpace(OptionalSubjectOverwrite)) forwardMail.Subject = OptionalSubjectOverwrite;

                SetBody(forwardMail, Body, IsBodyOverwrite);

                forwardMail.SendUsingAccount = account;
                forwardMail.Send();

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string ComposeNewMail()
        {
            try
            {
                Outlook.MailItem newMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

                newMail.Display();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string ComposeReplyMail(string ID)
        {
            return ComposeMail(ID, REPLYTYPE.Reply);
        }

        public string ComposeReplyAllMail(string ID)
        {
            return ComposeMail(ID, REPLYTYPE.ReplyAll);
        }

        private string ComposeMail(string ID, REPLYTYPE replyType)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                Outlook.MailItem replyMail = null;

                if (replyType == REPLYTYPE.Reply)
                    replyMail = mail.Reply() as Outlook.MailItem;
                else if (replyType == REPLYTYPE.ReplyAll)
                    replyMail = mail.ReplyAll() as Outlook.MailItem;

                replyMail.Display();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string ComposeForwardMail(string ID)
        {
            try
            {
                var result = SelectMailItem(ID);
                if (result != RETURNVALUE)
                    return result;

                Outlook.MailItem forwardMail = mail.Forward() as Outlook.MailItem;

                forwardMail.Display();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        private void PopulateCollectionForMail(Outlook.MailItem Mail, COLLECTIONTYPE CollectionType, string Items)
        {
            if (string.IsNullOrWhiteSpace(Items)) return;

            switch (CollectionType)
            { 

                case COLLECTIONTYPE.ToRecipients:

                    // To Recipients
                    string[] toList = Items.Split(';', ',', '|');
                    for (int i = 0; i < toList.Length; i++)
                    {
                        var recipient = Mail.Recipients.Add(toList[i].Trim());
                        recipient.Type = (int)Outlook.OlMailRecipientType.olTo;
                    }
                    break;

                case COLLECTIONTYPE.CcRecipients:

                    // Cc Recipients
                    string[] ccList = Items.Split(';', ',', '|');
                    for (int i = 0; i < ccList.Length; i++)
                    {
                        var recipient = Mail.Recipients.Add(ccList[i].Trim());
                        recipient.Type = (int)Outlook.OlMailRecipientType.olCC;
                    }
                    break;

                case COLLECTIONTYPE.BccRecipients:

                    // Bcc Recipients
                    string[] bccList = Items.Split(';', ',', '|');
                    for (int i = 0; i < bccList.Length; i++)
                    {
                        var recipient = Mail.Recipients.Add(bccList[i].Trim());
                        recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;
                    }
                    break;

                case COLLECTIONTYPE.ReplyRecipients:

                    // Reply Recipients
                    string[] replyList = Items.Split(';', ',', '|');
                    for (int i = 0; i < replyList.Length; i++)
                    {
                        var recipient = Mail.ReplyRecipients.Add(replyList[i].Trim());
                        recipient.Type = (int)Outlook.OlMailRecipientType.olOriginator;
                    }
                    break;

                case COLLECTIONTYPE.Attachments:

                    // Attachments
                    string[] attachList = Items.Split(';');
                    for (int i = 0; i < attachList.Length; i++)
                    {
                        Mail.Attachments.Add(attachList[i].Trim(), Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }
                    break;

            }
        }

        private void SetBody(Outlook.MailItem Mail, string Body, string IsBodyOverwrite)
        {
            // Body
            if (Mail.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
            {
                Body = Body.Replace(Environment.NewLine, @"<br>");

                if (IsBodyOverwrite.Trim().ToUpper() == "TRUE" || IsBodyOverwrite.Trim() == "1")
                    Mail.HTMLBody = Body;
                else
                {
                    int bodyStartIndex = Mail.HTMLBody.IndexOf("<body", 0, StringComparison.InvariantCultureIgnoreCase);
                    int bodyEndIndex = Mail.HTMLBody.IndexOf(">", bodyStartIndex + 5, StringComparison.InvariantCultureIgnoreCase);
                    Mail.HTMLBody = Mail.HTMLBody.Insert(bodyEndIndex + 1, Body + "<br>");
                }
            }
            else
            {
                if (IsBodyOverwrite.Trim().ToUpper() == "TRUE" || IsBodyOverwrite.Trim() == "1")
                    Mail.Body = Body;
                else
                    Mail.Body = Body + Environment.NewLine + Mail.Body;
            }
        }

        #endregion

        #region Get Calendar Appointments Collection/Array 

        public string[] GetAppointmentIDsArray(string NumberOfItems)
        {
            return PrepareAppointmentIDsArrayResponse(NumberOfItems, string.Empty);
        }

        public string[] GetAppointmentIDsArrayByFilter(string Filter)
        {
            return PrepareAppointmentIDsArrayResponse(Int32.MaxValue.ToString(), Filter);
        }

        public string[] GetAppointmentIDsArrayByDateRange(string StartDate, string EndDate)
        {
            List<string> IDs = new List<string>();

            try
            {
                var parsedStartDate = DateTime.Parse(StartDate);
                var parsedEndDate = DateTime.Parse(EndDate);

                if (parsedStartDate > parsedEndDate)
                {
                    IDs.Add(ERROR + "StartDate cannot be later than EndDate.");
                    return IDs.ToArray();
                }

                string Filter = string.Format("[Start] <= '{0} 11:59 PM' And [End] >= '{1} 12:00 AM'",
                    parsedEndDate.ToString("MM/dd/yyyy"), parsedStartDate.ToString("MM/dd/yyyy"));

                return PrepareAppointmentIDsArrayResponse(Int32.MaxValue.ToString(), Filter, parsedStartDate, parsedEndDate);
            }
            catch (Exception ex)
            {
                IDs.Add(ERROR + "StartDate and/or EndDate not in valid Date Time format.");
                return IDs.ToArray();
            }
        }

        private string[] PrepareAppointmentIDsArrayResponse(string NumberOfItems, string Filter, DateTime? startDate = null, DateTime? endDate = null)
        {
            List<string> IDs = new List<string>();

            try
            {
                Outlook.Items items = folder.Items;

                items.IncludeRecurrences = true;

                if (string.IsNullOrWhiteSpace(Filter))
                    items = folder.Items;
                else
                    items = folder.Items.Restrict(Filter);      // e.g. "[Start] > '11/13/2017 12:00 AM' And [Start] < '11/17/2017 12:00 AM'"

                items.Sort("[Start]");

                int numberOfItemsReturned = 0;
                for (int i = items.Count; (numberOfItemsReturned < Convert.ToInt32(NumberOfItems)) && (i > 0); i--)
                {
                    if (i <= items.Count && i > 0)
                    {
                        appointment = items[i] as Outlook.AppointmentItem;
                        if (appointment != null && !appointment.IsRecurring)
                        {
                            IDs.Add(appointment.EntryID);
                            numberOfItemsReturned++;
                        }
                        else if (appointment != null && appointment.IsRecurring)
                        {
                            if (startDate != null && endDate != null)
                            {
                                var recurPattern = appointment.GetRecurrencePattern();
                                var specificDate = startDate.Value;
                                specificDate = specificDate + appointment.Start.TimeOfDay;

                                while (true)
                                {
                                    try
                                    {
                                        var specificAppt = recurPattern.GetOccurrence(specificDate);
                                        IDs.Add(specificAppt.EntryID + "||" + specificDate.ToString());
                                        numberOfItemsReturned++;
                                        ReleaseComObject(specificAppt);
                                    }
                                    catch (Exception ex)
                                    {
                                        // Move on. 
                                    }

                                    specificDate = specificDate.AddDays(1);
                                    if (specificDate > endDate.Value)
                                        break;
                                }

                                ReleaseComObject(recurPattern);
                            }
                            else
                            {
                                IDs.Add(appointment.EntryID);
                                numberOfItemsReturned++;
                            }
                        }

                        ReleaseComObject(appointment);
                    }
                }

                return IDs.ToArray();
            }
            catch (Exception ex)
            {
                IDs.Add(ERROR + ex.ToString());
                return IDs.ToArray();
            }
        }

        #endregion

        #region Get AppointmentItem Data

        public string GetAppointmentOrganizerSMTPAddress()
        {
            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }

                var organizer = appointment.GetOrganizer();
                if (organizer.Type == "EX")
                {
                    //Now we have an AddressEntry representing the Sender
                    if (organizer.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry
                        || organizer.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                    {
                        //Use the ExchangeUser object PrimarySMTPAddress
                        Outlook.ExchangeUser exchUser = organizer.GetExchangeUser();
                        if (exchUser != null)
                        {
                            return exchUser.PrimarySmtpAddress;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    else
                    {
                        return organizer.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                    }
                }
                else
                {
                    return organizer.Address;
                }
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        private Outlook.AddressEntry GetMeetingOrganizer(Outlook.AppointmentItem appt)
        {
            if (appt == null)
            {
                throw new ArgumentNullException();
            }
            string organizerEntryID = appt.PropertyAccessor.BinaryToString(appt.PropertyAccessor.GetProperty(PR_SENT_REPRESENTING_ENTRYID));
            Outlook.AddressEntry organizer = application.Session.GetAddressEntryFromID(organizerEntryID);
            if (organizer != null)
            {
                return organizer;
            }
            else
            {
                return null;
            }
        }

        public string GetAppointmentLocation()
        {
            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }

                return appointment.Location;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetAppointmentSubject()
        {
            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }

                return appointment.Subject;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetAppointmentStartDateTime()
        {
            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }

                return appointment.Start.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetAppointmentEndDateTime()
        {
            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }

                return appointment.End.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetAppointmentDurationInMinutes()
        {
            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }

                return appointment.Duration.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetAppointmentBody()
        {
            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }

                return appointment.Body;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetAppointmentNumberOfAttachments()
        {
            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }
                return appointment.Attachments.Count.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string[] GetAppointmentRecipientsArray()
        {
            List<string> Recipients = new List<string>();

            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }

                Outlook.Recipients recipients = appointment.Recipients;
                if (recipients != null)
                {
                    foreach (Outlook.Recipient recipient in recipients)
                    {
                        if (recipient.AddressEntry.Type == "EX")
                        {
                            //Now we have an AddressEntry representing the Sender
                            if (recipient.AddressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry
                                || recipient.AddressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                            {
                                //Use the ExchangeUser object PrimarySMTPAddress
                                Outlook.ExchangeUser exchUser = recipient.AddressEntry.GetExchangeUser();
                                if (exchUser != null)
                                {
                                    Recipients.Add(exchUser.PrimarySmtpAddress);
                                }
                            }
                            else
                            {
                                Recipients.Add(recipient.AddressEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string);
                            }
                        }
                        else
                        {
                            Recipients.Add(recipient.Address);
                        }
                    }
                }

                return Recipients.ToArray();
            }
            catch (Exception ex)
            {
                Recipients.Add(ex.Message);
                return Recipients.ToArray();
            }
        }

        public string GetIsAppointmentRecurring()
        {
            try
            {
                if (appointment == null)
                {
                    throw new ArgumentNullException();
                }
                return appointment.IsRecurring.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        #endregion

        #region Calendar Appointment Actions

        public string DownloadAppointmentAttachments(string ID, string DownloadAttachmentPath)
        {
            try
            {
                var result = SelectAppointment(ID);
                if (result != RETURNVALUE)
                    return result;

                if (!Directory.Exists(DownloadAttachmentPath)) return ERROR + "Download file path for attachments does not exist.";

                Outlook.Attachments attachments = appointment.Attachments;
                foreach (Outlook.Attachment attachment in attachments)
                {
                    attachment.SaveAsFile(Path.Combine(DownloadAttachmentPath, attachment.FileName));
                }
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string DisplayAppointment(string ID)
        {
            try
            {
                var result = SelectAppointment(ID);
                if (result != RETURNVALUE)
                    return result;

                appointment.Display();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string CloseAppointment(string ID)
        {
            try
            {
                var result = SelectAppointment(ID);
                if (result != RETURNVALUE)
                    return result;

                appointment.Close(Outlook.OlInspectorClose.olDiscard);     // THIS KILLS ALL OUTLOOK OBJECTS. THEREFORE RECONNECT() METHOD BELOW IS REQUIRED. 
                ReleaseComObject(appointment);
                Reconnect();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string DeleteAppointment(string ID)
        {
            try
            {
                var result = SelectAppointment(ID);
                if (result != RETURNVALUE)
                    return result;

                appointment.Delete();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string CreateAppointment(string Subject, string Location, string StartDateTime, string EndDateTime, string Attachments, string Body)
        {
            try
            {
                Outlook.AppointmentItem appointment = application.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;

                // Attachments
                if (!string.IsNullOrWhiteSpace(Attachments))
                {
                    string[] attachList = Attachments.Split(';');
                    for (int i = 0; i < attachList.Length; i++)
                    {
                        appointment.Attachments.Add(attachList[i].Trim(), Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }
                }

                appointment.Subject = Subject;
                appointment.Location = Location;
                appointment.Start = DateTime.Parse(StartDateTime);
                appointment.End = DateTime.Parse(EndDateTime);
                appointment.Body = Body;
                appointment.Save();

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        #endregion

        #region Get Meeting Request Collection/Array 

        public string[] GetMeetingRequestIDsArray(string NumberOfItems)
        {
            return PrepareMeetingRequestIDsArrayResponse(NumberOfItems, string.Empty);
        }

        public string[] GetMeetingRequestIDsArrayByFilter(string Filter)
        {
            return PrepareMeetingRequestIDsArrayResponse(Int32.MaxValue.ToString(), Filter);
        }

        public string[] GetMeetingRequestIDsArrayByDateRange(string StartDate, string EndDate)
        {
            List<string> IDs = new List<string>();

            try
            {
                var parsedStartDate = DateTime.Parse(StartDate);
                var parsedEndDate = DateTime.Parse(EndDate);
                string Filter = string.Format("[ReceivedTime] >= '{0} 12:00 AM' And [ReceivedTime] <= '{1} 11:59 PM'",
                    parsedStartDate.ToString("MM/dd/yyyy"), parsedEndDate.ToString("MM/dd/yyyy"));

                return PrepareMeetingRequestIDsArrayResponse(Int32.MaxValue.ToString(), Filter);
            }
            catch (Exception ex)
            {
                IDs.Add(ERROR + "StartDate and/or EndDate not in valid Date Time format.");
                return IDs.ToArray();
            }
        }

        private string[] PrepareMeetingRequestIDsArrayResponse(string NumberOfItems, string Filter)
        {
            List<string> IDs = new List<string>();

            try
            {
                Outlook.Items items = null;

                if (string.IsNullOrWhiteSpace(Filter))
                    items = folder.Items;
                else
                    items = folder.Items.Restrict(Filter);      // e.g. "[ReceivedTime] > '11/13/2017 12:00 AM' And [ReceivedTime] < '11/17/2017 12:00 AM'"

                items.Sort("[ReceivedTime]", false);

                int numberOfItemsReturned = 0;
                for (int i = items.Count; (numberOfItemsReturned < Convert.ToInt32(NumberOfItems)) && (i > 0); i--)
                {
                    if (i <= items.Count && i > 0)
                    {
                        meeting = items[i] as Outlook.MeetingItem;
                        if (meeting != null)
                        {
                            IDs.Add(meeting.EntryID);
                            numberOfItemsReturned++;
                        }
                    }
                }

                return IDs.ToArray();
            }

            catch (Exception ex)
            {
                IDs.Add(ERROR + ex.ToString());
                return IDs.ToArray();
            }
        }

        #endregion

        #region Get MeetingItem Data

        public string GetMeetingRequestSenderSMTPAddress()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }

                if (meeting.SenderEmailType == "EX")
                {
                    Outlook.AddressEntry sender = GetMeetingOrganizer(meeting.GetAssociatedAppointment(true));
                    if (sender != null)
                    {
                        //Now we have an AddressEntry representing the Sender
                        if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry
                            || sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                        {
                            //Use the ExchangeUser object PrimarySMTPAddress
                            Outlook.ExchangeUser exchUser = sender.GetExchangeUser();
                            if (exchUser != null)
                            {
                                return exchUser.PrimarySmtpAddress;
                            }
                            else
                            {
                                return null;
                            }
                        }
                        else
                        {
                            return sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                        }
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return meeting.SenderEmailAddress;
                }
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetMeetingRequestLocation()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }

                return meeting.GetAssociatedAppointment(true).Location;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetMeetingRequestSubject()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }

                return meeting.GetAssociatedAppointment(true).Subject;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetMeetingRequestStartDateTime()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }

                return meeting.GetAssociatedAppointment(true).Start.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetMeetingRequestEndDateTime()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }

                return meeting.GetAssociatedAppointment(true).End.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetMeetingRequestDurationInMinutes()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }

                return meeting.GetAssociatedAppointment(true).Duration.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetMeetingRequestBody()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }

                return meeting.GetAssociatedAppointment(true).Body;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetMeetingRequestNumberOfAttachments()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }
                return meeting.Attachments.Count.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string[] GetMeetingRequestRecipientsArray()
        {
            List<string> Recipients = new List<string>();

            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }

                Outlook.Recipients recipients = meeting.GetAssociatedAppointment(true).Recipients;
                if (recipients != null)
                {
                    foreach (Outlook.Recipient recipient in recipients)
                    {
                        if (recipient.AddressEntry.Type == "EX")
                        {
                            //Now we have an AddressEntry representing the Sender
                            if (recipient.AddressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry
                                || recipient.AddressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                            {
                                //Use the ExchangeUser object PrimarySMTPAddress
                                Outlook.ExchangeUser exchUser = recipient.AddressEntry.GetExchangeUser();
                                if (exchUser != null)
                                {
                                    Recipients.Add(exchUser.PrimarySmtpAddress);
                                }
                            }
                            else
                            {
                                Recipients.Add(recipient.AddressEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string);
                            }
                        }
                        else
                        {
                            Recipients.Add(recipient.Address);
                        }
                    }
                }

                return Recipients.ToArray();
            }
            catch (Exception ex)
            {
                Recipients.Add(ex.Message);
                return Recipients.ToArray();
            }
        }

        public string GetMeetingRequestSentOnDate()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }
                return meeting.SentOn.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string GetIsMeetingRequestUnRead()
        {
            try
            {
                if (meeting == null)
                {
                    throw new ArgumentNullException();
                }
                return meeting.UnRead.ToString();
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        #endregion

        #region Meeting Request Actions

        public string MeetingRequestMarkAsRead(string ID)
        {
            try
            {
                var result = SelectMeetingRequest(ID);
                if (result != RETURNVALUE)
                    return result;

                meeting.UnRead = false;
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string MeetingRequestMarkAsUnRead(string ID)
        {
            try
            {
                var result = SelectMeetingRequest(ID);
                if (result != RETURNVALUE)
                    return result;

                meeting.UnRead = true;
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }
        
        public string DownloadMeetingRequestAttachments(string ID, string DownloadAttachmentPath)
        {
            try
            {
                var result = SelectMeetingRequest(ID);
                if (result != RETURNVALUE)
                    return result;

                if (!Directory.Exists(DownloadAttachmentPath)) return ERROR + "Download file path for attachments does not exist.";

                Outlook.Attachments attachments = meeting.Attachments;
                foreach (Outlook.Attachment attachment in attachments)
                {
                    attachment.SaveAsFile(Path.Combine(DownloadAttachmentPath, attachment.FileName));
                }
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string DisplayMeetingRequest(string ID)
        {
            try
            {
                var result = SelectMeetingRequest(ID);
                if (result != RETURNVALUE)
                    return result;

                meeting.Display();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string CloseMeetingRequest(string ID)
        {
            try
            {
                var result = SelectMeetingRequest(ID);
                if (result != RETURNVALUE)
                    return result;

                meeting.Close(Outlook.OlInspectorClose.olDiscard);     // THIS KILLS ALL OUTLOOK OBJECTS. THEREFORE RECONNECT() METHOD BELOW IS REQUIRED. 
                ReleaseComObject(meeting);
                Reconnect();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string DeleteMeetingRequest(string ID)
        {
            try
            {
                var result = SelectMeetingRequest(ID);
                if (result != RETURNVALUE)
                    return result;

                meeting.Delete();
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string AcceptMeetingRequest(string ID)
        {
            try
            {
                var result = SelectMeetingRequest(ID);
                if (result != RETURNVALUE)
                    return result;

                var associatedAppointment = meeting.GetAssociatedAppointment(true);
                if (associatedAppointment != null)
                {
                    var response = associatedAppointment.Respond(Outlook.OlMeetingResponse.olMeetingAccepted, true, false);
                    response.Send();
                    DeleteMeetingRequest(ID);
                }
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string TentativeMeetingRequest(string ID)
        {
            try
            {
                var result = SelectMeetingRequest(ID);
                if (result != RETURNVALUE)
                    return result;

                var associatedAppointment = meeting.GetAssociatedAppointment(true);
                if (associatedAppointment != null)
                {
                    var response = associatedAppointment.Respond(Outlook.OlMeetingResponse.olMeetingTentative, true, false);
                    response.Send();
                    DeleteMeetingRequest(ID);
                }
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string DeclineMeetingRequest(string ID)
        {
            try
            {
                var result = SelectMeetingRequest(ID);
                if (result != RETURNVALUE)
                    return result;

                var associatedAppointment = meeting.GetAssociatedAppointment(false);
                if (associatedAppointment != null)
                {
                    var response = associatedAppointment.Respond(Outlook.OlMeetingResponse.olMeetingDeclined, true, false);
                    response.Send();
                    DeleteMeetingRequest(ID);
                }
                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        public string SendMeetingRequest(string RequiredAttendees, string OptionalAttendees, string Subject, string Location, string StartDateTime, string EndDateTime, string Attachments, string Body)
        {
            try
            {
                Outlook.AppointmentItem appointment = application.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;

                // Required Attendees
                if (!string.IsNullOrWhiteSpace(RequiredAttendees))
                {
                    string[] reqList = RequiredAttendees.Split(';');
                    for (int i = 0; i < reqList.Length; i++)
                    {
                        var recipient = appointment.Recipients.Add(reqList[i].Trim());
                        recipient.Type = (int)Outlook.OlMeetingRecipientType.olRequired;
                    }
                }

                // Optional Attendees
                if (!string.IsNullOrWhiteSpace(OptionalAttendees))
                {
                    string[] optList = OptionalAttendees.Split(';');
                    for (int i = 0; i < optList.Length; i++)
                    {
                        var recipient = appointment.Recipients.Add(optList[i].Trim());
                        recipient.Type = (int)Outlook.OlMeetingRecipientType.olOptional;
                    }
                }

                appointment.Recipients.ResolveAll();

                // Attachments
                if (!string.IsNullOrWhiteSpace(Attachments))
                {
                    string[] attachList = Attachments.Split(';');
                    for (int i = 0; i < attachList.Length; i++)
                    {
                        appointment.Attachments.Add(attachList[i].Trim(), Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }
                }

                appointment.Subject = Subject;
                appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                appointment.Location = Location;
                appointment.Start = DateTime.Parse(StartDateTime);
                appointment.End = DateTime.Parse(EndDateTime);
                appointment.Body = Body;
                appointment.SendUsingAccount = account;
                appointment.Send();

                return RETURNVALUE;
            }
            catch (Exception ex)
            {
                return ERROR + ex.ToString();
            }
        }

        #endregion

    }
}