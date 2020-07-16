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
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
using System.Globalization;

namespace OutlookConsole
{
    public class Program
    {
        // This is a tester Class. The purpose of this is to test the OutlookWrapperV2 Class. Only to be used for internal purpose.
        static void Main(string[] args)
        {

            #region Test Meeting Requests

            OutlookWrapper.OutlookWrapper wrapper = new OutlookWrapper.OutlookWrapper();

            wrapper.SetPR_SMTP_ADDRESS("http://schemas.microsoft.com/mapi/proptag/0x39FE001E");
            wrapper.SetPR_SENT_REPRESENTING_ENTRYID("http://schemas.microsoft.com/mapi/proptag/0x00410102");

            Console.WriteLine(wrapper.LaunchOutlook());

            Console.WriteLine("Profile" + wrapper.SelectProfile());

            Console.WriteLine("Account" + wrapper.SelectAccount());

            Console.WriteLine("Inbox"+wrapper.SelectInbox());

            string[] ids = wrapper.GetAllMailIDsArray("1");

            for (int i = 0; i < ids.Length; i++)
            {
                Console.WriteLine("Subject"+wrapper.GetSubject());
                Console.WriteLine("SelectMail ID" + wrapper.SelectMailItem(ids[i]));
                Console.WriteLine("Recipient Array"+wrapper.GetSMTPAddressForToRecipientsArray());
                Console.WriteLine(wrapper.GetSMTPAddressForRecipientsString());
                Console.WriteLine(wrapper.GetSMTPAddressCCRecipientsString());
                Console.WriteLine(wrapper.GetImportance());
                Console.WriteLine(wrapper.GetReceivedTime());
                Console.WriteLine(wrapper.GetVotingResponse());
                Console.WriteLine(wrapper.GetReadReceiptRequested());

                Console.WriteLine(ids[i]);
                Console.ReadKey();

                wrapper.CloseMail(ids[i]);

            }

            Console.ReadKey();
        }

        #endregion

        //Commented out test regions only meant for testing purposes. Uncomment necessarily to call classes for test.

        #region Test Calendar Appointments

        //OutlookWrapper.OutlookWrapper wrapper = new OutlookWrapper.OutlookWrapper();

        //Console.WriteLine(wrapper.LaunchOutlook());

        //Console.WriteLine(wrapper.SelectProfile());

        //Console.WriteLine(wrapper.SelectAccount());

        //Console.WriteLine(wrapper.SelectCalendar());

        //Console.WriteLine(wrapper.CreateAppointment(
        //     "Test Create Appointment"
        //    , "Go to meeting. Dial 1-800-800-8000"
        //    , "01/19/2018 1:30 PM"
        //    , "01/19/2018 2:30 PM"
        //    , @"c:\temp\readme.txt;c:\temp\receipt image.jpg"
        //    , @"This is first line of body. 

        //This is second line of body."
        //    ));

        //string[] ids = wrapper.GetAppointmentIDsArrayByDateRange("6/25/2018", "6/25/2018");

        //for (int i = 0; i < ids.Length; i++)
        //{
        //    Console.WriteLine(ids[i]);

        //    wrapper.DisplayAppointment(ids[i]);

        //Console.ReadKey();

        // wrapper.CloseAppointment(ids[i]);

        //wrapper.SelectAppointment(ids[i]);

        //    Console.WriteLine(wrapper.GetAppointmentOrganizerSMTPAddress());

        //    Console.WriteLine(wrapper.GetAppointmentLocation());

        //    Console.WriteLine(wrapper.GetAppointmentSubject());

        //    Console.WriteLine(wrapper.GetAppointmentStartDateTime());

        //    Console.WriteLine(wrapper.GetAppointmentEndDateTime());

        //    Console.WriteLine(wrapper.GetAppointmentDurationInMinutes());

        //Console.WriteLine(wrapper.GetIsAppointmentRecurring());

        //    var ccRec = wrapper.GetAppointmentRecipientsArray();
        //    Console.WriteLine("Recepients:");
        //    foreach (var one in ccRec)
        //    {
        //        Console.WriteLine(one);
        //    }

        //    //Console.WriteLine(wrapper.GetAppointmentBody());

        //    // wrapper.DownloadAppointmentAttachments(ids[i], @"C:\Temp\Outlook\");

        // Console.ReadKey();
        // }

        //Console.WriteLine(wrapper.CloseOutlook());

        //Console.ReadKey();

        #endregion

        /*  #region Test Shared Folders

          OutlookWrapper.OutlookWrapper wrapper = new OutlookWrapper.OutlookWrapper();

          Console.WriteLine(wrapper.LaunchOutlook());

          Console.WriteLine(wrapper.SelectProfile());

          Console.WriteLine(wrapper.SelectAccount());

          Console.WriteLine(wrapper.SelectInbox());

          Console.WriteLine(wrapper.SelectFolderByPath(@"\\ssingh44@calstatela.edu\Inbox"));

          //tring[] ids = wrapper.GetSharedFoldersArray(); 

          string[] ids = wrapper.GetAllMailIDsArray("5");

          for (int i = 0; i < ids.Length; i++)
          {
              Console.WriteLine(wrapper.GetSubject());
              Console.WriteLine(wrapper.SelectMailItem(ids[i]));

              //Console.WriteLine(ids[i]);

              //wrapper.DisplayMail(ids[i]);

              //Console.ReadKey();

              //wrapper.CloseMail(ids[i]);

          }

          Console.ReadKey();

          //Console.WriteLine(wrapper.CloseOutlook());

          #endregion   */

    }
    }
//}
