using Spire.Email;
using Spire.Email.IMap;
using Spire.Email.Smtp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EMailTest
{
    class Program
    {
        static void Main(string[] args)
        {
            String mailFolder = "Inbox";
            int waitingTime = 3888000;
            var filepath = new DirectoryInfo("C:\\eNVenta-ERP\\BMECat\\eBay\\");

            while (1 == 1)  //Endless loop
            {
                //Connect to imap server
                Console.WriteLine("Connecting to imap...");
                MailMessage message = connectToImap(mailFolder);


                //Parse the message
                Console.WriteLine("Parsing the e-mail...");
                string mailText = parseEmailMessage(message);


                //Download the file, using firfox as a helper
                Console.WriteLine("Downloading file...");
                downloadFile(mailText);


                //Rename the downloaded file
                Thread.Sleep(60000); //Wait a minute
                Console.WriteLine("Renaming the file...");
                renameDownloadedFile(filepath);

                //Kill the firefox process
                Console.WriteLine("Killing the fox...");
                killFirefox();

                //Sleep for 1.08 hours and "restart" the program
                Console.WriteLine("See you in " + (waitingTime / 1000 / 60).ToString() + " minutes...");
                Console.WriteLine("");
                Thread.Sleep(waitingTime);
            }

        }

        private static void renameDownloadedFile(DirectoryInfo filepath)
        {
            var downloadedFile = filepath.GetFiles()
                                                 .OrderByDescending(f => f.LastWriteTime)
                                                 .First();
            downloadedFile.MoveTo(filepath.ToString() + "ebayOrder.csv");
        }

        private static void downloadFile(string mailText)
        {
            using (var client = new WebClient())
            {
                Process.Start(mailText);
            }
        }

        private static MailMessage connectToImap(string mailFolder)
        {
            //Create an IMAP client
            ImapClient imap = new ImapClient();
            // Set host, username, password etc. for the client
            imap.Host = "imap.udag.de";
            imap.Port = 143;
            imap.Username = "loechel-industriebedarfde-0025";
            imap.Password = "loechel123";
            imap.ConnectionProtocols = ConnectionProtocols.Ssl;
            //Connect the server
            imap.Connect();

            //Select Inbox folder
            imap.Select(mailFolder);

            Int64 messageCounter = imap.GetMessageCount(mailFolder);

            //Get the first message by its sequence number
            MailMessage message = imap.GetFullMessage(messageCounter.ToString());
            return message;
        }

        private static string parseEmailMessage(MailMessage message)
        {
            String mailTextOriginal = message.BodyText;
            //Cuts everything before the link
            String mailText = mailTextOriginal.Substring(0, mailTextOriginal.LastIndexOf(">"));
            //Cuts everything after the link
            mailText = mailText.Substring(mailText.LastIndexOf("<"), mailText.Length - mailText.LastIndexOf("<"));
            //Cuts first char (<)
            mailText = mailText.Substring(1, mailText.Length - 1);
            return mailText;
        }

        private static void killFirefox()
        {
            foreach (var process in Process.GetProcessesByName("firefox"))
            {
                process.Kill();
            }
        }
    }
}
