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
using System.Text.RegularExpressions;
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
                Console.WriteLine(DateTime.Now.ToString("hh:mm:ss tt") + " Connecting to imap...");
                MailMessage message = connectToImap(mailFolder);


                //Parse the message
                Console.WriteLine(DateTime.Now.ToString("hh:mm:ss tt") + " Parsing the e-mail...");
                String mailText = parseEmailMessage(message);

                //Get the name of the last downloaded file
                String lastDownload = Properties.Settings.Default["lastDownload"].ToString();

                if (lastDownload.Equals(mailText)) {
                    Console.WriteLine(DateTime.Now.ToString("hh:mm:ss tt") + " Nothing new... ");
                    Console.WriteLine(lastDownload);
                }
                else
                {
                    //Download the file, using firfox as a helper
                    Console.WriteLine(DateTime.Now.ToString("hh:mm:ss tt") + " Downloading the file... ");
                    Console.WriteLine(mailText);
                    downloadFile(mailText);

                    //Rename the downloaded file and remove "EUR" from the values
                    Console.WriteLine(DateTime.Now.ToString("hh:mm:ss tt") + " Waiting 30 seconds...");
                    Thread.Sleep(30000); //Wait a half-minute
                    Console.WriteLine(DateTime.Now.ToString("hh:mm:ss tt") + " Parsing the file...");
                    parseFile(filepath);




                    //Kill the firefox process
                    Console.WriteLine(DateTime.Now.ToString("hh:mm:ss tt") + " Killing the fox...");
                    killFirefox();

                    Properties.Settings.Default["lastDownload"] = mailText;
                    Properties.Settings.Default.Save();
                }

                //Sleep for 1.08 hours and "restart" the program
                Console.WriteLine(DateTime.Now.ToString("hh:mm:ss tt") + " See you in " + (waitingTime / 1000 / 60).ToString() + " minutes...");
                Console.WriteLine("");
                Thread.Sleep(waitingTime);
            }

        }

        private static void parseFile(DirectoryInfo filepath)
        {
            FileInfo downloadedFile = filepath.GetFiles()
                                .OrderByDescending(f => f.LastWriteTime)
                                .First();

            String csvText = File.ReadAllText(@downloadedFile.FullName);
            String pattern = @"\bEUR \b";
            String replace = "";
            String result = Regex.Replace(csvText, pattern, replace);
            String newFilePath = filepath.ToString() + "ebayOrder.csv";
            File.Delete(newFilePath);
            File.WriteAllText(newFilePath, @result);
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
