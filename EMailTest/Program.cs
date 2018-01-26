using Spire.Email;
using Spire.Email.IMap;
using Spire.Email.Smtp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
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
            while (1 == 1)  //Endless loop
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
                imap.Select("Inbox");

                Int64 messageCounter = imap.GetMessageCount("Inbox");

                //Get the first message by its sequence number
                MailMessage message = imap.GetFullMessage(messageCounter.ToString());

                //Parse the message
                String mailTextOriginal = message.BodyText;
                //Cuts everything before the link
                String mailText = mailTextOriginal.Substring(0, mailTextOriginal.LastIndexOf(">"));
                //Cuts everything after the link
                mailText = mailText.Substring(mailText.LastIndexOf("<"), mailText.Length - mailText.LastIndexOf("<"));
                //Cuts first char (<)
                mailText = mailText.Substring(1, mailText.Length - 1);

                //Download the file, using firfox as a helper
                using (var client = new WebClient()){
                    Process.Start(mailText);
                }

                Console.WriteLine(mailText);
                Thread.Sleep(60000); //Wait a minute

                //Kill the firefox process
                foreach (var process in Process.GetProcessesByName("firefox")) {
                    process.Kill();
                }


                Thread.Sleep(3888000);  //Sleep for 1.08 hours
            }
                
        }
    }
}
