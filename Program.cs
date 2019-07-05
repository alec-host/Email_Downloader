using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NationKeEmailDownload
{
    class Program
    {
        private static string connectionString = null;
        private static string tableName = null;
        private static string emailURL = null;
        private static string address = null;
        private static string password = null;

        static void Main(string[] args)
        {
            //-.create config file.
            CreateXmlFile xmlFile = new CreateXmlFile();
            xmlFile.CreateDirectoryAndFile();

            //-.read xml file.
            string config = xmlFile.ReadXmlFile();

            connectionString = config.Split('#')[0];
            tableName = config.Split('#')[1];
            emailURL = config.Split('#')[2];
            address = config.Split('#')[3];
            password = config.Split('#')[4];


            while (true)
            {
                EmailDownloader(connectionString, tableName, emailURL, address, password);
            }
        }
        //-.method to download mails.
        private static void EmailDownloader(string connectionString, string tableName, string emailURL, string address, string password)
        {

            NetworkConnection network = new NetworkConnection();

            Boolean InternetConnection = network.IsInternetAvailable();


            if (InternetConnection)
            {

                Db db = new Db(connectionString, tableName);

                ExchangeService _service;
                try
                {
                    Console.WriteLine(System.DateTime.Now + " Registering Exchange connection");

                    _service = new ExchangeService
                    {
                        Credentials = new WebCredentials(address, password)
                    };

                }
                catch
                {
                    Console.WriteLine(System.DateTime.Now + " ExchangeService failed. Press enter to exit:");
                    return;
                }

                _service.Url = new Uri(emailURL);


                foreach (EmailMessage email in _service.FindItems(WellKnownFolderName.Inbox, new ItemView(100)))
                {
                    email.Load(new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.TextBody));

                    // Then you can retrieve extra information like this:
                    string cc_recipients = "";

                    foreach (EmailAddress emailAddress in email.CcRecipients)
                    {
                        cc_recipients += ";" + emailAddress.Address.ToString();

                    }
                    string to_recipients = "";
                    foreach (EmailAddress emailAddress in email.ToRecipients)
                    {
                        to_recipients += ";" + emailAddress.Address.ToString();
                    }
                    //-.save emails.
                    db.SaveData(
                        email.InternetMessageId,
                        email.From.Address,
                        to_recipients,
                        cc_recipients,
                        email.Subject,
                        email.TextBody.ToString(),
                        email.DateTimeReceived.ToUniversalTime().ToString(),
                        System.DateTime.Now,
                        0);

                    Console.WriteLine(System.DateTime.Now + " Email downloaded to DB");
                }


                // Console.ReadLine();

                Console.WriteLine(System.DateTime.Now + " Terminating...");

                Environment.Exit(0);
            }
        }
    }
 }
