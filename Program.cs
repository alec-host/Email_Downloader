using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
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
        private static int loopCount = 0;

        private static string inputString;

        static void Main(string[] args)
        {
            String [] data;
            Console.WriteLine("========================");
            Console.WriteLine("Read External File");
            Console.WriteLine("========================");

            string destPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appConfigData.log");

            if (File.Exists(destPath))
            {
                inputString = File.ReadAllText(destPath);
            }
            else
            {
                Environment.Exit(0);
            }

            if (inputString.Trim() != "" && inputString.Contains("#"))
            {
                data = inputString.Split('#');

                CreateXmlFile xmlFile = new CreateXmlFile();

                //-.create config file.
                xmlFile.CreateDirectoryAndFile(data[0], data[1]);
                Console.WriteLine("------------------------");
                Console.WriteLine("File Created: " + data[0]);
                Console.WriteLine("========================");

                //-.read xml file.
                string config = xmlFile.ReadXmlFile(data[0], data[1]);

                connectionString = config.Split('#')[0];
                tableName        = config.Split('#')[1];
                emailURL         = config.Split('#')[2];
                address          = config.Split('#')[3];

                List<string> listOfEmail = new List<string>();

                if(address.Contains(",") && address.Contains("|"))
                {
                    Console.WriteLine("========================");
                    Console.Out.WriteLine("Accessing Email Information");
                    Console.WriteLine("========================");

                    string[] emailArray = address.Split(',');
                    //-.loop through the array and add the data to a list.
                    for (int j=0;j<=emailArray.Length-1;j++)
                    {
                        listOfEmail.Add(emailArray[j]);
                        Console.WriteLine("Email list  "+j+"  "+ emailArray[j]);
                    }

                    while (true)
                    { 
                        loopCount = (loopCount + 1);
                        //-.fetch each item in a list.
                        foreach (var fetchedEmail in listOfEmail)
                        {
                            EmailDownloader(connectionString, tableName, emailURL, fetchedEmail.Split('|')[0], fetchedEmail.Split('|')[1]);
                            Console.WriteLine("" + loopCount+ " "+ fetchedEmail.Split('|')[0]);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("========================");
                    Console.Out.WriteLine("Single Email");
                    Console.WriteLine("========================");
                    while (true)
                    {
                        loopCount = (loopCount + 1);
                        EmailDownloader(connectionString, tableName, emailURL, address, password);

                        Console.WriteLine("" + loopCount);
                    }
                }
            }
            else
            {
                Console.WriteLine("Missing File");
                Console.ReadLine();
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

                try
                {
                    foreach (EmailMessage email in _service.FindItems(WellKnownFolderName.Inbox, new ItemView(1)))
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
                            DateTime.Parse(email.DateTimeReceived.ToString()),
                            System.DateTime.Now,
                            0);

                        Console.WriteLine(System.DateTime.Now + " Email downloaded to DB");

                        if (loopCount >= 1000)
                        {

                            Environment.Exit(0);
                        }
                    }
                }
                catch (Exception e)
                {
                    Environment.Exit(0);
                }

                // Console.ReadLine();

                Console.WriteLine(System.DateTime.Now + " Terminating...");

                Environment.Exit(0);
            }
            else
            {
                if(loopCount >= 100)
                {
                    Environment.Exit(0);
                }
            }
        }
    }
 }
