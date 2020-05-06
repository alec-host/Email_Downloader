using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace NationKeEmailDownload
{
    class CreateXmlFile
    {
        private readonly string xml =
        "<?xml version = \"1.0\" encoding=\"utf-8\" ?>" + Environment.NewLine +
        "<nationmedia>" + Environment.NewLine +
            "<connectionString>Data Source = (localdb)\\MSSQLLocalDB;Initial Catalog = NMG; Integrated Security = True; Connect Timeout = 30; Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False</connectionString>" + Environment.NewLine +
            "<tableName>tbl_mail</tableName>" + Environment.NewLine  +
            "<emailUrl>https://outlook.office365.com/EWS/Exchange.asmx</emailUrl>" + Environment.NewLine +
            "<address>sochola@ke.nationmedia.com|1c1pE9494,one@two.com|password_1,three@two.com|password_2</address>" + Environment.NewLine +
        "</nationmedia>";

        private readonly string dir = @"C:\\";
        public  readonly string file = "AppFile.xml";
        private readonly string systemFile = "AppSystem.log";

        //-.method to create a folder & file.
        public void CreateDirectoryAndFile(String fileName, string dirName) {
            if (!Directory.Exists(dir + dirName))
            {
                Directory.CreateDirectory(dir + dirName);
                File.WriteAllText(Path.Combine(dir + dirName, fileName), xml);
            }
            else {
                File.WriteAllText(Path.Combine(dir + dirName, fileName), xml);
            }
        }
        //-read xml file.
        public string ReadXmlFile(string fileName,string dirName)
        {
            XDocument arr = XDocument.Load(dir + dirName + "\\"+ fileName);
            string[] connectionString = arr.Descendants("connectionString").Select(t => t.Value).ToArray();
            string[] tableName = arr.Descendants("tableName").Select(t => t.Value).ToArray();
            string[] emailUrl = arr.Descendants("emailUrl").Select(t => t.Value).ToArray();
            string[] address = arr.Descendants("address").Select(t => t.Value).ToArray();

            return connectionString[0]+"#"+tableName[0]+"#"+ emailUrl[0]+"#"+ address[0];
        }
        //-create a file.
        public void CreateSystemFile(String fileName, string dirName) {
            if (!Directory.Exists(dir + dirName))
            {
                Directory.CreateDirectory(dir + dirName);
                File.WriteAllText(Path.Combine(dir + dirName, systemFile), fileName);
            }
        }
    }
}
