using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace NationKeEmailDownload
{
    class NetworkConnection
    {
        //-.method to check internet connection.
        public bool IsInternetAvailable()
        {
            var request = (HttpWebRequest)WebRequest.Create("http://g.cn/generate_204");
            request.UserAgent = "Android";
            request.KeepAlive = false;
            request.Timeout = 1500;

            try
            {
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.ContentLength == 0 && response.StatusCode == HttpStatusCode.NoContent)
                    {
                        //Connection to internet available
                        return true;
                    }
                    else
                    {
                        //Connection to internet not available
                        return false;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(System.DateTime.Now + " No Internet Connection "+e.Message);

                return false;
            }
        }
    }
}
