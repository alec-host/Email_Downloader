using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace NationKeEmailDownload
{
    public class Db
    {
        private readonly string connectionString = null;
        private readonly string tableName = null;
        private SqlConnection conn;

        private SqlCommand cmd;
        private string sql = null;

        public Db(string connectionString, string tableName)
        {
            this.connectionString = connectionString;
            this.tableName = tableName;
        }
        //-.method to download emails & write to db.
        public void SaveData(string message_id,string from,string to, string cc, string subject, string body, DateTime recieve_date ,DateTime datetime, int flag)
        {
            conn = new SqlConnection(connectionString);

            NetworkConnection network = new NetworkConnection();

            Boolean InternetConnection = network.IsInternetAvailable();

            //-.check whether we have internet connection.
            if (InternetConnection)
            {
                try
                {
                    conn.Open();
                    sql = 
                        " IF NOT EXISTS (SELECT 1 FROM " + tableName + " WHERE message_id = @message_id) " +
                        " INSERT " +
                        " INTO "
                        + tableName +
                        " (message_id,_from,_to,_cc,subject,body,receive_time,date_created,has_attachment) " +
                        " VALUES " +
                        " (@message_id,@from,@to,@cc,@subject,@body,@receive_time,@date_created,@flag) ";

                    cmd = new SqlCommand(sql, conn);

                    cmd.Parameters.AddWithValue("@message_id", message_id.Trim());
                    cmd.Parameters.AddWithValue("@from", from.Trim());
                    cmd.Parameters.AddWithValue("@to", to.Trim());
                    cmd.Parameters.AddWithValue("@cc", cc.Trim());
                    cmd.Parameters.AddWithValue("@subject", subject.Trim());
                    cmd.Parameters.AddWithValue("@body", body.Trim());
                    cmd.Parameters.AddWithValue("@receive_time", recieve_date);
                    cmd.Parameters.AddWithValue("@date_created", datetime);
                    cmd.Parameters.AddWithValue("@flag", flag);

                    cmd.ExecuteNonQuery();
                    cmd.Dispose();

                    conn.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine(System.DateTime.Now + " No db connection " + e.Message);
                }
            }
            else {
                Console.WriteLine(System.DateTime.Now + " No Internet Connection");
            }
        }
    }
}
