using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// System.Data.OleDb необходимо для использования OleDbConnection
using System.Data.OleDb;

namespace MSAccessClient
{
    class DBCon
    {
        private OleDbConnection oledbcon = null;
        public bool opened = false;
        public string errormessage = "---";

        public DBCon(string connectionstring)
        {
            oledbcon = new OleDbConnection(connectionstring);
        }
        public void openConnection()
        { 
            try
            {
                if(!opened)
                {
                    oledbcon.Open();
                }                                             
                //Console.WriteLine("ServerVersion: {0} \nDataSource: {1}",oledbcon.ServerVersion, oledbcon.DataSource);
                opened = true;               
            }
            catch (Exception ex)
            {
                errormessage = ex.Message + "   " + ex.StackTrace;
                Console.WriteLine("Исключение в DBCon.openConnection(): " + ex.Message);
            }            
        }
        public void closeConnection()
        {
            try
            {
                if(opened)
                {
                    oledbcon.Close();
                }
                opened = false;
            }
            catch (Exception ex)
            {
                errormessage = ex.Message + "   " + ex.StackTrace;
                Console.WriteLine("Исключение в DBCon.closeConnection(): " + ex.Message);
            }
        }
        public OleDbConnection getConnection()
        {
            return oledbcon;
        }
    }
}
