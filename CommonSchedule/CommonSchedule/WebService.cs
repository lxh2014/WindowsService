using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using DBAccess;
using MyNewClass.Common;

namespace CommonSchedule
{
    public class WebService
    {
        static string connectionString;//獲取鏈接字符串
        static IDBAccess _access;
        public WebService()
        {
            try
            {
                //ConfigHelper fighelper = new ConfigHelper();
                //connectionString = fighelper.GetConnString(); 
                //_access = DBFactory.getDBAccess(DBType.MySql, connectionString);
            }
            catch(Exception ex)
            {
                FileOpetation.SaveRecord(ex.Message.ToString());
            }
        }
        public void GetExeScheduleServiceList()
        {
            string api = System.Configuration.ConfigurationManager.AppSettings["ScheduleApi"];
            HttpWebRequest httpRequest = (HttpWebRequest)WebRequest.Create(api);
            //httpRequest.Timeout = 5000;
            
            httpRequest.Method = "GET";
            HttpWebResponse httpResponse = (HttpWebResponse)httpRequest.GetResponse();
            
            //System.IO.StreamReader sr = new System.IO.StreamReader(httpResponse.GetResponseStream(), System.Text.Encoding.GetEncoding("UTF-8"));
            //string html = sr.ReadToEnd();
            
            httpResponse.Close();
            httpResponse = null;
            FileOpetation.SaveRecord(string.Format(@"WebService-->GetExeScheduleServiceList() 執行正常")); 
        }
    }
}
