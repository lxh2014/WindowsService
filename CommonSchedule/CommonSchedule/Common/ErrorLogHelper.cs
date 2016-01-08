using EmsDepRelation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyNewClass.Common
{
    
    public class ErrorLogHelper
    {
        public ErrorLogHelper(string content)
        {
            string EmailTitle = string.Empty;
            string GroupCode = string.Empty;
            try
            {
                EmailTitle = System.Configuration.ConfigurationManager.AppSettings["EmailTitle"].ToString();
                GroupCode = System.Configuration.ConfigurationManager.AppSettings["ErrorLogCode"].ToString();
                MailHelper helper = new MailHelper();
                helper.SendToGroup(GroupCode, EmailTitle, content, true, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine("提示錯誤(判斷是否存在錯誤,數據是否異常):");
                Console.WriteLine("EmailTitle值為:"+EmailTitle);
                Console.WriteLine("GroupCode值為:" + GroupCode);
                Console.WriteLine("系統提示錯誤:");
                Console.WriteLine(ex);
                Console.ReadKey();
            }
        }
    }
}
