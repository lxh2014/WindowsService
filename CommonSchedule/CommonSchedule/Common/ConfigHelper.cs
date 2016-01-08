using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Configuration;
namespace MyNewClass.Common
{
    public class ConfigHelper
    {
        public string Connstring = String.Empty;
        XmlDocument webConfigDoc = new XmlDocument();
        public string xmlFilePath = String.Empty; //web.config文件的存储路径
        
        public ConfigHelper()
        {
            xmlFilePath = ConfigurationManager.AppSettings["configFilePath"];
            //if (string.IsNullOrEmpty(xmlFilePath))
            //{
            //    Console.WriteLine("配置節點configFilePath內不存在,或內沒有值!");
            //    Console.ReadKey();
            //}
        }
        public string GetConnString()
        {
            string value = GetValueByKey("MySqlConnectionString");
            //if (string.IsNullOrEmpty(value))
            //{
            //    Console.WriteLine("配置節點中MySqlConnectionString不存在,或內沒有值!");
            //    Console.ReadKey();
            //}
            return value;
        }
        public string GetValueByKey(string key)
        {
            webConfigDoc.Load(xmlFilePath);
            string keyPath = "/configuration/appSettings/add[@key='?']";
            //将web.config文件加载到XmlDocument中
            
            XmlNode updateKey = webConfigDoc.SelectSingleNode((keyPath.Replace("?", key)));
            string value = updateKey.Attributes["value"].InnerText;
            
            //通过SelectSingleNode方法获得当前节点下的courses子节点

            return value;
        }
    }
}
