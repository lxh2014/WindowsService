using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CommonSchedule
{
    public class FileOpetation
    {
        /// <summary>
        /// 保存至本地文件
        /// </summary>
        /// <param name="ETMID"></param>
        /// <param name="content"></param>
        public static void SaveRecord(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return;
            }

            FileStream fileStream = null;

            StreamWriter streamWriter = null;

            try
            {
                string path = Path.Combine(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase, string.Format("log"));
                if (!Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                path = Path.Combine(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase, string.Format("log/" + "{0:yyyyMM}" , DateTime.Now));
                if (!Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                path = Path.Combine(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase, string.Format("log/" + "{0:yyyyMM}/"+"{0:yyyyMMdd}"+".txt", DateTime.Now));
                //path = Path.Combine(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase, string.Format("log/"+"{0:yyyyMMdd}"+".txt", DateTime.Now));

                using (fileStream = new FileStream(path, FileMode.Append, FileAccess.Write))
                {
                    using (streamWriter = new StreamWriter(fileStream))
                    {
                        streamWriter.Write(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+": ");
                        streamWriter.Write(content);
                        streamWriter.Write("\r\n");
                        if (streamWriter != null)
                        {
                            streamWriter.Close();
                        }
                    }

                    if (fileStream != null)
                    {
                        fileStream.Close();
                    }
                }
            }
            catch { }
        }
    }
}
