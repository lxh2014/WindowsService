using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Windows.Forms;
using System.ServiceProcess;

namespace MyInstallFirstWindowsService
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Service is uninstallging... ");
            Application.EnableVisualStyles();

            Application.SetCompatibleTextRenderingDefault(false);

            string sysDisk = System.Environment.SystemDirectory.Substring(0,3);

            string dotNetPath = sysDisk + @"WINDOWS\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe";//因为当前用的是4.0的环境

            string serviceEXEPath = Application.StartupPath + @"\gigadeWorkerSerives.exe";//把服务的exe程序拷贝到了当前运行目录下，所以用此路径

            string serviceName = "gigadeWorkerSerives";

            string serviceInstallCommand = string.Format(@"{0}  {1}", dotNetPath, serviceEXEPath);//安装服务时使用的dos命令

            string serviceUninstallCommand = string.Format(@"{0} -U {1}", dotNetPath, serviceEXEPath);//卸载服务时使用的dos命令
            //如果已經安裝Install，則先卸載；
            try
            {
                if (File.Exists(dotNetPath))
                {
                    Console.WriteLine("File :" + dotNetPath + " is exists ");
                }
                else
                {
                    Console.WriteLine("File :" + dotNetPath + " is not exists ");
                    Console.WriteLine("uninstall is failed !");
                    Thread.Sleep(3000);
                    return;
                }
                if (File.Exists(serviceEXEPath))
                {
                    Console.WriteLine("File :" + serviceEXEPath + " is removing... ");
                    string[] cmd = new string[] { serviceUninstallCommand };
                    string ss = Cmd(cmd);
                    CloseProcess("cmd.exe");
                    Console.WriteLine("uninstall is successed !");
                    Thread.Sleep(3000);
                }
                else
                {
                    Console.WriteLine("File :" + serviceEXEPath + "is not exists");
                    Console.WriteLine("uninstall is failed !");
                    Thread.Sleep(3000);
                    return;
                }
                

            }
            catch
            {
                Console.WriteLine("uninstall is failed !");
                Thread.Sleep(3000);
                return;
            }
            
            
        }

        /// <summary>
        /// 运行CMD命令
        /// </summary>
        /// <param name="cmd">命令</param>
        /// <returns></returns>
        public static string Cmd(string[] cmd)
        {
            string strRst=string.Empty;
            try
            {
                Process p = new Process();
                p.StartInfo.FileName = "cmd.exe";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardInput = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.CreateNoWindow = true;
                p.Start();
                p.StandardInput.AutoFlush = true;
                for (int i = 0; i < cmd.Length; i++)
                {
                    p.StandardInput.WriteLine(cmd[i].ToString());
                }
                p.StandardInput.WriteLine("exit");
                strRst = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
                p.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("You do not have permission to access this file. Please retry as a administrator !!!");
                Thread.Sleep(3000);
            }
                       
           
            return strRst;
        }

        /// <summary>
        /// 关闭进程
        /// </summary>
        /// <param name="ProcName">进程名称</param>
        /// <returns></returns>
        public static bool CloseProcess(string ProcName)
        {
            bool result = false;
            System.Collections.ArrayList procList = new System.Collections.ArrayList();
            string tempName = "";
            int begpos;
            int endpos;
            foreach (System.Diagnostics.Process thisProc in System.Diagnostics.Process.GetProcesses())
            {
                tempName = thisProc.ToString();
                begpos = tempName.IndexOf("(") + 1;
                endpos = tempName.IndexOf(")");
                tempName = tempName.Substring(begpos, endpos - begpos);
                procList.Add(tempName);
                if (tempName == ProcName)
                {
                    if (!thisProc.CloseMainWindow())
                        thisProc.Kill(); // 当发送关闭窗口命令无效时强行结束进程
                    result = true;
                }
            }
            return result;
        }
    }
}
