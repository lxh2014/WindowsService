using System.ServiceProcess;

namespace CommonSchedule
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
			{ 
				new gigadeWorkerService() 
			};
            ServiceBase.Run(ServicesToRun);
        }
    }
}
