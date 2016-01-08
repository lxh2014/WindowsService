using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CommonSchedule
{
    //partial class Service1 : ServiceBase
    //{
    //    public Service1()
    //    {
    //        InitializeComponent();
    //    }

    //    protected override void OnStart(string[] args)
    //    {
    //        // TODO: 在此处添加代码以启动服务。
    //    }

    //    protected override void OnStop()
    //    {
    //        // TODO: 在此处添加代码以执行停止服务所需的关闭操作。
    //    }
    //}
    public partial class gigadeWorkerService : ServiceBase
    {
        System.Threading.Timer recordTimer;

        public gigadeWorkerService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            IntialSaveRecord();
        }

        protected override void OnStop()
        {
            if (recordTimer != null)
            {
                recordTimer.Dispose();
            }
        }

        private void IntialSaveRecord()
        {
            TimerCallback timerCallback = new TimerCallback(CallbackTask);

            AutoResetEvent autoEvent = new AutoResetEvent(false);

            recordTimer = new System.Threading.Timer(timerCallback, autoEvent, 0, 1000 * 60); //60s遍歷一次
        }

        private void CallbackTask(Object stateInfo)
        {
            WebService webservice = new WebService();
            webservice.GetExeScheduleServiceList();
            //FileOpetation.SaveRecord(string.Format(@"当前记录时间：{0},状况：程序运行正常！", DateTime.Now));
        }
    }
}
