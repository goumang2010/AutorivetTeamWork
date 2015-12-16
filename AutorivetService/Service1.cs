using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using mysqlsolution;

namespace AutorivetService
{
    public partial class Service1 : ServiceBase
    {

      SocketServerService localServer = new SocketServerService();
        public Service1()
        {
            InitializeComponent();
            System.Timers.Timer t = new System.Timers.Timer();
            t.Interval = 1000 * 15;
            t.Elapsed += new System.Timers.ElapsedEventHandler(RunWork);
            t.AutoReset = true;//设置是执行一次（false),还是一直执行（true);
            t.Enabled = true;//是否执行
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                localServer.btnBeginListen();
            }
            catch (Exception ee)
            {

            }

        }

        protected override void OnStop()
        {

        }
        public void RunWork(object source, System.Timers.ElapsedEventArgs e)
        {
            //
        }
    }
}
