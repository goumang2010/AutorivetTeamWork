using GoumangToolKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronPython;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace AUTORIVET_KAOHE
{
    static class Program
    {
        public static bool ManagerActived = false;
        private static string userid;
       // private static string infopath;
        public static System.Data.DataTable prodTable;
        public static frmClient fff;
        public static Dictionary<string, string> procDic = new Dictionary<string, string>();
        // public static PRODLIST prodTable;

            public static string userID
        {
            get
            {
                return userid;
            }
            set
            {
                
                var ftchID = DbHelperSQL.getlist("select PRIVILEGE from People where ID='" + value + "';");
                if (ftchID.Count() > 0)
                {
                    if (ftchID.First() == "0")
                    {
                        Program.ManagerActived = true;
                        //MainTeamForm f = (MainTeamForm)FormMethod.GetForm("Form1");
                        //f.ActiveManager();

                    }
                }
                // FormMethod.get_Form("")
             
                userid = value;

            }
        }

        public static string InfoPath
        {
          
            get {

                ScriptEngine engine = Python.CreateEngine();
                ScriptScope scope = engine.CreateScope();
                engine.ExecuteFile(@"\\192.168.3.32\softwareTools\Autorivet_team_manage\settings\configuration.py", scope);
                return scope.GetVariable("InfoPath");
               

            }
            set
            {

                ScriptEngine engine = Python.CreateEngine();
                ScriptScope scope = engine.CreateScope();
                engine.ExecuteFile(@"\\192.168.3.32\softwareTools\Autorivet_team_manage\settings\configuration.py", scope);
                scope.SetVariable("InfoPath",value);


            }
        }

        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
           
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //向客户机添加凭据
            FormMethod.creatCredential();
            Program.prodTable = AutorivetDB.fullname_table("图号,名称,站位号,程序编号");
            Program.userID = FormMethod.GetMachineName();
            Application.Run(new MainTeamForm());
           //Application.Run(new AO());
        }















    }
}
