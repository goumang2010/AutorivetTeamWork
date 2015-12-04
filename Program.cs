using mysqlsolution;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace AUTORIVET_KAOHE
{
    static class Program
    {
        public static bool ManagerActived = false;
        public static string userid;
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



        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
           
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
       
         Application.Run(new MainTeamForm());
           //Application.Run(new AO());
        }















    }
}
