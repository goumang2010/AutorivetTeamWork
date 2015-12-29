using GoumangToolKit;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;


namespace AUTORIVET_KAOHE
{
    public partial class database_management : Form
    {
        public database_management()
        {
            InitializeComponent();
        }

        #region creat_table
        private void creatProductTableToolStripMenuItem_Click_1(object sender, EventArgs e)
        {

                AutorivetDB.product_table();


        }

        private void createArrayTableToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            AutorivetDB.task_table();
          
        }

        private void creatToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
          //  DbHelperSQL.ExecuteSql("Create table if not exists 任务模板(责任人 varchar(50) ,任务名称 varchar(100) NOT NULL PRIMARY KEY,任务类型 varchar(100),任务说明 longtext,绩效奖励 double);");
            AutorivetDB.taskmodel_table();
            
            Task_model f1 = new Task_model();
            f1.Show();
        }

        private void createMessageTableToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            AutorivetDB.pushinfo_table();

        }

        private void createToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            AutorivetDB.rnc_table();
        }

        private void createProductionTableToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            AutorivetDB.production_table();

        }

        private void creatPaperworkTableToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            AutorivetDB.paperwork_table();
        }

        private void creatCouponTableToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            AutorivetDB.coupon_table();
        }
        private void creatPartTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.part_table();
        }
        #endregion

        private void createCATIAMODELTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
           AutorivetDB.catiaModel_table();
        }

        private void createEveryDayTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.everyday_table();

        }

        private void createPeopleTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.people_table();
        }

        private void createMaterialTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.material_table();
        }

        private void createLogTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.log_table();

        }

        private void button1_Click(object sender, EventArgs e)
        {
           FormMethod.cleanbackup();
        }

        private void tVADatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {


     
        }

        private void fastenertableToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sqlstr = "select * from "+listBox1.SelectedItem.ToString();
            OFFICE_Method.excelMethod.SaveDataTableToExcel(DbHelperSQL.Query(sqlstr).Tables[0]);
        }

        private void fastenerTableToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            AutorivetDB.ini_fasttable();
        }

        private void createToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            AutorivetDB.tool_table();
        }

        private void createTroubleshootTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.trouble_table();

        }

        private void toolsTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.updateToolsTable();
        }

        private void iniProcessTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var lst = from DataRow pp in Program.prodTable.AsEnumerable()
                      let tablename = pp["图号"].ToString().Replace("-001", "process")
                      select "Create table if not exists " + tablename + "(ID INT(20), UUID varchar(30) UNIQUE PRIMARY KEY, 加工位置location varchar(100), 紧固件名称Fastener_Name varchar(100), 紧固件数量Fastener_Qty varchar(50), 钻头Drill varchar(100), 下铆头Lower_Anvil varchar(100), 上铆头Upper_Anvil varchar(100), 胶嘴Sealant_Tip varchar(100), 试片Coupon_used varchar(50), 参数号Process_NO varchar(50), 锪窝深度Countersink_depth varchar(50), 钻头转速Speed_of_drill varchar(50), 给进速率Feed_speed varchar(50), 夹紧力Clamp_force varchar(50), 夹紧释放力Clamp_relief_force varchar(50), 墩铆力Upset_force varchar(20), 墩铆位置Upset_position varchar(20), 注胶压力Seal_pres varchar(50), 注胶时间Seal_time varchar(50), 试片紧固件齐平度 double, 产品紧固件齐平度 double, 程序Program longtext);" +
                       "Create table if not exists " + tablename + "_backup(ID INT(20), UUID varchar(30) UNIQUE PRIMARY KEY, 加工位置location varchar(100), 紧固件名称Fastener_Name varchar(100), 紧固件数量Fastener_Qty varchar(50), 钻头Drill varchar(100), 下铆头Lower_Anvil varchar(100), 上铆头Upper_Anvil varchar(100), 胶嘴Sealant_Tip varchar(100), 试片Coupon_used varchar(50), 参数号Process_NO varchar(50), 锪窝深度Countersink_depth varchar(50), 钻头转速Speed_of_drill varchar(50), 给进速率Feed_speed varchar(50), 夹紧力Clamp_force varchar(50), 夹紧释放力Clamp_relief_force varchar(50), 墩铆力Upset_force varchar(20), 墩铆位置Upset_position varchar(20), 注胶压力Seal_pres varchar(50), 注胶时间Seal_time varchar(50), 试片紧固件齐平度 double, 产品紧固件齐平度 double, 程序Program longtext);";

                      



          var count=  DbHelperSQL.ExecuteSqlTran(lst);

            MessageBox.Show("成功的更新了" + count.ToString() + "条记录");


        }

        private void renumCouponsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach(DataRow dr in Program.prodTable.Rows)
            {
                AutorivetDB.rfcouponno(dr["图号"].ToString());

            }

            MessageBox.Show("执行成功！");
            



        }

        private void importPeopleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.updatePeopleTable();
        }

        private void createMBOMTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.EBOM_table();

        }

        private void eBOMTableToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            var dt = AutorivetDB.loadfromExcel();


            var d1 = from DataRow pp in dt.Rows
                     let partnum = pp["零件号"].ToString()
                     let dwg = partnum.Split('-')[0]
                     group pp by dwg into jj
                     select new
                     {
                         dwg = jj.Key,
                         ptls = from hh in jj
                                group hh by hh["零件号"].ToString() into qq
                                select new
                                {
                                    partnum = qq.Key,
                                    qty = qq.Sum(r => System.Convert.ToInt32(r["供件数量"].ToString()))
                                }

                     };
            

            DataTable gg = new DataTable();
            gg.Columns.Add("图号", typeof(string));
            gg.Columns.Add("零件号", typeof(string));
            gg.Columns.Add("数量", typeof(string));

            foreach (var bb in d1)
            {
                var dr = gg.Rows.Add();
                dr["图号"] = bb.dwg;
                string prt="";
                string prtqty = "";
                foreach (var cc in bb.ptls)
                {
                    prt = prt + cc.partnum.Replace(bb.dwg, "")+",";
                    prtqty = prtqty + cc.qty.ToString() + ",";
                
                }
                if (prt.Count()>0)
                    {
                    prt= prt.Remove(prt.Count() - 1);
                     }
                if (prtqty.Count() > 0)
                {
                    prtqty= prtqty.Remove(prtqty.Count() - 1);
                }
                dr["零件号"] = prt;
                dr["数量"] = prtqty;

            }

     
            OFFICE_Method.excelMethod.SaveDataTableToExcel(gg,iftext:true);




        }
    }
}
