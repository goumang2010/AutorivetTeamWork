using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using mysqlsolution;
using OFFICE_Method;



namespace AUTORIVET_KAOHE
{
    public partial class finish_task : Form,FormInterface
    {

        //设置全局datatable
        DataTable unidt;
        DataTable finaldt;
        public finish_task()
        {
            InitializeComponent();
        }
       public void rf_gridview(dynamic abc)
        {
           // unidt = abc;

            dataGridView1.DataSource = abc;

        }
       public DataTable get_datatable()
       {
           return (DataTable)dataGridView1.DataSource;

       }
        public void rf_default()
        {
            string fullstr = "select  责任人,任务名称,任务类型,任务说明,关联产品,节点日期,完成日期,绩效奖励,额外奖励,任务状态,流水号 from 任务管理 where TO_DAYS(NOW()) - TO_DAYS(生成时间) <= 30  order by 责任人;";
      
            unidt = DbHelperSQL.Query(fullstr).Tables[0];

            rf_gridview(unidt);
            rf_filter();
        }
         public void rf_filter()
        {
             
                  //如果包含请假及加班
        
         var tmpdt= from s in  unidt.AsEnumerable()
                    where s["任务类型"].ToString() != "请假" && s["任务类型"].ToString() != "加班"
                    select s;
             if (checkBox1.Checked == true)
        {
                 tmpdt=unidt.AsEnumerable();
        }
           
         
            //如果包含未完成任务
           if(checkBox2.Checked==false)
           {
                tmpdt= from s in  tmpdt
                    where s["任务状态"].ToString().Contains("完成")
                    select s;

           }

             //combobox1

             if(comboBox1.Text!="")
             {
                 tmpdt = from s in tmpdt
                         where s["责任人"].ToString().Contains(comboBox1.Text)
                         select s;
             }

             if (comboBox2.Text != "")
             {
                 tmpdt = from s in tmpdt
                         where s["任务类型"].ToString().Contains(comboBox2.Text)
                         select s;
             }


             if (tmpdt.Count()!=0)
             {
                 finaldt = tmpdt.CopyToDataTable();
                 dataGridView1.DataSource = finaldt;
             }
             else
             {
                 dataGridView1.DataSource = null;
             }
   


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            int k = e.ColumnIndex;
            int l = e.RowIndex;
            if (l>=0)
            {

           
            DataRowView temprow = (DataRowView)dataGridView1.Rows[l].DataBoundItem;
            //  string liushui = temprow[5].ToString();


            switch (k)
            {
                case 0:
                    textBox1.Text = temprow[1].ToString();

                    textBox2.Text = temprow[3].ToString();
                    textBox3.Text = temprow[8].ToString();
                    dateTimePicker1.Value = DateTime.Parse(temprow[6].ToString());
                    label7.Text = temprow[10].ToString();
                    break;
                case 1:

                    DbHelperSQL.ExecuteSql("update 任务管理 set 任务状态='待解决' where 流水号=" + temprow[10].ToString());
                    rf_default();

                    break;


            }
            }
        }

        private void finish_task_Load(object sender, EventArgs e)
        {
            var namelist = DbHelperSQL.getlist("select NAME from people group by NAME");
            namelist.Insert(0, "");
            comboBox1.DataSource = namelist;

            DataGridViewButtonColumn col_btn_insert = new DataGridViewButtonColumn();
            col_btn_insert.HeaderText = "操作1";
            col_btn_insert.Text = "编辑";//加上这两个就能显示
            col_btn_insert.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert);
            // dataviewfresh2();
            DataGridViewButtonColumn col_btn_insert3 = new DataGridViewButtonColumn();
            col_btn_insert3.HeaderText = "操作2";
            col_btn_insert3.Text = "打回";//加上这两个就能显示
            col_btn_insert3.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert3);
            rf_default();
          //  rf_filter();
          //  finishrf();

          //  dataGridView1.DataSource = DbHelperSQL.Query("select 责任人,任务名称,任务类型,任务说明,关联产品,节点日期,完成日期,绩效奖励,额外奖励,任务状态,流水号 from 任务管理 where 任务状态 like '%完成%' and TO_DAYS(NOW()) - TO_DAYS(生成时间) <= 30 and 任务类型<>'请假' and 任务类型<>'加班' order by 责任人;").Tables[0];
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            rf_filter();

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            rf_filter();
         
        }
        private void finishrf()
    {
        string fullstr = "select  责任人,任务名称,任务类型,任务说明,关联产品,节点日期,完成日期,绩效奖励,额外奖励,任务状态,流水号 from 任务管理 where TO_DAYS(NOW()) - TO_DAYS(生成时间) <= 30 and 任务状态 like '%完成%' and 任务类型<>'请假' and 任务类型<>'加班' order by 责任人;";
        string actstr =  fullstr;
            //如果包含请假及加班
        if (checkBox1.Checked == true)
        {
            actstr = fullstr.Replace(" and 任务类型<>'请假' and 任务类型<>'加班'", "");

        }
         
            //如果包含未完成任务
           if(checkBox2.Checked==true)
           {
               actstr = actstr.Replace(" and 任务状态 like '%完成%'", "");

           }

           dataGridView1.DataSource = DbHelperSQL.Query(actstr).Tables[0];

    }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            rf_filter();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DbHelperSQL.ExecuteSql("update 任务管理 set 任务名称='" + textBox1.Text + "',任务说明='"+textBox2.Text+"',额外奖励="+textBox3.Text+",完成日期='"+dateTimePicker1.Value.ToShortDateString()+"' where 流水号='" + label7.Text+"'");
            rf_default();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DbHelperSQL.ExecuteSql("update 任务管理 set 任务状态='已完成' where 流水号='" + label7.Text + "'");
            rf_default();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataTable punish = DbHelperSQL.Query("select 节点日期,流水号 from 任务管理 where 任务类型='请假' and 额外奖励=0").Tables[0];
            DataTable productdata = DbHelperSQL.Query("select 开始日期,结束日期,产品名称,产品架次 from 产品流水表 where 开始日期<>'' and 结束日期<>'' and 结束日期<>'0000-00-00'").Tables[0];

            foreach(DataRow task in punish.Rows)
            {
                DateTime qingjia = DateTime.Parse(task[0].ToString());
                foreach(DataRow prod in productdata.Rows)
                {
                    string starttimestr= prod[0].ToString();
                    string endtimestr = prod[1].ToString();

                    DateTime startdate = DateTime.Parse(starttimestr);
                    DateTime enddate = DateTime.Parse(endtimestr);
                     if (qingjia.CompareTo(startdate) >= 0 && qingjia.CompareTo(enddate) <= 0)
                    {
                        DbHelperSQL.ExecuteSql("update 任务管理 set 额外奖励=绩效奖励/2 where 流水号='" + task[1].ToString() + "'");
                        break;
                    }

                }

            }
            MessageBox.Show("执行成功");
            rf_default();

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            rf_filter();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataTable output = (DataTable)dataGridView1.DataSource;
            excelMethod.SaveDataTableToExcel(output);

            
            }

        private void button5_Click(object sender, EventArgs e)
        {
             DataTable punish = DbHelperSQL.Query("select * from (select 任务名称,节点日期,count(*) as qty,绩效奖励,责任人 from (select * from 任务管理 where 绩效奖励>0 and 任务类型<>'请假' and 任务类型<>'加班') kk group by 任务名称,责任人,节点日期) aa where aa.qty>1").Tables[0];
     //       DataTable productdata = DbHelperSQL.Query("select 开始日期,结束日期,产品名称,产品架次 from 产品流水表 where 开始日期<>'' and 结束日期<>'' and 结束日期<>'0000-00-00'").Tables[0];
             List<string> sqltran = new List<string>();

            foreach (DataRow task in punish.Rows)
            {
                double jiangli=System.Convert.ToDouble(task[3].ToString());
                string zerenren = task[4].ToString();
                int qty = System.Convert.ToInt16(task[2].ToString());
                double extrajl = jiangli/ qty-jiangli ;
                sqltran.Add("update 任务管理 set 额外奖励=" + extrajl + " where 任务名称='" + task[0].ToString() + "' and 节点日期='" + task[1] + "' and 绩效奖励>0 and 责任人='"+zerenren+"'");
                
           
            }

             DbHelperSQL.ExecuteSqlTran(sqltran);
            MessageBox.Show("执行成功");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DbHelperSQL.ExecuteSql("update (任务管理 a inner join 任务模板 b on a.任务名称=b.任务名称) set a.绩效奖励=b.绩效奖励");
            MessageBox.Show("执行成功");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = unidt.AsEnumerable().Where(p => System.Convert.ToDouble(p["绩效奖励"]) == 0).CopyToDataTable();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DataTable punish = DbHelperSQL.Query("select * from (select 任务名称,节点日期,count(*) as qty,绩效奖励,责任人 from (select * from 任务管理 where 绩效奖励>0 and 任务类型<>'请假' and 任务类型<>'加班') kk group by 任务名称,责任人,节点日期) aa where aa.qty=1").Tables[0];
            //       DataTable productdata = DbHelperSQL.Query("select 开始日期,结束日期,产品名称,产品架次 from 产品流水表 where 开始日期<>'' and 结束日期<>'' and 结束日期<>'0000-00-00'").Tables[0];
            List<string> sqltran = new List<string>();

            foreach (DataRow task in punish.Rows)
            {
                double jiangli = System.Convert.ToDouble(task[3].ToString());
                string zerenren = task[4].ToString();
              //  int qty = System.Convert.ToInt16(task[2].ToString());
               // double extrajl = jiangli / qty - jiangli;
                sqltran.Add("update 任务管理 set 额外奖励=0 where 任务名称='" + task[0].ToString() + "' and 节点日期='" + task[1] + "' and 绩效奖励>0 and 责任人='"+zerenren+"'");


            }

            DbHelperSQL.ExecuteSqlTran(sqltran);
            MessageBox.Show("执行成功");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataTable punish = DbHelperSQL.Query("select * from (select 任务名称,节点日期,count(*) as qty,绩效奖励 from (select * from 任务管理 where 绩效奖励>0 and 任务类型<>'请假' and 任务类型<>'加班' and 节点日期<='2015-6-1') kk group by 任务名称,节点日期) aa where aa.qty>1").Tables[0];
            //       DataTable productdata = DbHelperSQL.Query("select 开始日期,结束日期,产品名称,产品架次 from 产品流水表 where 开始日期<>'' and 结束日期<>'' and 结束日期<>'0000-00-00'").Tables[0];
            List<string> sqltran = new List<string>();

            foreach (DataRow task in punish.Rows)
            {
                double jiangli = System.Convert.ToDouble(task[3].ToString());
             
                int qty = System.Convert.ToInt16(task[2].ToString());
                double extrajl = jiangli / qty - jiangli;
                sqltran.Add("update 任务管理 set 额外奖励=" + extrajl + " where 任务名称='" + task[0].ToString() + "' and 节点日期='" + task[1] + "' and 绩效奖励>0");


            }

            DbHelperSQL.ExecuteSqlTran(sqltran);
            MessageBox.Show("执行成功");
        }
    }
}
