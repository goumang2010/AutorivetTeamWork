using System;
using System.Data;
using System.Windows.Forms;
using mysqlsolution;
using OFFICE_Method;
using CATIA_method;




namespace AUTORIVET_KAOHE
{
    public partial class MainTeamForm : Form
    {
        Form3 imp;

        public MainTeamForm()
        {
            InitializeComponent();
        }

        bool socketconnect = false;
     //   frmClient fff;
        private void updateproduct()
        {
            DataTable prod = DbHelperSQL.Query("select concat(产品名称,'_',产品架次) from 产品流水表 where 当前状态='正在进行'").Tables[0];
            if (prod.Rows.Count!=0)
            {
                listBox2.SelectedItem = prod.Rows[0][0].ToString();

                    
            }
            else
            {
                MessageBox.Show("请添加正在进行的产品");
            }
        }
        public string getinfo
        {
            get
            {
                return txtMsg.Text;
            }
            set
            {
                if (value.Contains("紧急"))
                {
                    MessageBox.Show(value);
                }
                txtMsg.Text = value;
            }
        }

        public string pushinfo
        {
            get
            {
                return txtMsg.Text;
            }
            set
            {
                if(socketconnect)
                {
                    Program.fff.sendmsg(value);
                }
                else
                {
                    MessageBox.Show("目前和服务器无连接，请确保关闭了防火墙");
                }
               
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            if (Program.ManagerActived==true)
            {
                ActiveManager();
            }

                

            Program.procDic.Add("FASTENER DRILL BEFORE AUTOMATED FASTENING", "架前钻孔—架后手铆");
            Program.procDic.Add("FASTENER INSTALLED AFTER AUTOMATED FASTENING", "不安装");
            Program.procDic.Add("FASTENER INSTALLED BEFORE AUTOMATED FASTENING", "架前手铆");
            Program.procDic.Add("FASTENER INSTALLED BY AUTOMATED FASTENING", "自动钻铆");
            Program.procDic.Add("RESYNCING AND FASTENER INSTALLED BY AUTOMATED FASTENING - Target type - Fast Tack", "手铆3号钉—自动钻铆");
            Program.procDic.Add("RESYNCING AND FASTENER INSTALLED BY AUTOMATED FASTENING - Target type - Pilot Holes", "透制导孔—自动钻铆");
            Program.procDic.Add("RESYNCING ONLY BY AUTOMATED FASTENING - Target type - Final", "手铆终钉");
            Program.procDic.Add("RESYNCING ONLY BY AUTOMATED FASTENING - Target type - Fast Tack", "手铆3号钉—补铆终钉");
            Program.procDic.Add("RESYNCING ONLY BY AUTOMATED FASTENING - Target type - Pilot Holes", "透制导孔—补铆终钉");
            Program.procDic.Add("DRILL ONLY BY AUTOMATED FASTENING", "自动钻铆仅钻孔");
            Program.procDic.Add("RESYNCING AND DRILL ONLY BY AUTOMATED FASTENING - Target type - Fast Tack", "手铆3号钉—自动钻铆钻孔");
            Program.procDic.Add("RESYNCING AND DRILL ONLY BY AUTOMATED FASTENING - Target type - Pilot Holes", "透制导孔—自动钻铆钻孔");


            //imp = (Form3)FormMethod.GetForm("Form3");
      

            //imp.Show();
            //imp.Visible = false;
            Program.fff = new frmClient(this);
            Program.fff.Show();
            Program.fff.Visible = false;
            socketconnect = Program.fff.connectserver();

            comboBox1.DataSource = DbHelperSQL.getlist("select NAME from people where PRIVILEGE=1 group by NAME");
            comboBox4.DataSource= DbHelperSQL.getlist("select NAME from people where PRIVILEGE<>1 group by NAME");
            listBox2.DataSource = DbHelperSQL.getlist("select concat(产品名称,'_',产品架次) from 产品流水表");






            comboBox3.Visible = false;
            //DbHelperSQL.ExecuteSql("Create table if not exists 任务管理(流水号 int(5) NOT NULL AUTO_INCREMENT PRIMARY KEY,责任人 varchar(50) ,生成时间 timestamp NOT NULL default CURRENT_TIMESTAMP,任务名称 varchar(100),任务类型 varchar(100),任务说明 longtext,节点日期 varchar(100),完成日期 varchar(100),任务状态 varchar(100),绩效奖励 double,额外奖励 double);");
         try
         {
                listBox1.DataSource = DbHelperSQL.getlist("select 任务名称 from 任务模板");
              //  updateproduct();
                dataviewfresh();
           

     
  DataGridViewButtonColumn col_btn_insert = new DataGridViewButtonColumn();
    col_btn_insert.HeaderText = "操作";
   col_btn_insert.Text = "完成";
col_btn_insert.UseColumnTextForButtonValue = true;

DataGridViewCheckBoxColumn col_btn_insert5 = new DataGridViewCheckBoxColumn();
col_btn_insert5.HeaderText = "重要";
col_btn_insert5.Name = "重要";





dataGridView2.Columns.Add(col_btn_insert);
dataGridView2.Columns.Add(col_btn_insert5);
dataviewfresh2();
             }
         catch
           {
             MessageBox.Show("请检查数据表是否存在！");
        }
            
 
       
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(comboBox1.Text==""||comboBox2.Text==""||textBox1.Text=="")
            {
                MessageBox.Show("请检查是否选择责任人、任务类型，填写了任务名称，选择了关联产品");
            }
            else
            {
                try
                {


                    DbHelperSQL.ExecuteSql("insert into 任务管理(责任人,任务类型,任务名称,任务说明,绩效奖励,任务状态,节点日期,关联产品) value('" + comboBox1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "','" + textBox2.Text + "'," + System.Convert.ToDouble(label6.Text) + ",'待审核','" + dateTimePicker1.Value.ToShortDateString() + "','" + listBox2.SelectedItem.ToString()+ "')");

                    dataviewfresh();
                }

                catch
                {
                    MessageBox.Show("网络出现问题，请稍后提交");
                }

            }

        }

      public void dataviewfresh()
        {
            dataGridView1.DataSource = DbHelperSQL.Query("select 责任人,任务名称,任务类型,任务说明,生成时间,节点日期,流水号 from 任务管理 where 任务状态='待审核' order by 流水号").Tables[0];
            dataGridView1.Columns[3].Width = 170;
            dataGridView1.Columns[0].Width = 55;
            dataGridView1.Columns[2].Width = 55;
            dataGridView1.Columns[6].Width = 55;
   



        }
       public void dataviewfresh2()
        {

            updateproduct();
           if(checkBox2.Checked==true)
           {
               dataGridView2.DataSource = DbHelperSQL.Query("select 责任人,任务名称,任务类型,任务说明,节点日期,流水号 from 任务管理 where (任务状态 not like '%完成%' and 任务状态<>'待审核')  order by 责任人,节点日期").Tables[0];
           }
           else
           {
               dataGridView2.DataSource = DbHelperSQL.Query("select 责任人,任务名称,任务类型,任务说明,节点日期,流水号 from 任务管理 where (任务状态 not like '%完成%' and 任务状态<>'待审核') and 责任人<>'李钟宁' and 责任人<>'于明洋' order by 责任人,节点日期").Tables[0];
           }
           
            dataGridView2.Columns[0].Width = 55;
            dataGridView2.Columns["重要"].Width = 50;
            dataGridView2.Columns["任务说明"].Width = 170;
            dataGridView2.Columns["责任人"].Width = 55;
            dataGridView2.Columns["流水号"].Width = 55;
            dataGridView2.Columns["任务类型"].Width = 55;
           if(checkBox1.Checked==true)
           {
               dataGridView3.DataSource = DbHelperSQL.Query("select 责任人,任务名称,节点日期,完成日期,绩效奖励,额外奖励,任务状态,流水号 from 任务管理 where 任务状态 like '%完成%' and TO_DAYS(NOW()) - TO_DAYS(生成时间) <= 30 order by 责任人;").Tables[0];
           }
           else
           {
               dataGridView3.DataSource = DbHelperSQL.Query("select 责任人,任务名称,节点日期,完成日期,绩效奖励,额外奖励,任务状态,流水号 from 任务管理 where 任务状态 like '%完成%' and TO_DAYS(NOW()) - TO_DAYS(生成时间) <= 30 and 任务类型<>'请假' and 任务类型<>'加班' order by 责任人;").Tables[0];
           }
         
            dataGridView3.Columns[0].Width = 50;
            dataGridView3.Columns[1].Width = 100;
            dataGridView3.Columns[2].Width = 80;
            dataGridView3.Columns[3].Width = 80;
            dataGridView3.Columns[4].Width = 40;
            dataGridView3.Columns[5].Width = 40;
            dataGridView3.Columns[6].Width =50;
            dataGridView3.Columns[7].Width = 40;

           
           DataTable tmpinfo=DbHelperSQL.Query("select 信息 from 推送信息 where 信息状态='有效'").Tables[0];
           if(tmpinfo.Rows.Count!=0)
           {
               label14.Text = "推送通知："+tmpinfo.Rows[0][0].ToString();
           }
           


        }
        private void createArrayTableToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void creatToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            string taskname = listBox1.SelectedItem.ToString();
            DataTable tianchong = DbHelperSQL.Query("select 责任人,任务类型,任务名称,任务说明,绩效奖励 from 任务模板 where 任务名称='"+taskname+"'").Tables[0];
            comboBox1.Text = tianchong.Rows[0][0].ToString();
            comboBox2.Text = tianchong.Rows[0][1].ToString();
           textBox1.Text = tianchong.Rows[0][2].ToString();
           textBox2.Text = tianchong.Rows[0][3].ToString();
           label6.Text = tianchong.Rows[0][4].ToString();

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
          //  if(label6.Text=="")
          //  {
            comboBox3.Items.Clear();
            comboBox3.Items.Add("半天");
            comboBox3.Items.Add("一天");
            comboBox3.Items.Add("两天");
            comboBox3.SelectedIndex = 1;
            comboBox3.Visible = false;
                switch (comboBox2.Text)
                {
                    case "请假":                     
                        comboBox3.Visible = true;
                        textBox1.Text = "请假" + comboBox3.Text;

                        comboBox3.Focus();
                        break;
                    case "加班":
                        textBox1.Text = "加班" + comboBox3.Text;
                        comboBox3.Visible = true;
                        comboBox3.Focus();
                        break;
                    case "计划内":
                       // comboBox3.Visible = false;
                        label6.Text = "0.00";
                        textBox2.Text = "填写计划内需完成的内容:";
                        break;
                    case "申请":
                        //comboBox3.Visible = false;
                        label6.Text = "0.00";
                        textBox2.Text = "填写申请需完成的内容:";
                        break;
                    case "迟到":
                        //comboBox3.Visible = false;

                        textBox1.Text = "迟到 " + DateTime.Now.ToShortDateString();
                        label6.Text = "-0.005";
                        textBox2.Text = "";
                        break;
                    case "已完成":
                      //  comboBox3.Visible = false;
                        textBox2.Text = "在此填写完成日期和情况说明";
                       // label6.Text = "0.00";
                        break;

                }
          //  }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "加班" || comboBox2.Text == "请假")
            {

          
            textBox2.Text = "填写加班/请假区间。注：加班时间需大于3小时才可算为加班半天";
            switch (comboBox3.Text)
            {
                case "半天":
                    label6.Text = "0.01";
                    textBox1.Text = comboBox2.Text + "半天";
                    break;
                case "一天":
                    label6.Text = "0.02";
                    textBox1.Text = comboBox2.Text + "一天";
                    break;
                case "两天":
                    label6.Text = "0.03";
                    textBox1.Text = comboBox2.Text + "两天";
                    break;

            }
            if(comboBox2.Text=="请假")
            {
                label6.Text = "-" + label6.Text;
            }
        }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            { 


                            int sum = dataGridView1.SelectedRows.Count;
            for(int i=0;i<sum;i++)
            {
                DataRow oprow = ((DataRowView)dataGridView1.SelectedRows[i].DataBoundItem).Row;

                DbHelperSQL.ExecuteSql("delete from 任务管理 where 流水号='" + oprow[6] + "'");
            }
             
            dataviewfresh();
                }
            catch
            {
                MessageBox.Show("请先选中一行！");
            }
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
          //  dataviewfresh();
            rfall();
        }

        private void rfall()
        {
            try
            {
                if ((!socketconnect)||(getinfo.Contains("异常")))
            {
               socketconnect= Program.fff.connectserver();
            }
              
                dataviewfresh();
                dataviewfresh2();

                label13.Text = "通讯状态:正常！更新时间：" + DateTime.Now.ToShortTimeString();

            }
                catch
                {
                    label13.Text = "通讯状态:断网啦！";

                }
        }


        private void button3_Click(object sender, EventArgs e)
        {

            int sum = dataGridView1.SelectedRows.Count;
            for(int i=0;i<sum;i++)
            {
                DataRow oprow = ((DataRowView)dataGridView1.SelectedRows[i].DataBoundItem).Row;
                DbHelperSQL.ExecuteSql("update 任务管理 set 任务状态='待解决' where 流水号='" + oprow[6] + "'");
            }
          

            dataviewfresh();
            dataviewfresh2();
            
             
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "" || comboBox2.Text == "" || textBox1.Text == "" )
            {
                MessageBox.Show("请检查是否选择责任人、任务类型，填写了任务名称，选择了关联产品");
            }
            else
            {
                DbHelperSQL.ExecuteSql("insert into 任务管理(责任人,任务类型,任务名称,任务说明,绩效奖励,任务状态,节点日期,关联产品) value('" + comboBox1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "','" + textBox2.Text + "'," + System.Convert.ToDouble(label6.Text) + ",'待解决','" + dateTimePicker1.Value.ToShortDateString() + "','" + label15.Text + "')");

                dataviewfresh2();
            }
        }
        public void ActiveManager()
        {
            button3.Visible = true;
            button4.Visible = true;
            button5.Visible=true;
            button6.Visible = true;
            button7.Visible = true;
            button8.Visible = true;
            button10.Visible = true;
            comboBox4.Visible = true;
            button16.Visible = true;
           // Program.ManagerActived = true;
            
           // checkBox2.Checked = true;

        }
        private void activateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            user_key f = new user_key(this);
            this.Hide();
            f.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            int sum = dataGridView2.SelectedRows.Count;
            for (int i = 0; i < sum; i++)
            {
                DataRow oprow = ((DataRowView)dataGridView2.SelectedRows[i].DataBoundItem).Row;

                DbHelperSQL.ExecuteSql("delete from 任务管理 where 流水号='" + oprow[5] + "'");
            }


            dataviewfresh2();
        }

        private void generateToolStripMenuItem_Click(object sender, EventArgs e)
        {

           DataTable tianchong = new DataTable();

           // List<string> tianchongname = new List<string>();


            tianchong=DbHelperSQL.Query("select * from 任务管理").Tables[0];
         //   tianchongname.Add("任务总表");



            excelMethod.SaveDataTableToExcel(tianchong);
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int k = e.ColumnIndex;
            int l = e.RowIndex;
            DataRowView temprow=(DataRowView) dataGridView2.Rows[l].DataBoundItem;
           string liushui=temprow[5].ToString();
           string zerenren= temprow[0].ToString();

            switch(k)
            {
                case 0:
                     if(zerenren!="李钟宁"||Program.ManagerActived==true)
             
        
                {
                         //传入0
                    Form2 f = new Form2(this);
                    f.liushui = liushui;
                    f.desc = temprow[3].ToString();
                    f.Show();
                }
                else
                {
                    MessageBox.Show("请仅完成自己的任务！");
                }
             
                    break;


                case 1:

                    DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView2.Rows[l].Cells["重要"];
                          Boolean flag = Convert.ToBoolean(checkCell.Value); 
               if (flag == true)     //查找被选择的数据行 
                {    
                    
                         AutorivetDB.update_taskprop(liushui, "重要");


                     }
               else
               {
                   
                   
                   
                   
                   AutorivetDB.update_taskprop(liushui, "待解决");
               }
                   imp.rfgrid();
              
                    break;




            }




        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           if(textBox1.Text.Contains("编程")|| textBox1.Text.Contains("TVA"))
            {

            }
            else
           {
               switch (comboBox2.Text)
               {

                   case "计划内":
                       // comboBox3.Visible = false;
                       label6.Text = "0.01";
                       break;
                   case "申请":
                       //comboBox3.Visible = false;
                       label6.Text = "0.01";
                       break;
                   case "已完成":
                       //  comboBox3.Visible = false;
                       textBox2.Text = "在此填写完成日期和情况说明";
                       label6.Text = "0.01";
                       break;

               }
           }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" )
            {
                MessageBox.Show("请检查是否选择责任人、任务类型，填写了任务名称，选择了关联产品");
            }
            else
            {
                string zerr=comboBox4.Text;

                DbHelperSQL.ExecuteSql("insert into 任务管理(责任人,任务类型,任务名称,任务说明,绩效奖励,任务状态,节点日期,关联产品) value('" + zerr + "','自理','" + textBox1.Text + "','" + textBox2.Text + "',0,'待解决','" + dateTimePicker1.Value.ToShortDateString() + "','" + label15.Text + "')");

                dataviewfresh2();
            }


        }

        private void button7_Click(object sender, EventArgs e)
        {
            int sum = dataGridView3.SelectedRows.Count;
            for (int i = 0; i < sum; i++)
            {
                DataRow oprow = ((DataRowView)dataGridView3.SelectedRows[i].DataBoundItem).Row;
                DbHelperSQL.ExecuteSql("update 任务管理 set 任务状态='已完成' where 流水号='" + oprow[7] + "'");
            }


            dataviewfresh();
            dataviewfresh2();
        }

        private void rNCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RNC f = new RNC();
            f.Show();

        }

        private void button9_Click(object sender, EventArgs e)
        {
          
         
            try
            {
                DataRowView temprow = (DataRowView)dataGridView2.SelectedRows[0].DataBoundItem;
                Form2 f = new Form2(this, 1);
                string liushui = temprow[5].ToString();
                //处理代码
                f.jiedianstr = temprow["节点日期"].ToString();

                f.liushui = liushui;
                f.desc = temprow[3].ToString();
                

                f.Show();
            }
           catch
            {
                MessageBox.Show("请先选中一行！");

            }





        }

        private void button10_Click(object sender, EventArgs e)
        {

          //  DataRow oprow = ((DataRowView)dataGridView1.SelectedRows[0].DataBoundItem).Row;
            DbHelperSQL.ExecuteSql("update 任务管理 set 任务状态='待解决' where 任务状态='待审核'");
            dataviewfresh();
            dataviewfresh2();
        }

        private void sendMessageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(Program.ManagerActived==true)
            {
                Form2 f = new Form2(this, 2);
                f.Show();

            }
            else
            {
                MessageBox.Show("缺少权限");
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            dataviewfresh2();
        }

        private void createMessageTableToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        private void createToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void createProductionTableToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void programToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void productionToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void peopleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable pr = DbHelperSQL.Query("select 责任人,sum(绩效奖励),sum(额外奖励),sum(绩效奖励+额外奖励) 总成绩 from 任务管理 where 任务状态 like '%完成%' group by 责任人 order by 总成绩 desc").Tables[0];
            DataColumn newColumn = new DataColumn();
            newColumn.DataType = System.Type.GetType("System.Decimal");
            newColumn.ColumnName = "申请K2";
            newColumn.Expression = "0.35/sum(总成绩)*总成绩";

/*
            DataColumn idColumn = new DataColumn();
            idColumn.DataType = System.Type.GetType("System.Int32");
            idColumn.ColumnName = "排名";
            idColumn.AutoIncrement = true;
            idColumn.AutoIncrementSeed = 1;
            idColumn.AutoIncrementStep = 1;

            pr.Columns.Add(idColumn);
 * */
            pr.Columns.Add("排名", Type.GetType("System.Int32"));
           for (int i = 0; i < pr.Rows.Count; i++)  
                {  
                    pr.Rows[i][4] = i + 1;  
            }  

          
            pr.Columns.Add(newColumn);
            var wSheet = excelMethod.SaveDataTableToExcel(pr);
            excelMethod.setNumberFormat(wSheet, 7, "0.00_ ");
            //Microsoft.Office.Interop.Excel.Range allColumn = (Microsoft.Office.Interop.Excel.Range)wSheet.Columns[6];
            //allColumn.NumberFormatLocal = "0.00_ ";

        }

        private void finishTaskToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Program.ManagerActived == true)
            {
                finish_task f = new finish_task();
                f.Show();

            }
            else
            {
                MessageBox.Show("缺少权限");
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            dataviewfresh2();

        }

        private void aAOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Program.ManagerActived == true)
            {
                AO f = new AO();
                f.Show();

            }
            else
            {
                MessageBox.Show("缺少权限");
            }
        }

        private void aLLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Production f = new Production();
            f.Show();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DataTable output = (DataTable)dataGridView2.DataSource;
            excelMethod.SaveDataTableToExcel(output);
           
        }

        private void allToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ProductAll ppallf = new ProductAll();
            ppallf.Show();
        }

        private void sortToolStripMenuItem_Click(object sender, EventArgs e)
        {
            paperWork f = new paperWork();
            f.Show();
        }

        private void creatPaperworkTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
   
           
        }

        private void creatProductTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void creatCouponTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
         
        }

        private void msgServerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Program.ManagerActived == true)
            {

                frm_server f = new frm_server();
                f.Show();


            }
            else
            {
                MessageBox.Show("该操作需要权限");
            }
        }



        private void btnSend_Click(object sender, EventArgs e)
        {
            Program.fff.sendmsg(txtSendMsg.Text);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Program.fff.Visible = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            rfall();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            DataTable output = (DataTable)dataGridView3.DataSource;
            excelMethod.SaveDataTableToExcel(output);
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void datebaseToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (Program.ManagerActived == true)
            {

                database_management f = new database_management();
                f.Show();


            }
            else
            {
                MessageBox.Show("该操作需要权限");
            }
        }

        private void partListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Part_list f = new Part_list();
            f.Show();

        }

        private void button15_Click(object sender, EventArgs e)
        {

            excelMethod.SaveDataTableToExcel(DbHelperSQL.Query("select * from 任务管理 where 任务类型 <> '自理' and TO_DAYS(NOW()) - TO_DAYS(生成时间) <= 7 order by 节点日期;").Tables[0]);

        }

        private void txtMsg_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
          //  dataGridView2.CommitEdit();
            if (dataGridView2.IsCurrentCellDirty)
            {
                dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void modelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CATIA_model f = new CATIA_model();
            f.Show();

        }

        private void makePlanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            planMaker f = new planMaker();
            f.Show();
        }

        private void materialToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Material f = new Material();
            f.Show();
        }

        private void tVAToolsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainTools f = new MainTools();
            f.Show();
        }

        private void killWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int k = FormMethod.killProcess("WINWORD");

            MessageBox.Show("执行成功，杀死" + k.ToString() + "个word进程");
        }

        private void killExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
                int k = FormMethod.killProcess("EXCEL");

            MessageBox.Show("执行成功，杀死" + k.ToString() + "个Excel进程");
        }

        private void fileManagerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileManagerNew.MainPage f = new FileManagerNew.MainPage();
            f.Show();
        }

        private void randomGENToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RandomGEN f = new RandomGEN();
            f.Show();
        }

        private void allToolToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Tools f = new Tools();
            f.Show();

        }

        private void troubleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Troubles f = new Troubles();
            f.Show();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            DbHelperSQL.ExecuteSql("update 任务管理 set 任务状态='已完成待审核' where (任务状态='待解决' or 任务状态='重要') and 责任人<>'李钟宁'");
            dataviewfresh();
            dataviewfresh2();
        }

        private void fakeCreationToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void couponTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            couponTest f = new couponTest();
            f.Show();
        }

        private void nCToolToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var f1 = new NC_TOOL.program_input();
          f1.inputValue = "";
            f1.Show();
        }

        private void aboutFeedbackToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Feedback f1 = new Feedback();
            f1.Show();
        }

        private void sACISearcherToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SACISearcher.MainSearchForm f = new SACISearcher.MainSearchForm();
            f.Show();
        }
    }
}
