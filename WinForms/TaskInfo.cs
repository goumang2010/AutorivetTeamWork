using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using mysqlsolution;


namespace AUTORIVET_KAOHE
{
    public partial class Form2 : Form
    {
        public string liushui;
        public string jiedianstr;
        public string desc;
        int stateform;
        MainTeamForm f;
        public Form2( MainTeamForm ftemp,int state=0)
        {
            InitializeComponent();
            f = ftemp;
            stateform = state;
            switch (state)
            {
                case 1:

                 if (Program.ManagerActived)
                 {
                     label1.Text = "请修改节点日期";
                     dateTimePicker1.Value = DateTime.Parse(jiedianstr);
                 }
                 else
                 {
                     label1.Visible = false;
                     dateTimePicker1.Visible = false;
                 }
       
               
                    break;
                case 2:
                label1.Visible = false;
                dateTimePicker1.Visible = false;
                label4.Text = "请填写信息";
                comboBox1.Visible = true;

                    break;

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            
           // DataRow oprow = ((DataRowView)dataGridView1.SelectedRows[0].DataBoundItem).Row;

            switch (stateform)
            {
                    //正常完成任务
                case 0:
                    
                    DbHelperSQL.ExecuteSql("update 任务管理 set 任务状态='已完成待审核',完成日期='" + DateTime.Now.Date.ToShortDateString() + "',任务说明='" + desc + "' where 流水号='" + liushui + "'");
                    Form3 f3 = (Form3)FormMethod.GetForm("Form3");
                    f3.rfgrid();
                    break;

                  

                case 1:
                    DbHelperSQL.ExecuteSql("update 任务管理 set 任务说明='" + textBox2.Text + "' where 流水号='" + liushui + "'");

                    if(Program.ManagerActived)
                    {
                        DbHelperSQL.ExecuteSql("update 任务管理 set 节点日期='" + dateTimePicker1.Value.ToShortDateString() + "' where 流水号='" + liushui + "'");

                    }


                    break;
                case 2:
                    DbHelperSQL.ExecuteSql("update 推送信息 set 信息状态='过期' where 信息状态='有效'");
                    DbHelperSQL.ExecuteSql("Insert into 推送信息(责任人,信息状态,信息) Values('" + comboBox1.Text + "','有效','" + textBox2.Text + "');");
                    break;


            }





            f.dataviewfresh2();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

            textBox2.Text = desc;
            
        }
    }
}
