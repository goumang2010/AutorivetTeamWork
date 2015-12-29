using MySql.Data.MySqlClient;
using GoumangToolKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AUTORIVET_KAOHE
{
    public partial class Troubles : Form, FormInterface
    {

        int currentrow;
        int currentpos;

        private MySqlDataAdapter daMySql = new MySqlDataAdapter();
        DataTable unidt;
        public Troubles()
        {
         
            InitializeComponent();
            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int k = e.ColumnIndex;
            int l = e.RowIndex;
            if (l >= 0)
            {


                currentrow = l;
                dataGridView1.Rows[currentrow].Selected = true;
                currentpos = dataGridView1.FirstDisplayedScrollingRowIndex;

                DataRowView temprow = (DataRowView)dataGridView1.Rows[l].DataBoundItem;
                //  string liushui = temprow[5].ToString();
                switch (k)
                {
                    case 0:
                        //名称
                        SEQ_label.Text = temprow["流水号"].ToString();
                        textBox1.Text = temprow["名称"].ToString();
                     
                         comboBox2.Text = temprow["类别"].ToString();
                        textBox4.Text = temprow["备注"].ToString();
                        textBox5.Text = temprow["当前状态"].ToString();


                     

                        string startdatestr = temprow["发生日期"].ToString();
                        string closedatestr = temprow["解决日期"].ToString();
                        
                        if (closedatestr != "" && !closedatestr.Contains("0000"))
                        {
                            dateTimePicker1.Value = DateTime.Parse(closedatestr);
                            checkBox1.Checked = false;
                        }
                        if (startdatestr != "" && !startdatestr.Contains("0000"))
                        {
                            dateTimePicker2.Value = DateTime.Parse(startdatestr);
                           // checkBox1.Checked = false;
                        }


                        break;

                    case 1:
                        //流水号
                        string liushui = temprow["流水号"].ToString();

                        string folderpath = Program.InfoPath + "Troubles" + "\\" + liushui + "\\";
                        localMethod.creatDir(folderpath);
                        System.Diagnostics.Process.Start("explorer.exe", folderpath);
                       





                        break;

                }

            }



                  
        }

        private void Troubles_Load(object sender, EventArgs e)
        {
      
            DataGridViewButtonColumn col_btn_insert = new DataGridViewButtonColumn();
            col_btn_insert.HeaderText = "操作1";
            col_btn_insert.Text = "编辑";//加上这两个就能显示
            col_btn_insert.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert);
            //DataGridViewButtonColumn col_btn_insert2 = new DataGridViewButtonColumn();
            //col_btn_insert2.HeaderText = "操作2";
            //col_btn_insert2.Text = "返修";//加上这两个就能显示
            //col_btn_insert2.UseColumnTextForButtonValue = true;
            //dataGridView1.Columns.Add(col_btn_insert2);

            // dataviewfresh2();
            DataGridViewButtonColumn col_btn_insert3 = new DataGridViewButtonColumn();
            col_btn_insert3.HeaderText = "操作2";
            col_btn_insert3.Text = "资料";//加上这两个就能显示
            col_btn_insert3.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert3);
            
            
            rf_default();
        }

        #region interface

        public void rf_default()
        {
            unidt = new DataTable();

            MySqlConnection MySqlConn = new MySqlConnection(PubConstant.ConnectionString);
            MySqlConn.Open();
            String sql = "select SEQ as 流水号,NAME as 名称,CATEGORY as 类别,NOTE AS 备注,STATE AS 当前状态,OCCUR_DATE AS 发生日期,SOLVE_DATE AS 解决日期 from Troubles";
            daMySql = new MySqlDataAdapter(sql, MySqlConn);
            // DataSet OleDsyuangong = new DataSet();

            daMySql.Fill(unidt);

            rf_gridview(unidt);
          
        }
        public void rf_gridview(dynamic dt)
        {
            
            dataGridView1.DataSource = dt;
            if(dt!=null)
            {

          
            dataGridView1.Columns[0].Width =50;
            dataGridView1.Columns[1].Width =50;
            dataGridView1.Columns[2].Width = 65;
            }
        }

        public void rf_filter()
        {
            // rf_default();
            DataTable curdt = unidt;
            if (textBox6.Text != "")
            {


                var kkk = from DataRow pp in unidt.Rows
                          where pp["名称"].ToString().Contains(textBox6.Text)
                          select pp;
                if (kkk.Count() > 0)
                {
                    curdt = kkk.CopyToDataTable();

                }
                else
                {
                    curdt = null;
                }
            }

         







            rf_gridview(curdt);


        }


        public DataTable get_datatable()
        {
            return (DataTable)dataGridView1.DataSource;

        }



        #endregion

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string closedatestr = "";
            if (!checkBox1.Checked)
            {
                closedatestr = dateTimePicker2.Value.ToShortDateString();
                //startdatestr = dateTimePicker1.Value.ToShortDateString();
            }

       


            DbHelperSQL.ExecuteSql("insert into Troubles(NAME,CATEGORY,NOTE,STATE,OCCUR_DATE,SOLVE_DATE) values('" + textBox1.Text  + "','" +  comboBox2.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + dateTimePicker1.Value.ToShortDateString() + "','" + closedatestr+ "')");
           // rncrf();
            MessageBox.Show("执行成功");

            rf_default();

        }
        private void getupdate(DataTable dt)
        {
            //dt = dataGridView1.DataSource as DataTable;//把DataGridView绑定的数据源转换成DataTable 

            MySqlCommandBuilder cb = new MySqlCommandBuilder(daMySql);

            daMySql.Update(dt);

        }
        private void button3_Click(object sender, EventArgs e)
        {
            var temprow = unidt.Select("流水号 = " + SEQ_label.Text).First();

          


           
            temprow["名称"]=textBox1.Text;
        
           temprow["类别"]= comboBox2.Text;
          temprow["备注"]= textBox4.Text;
            temprow["当前状态"] = textBox5.Text;


       

            temprow["发生日期"] = new MySql.Data.Types.MySqlDateTime(dateTimePicker1.Value);
            temprow["解决日期"]= new MySql.Data.Types.MySqlDateTime(dateTimePicker2.Value);








            temprow.EndEdit();
            getupdate(unidt);
            rf_filter();



        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            rf_filter();
        }
    }
}
