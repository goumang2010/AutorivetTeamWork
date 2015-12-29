using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GoumangToolKit;
using OFFICE_Method;

using MySql.Data.MySqlClient;

namespace AUTORIVET_KAOHE
{
    public partial class RNC : Form,FormInterface
    {
       
                //当前操作的位置
        int currentrow=0;
        int currentpos = 0;
        bool querystate = false;

        bool jushoufilter = false;
        bool AAOfilter = false;

        DataTable unidt=new DataTable();

        private MySqlDataAdapter daMySql = new MySqlDataAdapter();
        MySqlConnection MySqlConn;
        public RNC()
        {
            InitializeComponent();
        }


        #region interface

       public void rf_gridview(dynamic dt)
        {

            dataGridView1.DataSource = dt;
     
           if(dt!=null)
           {

       
            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[2].Width = 50;
            dataGridView1.Columns["拒收原因"].Width = 210;
            dataGridView1.Columns["原因类型"].Width = 70;
            dataGridView1.Columns["责任人"].Width = 60;
            dataGridView1.Columns["文件"].Width = 250;
            dataGridView1.Columns["流水号"].Width = 65;
                var viewcount = dataGridView1.Rows.Count;
            if (viewcount > currentrow && viewcount > currentpos)
            {
                dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows[currentpos].Index;
                dataGridView1.Rows[currentrow].Selected = true;
            }

           }
        }

        public string product_filter
       {
            get
           {
               return listBox2.SelectedItem.ToString();
           }
            set
           {
               listBox2.SelectedItem = value;
               querystate = true;
               rf_filter();

           }



       }


        private void getupdate(DataTable dt)
        {
            //dt = dataGridView1.DataSource as DataTable;//把DataGridView绑定的数据源转换成DataTable 

            MySqlCommandBuilder cb = new MySqlCommandBuilder(daMySql);

            daMySql.Update(dt);

        }

        public void rf_default()
       {
            unidt = new DataTable();
            using (MySqlConn = new MySqlConnection(PubConstant.ConnectionString))
            {

          
                MySqlConn.Open();
            String sql = "select 外部拒收号, 内部拒收号, 拒收原因, 纠正措施, 当前状态, 文件, 关联产品, 产品架次, 发生日期, 关闭日期, 流水号, 原因类型, 责任人 from RNC总表 order by 发生日期";
            daMySql = new MySqlDataAdapter(sql, MySqlConn);
            // DataSet OleDsyuangong = new DataSet();

            daMySql.Fill(unidt);
            currentpos = unidt.Rows.Count - 1;
           rf_gridview( unidt);
           

          AAOfilter = false;
          jushoufilter = false;
          querystate = false;

            }

        }
      public  void rf_filter()
       {
            // rf_default();
            dynamic curdt;

          curdt = unidt;
         if( textBox7.Text!="")
         {

       
           var kkk = from DataRow pp in unidt.AsEnumerable()
                     where pp["当前状态"].ToString().Contains(textBox7.Text)
                     select pp;
           if (kkk.Count() > 0)
           {
               curdt = kkk.AsDataView();

           }
           else
           {
               curdt = null;
           }
         }

         if (textBox6.Text != "")
         {

             var kkk = from DataRow pp in unidt.AsEnumerable()
                       where pp["外部拒收号"].ToString().Contains(textBox6.Text)
                       select pp;
             if (kkk.Count() > 0)
             {
                 curdt = kkk.AsDataView();

             }
             else
             {
                 curdt = null;
             }


         }

         if (listBox2.SelectedIndex != -1 && querystate==true)
         {
             querystate = false;
             var kkk = from DataRow pp in unidt.AsEnumerable()
                       where (pp["关联产品"].ToString()+"_" + pp["产品架次"].ToString()) == listBox2.SelectedItem.ToString()
                       select pp;
             if (kkk.Count() > 0)
             {
                 curdt = kkk.AsDataView();

             }
             else
             {
                 curdt = null;
             }


         }
       if(  AAOfilter)
       {

    

       //  AAOfilter = false;
         var ggg = from DataRow pp in unidt.AsEnumerable()
                   where pp["文件"].ToString().Contains("AAO:" + comboBox6.Text+",")
                   select pp;
         if (ggg.Count() > 0)
         {
             curdt = ggg.AsDataView();

         }
         else
         {
             curdt = null;
         }

       }

       if (jushoufilter)
       {

           var kkk = from DataRow pp in unidt.AsEnumerable()
                     where pp["文件"].ToString().Contains("内部拒收:" + comboBox3.Text + ",")
                     select pp;
           if (kkk.Count() > 0)
           {
               curdt = kkk.AsDataView();

           }
           else
           {
               curdt = null;
           }



       }







         rf_gridview(curdt);

       }
      public  DataTable get_datatable()
        {

        //   unidt =      autorivet_op.RNC_view();
            return unidt;

        }



        #endregion








        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            //
        }
        //private void rncrf()
        //{
           
        //        unidt = autorivet_op.RNC_view();
        //        dataGridView1.DataSource = unidt;
           

        //    dataGridView1.Columns[0].Width = 50;
        //    dataGridView1.Columns[1].Width = 50;
        //    dataGridView1.Columns[2].Width = 50;
        //    dataGridView1.Columns["责任人"].Width = 50;
        //    dataGridView1.Columns["流水号"].Width = 40;
        //   // dataGridView1.Columns[2].Width = 50;
        //    dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows[currentpos].Index;
        //    dataGridView1.Rows[currentrow].Selected = true;
        //}
        private void RNC_Load(object sender, EventArgs e)
        {

            //MySqlConnection MySqlConn = new MySqlConnection(PubConstant.ConnectionString);
            //MySqlConn.Open();
            //String sql = "select 外部拒收号, 内部拒收号, 拒收原因, 纠正措施, 当前状态, 文件, 关联产品, 产品架次, 发生日期, 关闭日期, 流水号, 原因类型, 责任人 from RNC总表 order by 发生日期";
            //daMySql = new MySqlDataAdapter(sql, MySqlConn);
            //// DataSet OleDsyuangong = new DataSet();

            //daMySql.Fill(unidt);
            // listBox1.DataSource = DbHelperSQL.getlist("select 产品架次 from 产品流水表 group by 产品架次");

            listBox2.DataSource = DbHelperSQL.getlist("select concat(产品名称,'_',产品架次) from 产品流水表");

            DataTable prod = DbHelperSQL.Query("select concat(产品名称,'_',产品架次) from 产品流水表 where 当前状态='正在进行'").Tables[0];
            if (prod.Rows.Count != 0)
            {
                listBox2.SelectedItem = prod.Rows[0][0].ToString();


            }
            else
            {
                MessageBox.Show("请添加正在进行的产品");
            }


            DataGridViewButtonColumn col_btn_insert = new DataGridViewButtonColumn();
            col_btn_insert.HeaderText = "操作1";
            col_btn_insert.Text = "编辑";//加上这两个就能显示
            col_btn_insert.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert);
            DataGridViewButtonColumn col_btn_insert2 = new DataGridViewButtonColumn();
            col_btn_insert2.HeaderText = "操作2";
            col_btn_insert2.Text = "AAO";//加上这两个就能显示
            col_btn_insert2.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert2);

            // dataviewfresh2();
            DataGridViewButtonColumn col_btn_insert3 = new DataGridViewButtonColumn();
            col_btn_insert3.HeaderText = "操作3";
            col_btn_insert3.Text = "打开";//加上这两个就能显示
            col_btn_insert3.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert3);




            rf_default();
        }

        //private void checkBox2_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (checkBox2.Checked == true)
        //    {
        //        closedate = false;
        //    }
        //    else
        //    {
        //        closedate = true;
        //    }
        //}

        private void button3_Click(object sender, EventArgs e)
        {
            string closedatestr = "";
            if (!checkBox2.Checked)
            {
                closedatestr = dateTimePicker1.Value.ToShortDateString();
                //startdatestr = dateTimePicker1.Value.ToShortDateString();
            }

            string[] prod = listBox2.SelectedValue.ToString().Split('_');






            DbHelperSQL.ExecuteSql("insert into RNC总表(外部拒收号,内部拒收号,拒收原因,纠正措施,当前状态,文件,关联产品,产品架次,发生日期,关闭日期,原因类型) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + filetrackstr + "','" + prod[0] + "','" + prod[1] + "','" + dateTimePicker1.Value.ToShortDateString() + "','" + closedatestr + "','" + comboBox1.Text + "')");
            rf_default();

            //填充Production表备注
            update_note(prod[0], prod[1]);


            FormMethod.notifyServers("插入了RNC:" + textBox1.Text+ ",产品：" + prod[0]+"，架次：" + prod[1]);

            MessageBox.Show("执行成功");
        
        }


        private string filetrackstr
        {
            get
            {
                //生成文件字符串

                string jushou = "内部拒收:" + comboBox3.Text + "," + comboBox4.Text;

                string AAO = "AAO:" + comboBox6.Text + "," + comboBox5.Text;


                //默认报废单不存在
                string scrap = "";
                if (comboBox8.Text!="")
                {
                  scrap = "报废单:" + comboBox8.Text + "," + comboBox7.Text;
                }

              


                string filestr = jushou + ";" + AAO + ";" + scrap;
                return filestr;
            }
            set
            {

                if (value=="")
                {

                    comboBox3.Text="未编制";
                    comboBox4.Text = "检验";

                    comboBox6.Text = "未编制";
                    comboBox5.Text = "工艺";

                    comboBox8.Text = "";
                    comboBox7.Text = "检验";


                }
                else
                {

                    string[] aa = value.Split(';');

                    string[] jushou = aa[0].Split(':')[1].Split(',');

                    comboBox3.Text = jushou[0];
                    comboBox4.Text = jushou[1];



                    string[] AAO = aa[1].Split(':')[1].Split(',');

                    comboBox6.Text = AAO[0];
                    comboBox5.Text = AAO[1];

                    if(aa[2]!="")
                    {
                        string[] scrap = aa[2].Split(':')[1].Split(',');


                        comboBox8.Text = scrap[0];
                        comboBox7.Text = scrap[1];


                    }





                }



            }








        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int k = e.ColumnIndex;
            int l = e.RowIndex;
            if(l>=0)
            {

     
           currentrow = l;
       // dataGridView1.Rows[currentrow].Selected = true;
            currentpos = dataGridView1.FirstDisplayedScrollingRowIndex; 
      
            DataRowView temprow = (DataRowView)dataGridView1.Rows[l].DataBoundItem;
            //  string liushui = temprow[5].ToString();



            var bb = from kk in Program.prodTable.AsEnumerable()
                     where kk["名称"].ToString() == temprow["关联产品"].ToString()
                     select kk["名称"].ToString() + "_" + kk["图号"].ToString();



            //直接获取目录

            string path = Program.InfoPath + bb.First() + "\\" + temprow["产品架次"].ToString() + "\\RNC\\" + temprow["外部拒收号"].ToString();


            localMethod.creatDir(path);
                
                
                
                
                
                switch (k)
            {
                case 0:
                    //产品-汉字
                    listBox2.SelectedItem = temprow[6].ToString() + "_" + temprow[7].ToString();

                    //架次
                  //  listBox1.SelectedItem = temprow[7].ToString();
                   
                //6个文本框
                    //外部拒收号
                textBox1.Text = temprow["外部拒收号"].ToString();

                //内部拒收号
                textBox2.Text = temprow[1].ToString();
                //拒收原因
                textBox3.Text = temprow[2].ToString();
                textBox4.Text = temprow[3].ToString();
                textBox5.Text = temprow[4].ToString();

                    //AAO
                filetrackstr = temprow["文件"].ToString();

                    string startdatestr = temprow[8].ToString();
                    string closedatestr = temprow[9].ToString();
                    //string transferdatestr = temprow[6].ToString();

                    if (startdatestr != "" && !startdatestr.Contains( "0000"))
                    {
                        dateTimePicker1.Value = DateTime.Parse(startdatestr);


                    }
                    if (closedatestr != "" && !closedatestr.Contains("0000"))
                    {
                        dateTimePicker2.Value = DateTime.Parse(closedatestr);
                        checkBox2.Checked =false;
                    }
                    else
                    {
                        checkBox2.Checked = true;
                    }
                   

                label1.Text = temprow[10].ToString();
                comboBox1.Text = temprow[11].ToString();
                comboBox2.Text = temprow["责任人"].ToString();
                break;
                case 1:

                 Dictionary<string, string> tmp = new Dictionary<string, string>();
                     //名称
                 string prodname = temprow[6].ToString();

                 tmp.Add("中文名称", prodname+"壁板");
                 tmp.Add("名称", prodname );
                //架次
                string jiaci = temprow[7].ToString();
                tmp.Add("架次", jiaci);
                 //填充标题
                 string neibujushou=temprow[0].ToString();
                     //现改为填充外部拒收号
                 tmp.Add("内部拒收号", neibujushou);
                 string folderpath = path;
                 tmp.Add("保存地址",folderpath+"\\"+temprow[0].ToString()+"_AAO.doc");
                 tmp.Add("类型", "AO");

                     AO f=new AO();
                     f.rncaao=tmp;
                     f.Show();
            //现已改为AAO

                    paperWork f2 =new paperWork();
                    f2.Show();
                    f2.filter_filename = temprow["外部拒收号"].ToString();
                    f2.rf_filter();
                break;
                case 2:
                    //获取目录名称


                try
                {
                      System.Diagnostics.Process.Start("explorer.exe", path);

                }
        
                    catch
                    {
                        FormMethod.creatCredential();
                    }
                break;




            }

            }
          
        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //   // int productindex = listBox2.SelectedIndex;
        //  //  int jiaciindex = listBox1.SelectedIndex;

        //    string[] prod = listBox2.SelectedValue.ToString().Split('_');
         
        //        dataGridView1.DataSource = DbHelperSQL.Query("select 外部拒收号,内部拒收号,拒收原因,纠正措施,当前状态,AAO,关联产品,产品架次,发生日期,关闭日期,流水号,原因类型 from RNC总表 where 关联产品='" + prod[0] + "' and 产品架次='" + prod[1] + "'").Tables[0];


         
        
        //}
        public void rncrf(string prodectname,string jiaci)
        {
            dataGridView1.DataSource = DbHelperSQL.Query("select 外部拒收号,内部拒收号,拒收原因,纠正措施,当前状态,AAO,关联产品,产品架次,发生日期,关闭日期,流水号,原因类型 from RNC总表 where 关联产品='" + prodectname + "' and 产品架次='" + jiaci + "'").Tables[0];
        }
        //public void rncrf(string state)
        //{
        //    var kkk = from DataRow pp in unidt.Rows
        //              where pp["当前状态"].ToString().Contains(state)
        //              select pp;
        //    if (kkk.Count()>0)
        //    {
        //        dataGridView1.DataSource = kkk.CopyToDataTable();
               
        //    }
        //    else
        //    {
        //        dataGridView1.DataSource = null;
        //    }
            
        //}

        private void button4_Click(object sender, EventArgs e)
        {
            textBox6.Text = "";
            textBox7.Text = "";
           // listBox2.SelectedIndex = -1;
            rf_default();
        }

        private void button2_Click(object sender, EventArgs e)
        {
     string closedatestr = "0000-00-00";
            if (!checkBox2.Checked)
            {
                closedatestr = dateTimePicker2.Value.ToShortDateString();
                //startdatestr = dateTimePicker1.Value.ToShortDateString();
            }


            string[] prod = listBox2.SelectedValue.ToString().Split('_');

           var pp= unidt.Select("流水号 = " + label1.Text).First();

            pp["外部拒收号"] = textBox1.Text;
            pp["内部拒收号"] = textBox2.Text;
            pp["拒收原因"] = textBox3.Text;
            pp["纠正措施"] = textBox4.Text;
            pp["当前状态"] = textBox5.Text;
           
            pp["文件"] = filetrackstr;
            pp["关联产品"] = prod[0];
            pp["产品架次"] = prod[1];
      

            pp["发生日期"] =new MySql.Data.Types.MySqlDateTime(dateTimePicker1.Value);
            pp["关闭日期"] = new MySql.Data.Types.MySqlDateTime(closedatestr);
            pp["原因类型"] = comboBox1.Text;
            pp["责任人"] = comboBox2.Text;

            pp.EndEdit();
            getupdate(unidt);



            //填充Production表备注
           update_note(prod[0], prod[1]);
            FormMethod.notifyServers("更新了RNC:" + textBox1.Text + ",产品：" + prod[0] + "，架次：" + prod[1]);

            MessageBox.Show("执行成功");
        }

        private void update_note(string prodname,string fuseno)
        {




            var kk = from cc in unidt.AsEnumerable()
                     where cc["关联产品"].ToString() == prodname && cc["产品架次"].ToString() == fuseno
                     select cc;


            string beizhu = "";
            if (kk.Count() != 0)
            {
                beizhu = "RNC:";
                foreach (DataRow pp in kk)
                {
                    beizhu = beizhu + "[" + pp["外部拒收号"].ToString() + ":" + pp["当前状态"].ToString() + ";" + pp["文件"].ToString() + "]";
                }
                DbHelperSQL.ExecuteSql("update 产品流水表 set 备注='" + beizhu + "' where 产品名称='" + prodname+"' and 产品架次='"+fuseno+"';");
            }

        }




        private void button6_Click(object sender, EventArgs e)
        {
          
           try
         {
             DataTable output = (DataTable)dataGridView1.DataSource;
            var abc = excelMethod.SaveDataTableToExcel(output);
           //  abc.Rows.AutoFit();
             if (checkBox1.Checked == true)
             {
                 string path = @"\\192.168.3.32\Autorivet\output\INFO\backup\";
                    excelMethod.SaveAs(abc,path + "RNC.xlsx");
                 System.Diagnostics.Process.Start("explorer.exe", path);
             }




           }
         catch
          {
              string path = @"D:\Autorivet_INFO\";
              localMethod.creatDir(path);


                foreach (Form Frm in Application.OpenForms)
                {
                    if (Frm.Name == "Form1")
                    {
                        ((MainTeamForm)Frm).pushinfo = "output_RNC";
                   System.Threading.Thread.Sleep(10000);
                   if (!((MainTeamForm)Frm).getinfo.Contains("文件保存成功"))
                   {
                       System.Threading.Thread.Sleep(5000);
                   }
                     System.Diagnostics.Process.Start(path+"RNC总表.xlsx");
                     // System.Diagnostics.Process.Start("explorer.exe",@"D:\");
            
                    }
                }
      }
           
        }

       

        private void button8_Click(object sender, EventArgs e)
        {
            DataTable productliushui = DbHelperSQL.Query("select 外部拒收号 from rnc总表").Tables[0];
            foreach (DataRow p in productliushui.Rows)
            {
                DataTable rnc = DbHelperSQL.Query("select 文件编号,编制日期 from paperwork where 文件名 like '" + p[0].ToString() + "%'").Tables[0];
                string beizhu = "";
                if (rnc.Rows.Count != 0)
                {
                    beizhu = "AAO:";
                    foreach (DataRow pp in rnc.Rows)
                    {
                        beizhu = beizhu + pp[1].ToString()+"提出" + pp[0].ToString() ;
                    }
                    DbHelperSQL.ExecuteSql("update rnc总表 set AAO='" + beizhu + "' where 外部拒收号='" + p[0].ToString()+"'");
                }
            }

            MessageBox.Show("执行成功");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataTable output = DbHelperSQL.Query("select replace(拒收原因,'\r\n\r\n','\r\n') as 问题描述,发生日期 as 日期,replace(concat(AAO,';',当前状态),'\r\n\r\n','\r\n') as 状态更新,关闭日期 as 实际完成,concat(关联产品,产品架次,';RNC:',外部拒收号) as 备注 from rnc总表 order by 日期").Tables[0];
            excelMethod.SaveDataTableToExcel(output);


        }

        private void button10_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {


            rf_filter();

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            
            rf_filter();
        }

        private void listBox2_DoubleClick(object sender, EventArgs e)
        {
            textBox7.Text = "";

            textBox6.Text = "";

            querystate = true;
            rf_filter();



        }

        private void textBox6_Click(object sender, EventArgs e)
        {
           listBox2.SelectedIndex = -1;
            textBox7.Text = "";
        }

        private void textBox7_Click(object sender, EventArgs e)
        {
           listBox2.SelectedIndex = -1;
            textBox6.Text = "";
        }

        private void label22_Click(object sender, EventArgs e)
        {
            AAOfilter = true;
            rf_filter();


        }

        private void label19_Click(object sender, EventArgs e)
        {
            jushoufilter = true;
            rf_filter();
   

        }

        private void aAOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable productliushui = DbHelperSQL.Query("select 外部拒收号 from rnc总表").Tables[0];
            foreach (DataRow p in productliushui.Rows)
            {
                DataTable rnc = DbHelperSQL.Query("select 文件编号,编制日期 from paperwork where 文件名 like '" + p[0].ToString() + "%'").Tables[0];
                string beizhu = "";
                if (rnc.Rows.Count != 0)
                {
                    beizhu = "AAO:";
                    foreach (DataRow pp in rnc.Rows)
                    {
                        beizhu = beizhu + pp[1].ToString() + "提出" + pp[0].ToString();
                    }
                    DbHelperSQL.ExecuteSql("update rnc总表 set AAO='" + beizhu + "' where 外部拒收号='" + p[0].ToString() + "'");
                }
            }

            MessageBox.Show("执行成功");
        }

        private void 问题记录单ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable output = DbHelperSQL.Query("select replace(拒收原因,'\r\n\r\n','\r\n') as 问题描述,发生日期 as 日期,replace(concat(AAO,';',当前状态),'\r\n\r\n','\r\n') as 状态更新,关闭日期 as 实际完成,concat(关联产品,产品架次,';RNC:',外部拒收号) as 备注 from rnc总表 order by 日期").Tables[0];
            excelMethod.SaveDataTableToExcel(output);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var pp = ((DataRowView)(dataGridView1.SelectedRows[0].DataBoundItem)).Row;
            FormMethod.notifyServers("删除了RNC:" + pp["外部拒收号"].ToString() + ",产品：" + pp["关联产品"].ToString() + "，架次：" + pp["产品架次"].ToString());
            pp.Delete();


            pp.EndEdit();
            getupdate(unidt);

            //  rf_gridview(unidt);
            //    DbHelperSQL.ExecuteSql("update RNC总表 set 外部拒收号='" + textBox1.Text + "',内部拒收号='" + textBox2.Text + "',拒收原因='" + textBox3.Text + "',纠正措施='" + textBox4.Text + "',当前状态='" + textBox5.Text + "',文件='" + filetrackstr + "',关联产品='" + prod[0] + "',产品架次='" + prod[1] + "',发生日期='" + dateTimePicker1.Value.ToShortDateString() + "',关闭日期='" + closedatestr + "',原因类型='" + comboBox1.Text + "',责任人='" + comboBox2.Text + "' where 流水号=" + label1.Text);
            //   dataGridView1.SelectedRows=
            //  dataGridView1.Rows[dataGridView1.Rows.Count-1].Selected = true;
            rf_filter();
            //  dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows[currentpos].Index;
            //  dataGridView1.Rows[currentrow].Selected = true;
            //   productionrf();

            //填充Production表备注
       


            MessageBox.Show("执行成功");


        }
    }
}
