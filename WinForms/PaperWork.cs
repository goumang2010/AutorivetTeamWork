using mysqlsolution;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FileManagerNew;
using System.Threading.Tasks;
using OFFICE_Method;

namespace AUTORIVET_KAOHE
{
    public partial class paperWork : Form,FormInterface
    {

      //  private Microsoft.Office.Interop.Word.Application wordApp;
        private Microsoft.Office.Interop.Word.Document myWordfile;

        private object defaultV = System.Reflection.Missing.Value;
     //   private object documentType;


       private DataTable unidt=new DataTable();



        public paperWork()
        {
            InitializeComponent();
        }
        public void ppworkrf(string filterstr, string jiaci = "")
        {
            string sqlstr="";
         
              
                   //传递产品中文名的情况
                   if (jiaci == "")
                   {
                       sqlstr = "select 文件类型 as 类型,文件名,文件编号,文件名称,状态,位置,版次,关联产品,关联架次,编制日期,修改日期,备注,文件地址 as 地址 from paperWork where 关联产品='" + filterstr + "' order by 编制日期";
                   }
                   else
                   {
                       sqlstr = "select 文件类型 as 类型,文件名,文件编号,文件名称,版次,关联产品,关联架次,编制日期,修改日期,备注,文件地址 as 地址 from paperWork where 关联产品='" + filterstr + "' and (关联架次='" + jiaci + "' or 关联架次='--') order by 编制日期";
                   }

            dataGridView1.DataSource = DbHelperSQL.Query(sqlstr).Tables[0];
        }

        public DataTable get_datatable()
        {
          

              unidt = AutorivetDB.paperwork_view();

              return unidt;
        }

            
        public void rf_filter()
        {
            //根据窗体的控件状态进行筛选
            DataTable curdt = unidt;
          if(unidt.Rows.Count>0)
          {

              if (comboBox2.SelectedIndex>0)
             {
                 string papersort = comboBox2.Text;

                 var cc = curdt.AsEnumerable().Where(p => p["类型"].ToString() == papersort);

                 if (cc.Count() > 0)
                 {
                     curdt = cc.CopyToDataTable();
                     rf_gridview(curdt);
                 }


                 else
                 {
                     curdt = null;
                 }




             }
           


               if (listBox2.SelectedIndex>=0)
               {

                   string prodname = listBox2.SelectedValue.ToString();

                   var cc = curdt.AsEnumerable().Where(p => p["关联产品"].ToString() == prodname);

                   if (cc.Count() > 0)
                   {
                       curdt = cc.CopyToDataTable();
                       rf_gridview(curdt);
                   }


                   else
                   {
                       curdt = null;
                   }



               }
             

               if (textBox5.Text != "" || textBox6.Text != "" || textBox7.Text != "")
               {

                  // string prodname = listBox2.SelectedValue.ToString();

                   var cc = curdt.AsEnumerable().Where(p => p["文件名"].ToString().Contains(textBox5.Text) && p["文件编号"].ToString().Contains(textBox6.Text) && p["备注"].ToString().Contains(textBox7.Text));

                   if (cc.Count() > 0)
                   {
                       curdt = cc.CopyToDataTable();
                       rf_gridview(curdt);
                   }


                   else
                   {
                       curdt = null;
                   }



               }

               rf_gridview(curdt);











          }


        }


         public void ppworkrf(string filterstr, int k)
        {
            string sqlstr = "";
            switch (k)
            {
                case 1:
                    //文件编号
                    sqlstr = "select 文件类型 as 类型,文件名,文件编号,文件名称,版次,关联产品,关联架次,编制日期,修改日期,备注,文件地址 as 地址 from paperWork where 文件编号='" + filterstr + "' order by 编制日期";

                    break;
                case 2:
                    //文件类型
                    sqlstr = "select 文件类型 as 类型,文件名,文件编号,文件名称,版次,关联产品,关联架次,编制日期,修改日期,备注,文件地址 as 地址 from paperWork where 文件类型='" + filterstr + "' order by 编制日期";
                    break;
            }
            dataGridView1.DataSource = DbHelperSQL.Query(sqlstr).Tables[0];
        }

         public void ppworkrf(int k)
         {
             //用于判断筛选的文件类型
             switch (k)
             {
                  

                 case 1:
                     //AO
                     ppworkrf("AO", 2);
                     break;
                 case 2:
                     //补铆_AAO
                     ppworkrf("补铆_AAO", 2);
                     break;
                 case 3:
                     //RNC_AAO
                     ppworkrf("RNC_AAO", 2);
                     break;
                 case 4:
                     //COS
                     ppworkrf("COS", 2);
                     break;
                 case 5:
                     //技术单
                     ppworkrf("TS", 2);
                     break;
                 case 6:
                     //大纲索引
                     ppworkrf("INDEX", 2);
                     break;
                 default:
                     //显示全部
                     rf_default();
                     break;
             }
         }


        public string filter_filename
         {
            get
             {
                 return textBox5.Text;
             }

            set
             {
                textBox5.Text=value;

             }
         }






        public  void rf_gridview(dynamic dt)
         {
             dataGridView1.DataSource = dt;
         }
        private void paperWork_Load(object sender, EventArgs e)
        {
            
            listBox2.DataSource = DbHelperSQL.getlist("select 名称 from 产品列表");
            listBox2.SelectedIndex = -1;
          
            DataGridViewButtonColumn col_btn_insert = new DataGridViewButtonColumn();
            col_btn_insert.HeaderText = "操作1";
            col_btn_insert.Text = "Open";//加上这两个就能显示
            col_btn_insert.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert);
            // dataviewfresh2();
            DataGridViewButtonColumn col_btn_insert2 = new DataGridViewButtonColumn();
            col_btn_insert2.HeaderText = "操作2";
            col_btn_insert2.Text = "Folder";//加上这两个就能显示
            col_btn_insert2.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert2);
            DataGridViewButtonColumn col_btn_insert3 = new DataGridViewButtonColumn();
            col_btn_insert3.HeaderText = "操作3";
            col_btn_insert3.Text = "Edit";//加上这两个就能显示
            col_btn_insert3.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert3);
            DataGridViewButtonColumn col_btn_insert4 = new DataGridViewButtonColumn();
            col_btn_insert4.HeaderText = "操作4";
            col_btn_insert4.Text = "Prod";//加上这两个就能显示
            col_btn_insert4.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert4);
            rf_default();

            DataGridViewCheckBoxColumn col_btn_insert5 = new DataGridViewCheckBoxColumn();
            col_btn_insert5.HeaderText = "选择";
            col_btn_insert5.Name = "choose";
           // col_btn_insert5.Text = "Prod";//加上这两个就能显示
           // col_btn_insert5. = true;
            dataGridView1.Columns.Add(col_btn_insert5);

        }


       public  void rf_default()
        {
            rf_gridview(get_datatable());
            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[2].Width = 50;
            dataGridView1.Columns[3].Width = 50;
            dataGridView1.Columns[4].Width = 70;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(listBox2.SelectedIndex!=-1)
            {

       
            //fileInfo allfiles = fileOP.WalkTree(Properties.Settings.Default.filepath.Remove(Properties.Settings.Default.filepath.Count()-1));
            fileInfo allfiles = fileOP.WalkTree(FormMethod.get_storefolder(listBox2.SelectedItem.ToString()));
            FormMethod.scanfiledoc(allfiles.pathfilter("", "old").pathfilter("", "backup").extfilter("doc").namefilter("AAO", "~$"), checkBox2.Checked);
           // scanfilepdf(allfiles.pathfilter("","old").extfilter("pdf").namefilter("V0"));
            rf_default();
            }
            else
            {
                MessageBox.Show("请选择1个产品");
            }

        }



          private void scanfilepdf(fileInfo allfiles)
          {
              int count = allfiles.count;

              for (int i = 0; i < count; i++)
              {
                  string filename = allfiles.fileName[i];
                  string biaoshi = filename.Split('_').Last();

                  switch (biaoshi)
                  {
                      case "V0":

                          break;

                  }



              }

          }

          private void button2_Click(object sender, EventArgs e)
          {
              string creatdate = "";
              if (!checkBox1.Checked)
              {
                  creatdate = dateTimePicker1.Value.ToShortDateString();
              }



              DbHelperSQL.ExecuteSql("update paperwork set 编制日期='" + creatdate + "',备注='" + textBox1.Text + "',状态='" + comboBox3.Text + "',位置='" + comboBox4.Text + "',版次='" + textBox2.Text + "' where 文件名='" + label1.Text + "';");
            //  ppworkrf(comboBox2.SelectedIndex);
              get_datatable();
              rf_filter();


              //对编辑信息的行为进行记录

              AutorivetDB.log_insert(Program.userID, "编辑文件信息:" + label1.Text);







              MessageBox.Show("执行成功");
          }

          private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
          {
                          int k = e.ColumnIndex;
            int l = e.RowIndex;

              if(l<0)
              {
                  return;

              }

            DataRowView temprow = (DataRowView)dataGridView1.Rows[l].DataBoundItem;
          //  string liushui = temprow[5].ToString();
            bool manager = Program.ManagerActived;
            string filepath=temprow["地址"].ToString();
              string filename=temprow["文件名"].ToString();

              switch (k)
              {
                  case 0:
              System.Diagnostics.Process.Start(filepath);
                     
                      
                      
                   
                      
                      
                      
                      
                      
                      break;


                  case 1:
                      string path = filepath.Replace(filename,"");
                   
                          System.Diagnostics.Process.Start("explorer.exe", path);
                    
                      break;
                  case 2:

                     




             string creatdate = temprow["编制日期"].ToString();
              textBox1.Text = temprow["备注"].ToString();
             //版次
              textBox2.Text= temprow["版次"].ToString();
                      label1.Text = filename;
                      if (creatdate != ""&& !creatdate.Contains("0000"))
                      {
                          dateTimePicker1.Value = DateTime.Parse(creatdate);
                          checkBox1.Checked = false;
                      }
                      else
                      {
                          checkBox1.Checked = true;
                      }

                      comboBox3.Text = temprow["状态"].ToString();
                      comboBox4.Text = temprow["位置"].ToString();








                      break;
                  case 3:
                      string prod=temprow["关联产品"].ToString();
                      string jiaci=temprow["关联架次"].ToString();
                      Production f = new Production();
                     
                      f.Show();
                      f.shaixuan(prod, jiaci);
                      break;
              }
          }

          private void button3_Click(object sender, EventArgs e)
          {
              DataTable output = (DataTable)dataGridView1.DataSource;
              excelMethod.SaveDataTableToExcel(output);
          }

          private void button4_Click(object sender, EventArgs e)
          {
              if(listBox2.SelectedIndex==-1||comboBox1.SelectedIndex==-1)
              {
                  MessageBox.Show("请选择产品及类型");
              }
              else
              {
                  OpenFileDialog ofd = new OpenFileDialog();
                  ofd.ShowDialog();
                  if (ofd.FileName != "")
                  {
                      string prodname = listBox2.SelectedValue.ToString();
                      DataRow pp = DbHelperSQL.Query("select 图号,站位号 from 产品列表 where 名称='" + prodname+"'").Tables[0].Rows[0];
                    
                      string filepath="";
                      string filefullname = "";
                      string filename = "";
                      string fileExt = "";
                      switch(comboBox1.Text)
                      {
                        

                    case "COS":
                              
                              filepath = @"\\192.168.3.32\Autorivet\prepare\INFO\" + prodname + "_" + pp[0] + @"\COS";
                              filename = pp[0] + "_" + pp[1] + "_" + "COS.doc";
                              filefullname = filepath + "\\" + filename;
                              fileExt = ".doc";
                              
                             
                        break;
                    case "AO":

                               filepath = @"\\192.168.3.32\Autorivet\prepare\INFO\" + prodname + "_" + pp[0] + @"\AO";
                            
                              filename = pp[0] + "_" + pp[1] + "_" + "AO.doc";
                              filefullname = filepath + "\\" + filename;
                              fileExt = ".doc";

                     
                        break;
                    case "AOI":

                        filepath = @"\\192.168.3.32\Autorivet\prepare\INFO\" + prodname + "_" + pp[0] + @"\AO";

                        filename = pp[0] + "_" + pp[1] + "_" + "AOI.doc";
                        filefullname = filepath + "\\" + filename;
                        fileExt = ".doc";


                        break;
                    case "PACR":

                        filepath = @"\\192.168.3.32\Autorivet\prepare\INFO\" + prodname + "_" + pp[0] + @"\AO";

                        filename = pp[0] + "_" + pp[1] + "_" + "PACR.doc";
                        filefullname = filepath + "\\" + filename;
                        fileExt = ".doc";


                        break;
                    case "TS":

                        filepath = @"\\192.168.3.32\Autorivet\prepare\INFO\" + prodname + "_" + pp[0] + @"\TS";

                        filename = pp[0] + "_" + pp[1] + "_" + "TS.doc";
                        filefullname = filepath + "\\" + filename;
                        fileExt = ".doc";


                        break;




                    default:
                        
                        break; 


                      }

                      localMethod.creatDir(filepath);
                      try
                      {
                          File.Copy(ofd.FileName, filefullname);
                      }
                      catch(SystemException kk)
                      {
                          MessageBox.Show(kk.Message);
                      }
                      
                      fileInfo tempfile = new fileInfo();
                      tempfile.add(filepath, filefullname, filename, fileExt);
                      //扫描该文档
                     FormMethod.scanfiledoc(tempfile);
                  }
              }

          }

          private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
          {
              listBox2.SelectedIndex = -1;
              rf_filter();

                    //  ppworkrf(comboBox2.SelectedIndex);
                     
              
          }

          private void button5_Click(object sender, EventArgs e)
          {
            DataTable temdt= (DataTable) dataGridView1.DataSource;
            List<DataRow> ccc = new List<DataRow>();
            int count = temdt.Rows.Count;
              for(int i=0;i<count;i++)
              {
                  string filepath = temdt.Rows[i]["地址"].ToString();

                  if(!File.Exists(filepath))
                  {

                     ccc.Add(temdt.Rows[i]);
                      DbHelperSQL.ExecuteSql("delete from paperwork where 文件名='" + temdt.Rows[i][1].ToString() + "';");
                  }
               

              }
              if(ccc.Count()>0)
              {
                   get_datatable();
              rf_filter();
              excelMethod.SaveDataTableToExcel(ccc.CopyToDataTable());
              }
             
              MessageBox.Show("清理完毕");
          }

          private void checkBox3_CheckedChanged(object sender, EventArgs e)
          {
              if(checkBox3.Checked==true)
              {
                     int count = Convert.ToInt16( dataGridView1.Rows.Count.ToString()); 
            for (int i = 0; i < count; i++) 
            { 
                DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells["choose"];
              checkCell.Value=true;

               
            } 
      
              }
              else
              {
                  int count = Convert.ToInt16(dataGridView1.Rows.Count.ToString());
                  for (int i = 0; i < count; i++)
                  {
                      DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells["choose"];
                      checkCell.Value = false;


                  } 


              }
          }

          private void button6_Click(object sender, EventArgs e)
          {
             // wordMethod.
              string outputpath=@"\\192.168.3.32\Autorivet\output\INFO\files\";
               int count = Convert.ToInt16(dataGridView1.Rows.Count.ToString()); 
            for (int i = 0; i < count; i++) 
            { 
                
               DataRowView temprow = (DataRowView)dataGridView1.Rows[i].DataBoundItem;
                DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells["choose"]; 
                Boolean flag = Convert.ToBoolean(checkCell.Value); 
               if (flag == true)     //查找被选择的数据行 
                {       
                    string sourcepath=temprow["地址"].ToString();
                   
                    string filename=temprow["文件名"].ToString();

                    string targetpath = outputpath + filename.Replace(filename.Split('.').Last(), "pdf");
                    
                    wordMethod.WordToPDF(sourcepath,targetpath);
                   checkCell.Value = false; 

                } 
                else 
                   continue; 
            } 

               System.Diagnostics.Process.Start("explorer.exe", outputpath);
      }

        private void groupBox3_Enter(object sender, EventArgs e)
        {
        
        }

        private void cOSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable coslist=DbHelperSQL.Query("select 文件地址,文件名 from paperwork where 文件类型='COS'").Tables[0];
            List<string> creatName = new List<string>();

            foreach (DataRow pp in coslist.Rows)
            {
                string drawingno = pp[1].ToString().Split('_')[0].Replace("\r\a", "");
                myWordfile = wordMethod.opendoc(pp[0].ToString(),false);

                Microsoft.Office.Interop.Word.Table parttable = myWordfile.Tables[1];
                int ptc = parttable.Rows.Count;
                for (int i=2;i<=ptc;i++)
                {
                    int k = 0;
                    if (parttable.Cell(1, 1).Range.Text.Contains("序号"))
                    {
                        k = 1;
                    }
                    string partname = parttable.Cell(i, 3+k).Range.Text.Replace("\r\a", "");
                    string partno = parttable.Cell(i, 1+k).Range.Text.Replace("\r\a", "");
                    string partqty = parttable.Cell(i, 4+k).Range.Text.Replace("\r\a", "");
                    StringBuilder strSqlname = new StringBuilder();
                    strSqlname.Append("INSERT INTO 零件列表 (");

                    strSqlname.Append("零件名称,零件图号,零件数量,产品图号");

                    strSqlname.Append(string.Format(") VALUES ('{0}','{1}',{2},'{3}') ON DUPLICATE KEY UPDATE 零件数量={2}", partname, partno, partqty, drawingno));

                    

                    creatName.Add(strSqlname.ToString());
                    
                }
            }

            DbHelperSQL.ExecuteSqlTran(creatName);
            MessageBox.Show("执行成功");
          
        }

        private void otherPaperworkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutorivetDB.otherpaperwork_table();
            otherPaperwork f1 = new otherPaperwork();
            f1.Show();



        }

        private void button7_Click(object sender, EventArgs e)
        {
            System.Data.DataTable allfiles = DbHelperSQL.Query("select 文件名,文件地址 from Paperwork").Tables[0];
            string outpufolder = @"\\192.168.3.32\Autorivet\output\INFO\files\Autorivet\";
            foreach (DataRow pp in allfiles.Rows)
            {
                string targetpath = outpufolder + pp[0].ToString().Split('.')[0] + ".pdf";
                wordMethod.WordToPDF(pp[1].ToString(), targetpath);
            }


            MessageBox.Show("完成");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("工具会首先杀死所有的word进程，请先保存当前的word工作，然后点击确定", "警告", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                FormMethod.killProcess("WINWORD");
                int count = Convert.ToInt16(dataGridView1.Rows.Count.ToString());
               
                for (int i = 0; i < count; i++)
                {

                    DataRowView temprow = (DataRowView)dataGridView1.Rows[i].DataBoundItem;
                    DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells["choose"];
                    Boolean flag = Convert.ToBoolean(checkCell.Value);
                    if (flag == true)     //查找被选择的数据行 
                    {
                        string[] strobj = new string[3];
                        strobj[0] = temprow["地址"].ToString();
                        strobj[1] = textBox3.Text;
                        strobj[2] = textBox4.Text;

                        FormMethod.backupfile(temprow["地址"].ToString());
                        //Thread t1 = new Thread(new ParameterizedThreadStart(replacethread));
                        //t1.Start(strobj);

                        Task.Factory.StartNew(replacethread, strobj);

                        
                        
                        //  wordMethod.SearchReplace(sourcepath, textBox3.Text, textBox4.Text);
                     //   checkCell.Value = false;

                    }
                    else
                        continue;
                }

            }
          //  checkBox3.Checked = false;
            Task.WaitAll();
            MessageBox.Show("完成");

        } 

        private void replacethread(object obj)
        {
            wordMethod.SearchReplace(((string[])obj)[0], ((string[])obj)[1], ((string[])obj)[2]);


        }

        private void listBox2_DoubleClick(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex=0;
            rf_filter();
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
           //comboBox2.SelectedIndex = 0;
            listBox2.SelectedIndex = -1;
            rf_filter();
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            //comboBox2.SelectedIndex = 0;
            listBox2.SelectedIndex = -1;
            rf_filter();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
           // comboBox2.SelectedIndex = 0;
            listBox2.SelectedIndex = -1;
            rf_filter();
        }

        private void label15_Click(object sender, EventArgs e)
        {
            textBox1.Text += label15.Text;
        }

        private void label14_Click(object sender, EventArgs e)
        {

            textBox1.Text += label14.Text;
        }

        private void label16_Click(object sender, EventArgs e)
        {

            textBox1.Text += label16.Text;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where Convert.ToBoolean(dd.Cells["choose"].Value) == true 
                      select dd.Cells["地址"].Value.ToString();

            foreach (string pp in kkk)
            {

             
                if (File.Exists(pp))
                {
                    FormMethod.backupfile(pp);
                    File.Delete(pp);


                }
                DbHelperSQL.ExecuteSql("delete from paperwork where 文件地址='" + pp.Replace("\\","\\\\") + "';");
                

            }


            MessageBox.Show("执行成功");
            get_datatable();
            rf_filter();


        }

        private void button10_Click(object sender, EventArgs e)
        {
            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where Convert.ToBoolean(dd.Cells["choose"].Value) == true 
                      select dd.Cells["地址"].Value.ToString();

            string outputfolder = Properties.Settings.Default.filepath.Replace("prepare","output") + "files\\";
            foreach (string pp in kkk)
            {
                File.Copy(pp, outputfolder + pp.Split('\\').Last(), true);


            }

            System.Diagnostics.Process.Start("explorer.exe", outputfolder);

        }

        private void button11_Click(object sender, EventArgs e)
        {
            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where Convert.ToBoolean(dd.Cells["choose"].Value) == true
                      select dd.Cells["地址"].Value.ToString();

           
            foreach (string pp in kkk)
            {
                var doc = wordMethod.opendoc(pp);
                string range = textBox8.Text;
                if(range!="")
                {
                    doc.PrintOut(Range:4, Pages: range);
                }
                else
                {
                    doc.PrintOut();
                }
            
                doc.Close();

            }

            
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where dd.Cells["文件名"].Value.ToString().Contains("C017")
                      select dd.Cells["choose"];
            if (checkBox9.Checked == true)
            {
                //  int count = Convert.ToInt16(dataGridView1.Rows.Count.ToString());


                foreach (var dd in kkk)
                {
                    dd.Value = true;

                }

            }
            else
            {
                foreach (var dd in kkk)
                {
                    dd.Value = false;

                }


            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where dd.Cells["文件名"].Value.ToString().Contains("C023")
                      select dd.Cells["choose"];
            if (checkBox10.Checked == true)
            {
                //  int count = Convert.ToInt16(dataGridView1.Rows.Count.ToString());


                foreach (var dd in kkk)
                {
                    dd.Value = true;

                }

            }
            else
            {
                foreach (var dd in kkk)
                {
                    dd.Value = false;

                }


            }
        }
    }
}
