using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GoumangToolKit;
using System.IO;
using OfficeMethod;
using FileManagerNew;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;

namespace AUTORIVET_KAOHE
{
    public partial class Production : Form,FormInterface
    {
        //bool startdate;
       // bool enddate;
       // bool transferdate;

        //当前操作的位置
        int currentrow=0;
        int currentpos = 0;
        DataTable unidt;
        private MySqlDataAdapter daMySql = new MySqlDataAdapter();
      Dictionary<string,string[]> prodDic= Program.prodTable.Select().ToDictionary(aa => aa["名称"].ToString(), bb =>new string[2] { bb["图号"].ToString() ,bb["站位号"].ToString() }  );
        
        public Production()
        {
            InitializeComponent();
        }

        private void getupdate(DataTable dt)
        {
            //dt = dataGridView1.DataSource as DataTable;//把DataGridView绑定的数据源转换成DataTable 

            MySqlCommandBuilder cb = new MySqlCommandBuilder(daMySql);

            daMySql.Update(dt);

        }
        #region interface
        public  void rf_gridview(dynamic dt)
        {

            dataGridView1.DataSource = dt;

        }

     public  void rf_default()
       {
            unidt = new DataTable();
            var drawingcol = unidt.Columns.Add("图号", typeof(string));
          var stationcol = unidt.Columns.Add("站位号", typeof(string));
            MySqlConnection MySqlConn = new MySqlConnection(PubConstant.ConnectionString);
            MySqlConn.Open();
            String sql = "select   产品名称, 产品架次, 开始日期, 结束日期, 移交日期, 当前状态, 备注, 流水号 as id,状态说明 from 产品流水表 ORDER BY 开始日期";
            daMySql = new MySqlDataAdapter(sql, MySqlConn);
            // DataSet OleDsyuangong = new DataSet();

            daMySql.Fill(unidt);

            foreach(DataRow dd in unidt.Rows)
            {
                dd["图号"] = prodDic[dd["产品名称"].ToString()][0];
                dd["站位号"] = prodDic[dd["产品名称"].ToString()][1];
            }
      
              
       

           rf_gridview(unidt);
           // startdate = false;
           // enddate = false;
           //transferdate = false;
           dataGridView1.Columns[0].Width = 50;
           dataGridView1.Columns[1].Width = 50;
           dataGridView1.Columns[2].Width = 50;
           dataGridView1.Columns[3].Width = 50;
           dataGridView1.Columns[4].Width = 50;
           dataGridView1.Columns["id"].Width = 20;
            //  dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows[currentpos].Index;
             dataGridView1.Rows[currentrow].Selected = true;
            dataGridView1.FirstDisplayedScrollingRowIndex = unidt.Rows.Count - 1;

        }

    public  void rf_filter()
     {
        
        if (listBox2.SelectedIndex!=-1)
        {

            var kk = from pp in unidt.AsEnumerable()
                     where pp["产品名称"].ToString() == listBox2.SelectedValue.ToString()
                     select pp;


                if (kk.Count()>0)
            {
                rf_gridview(kk.AsDataView());
            }
            else
            {
                dataGridView1.DataSource = null;

            }
       
        }


     }


   public  DataTable get_datatable()
     {

         return AutorivetDB.production_view();
     }






        #endregion















        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int k = e.ColumnIndex;
            int l = e.RowIndex;
            if(l>=0)
            {

         
            currentrow = l;
            dataGridView1.Rows[currentrow].Selected = true;
            currentpos = dataGridView1.FirstDisplayedScrollingRowIndex; 
            DataRowView temprow = (DataRowView)dataGridView1.Rows[l].DataBoundItem;
          //  string liushui = temprow[5].ToString();
            bool manager = Program.ManagerActived;

            switch (k)
            {
                case 0:
                    if(checkBox1.Checked==false)
                {
                        //产品
                    listBox2.SelectedItem = temprow[2].ToString();
                    textBox1.Text = temprow[3].ToString();
                    string startdatestr = temprow[4].ToString();
                    string enddatestr = temprow[5].ToString();
                    string transferdatestr = temprow[6].ToString();

                    if (startdatestr != "" && !startdatestr.Contains("0000"))
                    {
                        dateTimePicker1.Value = DateTime.Parse(startdatestr);
                        checkBox4.Checked = false;
                    }
                    else
                    {
                        checkBox4.Checked = false;
                    }
                    if (enddatestr != "" && !enddatestr.Contains( "0000"))
                    {
                        dateTimePicker2.Value = DateTime.Parse(enddatestr);
                        checkBox3.Checked = false;
                    }
                    else
                    {
                        checkBox3.Checked = true;
                    }


                    if (transferdatestr != "" && !transferdatestr.Contains("0000"))
                    {
                        dateTimePicker3.Value = DateTime.Parse(transferdatestr);
                        checkBox2.Checked = false;
                    }
                        else
                    {
                        checkBox2.Checked = true;
                    }
                    comboBox1.Text = temprow[7].ToString();
                    textBox3.Text = temprow[8].ToString();
                    textBox2.Text = temprow[10].ToString();
                }

                label8.Text = temprow[9].ToString();
               // startdate = true;
              //  enddate = true;
               // transferdate = true;

                  
                    break;
                case 1:
                    RNC f = new RNC();
                    f.Show();
                    f.product_filter=temprow["产品名称"].ToString()+"_"+temprow["产品架次"].ToString();
                   
                  //  f.rncrf(temprow[2].ToString(), temprow[3].ToString());


                    break;
                case 2:

                        //Get the NC program
                        //Auto-copy it to G:\

                        //System.IO.DriveInfo[] allDrivess = System.IO.DriveInfo.GetDrives();
                        //Console.Write(allDrivess.First().VolumeLabel);

                        if (Program.userID.Split('_')[0]=="192.168.13.68"||Program.ManagerActived==true)
                        {

                      

       
       string proname = temprow[0].ToString();

       //     proname = proname + "process";

            string folderpath = Program.InfoPath + temprow["产品名称"].ToString()+"_" + proname+"\\NC\\";

            string progname=   Program.prodTable.Select("图号='" + proname + "'").First()["程序编号"].ToString();
             string filepath = folderpath + "\\" + progname;




            if (System.IO.Directory.Exists(folderpath)&& File.Exists(filepath))

            {
                            //校验MD5
                            
                            string curMD5 = localMethod.GetMD5HashFromFile(filepath);
                            Dictionary<string, string> p1=new Dictionary<string, string>();
                            string preMD5;
                            try
                            {                              
                                string preJson = DbHelperSQL.getlist("select 备注 from 产品数模 where 产品名称='" + temprow["产品名称"].ToString() + "' and 文件类型='Process'").First();

                                JsonSerializer serializer = new JsonSerializer();
                                // JsonReader reader = new JsonTextReader(new StringReader(preJson));
                                StringReader sr = new StringReader(preJson);
                                p1 = (Dictionary<string, string>)serializer.Deserialize(new JsonTextReader(sr), typeof(Dictionary<string, string>));
                                preMD5 = p1["MD5"];
                            }
                            catch(Exception ke)
                            {

                                MessageBox.Show("编程者未正确输出程序，请联系编程者！\r额外信息：\r" + ke.Message);
                                return;
                            }     
                            if (preMD5 == curMD5)
                            {
                                    //Copy all files into the CF card
                                    System.IO.DriveInfo[] allDrives = System.IO.DriveInfo.GetDrives();
                                  var targetDri=  allDrives.Where(x =>x.DriveType.ToString().ToUpper()!="CDROM"&& x.VolumeLabel == "NC_PROGRAM");


                                    if (targetDri.Count()>0)
                            {
                                        
                                    string newfoldername = targetDri.First().RootDirectory.FullName;

                         
                                        BackupOperation.backupfolder(newfoldername);

                                    localMethod.creatDir(newfoldername);
                                  
                                    List<FileInfo> files = new List<FileInfo>();
                                    files.WalkTree(folderpath, false);
                                    files.copyto(newfoldername);
                                    if(p1.ContainsKey("COPYTIME"))
                                    {
                                        p1["COPYTIME"] = DateTime.Now.ToString();
                                    }
                                    else
                                    {
                                        p1.Add("COPYTIME", DateTime.Now.ToString());
                                    }
                               
                               
                                    DbHelperSQL.ExecuteSql("update 产品数模 set 备注='" + localMethod.ToJson(p1) + "' where 产品名称='" + temprow["产品名称"].ToString() + "' and 文件类型='Process'");

                                    MessageBox.Show("请在操作记录或相关文件上记录MD5值：\r"+curMD5+"\r\t该程序的生成时间为：\r"+p1["TIME"]);
                                    System.Diagnostics.Process.Start("explorer.exe", newfoldername);


                                }

                          else
                                {

                                    MessageBox.Show("未连接到CF卡（NC_PROGRAM），请确保你在我的电脑中已看到CF卡的盘符！");
                                    return;

                                }
                                  



                            }
                            else
                            {
                                MessageBox.Show("MD5校验未通过，联系编程者重新输出！");
                                return;
                            }

                        }
            else
                        {

                            MessageBox.Show("NC程序不存在，联系工艺员或编程者！");
                            return;
                        }

                            //  if (manager)
                            //   {



                            //f1.inputValue = proname;
                            //f1.Show();
                        }
         else
                        {
                            MessageBox.Show("必须在自动钻铆现场工作站上使用，并插入CF卡！");

                        }
                        break;

                case 3:

                    paperWork paperf = new paperWork();
                    paperf.Show();
                    paperf.ppworkrf(temprow[2].ToString(), temprow[3].ToString());
                    break;


                case 4:

              
     






           
          //查看试片列表

                   string prodname=temprow["图号"].ToString();
                   string chnname=temprow["产品名称"].ToString();
                   string fuseno=temprow["产品架次"].ToString();
                    string foldername=Program.InfoPath+chnname+ "_"+prodname+"\\"+fuseno+"\\Inspection\\";

                        //刷新试片编号

                      DbHelperSQL.ExecuteSqlTran( AutorivetDB.rfcouponno(prodname));


                    localMethod.creatDir(foldername);


                    System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(foldername);
            int fileNum = dir.GetFiles().Length;

                    if(fileNum>2)
                    {
                        System.Diagnostics.Process.Start("explorer.exe", foldername);
                        return;
                    }

            //               Microsoft.Office.Interop.Excel.Application app =
            //  new Microsoft.Office.Interop.Excel.ApplicationClass();


            //app.Visible = false;
            string processtb = prodname.Replace("-001", "Process");
            DataTable coupon_table = DbHelperSQL.Query("select * from (select * from 试片列表 where 产品图号='" + prodname + "') aa inner join " + processtb + " bb  on aa.程序段编号=bb.UUID").Tables[0];
                //查询程序编号

         string programno=   AutorivetDB.queryno(prodname, "程序编号");


                foreach (DataRow pp in coupon_table.Rows)
                {
                    string progpart=pp["程序段编号"].ToString();

                   
                    //取回程序段的ID号
                    string progid = pp["ID"].ToString();



                    string couponno = "T"+pp["编号"].ToString();
                    string fstname=progpart.Split('_')[1];

                    string skinthk = pp["蒙皮厚度"].ToString();
                    string secmat = pp["二层材料"].ToString();
                    string secthk = pp["二层厚度"].ToString();
                    if (skinthk.Length < 3)
                    {
                        skinthk = "0" + skinthk;

                    }
                    if (secthk.Length < 3)
                    {
                        secthk = "0" + secthk;

                    }
                
                    if(fstname.Contains("B0206002"))
                    {

                       var wBook =new excelMethod(Program.InfoPath + "SAMPLE\\COUPON\\HI_LITE.xls");
                        string filename = foldername + progid + "_" + couponno + "_" + progpart + "_" + secthk + ".xls";
                        if(File.Exists(filename))
                        {
                            continue;
                        }


                        wBook.SaveAs(filename);

                         wBook.Set_CellValue(2,8, (programno + "/" + progid + "/" + progpart)) ;
                                wBook.Set_CellValue(5, 8, prodname);
                              //  wBook.Set_CellValue(7, 8, fuseno);
                                wBook.Set_CellValue(9, 8,couponno);
                                wBook.Set_CellValue(11, 4,fstname);
                                //蒙皮试片


                                wBook.Set_CellValue(9, 5, "C1-SKIN-" + skinthk);



                                //二层材料


                                wBook.Set_CellValue(9, 6,"C1-" + secmat + "-" + secthk);
                       




                        wBook.Save();
                                wBook.Quit();

                    }

                    else
                    {


                    var wBook =new excelMethod(Program.InfoPath + "SAMPLE\\COUPON\\RIVET.xls");
                  //   progid = pp["ID"].ToString();
                        string filename = foldername + progid + "_" + couponno + "_" + progpart + "_" + secthk + "_BEFORE.xls";

                        if (File.Exists(filename))
                        {
                            continue;
                        }

                        wBook.SaveAs(filename);
                                //  Microsoft.Office.Interop.Excel.Worksheet wSheet = wBook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                                wBook.Set_CellValue(2, 10,programno + "/" + progid+"/"+progpart);
                                wBook.Set_CellValue(5, 10, prodname);
                                //wBook.Set_CellValue(7, 10, fuseno);
                                wBook.Set_CellValue(12, 10, couponno+"_BEFORE");
                                wBook.Set_CellValue(14, 6, fstname);
                                //蒙皮试片

                                wBook.Set_CellValue(12, 7,"C1-SKIN-" + skinthk);



                                //二层材料


                                wBook.Set_CellValue(12, 8, "C1-" + secmat + "-" + secthk);
                        




                        wBook.Save();

                                wBook.Set_CellValue(12, 10,couponno + "_AFTER");

                        filename = foldername + progid + "_" + couponno + "_" + progpart + "_" + secthk + "_AFTER.xls";
                        if (File.Exists(filename))
                        {
                            continue;
                        }

                        wBook.SaveAs(filename);

                                wBook.Quit();





                            }


                

                }

                        //生成产品检查单

                        string newpath = foldername+prodname+ "_INSPECTION.doc";
                        File.Copy(Program.InfoPath + "SAMPLE\\PRODUCT_INSPECTION.docx", newpath, true);
                        var myAO = wordMethod.opendoc(newpath);

                        wordMethod.SearchReplace(myAO, "[1]", prodname);
                       // wordMethod.SearchReplace(myAO, "[2]", fuseno);
                        wordMethod.SearchReplace(myAO, "[3]", chnname);
                        myAO.Save();

                        myAO.Application.Quit();

                        //生成参数检查单
                        newpath = foldername + prodname + "_PARA_INSPECTION.xlsx";
                        var tianchong = new Dictionary<string, DataTable>();

                        // List<string> tianchongname=new List<string> ();


                        tianchong.Add(prodname, AutorivetDB.getparatable(prodname));
                        //  tianchongname.Add(productnametrim);

                        OfficeMethod.excelMethod.SaveDataTableToExcelTran(tianchong, newpath);


                        System.Diagnostics.Process.Start("explorer.exe", foldername);

                   
                    break;
                
                
                
                
                
                
                
                //case 3:



               //     Dictionary<string, string> tmp = new Dictionary<string, string>();
               //     tmp.Add("类型", "补铆");
               //     //建立补铆文件夹
               //     string rootdir =Program.InfoPath  + temprow[2].ToString() + "_" + temprow[0].ToString() + "\\" + temprow[3].ToString() + "\\补铆\\" ;
                  
                    
               //     if (!System.IO.Directory.Exists(rootdir))
               //     {
               //         localMethod.creatDir(rootdir);

               //     }
               //     System.Diagnostics.Process.Start("explorer.exe", rootdir);
               //     tmp.Add("目录地址", rootdir);
               //     //用于信息传递

                    
               //     //名称
               // string prodname = temprow[2].ToString();
                
               // tmp.Add("中文名称", prodname+"壁板");
               // tmp.Add("名称", prodname );



               ////架次
               //string jiaci = temprow[3].ToString();
               //tmp.Add("架次", jiaci);

  



               //     tmp.Add("保存地址", rootdir  + temprow[0].ToString() +"_"+ temprow[1].ToString() + "_补铆.doc");


               //     AO f2=new AO();

               //     f2.rncaao=tmp;
               //     f2.tianchong();
               //     f2.Show();



               //     break;


       

            }











            }   



        }

        public void shaixuan(string prodectname, string jiaci="")
        {
            string sqlstr;
            if(jiaci=="")
            {
                sqlstr = "select 图号,站位号,产品名称,产品架次,开始日期,结束日期,移交日期,当前状态,备注,流水号,状态说明 from 产品流水表 a left join 产品列表 b on a.产品名称=b.名称 where 产品名称='" + prodectname + "'";
            }
            else
            {
                sqlstr = "select 图号,站位号,产品名称,产品架次,开始日期,结束日期,移交日期,当前状态,备注,流水号,状态说明 from 产品流水表 a left join 产品列表 b on a.产品名称=b.名称 where 产品名称='" + prodectname + "' and 产品架次='" + jiaci + "'";
            }
            dataGridView1.DataSource = DbHelperSQL.Query(sqlstr).Tables[0];
        }

        private void Production_Load(object sender, EventArgs e)
        {


            listBox2.DataSource = prodDic.Keys.ToList();
            listBox2.SelectedIndex = -1;

            DataGridViewButtonColumn col_btn_insert = new DataGridViewButtonColumn();
            col_btn_insert.HeaderText = "操作1";
            col_btn_insert.Text = "编辑";//加上这两个就能显示
            col_btn_insert.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert);
           // dataviewfresh2();
            DataGridViewButtonColumn col_btn_insert2 = new DataGridViewButtonColumn();
            col_btn_insert2.HeaderText = "操作2";
         
            col_btn_insert2.Text = "RNC";//加上这两个就能显示
            col_btn_insert2.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert2);
            DataGridViewButtonColumn col_btn_insert3 = new DataGridViewButtonColumn();
            col_btn_insert3.HeaderText = "操作3";
            col_btn_insert3.Text = "程序";//加上这两个就能显示
            col_btn_insert3.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert3);
            //DataGridViewButtonColumn col_btn_insert4 = new DataGridViewButtonColumn();
            //col_btn_insert4.HeaderText = "操作4";
            //col_btn_insert4.Text = "补铆";//加上这两个就能显示
            //col_btn_insert4.UseColumnTextForButtonValue = true;
            //dataGridView1.Columns.Add(col_btn_insert4);
            DataGridViewButtonColumn col_btn_insert5 = new DataGridViewButtonColumn();
            col_btn_insert5.HeaderText = "操作4";
            col_btn_insert5.Text = "资料";//加上这两个就能显示
            col_btn_insert5.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert5);


            DataGridViewButtonColumn col_btn_insert6 = new DataGridViewButtonColumn();
            col_btn_insert6.HeaderText = "操作5";
            col_btn_insert6.Text = "打印";//加上这两个就能显示
            col_btn_insert6.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert6);



            rf_default();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
       

        private void dateTimePicker1_MouseDown(object sender, MouseEventArgs e)
        {
           // startdate = true;

        }

        private void dateTimePicker2_MouseDown(object sender, MouseEventArgs e)
        {
          //  enddate = true;

        }

        private void dateTimePicker3_MouseDown(object sender, MouseEventArgs e)
        {
          //  transferdate = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string startdatestr = "";
            if (!checkBox4.Checked)
            {
                startdatestr = dateTimePicker1.Value.ToShortDateString();
                //startdatestr = dateTimePicker1.Value.ToShortDateString();
            }
            string enddatestr = "";
            if (!checkBox3.Checked)
            {
                enddatestr = dateTimePicker2.Value.ToShortDateString();
            }
            string transferdatestr = "";
            if (!checkBox2.Checked)
            {
                transferdatestr = dateTimePicker3.Value.ToShortDateString();
            }

           

            DbHelperSQL.ExecuteSql("insert into 产品流水表(产品名称,产品架次,开始日期,结束日期,移交日期,状态说明,当前状态) values('" + listBox2.SelectedValue + "','" + textBox1.Text + "','" + startdatestr + "','" + enddatestr + "','" + transferdatestr + "','" + textBox3.Text + "','" + comboBox1.Text + "')");
            rf_default();
            FormMethod.notifyServers("插入了新产品:" + listBox2.SelectedValue+",架次："+ textBox1.Text);
            MessageBox.Show("执行成功");
        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string startdatestr = "0000/0/0";
            if (!checkBox4.Checked)
            {
                startdatestr = dateTimePicker1.Value.ToShortDateString();
                //startdatestr = dateTimePicker1.Value.ToShortDateString();
            }
            string enddatestr = "0000/0/0";
            if (!checkBox3.Checked)
            {
                enddatestr = dateTimePicker2.Value.ToShortDateString();
            }
            string transferdatestr = "0000/0/0";
            if (!checkBox2.Checked)
            {
                transferdatestr = dateTimePicker3.Value.ToShortDateString();
            }
         //  var curdt = (DataView)dataGridView1.DataSource;

            //var temprow = (from DataRow dd in curdt
            //               where dd["id"].ToString() == label8.Text
            //               select dd).First();

           var temprow = unidt.Select("id=" + label8.Text).First();
          




            temprow["产品名称"] = listBox2.SelectedValue;
            temprow["产品架次"] = textBox1.Text;
            
            temprow["备注"] = textBox3.Text;
            temprow["状态说明"] = textBox2.Text;
            temprow["当前状态"] = comboBox1.Text;



            temprow["开始日期"] = new MySql.Data.Types.MySqlDateTime(startdatestr);
            temprow["结束日期"] = new MySql.Data.Types.MySqlDateTime(enddatestr);
            temprow["移交日期"] = new MySql.Data.Types.MySqlDateTime(transferdatestr);

            FormMethod.notifyServers("更新产品:" + listBox2.SelectedValue + ",架次：" + textBox1.Text);


            //DbHelperSQL.ExecuteSql("update 产品流水表 set 产品名称='" + listBox2.SelectedValue + "',产品架次='" + textBox1.Text + "',开始日期='" + startdatestr + "',结束日期='" + enddatestr + "',移交日期='" + transferdatestr + "',备注='" + textBox3.Text + "',当前状态='" + comboBox1.Text + "',状态说明='" + textBox2.Text + "' where 流水号=" + label8.Text);

            temprow.EndEdit();

            getupdate(unidt);
         // rf_filter();
             //   rf_default();
          
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker3.Value;
        }


        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataTable output = (DataTable)dataGridView1.DataSource;
       var abc = excelMethod.SaveDataTableToExcel(output);
           // abc.Rows.AutoFit();
            if(checkBox5.Checked==true)
            {
    
                 string path = @"\\192.168.3.32\Autorivet\output\INFO\backup\";
                excelMethod.SaveAs(abc,path + "production.xlsx");
                 System.Diagnostics.Process.Start("explorer.exe", path);

            }


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox2_DoubleClick(object sender, EventArgs e)
        {
            rf_filter();
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void button5_Click(object sender, EventArgs e)
        {
            rf_gridview(unidt);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if(dataGridView1.SelectedRows.Count>0)
            {
                int row = dataGridView1.SelectedRows[0].Index;
                DataRow temprow = ((DataRowView)dataGridView1.SelectedRows[0].DataBoundItem).Row;
               
                FormMethod.notifyServers("删除了产品:" + temprow["产品名称"].ToString() + ",架次：" + temprow["产品架次"].ToString());
                temprow.Delete();
                temprow.EndEdit();
                getupdate(unidt);

            }
          
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var dd = from pp in unidt.AsEnumerable()
                     group pp by pp["产品架次"].ToString() into m
                     select m;
            var ff = from pp in unidt.AsEnumerable()
                     group pp by pp["图号"].ToString() into n
                     select n;
                    
                    
           
            DataTable matrix = new DataTable();
            matrix.Columns.Add("图号");
            matrix.Columns.Add("名称");

            foreach (var item in dd.OrderBy(o=>o.Key))
            {
                matrix.Columns.Add(item.Key);
               
            }
            foreach (var item in ff.OrderBy(o=>o.Key))
            {
                var dr = matrix.Rows.Add();
                dr["图号"] =item.Key;
             
                foreach (var pp in item)
                {
                    dr["名称"] = pp["产品名称"].ToString();
                    dr[pp["产品架次"].ToString()] = pp["开始日期"].ToString();
                }
            }

            excelMethod.SaveDataTableToExcel(matrix);


        }
    }
}
