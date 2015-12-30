using FileManagerNew;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Threading;
using GoumangToolKit;
using OfficeMethod;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using MySql.Data.MySqlClient;
using CATIA_method;

namespace AUTORIVET_KAOHE
{
    public partial class CATIA_model : Form,FormInterface
    {
        DataTable unidt;
        Thread CATIA_task;
       
        private MySqlDataAdapter daMySql = new MySqlDataAdapter();
       
        public CATIA_model()
        {
         
            InitializeComponent();
        }

  
    

        private void CATIA_model_Load_1(object sender, EventArgs e)
        {

            

            DataGridViewButtonColumn col_btn_insert = new DataGridViewButtonColumn();
            col_btn_insert.HeaderText = "操作1";
            col_btn_insert.Text = "打开";//加上这两个就能显示
            col_btn_insert.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert);
            // dataviewfresh2();
            DataGridViewButtonColumn col_btn_insert2 = new DataGridViewButtonColumn();
            col_btn_insert2.HeaderText = "操作2";
            col_btn_insert2.Text = "目录";//加上这两个就能显示
            col_btn_insert2.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert2);
            //DataGridViewButtonColumn col_btn_insert3 = new DataGridViewButtonColumn();
            //col_btn_insert3.HeaderText = "操作3";
            //col_btn_insert3.Text = "检查";//加上这两个就能显示
            //col_btn_insert3.UseColumnTextForButtonValue = true;
            //dataGridView1.Columns.Add(col_btn_insert3);
            DataGridViewButtonColumn col_btn_insert4 = new DataGridViewButtonColumn();
            col_btn_insert4.HeaderText = "操作4";
            col_btn_insert4.Text = "替换"; //加上这两个就能显示
            col_btn_insert4.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert4);

            DataGridViewCheckBoxColumn col_btn_insert5 = new DataGridViewCheckBoxColumn();
            col_btn_insert5.HeaderText = "选择";
            col_btn_insert5.Name = "choose";
            dataGridView1.Columns.Add(col_btn_insert5);
         // Form_method.gridview_unirf(this);
            listBox1.DataSource = AutorivetDB.foldername_list();
            listBox1.SelectedIndex = -1;
            listBox2.DataSource = DbHelperSQL.getlist("select concat(产品名称,'_',产品架次) from 产品流水表");
            listBox2.SelectedIndex = listBox2.Items.Count-1;
            rf_default();

        }
        public void rf_default()
        {
            unidt = new DataTable();
            using (MySqlConnection MySqlConn = new MySqlConnection(PubConstant.ConnectionString))
            {

           
                MySqlConn.Open();
            string sql = "select * from 产品数模";
            daMySql = new MySqlDataAdapter(sql, MySqlConn);
            // DataSet OleDsyuangong = new DataSet();

            daMySql.Fill(unidt);
            rf_gridview(unidt);
                MySqlConn.Close();
            }
        }
       public void rf_gridview(dynamic dt)
        {
            
            dataGridView1.DataSource = dt;
            if (dt!=null)
            {


            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[2].Width = 50;
            dataGridView1.Columns[3].Width = 50;
            dataGridView1.Columns[4].Width = 200;
            }
            // dataGridView1.Columns[5].Width =100;
        }

       public void rf_filter()
       {
            // DataTable curdt = unidt;
            DataView curdtView=unidt.AsDataView();
           if (listBox1.SelectedIndex > -1)
           {


               string prodname = listBox1.SelectedItem.ToString().Split('_')[0];

               var cc = unidt.AsEnumerable().Where(p => p["产品名称"].ToString() == prodname);

               if (cc.Count() > 0)
               {
                    curdtView = cc.AsDataView();
               }


               else
               {
                    curdtView = null;
               }


           }
           if (comboBox4.SelectedIndex > -1)
           {


               string filetype = comboBox4.Text;

               var cc = unidt.AsEnumerable().Where(p => p["文件类型"].ToString() == filetype);

               if (cc.Count() > 0)
               {
                    curdtView = cc.AsDataView();
               }


               else
               {
                    curdtView = null;
               }

           }

           rf_gridview(curdtView);


       }


       public DataTable get_datatable()
       {
           unidt = AutorivetDB.catiaModel_show();
           return unidt;

       }












       private void button2_Click(object sender, EventArgs e)
       {
           if (listBox1.SelectedIndex == -1 || comboBox1.SelectedIndex == -1)
           {
               MessageBox.Show("请选择产品及类型");
           }
           else
           {
               OpenFileDialog ofd = new OpenFileDialog();
               ofd.ShowDialog();
               string orgname = ofd.SafeFileName;
               if (orgname != "")
               {
                   string foldername=listBox1.SelectedValue.ToString();
                   string prodname = foldername.Split('_')[0];
                   string drawingname = foldername.Split('_')[1];
                  // DataRow pp = DbHelperSQL.Query("select 图号,站位号 from 产品列表 where 名称='" + prodname + "'").Tables[0].Rows[0];

                   string filepath = "";
                   filepath = @"\\192.168.3.32\Autorivet\prepare\INFO\" + foldername + @"\MODEL";
                   string filefullname = "";
                   string filename = "";
                   string fileExt = comboBox1.Text;
                   switch (comboBox1.Text)
                   {

                            
                       case "TVA":


                           filepath = filepath + @"\TVA";
                           filename = orgname.Substring(0, 13) + "_TVA.CATPart";
                           filefullname = filepath + "\\" + filename;
                         //  fileExt = ".CATPart";


                           break;
                       case "Product":


                           filepath = filepath + @"\Product";
                           filename = drawingname + "_"+comboBox3.Text+".CATProduct";
                           filefullname = filepath + "\\" + filename;
                         //  fileExt = ".CATProduct";


                           break;
                       case "Process":

                           filepath = filepath + @"\Process";

                           filename = drawingname + "_Process.CATProcess";
                           filefullname = filepath + "\\" + filename;
                          // fileExt = ".CATProcess";


                           break;
                       case "cgr":

                           filepath = filepath + @"\cgr";


                           filename = drawingname + "_" + comboBox3.Text + ".cgr";
                           filefullname = filepath + "\\" + filename;
                        //   fileExt = ".cgr";


                           break;
                       case "Drawing":

                           filepath = filepath + @"\TVA";


                           filename = drawingname + "_Drawing"  + ".CATDrawing";
                           filefullname = filepath + "\\" + filename;
                           //   fileExt = ".cgr";


                           break;


                       default:

                           break;


                   }

                   localMethod.creatDir(filepath);
                   try
                   {
                       File.Copy(ofd.FileName, filefullname);
                   }
                   catch (SystemException kk)
                   {
                       MessageBox.Show(kk.Message);
                   }

                   List<FileInfo> tempfile = new List<FileInfo>();
                   tempfile.Add(new FileInfo(filefullname));
                   //扫描该文档
                   FormMethod.scanfileCatia(tempfile);
                   get_datatable();
                   rf_filter();
               }
           }

       }









       void dataGridView1_CellContentClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
       {

           int k = e.ColumnIndex;
           int l = e.RowIndex;
          if(k<3&&l>=0)
          {
              
                DataRowView temprowview = (DataRowView)dataGridView1.Rows[l].DataBoundItem;
                var temprow = temprowview.Row;
           string filepath = temprow["文件地址"].ToString();
           string filename = temprow["文件名"].ToString();
           string folderpath = filepath.Replace(filename, "");
                //JObject jObjectorg = new JObject();
                //if(temprow["历史"].ToString()!="")
                //{
                //    jObjectorg= JObject.Parse(temprow["历史"].ToString());
                //}
                //else
                //{
                //    jObjectorg.Add("history",null);
                //}
                JsonSerializer serializer = new JsonSerializer();
                StringReader sr = new StringReader(temprow["历史"].ToString());

                List<historyJsonModel> p1 =jsonMethod.ReadFromStr(temprow["历史"].ToString());
               // dataGridView1.Rows[l].Cells["历史"].ToolTipText = p1.JsonToString();



                var jitem = new historyJsonModel()
                {
                    writer = Program.userID,
                    date = DateTime.Now.ToLongDateString(),
                    time = DateTime.Now.ToLongTimeString(),
                    //filepath = filepath,
                };

                switch (k)
           {
               case 0:

                        //将该操作记录于历史当中
                        //使用json

                        jitem.operation = "打开文件";

           
                        //string sb;
                    
                  
                   TVA_Method.setVis(true);
                   if(temprow["文件类型"].ToString()=="Process")
                   {
                       //打开Process前先备份
                       FormMethod.open_process(filepath);
                   }
                   else
                   {
                      FormMethod.open_file(filepath);
                   }
                 break;

               case 1:
                        jitem.operation = "打开目录";
                        System.Diagnostics.Process.Start("explorer.exe", folderpath);
                   break;

               case 2:
                        
                        //替换文件
                        OpenFileDialog ofd = new OpenFileDialog();
              if( ofd.ShowDialog()==DialogResult.OK)
              {
                  //备份原有文件
                  //已制作方法
       
                 BackupOperation.backupfile(filepath);
                 // File.Copy(filepath, backupfolder + backupname,true);
                  string sourcepath = ofd.FileName;
                  File.Copy(sourcepath, filepath, true);
                  MessageBox.Show("替换完成");
                  jitem.operation = "替换文件为:"+ sourcepath;

                        }             
                   break;



                     //   break;


                }
                p1.Add(jitem);
                p1.Sort((a, b) => DateTime.Parse(b.date).CompareTo(DateTime.Parse(a.date)));
                
                StringWriter sb = new StringWriter();
                serializer.Serialize(new JsonTextWriter(sb), p1.Take(8));

                temprow["历史"] = sb.ToString();
               dataGridView1.Rows[l].Cells["历史"].ToolTipText = p1.JsonToString();
                temprow.EndEdit();

                getupdate(unidt);




            }
       }

        private void getupdate(DataTable dt)
        {
            //dt = dataGridView1.DataSource as DataTable;//把DataGridView绑定的数据源转换成DataTable 

            MySqlCommandBuilder cb = new MySqlCommandBuilder(daMySql);

            daMySql.Update(dt);

        }



        private  void check_product(object filepathpara)
       {

            button17.Enabled = false;

           var unboxlist = (List<string []>)filepathpara;

            foreach (string [] kk in unboxlist)
            {

        
           //var aa = new TVA_Method(kk[1], kk[2]);

           //excelMethod.SaveDataTableToExcel(aa.TVAPoints.outputtable());

         
             
             Process_Method aa = new Process_Method(kk[0]);
  

      try
                {

         
           aa.iniProc();
           aa.iniTVA(kk[1]);
                //var pp=aa.TVA;
                //    pp.coordswitch=true;
                //      pp.ifvec=false;
                //aa.GET_OPNAMES();
           var tt = aa.PFpoints;

         //  excelMethod.SaveDataTableToExcel(tt.outputtable());
        
             tt.outputdb(kk[0].Split('\\').Last().Replace("-001_Process.CATProcess","PF"));



              if  ( unboxlist.Count>2)
                {
                    aa.close();
                }
              else

                {
                    MessageBox.Show("成功更新PF数据库");
                }



                }
                catch
                {
                    MessageBox.Show("同步PF数据库出错!");
                }


            }
            button17.Enabled = true;


        }
        private void open_process(object filepathpara)
        {

            button17.Enabled = false;

            var unboxlist = (List<string[]>)filepathpara;

            foreach (string[] kk in unboxlist)
            {


                //var aa = new TVA_Method(kk[1], kk[2]);

                //excelMethod.SaveDataTableToExcel(aa.TVAPoints.outputtable());

              

                Process_Method aa = new Process_Method(kk[0]);




            }
            button17.Enabled = true;


        }
        private void process_checklist(object filepathpara)
        {

            button17.Enabled = false;

            var unboxlist = (List<string[]>)filepathpara;

            foreach (string[] kk in unboxlist)
            {


                //var aa = new TVA_Method(kk[1], kk[2]);

                //excelMethod.SaveDataTableToExcel(aa.TVAPoints.outputtable());



                Process_Method aa = new Process_Method(kk[0]);

                aa.iniProc();
                aa.iniTVA(kk[1]);
                //var pp=aa.TVA;
                //    pp.coordswitch=true;
                //      pp.ifvec=false;
                //aa.GET_OPNAMES();


                //  myAO.SaveAs2(kk[0].Replace(kk[0].Split('_').Last(), "CheckList.doc"));

                //当前涉及的数模文件：

                var pdlist = aa.GET_PRODUCTSNAMES();
                string pdlststr = "";
                foreach(var pp in pdlist)
                {
                    pdlststr = pdlststr + pp + "^p";

                }




   


            string  prodname  = kk[0].Split('\\').Last().Split('_').First();

                var progls = DbHelperSQL.getlist("select UUID from " + prodname.Replace("-001" , "Process"));
                string proglist = "";
                foreach( string pp in progls)
                {
                    proglist = proglist + pp + "^p";
                }
                //程序分段:
           


                //Delmia Program Task:
                string taskliststr = "";
                    var ttls = aa.GET_TASKNAMES();
                foreach (string pp in ttls)
                {
                    taskliststr = taskliststr + pp + "^p";
                }

           

                var tt = aa.PFvsTVA();

                //  excelMethod.SaveDataTableToExcel(tt.outputtable());
                DataTable tvaonly = tt[0].outputtable();
                DataTable proconly = tt[1].outputtable();
                string comparetext = "";
                if (tvaonly.Rows.Count>0)
                {

             
   
                  if   (tvaonly.Rows.Count<10)
                    {
                        comparetext = "TVA存在Process不存在的点位：\r\nID,UUID,PFName,FrameName,FastenerName,ProcessType^p";

                        foreach ( DataRow dr in tvaonly.Rows)
                {
                    comparetext = comparetext + dr["ID"].ToString() + "," + dr["UUID"].ToString() + "," + dr["PFName"].ToString() + "," + dr["FrameName"].ToString() + "," + dr["FastenerName"].ToString() + "," + dr["ProcessType"].ToString()+"^p";


                }

   

                    }
                  else
                    {

                        
                        comparetext = comparetext + "未更新的点位过多，请更新后检查,具体请看输出窗口";

                    }
                    TVA_locate dd = new TVA_locate();
                    dd.expl = " TVA存在Process不存在的点位：";
                    dd.showPoints = tt[0];
                    dd.Show();


                }

                if (proconly.Rows.Count>0)
                {
                    if (proconly.Rows.Count <10)
                    {


                        comparetext = comparetext + "Process与TVA不同的点位：^p";

                        foreach (DataRow dr in proconly.Rows)
                        {
                            comparetext = comparetext + dr["ID"].ToString() + "," + dr["UUID"].ToString() + "," + dr["PFName"].ToString() + "," + dr["FrameName"].ToString() + "," + dr["FastenerName"].ToString() + "," + dr["ProcessType"].ToString() + "^p";


                        }
                    }
                    else
                    {
                        //MessageBox.Show("未更新的点位过多，请更新后检查");
                        //break;
                       
                        comparetext = comparetext + "未更新的点位过多，请更新后检查,具体请看输出窗口";
                    }

                    TVA_locate dd = new TVA_locate();
                    dd.expl = " Process与TVA加工类型不同的点位：";
                    dd.showPoints = tt[1];
                    dd.Show();


                }






                if  (comparetext=="")
                {
                    comparetext = "经系统比对，TVA与ProcessFeature完全一致^p";
                }


                var cc = aa.PFvsPath();

                if (cc[0].count()>0)
                {
                   

                    if (cc[0].count()<10)
                    {
                        comparetext += "需要编程的ProcessFeature：^p";

                        foreach (DataRow dr in cc[0].outputtable().Rows)
                    {
                        comparetext = comparetext + dr["ID"].ToString() + "," + dr["UUID"].ToString() + "," + dr["PFName"].ToString() + "," + dr["FastenerName"].ToString()  + "^p";


                    }
                    }
                    else
                    {
                        TVA_locate dd = new TVA_locate();
                        dd.expl = " 需要编程的ProcessFeature：";
                        dd.showPoints = cc[0];
                        dd.Show();
                        comparetext = comparetext + "需要的ProcessFeature，请更新后检查,具体请看输出窗口";
                    }

                }

                if (cc[1].count() > 0)
                {
                   

                    if (cc[1].count() < 10)
                    {
                        comparetext += "多余编程的ProcessFeature：^p";
                        foreach (DataRow dr in cc[1].outputtable().Rows)
                        {
                            comparetext = comparetext + dr["ID"].ToString() + "," + dr["UUID"].ToString() + "," + dr["PFName"].ToString()  + "," + dr["FastenerName"].ToString() + "^p";


                        }
                    }
                    else
                    {
                      
                        comparetext = comparetext + "多余编程的ProcessFeature，请更新后检查,具体请看输出窗口";

                    }
                    TVA_locate dd = new TVA_locate();
                    dd.expl = "多余编程的ProcessFeature：";
                    dd.showPoints = cc[1];
                    dd.Show();
                }


         if(cc[0].count()==0&& cc[1].count() == 0)
                {
                    comparetext += "经系统比对，ProcessFeature铆接及钻孔点位均已编程^p";
                }



                string newpath = kk[0].Replace(kk[0].Split('_').Last(), "_PROCESS_CHECKLIST.doc");
                File.Copy(@"\\192.168.3.32\Autorivet\Prepare\INFO\SAMPLE\NC_INSPECTION.docx", newpath, true);
                var myAO = wordMethod.opendoc(newpath);

                try
                {

               
                wordMethod.SearchReplace(myAO, "[1]", pdlststr);
                wordMethod.SearchReplace(myAO, "[2]", proglist);
                wordMethod.SearchReplace(myAO, "[3]", taskliststr);
                wordMethod.SearchReplace(myAO, "[4]", comparetext);
                }
                catch
                {

                    MessageBox.Show("因为你太坑了，寡人决定切了你的小JJ");
                }





            }
            button17.Enabled = true;


        }
        private void show_PF(object filepathpara)
         {

             var kk = (List<string>)filepathpara;

             //kk[0]作为产品图号


           //var aa = new TVA_Method()

           //  //excelMethod.SaveDataTableToExcel(aa.TVAPoints.outputtable());



           //  Process_Method aa = new Process_Method(kk[0]);

           //  aa.iniProc();
           //  aa.iniTVA(kk[1]);
           //  //var pp=aa.TVA;
           //  //    pp.coordswitch=true;
           //  //      pp.ifvec=false;

           //  var tt = aa.PFpoints();

           //  //  excelMethod.SaveDataTableToExcel(tt.outputtable());

           //  tt.outputdb(kk[0].Split('\\').Last().Replace("-001_Process.CATProcess", "PF"));

           //  MessageBox.Show("成功更新PF数据库");

         }
       private  void check_tva(object filepath)
       {



        var skin_prod = from aa in unidt.AsEnumerable()
                             where aa["文件类型"].ToString() == "TVA"
                             join bb in Program.prodTable.AsEnumerable()
                             on aa["产品名称"].ToString() equals bb["名称"].ToString()
                             select new
                             {
                                 TVA = aa["文件名"].ToString(),
                                 TVAPath = aa["文件地址"].ToString(),
                                 DWG = bb["图号"].ToString().Replace("-001", ""),
                                 CHN= bb["名称"].ToString()
                             }
                  ;


            button17.Enabled = false;
            button13.Enabled = true;
           button12.Enabled = false;
           bool ifcheck = this.ifcheck.Checked;
           bool ifcolor = this.ifcolor.Checked;
           bool ifdatabase = ifdt.Checked;
            bool ifstac = ifstatistic.Checked;
           bool fix = checkBox2.Checked;
            bool ifrebuild = checkBox6.Checked;
            bool ifreset = checkBox9.Checked;
            DataTable perTable = new DataTable();
            perTable.Columns.Add("PRODUCT", typeof(string));
            perTable.Columns.Add("PRODUCT NAME", typeof(string));
            perTable.Columns.Add("SUM", typeof(string));
            perTable.Columns.Add("BY MACHINE", typeof(string));
            perTable.Columns.Add("PERCENTAGE", typeof(string));
            perTable.Columns.Add("DOU/WIN", typeof(string));
            foreach (string pp in (List<string>)filepath)
            {
                var infoRow = perTable.Rows.Add();
                var prodinfo = skin_prod.Where(bb => bb.TVA.Substring(0,13) == pp.Split('\\').Last().Substring(0,13)).First();
             //   Console.Write(skin_prod.Count().ToString());
              //  MessageBox.Show(pp.Split('\\').Last());
               
              //  Console.Write(prodinfo.CHN);
                string prodname = prodinfo.DWG;
                TVA_Method tvamodel = new TVA_Method(pp)

                {
                    ifvec = true,
                    coordswitch = false
                };
                TVA_Method.setVis(true);




                    if (ifreset)
                {
                    tvamodel.TreeListtoPointsList(tvamodel.TVATreeList);
                }
                 

                if (fix)
                {
                    //做修复前需要备份文件
                    BackupOperation.backupfile(pp);


                    tvamodel.fix_all();
                }

                if (ifrebuild)
                {

                    //重制前备份文件
                    BackupOperation.backupfile(pp);

                    //同步数据库
                    tvamodel.updatedt();
                    //更新位置


                    AutorivetDB.updateTVAlocation(prodname);
                    tvamodel.rebuild(prodname);



                }



                if (ifcheck)
                {
                    tvamodel.CheckTVA(color: ifcolor, database: ifdatabase);

                    infoRow.ItemArray = (new string[] {prodname, prodinfo.CHN, tvamodel.infoBag["SUM"], tvamodel.infoBag["BY MACHINE"], tvamodel.infoBag["PERCENTAGE"] });


                }
                else
                {
                    if (ifdatabase)
                    {
                        tvamodel.updatedt();
                        AutorivetDB.updateTVAlocation(prodname);

                    }
                    if (ifcolor)
                    {

                        tvamodel.setcolor();

                    }
                    if (ifstac)
                    {
                        var sumcount = DbHelperSQL.getlist("select count(*) from " + prodinfo.DWG).First();
                        string mchesql = "select count(*) from " + prodname + " where ProcessType like '%INSTALLED BY%' or ProcessType like '%DRILL ONLY BY%'";
                        var machinecount = DbHelperSQL.getlist(mchesql).First();

                        infoRow.ItemArray=(new string[] { prodname, prodinfo.CHN, sumcount, machinecount});




                    }


                    }
                if (ifstac)
                {
                   // autorivet_op.updateTVAlocation(prodname);
                    //Get DOU/WIN count
                    string dousql = "select count(*) from " + prodname + " where location='WIN' or location= 'DOU'";
                    infoRow["DOU/WIN"] = DbHelperSQL.getlist(dousql).First();

                }




                }
            if(ifstac)
            {




                excelMethod.SaveDataTableToExcel(perTable);
             
            }

           button12.Enabled = true;
           button13.Enabled = false;
            button17.Enabled = true;
        }
       //private  void  check_tva_finish(Task aa)
       //{

          
       //    MessageBox.Show("Check完成");
       //}



       private void button3_Click(object sender, EventArgs e)
       {
             
           string sestr = listBox2.SelectedItem.ToString();

           //找出产品所用的TVA
           string productname = sestr.Split('_')[0];
           string jiaci = sestr.Split('_')[1];
           //获取产品一级目录

        string firstfolder=   ((List<string>)listBox1.DataSource).Where(k => k.Contains(productname)).First();

     


           DataTable dtmp = unidt;
           var kk = dtmp.AsEnumerable().Where(p => p["产品名称"].ToString() == productname && p["文件类型"].ToString() == "TVA" && p["产品架次"].ToString() == "");
           if(kk.Count()==0)
           {
               MessageBox.Show("尚未输入该产品的TVA");
           }
           else
           {

               DataRow tmprow = kk.First();
               string newfilename = tmprow["文件名"].ToString().Replace(".", "_" + jiaci + ".");
               string folderpath = Program.InfoPath + firstfolder + "\\" + jiaci + "\\";
               string newTVApath = folderpath + tmprow["文件名"].ToString().Replace(".", "_" + jiaci + ".");

               localMethod.creatDir(folderpath);

               if (!File.Exists(newTVApath))
               {
                   File.Copy(tmprow["文件地址"].ToString(), newTVApath);
               }
                  
                  System.Diagnostics.Process.Start(newTVApath);
                  System.Diagnostics.Process.Start("explorer.exe", folderpath);

                  DbHelperSQL.ExecuteSql("insert ignore into 产品数模 (文件名,文件类型,产品名称,产品架次,文件地址) values ('" + newfilename + "','Fuse_TVA','"+tmprow["产品名称"].ToString()+"','"+jiaci+"','"+newTVApath.Replace("\\", "\\\\")+"');");

                //记录该行为
                //记录于历史
                //     autorivet_op.log_insert(Program.userID, "要修改TVA，并生成" + newfilename);
                //通知服务器

                FormMethod.notifyServers("正在生成、修改TVA:" + newfilename);
                  get_datatable();
                  rf_filter();
                //currentfilepath = newTVApath;
                //取得该条目的历史
                var temprow = (from  pp in unidt.AsEnumerable()
                          where pp["文件名"].ToString() == newfilename
                          select pp).First();

           

                List < historyJsonModel > p1 = jsonMethod.ReadFromStr(temprow["历史"].ToString());
                p1.Add(new historyJsonModel
                {
                    operation="创建文件",
                    writer=Program.userID,
                    date=DateTime.Now.ToLongDateString(),
                    time=DateTime.Now.ToLongTimeString()

                });

                temprow["历史"] = p1.ToJson();
                temprow.EndEdit();
                getupdate(unidt);


            }


       }

       private void button4_Click(object sender, EventArgs e)
       {
           if (MessageBox.Show("确定要覆盖当前的TVA么？", "警告", MessageBoxButtons.OKCancel) == DialogResult.OK)
           {

             

      

           string sestr = listBox2.SelectedItem.ToString();

           //找出产品所用的TVA
           string productname = sestr.Split('_')[0];
           string jiaci = sestr.Split('_')[1];
           //获取产品一级目录

       //    string firstfolder = ((List<string>)listBox1.DataSource).Where(k => k.Contains(productname)).First();




           //DataTable dtmp = unidt;
               //查找当前的TVA
           var kk = from p in unidt.AsEnumerable()
                    where p["产品名称"].ToString() == productname && p["文件类型"].ToString() == "TVA" && p["产品架次"].ToString() == ""
                    select p;

               
         //      dtmp.AsEnumerable().Where(p => p["产品名称"].ToString() == productname && p["文件类型"].ToString() == "TVA" && p["产品架次"].ToString() == "");
          
               
               
               if (kk.Count() != 1)
           {
               MessageBox.Show("尚未输入该产品的TVA或有多个TVA,请核查");
               return;
           }
           else
           {

               var kk2 = from p in unidt.AsEnumerable()
                         where p["产品名称"].ToString() == productname && p["文件类型"].ToString() == "Fuse_TVA" && p["产品架次"].ToString() == jiaci
                        select p["文件地址"].ToString();

               if (kk.Count() != 1)
               {
                   MessageBox.Show("尚未输入该产品的架次TVA或有多个架次TVA,请核查");
                   return;
               }

                    // string folderpath = Form_method.get_storefolder(productname ) + jiaci + "\\";
                    var temprow = kk.First();
               string newTVApath = kk2.First().ToString();
               string currentpath = temprow["文件地址"].ToString().ToString();
               //localMethod.creatDir(folderpath);

               if (File.Exists(newTVApath))
               {
                   //备份TVA文件
                   BackupOperation.backupfile(currentpath);

                   File.Copy(newTVApath, currentpath, true);

                    FormMethod.notifyServers("已替换TVA，源文件:" + newTVApath);
                        //写入TVA历史

                        List<historyJsonModel> p1 = jsonMethod.ReadFromStr(temprow["历史"].ToString());
                        p1.Add(new historyJsonModel
                        {
                            operation = "架次替换",
                            writer = Program.userID,
                            date = DateTime.Now.ToLongDateString(),
                            time = DateTime.Now.ToLongTimeString()

                        });

                        temprow["历史"] = p1.ToJson();
                        temprow.EndEdit();
                        getupdate(unidt);



                    }
               else
               {
                   MessageBox.Show("根本不存在该架次的TVA文件，请核对后操作");
                   return;

               }
               //System.Diagnostics.Process.Start(newTVApath);
              // System.Diagnostics.Process.Start("explorer.exe", currentpath.Replace(tmprow["文件名"].ToString(),""));
               MessageBox.Show("替换成功");


           }
       }
       }

    

       private void button1_Click(object sender, EventArgs e)
       {
           excelMethod.SaveDataTableToExcel(get_datatable());

       }

       private void button5_Click(object sender, EventArgs e)
       {
           int sum = dataGridView1.SelectedRows.Count;
           if (sum!=1)
           {
               MessageBox.Show("请选中一行进行操作！");
           }
           else
           {
               DataRow oprow = ((DataRowView)dataGridView1.SelectedRows[0].DataBoundItem).Row;
               string folderpath = FormMethod.get_storefolder(oprow["产品名称"].ToString(), "COS") + comboBox2.Text + "\\";
               localMethod.creatDir(folderpath);
               string newfilename=oprow["文件名"].ToString().Replace(".", "_COS_" + comboBox2.Text + ".");
               string newfilepath = folderpath + newfilename;
               if (File.Exists(newfilepath))
               {
                   if (!(MessageBox.Show("已存在归档的TVA，确定要覆盖已归档的TVA么？", "警告", MessageBoxButtons.OKCancel) == DialogResult.OK))
                    {
                        return;
                    }

               }

             

               if(oprow["文件类型"].ToString()=="TVA")
               {
                   DbHelperSQL.ExecuteSql("replace into 产品数模 (文件名,文件类型,产品名称,文件地址) values ('" + newfilename + "','COS_TVA','" + oprow["产品名称"].ToString() + "','" + newfilepath.Replace("\\", "\\\\") + "');");

               }

               else
               {

                   if(oprow["文件类型"].ToString()=="Drawing")
                   {

                       DbHelperSQL.ExecuteSql("replace into 产品数模 (文件名,文件类型,产品名称,文件地址) values ('" + newfilename + "','COS_Drawing','" + oprow["产品名称"].ToString() + "','" + newfilepath.Replace("\\", "\\\\") + "');");

                   }
                   else
                   {
                       MessageBox.Show("请选择有效的TVA或Drawing！");
                       return;
                   }
            
               }

               File.Copy(oprow["文件地址"].ToString(), newfilepath, true);

              
               rf_filter();
               MessageBox.Show("归档成功！");


           }
       }

       private void button6_Click(object sender, EventArgs e)
       {
             int sum = dataGridView1.SelectedRows.Count;
             if (sum != 1)
             {
                 MessageBox.Show("请选中一行进行操作！");
             }
             else
             {
                 if (MessageBox.Show("确定要删除该项及相关的文件吗？", "警告", MessageBoxButtons.OKCancel) == DialogResult.OK)
                 {
                   
               
                 DataRow oprow = ((DataRowView)dataGridView1.SelectedRows[0].DataBoundItem).Row;
                     string filepath=oprow["文件地址"].ToString();

                     if(File.Exists(filepath))
                     {
                         File.Delete(filepath);
                     }

                     DbHelperSQL.ExecuteSql("delete from 产品数模 where 文件地址='" + filepath.Replace("\\","\\\\") + "';");


                     MessageBox.Show("删除成功！");

                     get_datatable();
                     rf_filter();
                 }

                



             }



       }

       private void button7_Click(object sender, EventArgs e)
       {

           if ((listBox1.SelectedIndex != -1) && (get_datatable() != null))
           {

               string foldername = listBox1.SelectedValue.ToString();
               string prodname = foldername.Split('_')[0];
               //    string drawingname = foldername.Split('_')[1];



               var pathlist = (from p in get_datatable().AsEnumerable()
                               where (p["产品名称"].ToString() == prodname) && (p["文件类型"].ToString() == "Process")
                               select p["文件地址"]);
               if (pathlist.Count() != 1)
               {

                   MessageBox.Show("请确认表格中有且仅有1个该产品的Process文件！");
               }

               else
               {

                    

                    //BACKUP FILE
                

                   string filepath = pathlist.First().ToString();
                    BackupOperation.backupfile(filepath);
                   pathlist = (from p in get_datatable().AsEnumerable()
                               where (p["产品名称"].ToString() == prodname) && (p["文件类型"].ToString() == "TVA")
                               select p["文件地址"]);



                    var para = new string[2];
                    para[0] = (filepath);
                    para[1] = (pathlist.First().ToString());
                    //pathlist = (from p in get_datatable().AsEnumerable()
                    //            where (p["产品名称"].ToString() == prodname) && (p["文件类型"].ToString() == "ALL_Product")
                    //            select p["文件地址"]);
                    //para.Add(pathlist.First().ToString());
                    if (checkBox10.Checked)
                    {
                        CATIA_task = new Thread(new ParameterizedThreadStart(check_product));
                    }
                    else
                    {
                        CATIA_task = new Thread(new ParameterizedThreadStart(open_process));

                    }
                    var aa = new List<string[]>();
                    aa.Add(para);
                    CATIA_task.Start(aa);




                    // ss.Start(para);
                    Thread.Sleep(500);
                   // Form_method.open_process(filepath);
               }


           }
           else
           {
               MessageBox.Show("请选择一个产品！");
           }























       }

       private void button8_Click(object sender, EventArgs e)
       {
           int sum = dataGridView1.SelectedRows.Count;
           if (sum != 1)
           {
               MessageBox.Show("请选中一行进行操作！");
           }
           else
           {
               DataRow oprow = ((DataRowView)dataGridView1.SelectedRows[0].DataBoundItem).Row;

               var signtext= oprow["签名"].ToString();
                var bb = signtext.Split(';');
                if (bb.Count()>=3)
                {
                    if (bb[0]!=""&&! bb[0].Contains("null"))
                    {
                        comboBox5.Text = bb[0].Split('_')[0];
                        comboBox5.Enabled = false;
                    }

                    if (bb[1] != ""&&!bb[1].Contains("null"))
                    {
                        comboBox6.Text = bb[0].Split('_')[0];
                        comboBox6.Enabled = false;
                    }


                 

                    textBox1.Text = bb[2];
                }

           }
       }

       private void button9_Click(object sender, EventArgs e)
       {
           int sum = dataGridView1.SelectedRows.Count;
           if (sum != 1)
           {
               MessageBox.Show("请选中一行进行操作！");
           }
           else
           {
                DataRow oprow = ((DataRowView)dataGridView1.SelectedRows[0].DataBoundItem).Row;

                var signtext = oprow["签名"].ToString();

                var bb = signtext.Split(';');
             while(bb.Count()<=3)
                {
                    signtext += "null"+ bb.Count()+";";
                    bb = signtext.Split(';');
                }
                string writer;
                string checker;
                if (comboBox5.Enabled!=false&& comboBox5.Text!="")
                {
                  writer = comboBox5.Text + "_" + Program.userID + "_" + DateTime.Now.ToString("yyyy-MM-dd_hh-mm");
                    signtext= signtext.Replace(bb[0], writer);

                }


                if (comboBox6.Enabled != false && comboBox6.Text != "")
                {
                   checker = comboBox6.Text + "_" + Program.userID + "_" + DateTime.Now.ToString("yyyy-MM-dd_hh-mm");
                    signtext= signtext.Replace(bb[1],checker);
                }

                signtext= signtext.Replace(bb[2], textBox1.Text);
                dataGridView1.SelectedRows[0].Cells["签名"].Value = signtext;

               DbHelperSQL.ExecuteSql("update 产品数模 set 签名='" + signtext + "' where 文件名='" + oprow["文件名"].ToString() + "';");

             //  rf_filter();

           }
       }

       private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
       {

           comboBox4.SelectedIndex = -1;
           rf_filter();

            selectTVA();

        }
        private void selectTVA()
        {
            //默认选中该产品的Process及TVA
            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where dd.Cells["文件类型"].Value.ToString() == "TVA" || dd.Cells["文件类型"].Value.ToString() == "Process"
                      select dd;
            foreach (var dd in kkk)
            {
                dd.Cells["choose"].Value = true;
            }
        }

       private void button10_Click(object sender, EventArgs e)
       {
           rf_default();
           listBox1.SelectedIndex = -1;
       }

       private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
       {

       }

       private void button11_Click(object sender, EventArgs e)
       {
           if ((listBox1.SelectedIndex != -1) && (get_datatable() != null))
           {

               string foldername = listBox1.SelectedValue.ToString();
               string prodname = foldername.Split('_')[0];
               //    string drawingname = foldername.Split('_')[1];



               var pathlist = (from p in get_datatable().AsEnumerable()
                               where (p["产品名称"].ToString() == prodname) && (p["文件类型"].ToString() == "TVA")
                               select p["文件地址"]);
               if (pathlist.Count() != 1)
               {

                   MessageBox.Show("请确认表格中有且仅有1个该产品的TVA文件！");
               }

               else
               {

                   string filepath = pathlist.First().ToString();
                    FormMethod.notifyServers("正在修改TVA:" + filepath);
                    FormMethod.open_file(filepath);
                   // Form_method.open_process(filepath);
               }


           }
           else
           {
               MessageBox.Show("请选择一个产品！");
           }








          
       }

       private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
       {
           listBox1.SelectedIndex = -1;

           rf_filter();
       }

       private void checkBox3_CheckedChanged(object sender, EventArgs e)
       {
           if (checkBox3.Checked == true)
           {
               int count = Convert.ToInt16(dataGridView1.Rows.Count.ToString());
               for (int i = 0; i < count; i++)
               {
                    
                    DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells["choose"];
                   checkCell.Value = true;


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

       private void button12_Click(object sender, EventArgs e)
       {
           
           var kkk = from DataGridViewRow dd in dataGridView1.Rows
                     where Convert.ToBoolean(dd.Cells["choose"].Value) == true && dd.Cells["文件类型"].Value.ToString().Contains("TVA")
                     select dd.Cells["文件地址"].Value.ToString();
           //dataGridView1.Rows = kkk;

    //  var token = CATIA_task_cancel.Token;
         //  Thread t1 = new Thread(new ParameterizedThreadStart(replacethread));
          // t1.Start(strobj);
           CATIA_task = new Thread(new ParameterizedThreadStart(check_tva)); 
             //  new Thread(check_tva, kkk.ToList());
       
         //  CATIA_task = new Task(check_tva,CATIA_task_cancel);
           CATIA_task.Start(kkk.ToList());
           Thread.Sleep(500);
         //  CATIA_task.ContinueWith(check_tva_finish, TaskContinuationOptions.OnlyOnRanToCompletion);

           //CATIA_task_cancel.Cancel();
         
         


       }

       private void button13_Click(object sender, EventArgs e)
       {
           //CATIA_task_cancel.Cancel();

           //try
           //{
           //    CATIA_task.Wait(CATIA_task_cancel.Token);
           //}
           //catch (Exception ee)
           //{
           //    throw ee;
           //    //MessageBox.Show(ee.Message);
           //}

           CATIA_task.Abort();
           button12.Enabled = true;
           button13.Enabled = false;
       }

       private void button14_Click(object sender, EventArgs e)
       {
           DataTable temdt = (DataTable)dataGridView1.DataSource;
           //DataTable deldt = temdt.Copy();
           List<DataRow> ccc = new List<DataRow>();
         
           int count = temdt.Rows.Count;
           for (int i = 0; i < count; i++)
           {
               string filepath = temdt.Rows[i]["文件地址"].ToString();

               if (!File.Exists(filepath))
               {
                   ccc.Add(temdt.Rows[i]);
                   DbHelperSQL.ExecuteSql("delete from 产品数模 where 文件名='" + temdt.Rows[i]["文件名"].ToString() + "';");
               }
           

           }
          // ccc.CopyToDataTable();
           if(ccc.Count()>0)
           {

         
           get_datatable();
           rf_filter();
           excelMethod.SaveDataTableToExcel(ccc.CopyToDataTable());
           }
           
           
           MessageBox.Show("清理完毕");
       }
        private void allPROCToolStripMenuItem_Click(object sender, EventArgs e)
        {

            var TVAkkk = from DataRow dd in unidt.Rows
                      where dd["文件类型"].ToString() == "TVA"
                      select new
                      {
                         aa= dd["文件地址"].ToString(),
                         bb= dd["产品名称"].ToString()
                      };



            var Prockkk = from DataRow dd in unidt.Rows
                         where dd["文件类型"].ToString() == "Process"
                          select new
                          {
                              aa = dd["文件地址"].ToString(),
                              bb = dd["产品名称"].ToString()
                          };

            var kkk = from tk in TVAkkk
                      from pk in Prockkk
                      where tk.bb == pk.bb
                      select new string[2]
                      {
                         pk.aa, tk.aa
                      };

            if (checkBox10.Checked)
            {
                CATIA_task = new Thread(new ParameterizedThreadStart(check_product));
            }
            else
            {
                CATIA_task = new Thread(new ParameterizedThreadStart(open_process));

            }
            CATIA_task.Start(kkk.ToList());
            Thread.Sleep(500);


        }
        private void TVATOOLSToolStripMenuItem_Click(object sender, EventArgs e)
        {

            MainTools f = new MainTools();
            f.Show();
        }
        private void allTVAToolStripMenuItem_Click(object sender, EventArgs e)
       {
            var kkk = from DataRow dd in unidt.Rows
                      where dd["文件类型"].ToString()=="TVA"
                      select dd["文件地址"].ToString();
            //dataGridView1.Rows = kkk;

            //  var token = CATIA_task_cancel.Token;
            //  Thread t1 = new Thread(new ParameterizedThreadStart(replacethread));
            // t1.Start(strobj);
            CATIA_task = new Thread(new ParameterizedThreadStart(check_tva));
            //  new Thread(check_tva, kkk.ToList());

            //  CATIA_task = new Task(check_tva,CATIA_task_cancel);
            CATIA_task.Start(kkk.ToList());
            Thread.Sleep(500);

        }
       private void button15_Click(object sender, EventArgs e)
       {

           ifcolor.Checked = true;
           ifcheck.Checked = true;
           checkBox2.Checked = false;
            ifdt.Checked = false;

          

      

           string sestr = listBox2.SelectedItem.ToString();

           //找出产品所用的TVA
           string productname = sestr.Split('_')[0];
           string jiaci = sestr.Split('_')[1];
           //获取产品一级目录
           var kk = from p in unidt.AsEnumerable()
                    where p["产品名称"].ToString() == productname && p["文件类型"].ToString() == "Fuse_TVA" && p["产品架次"].ToString() == jiaci
                    select p["文件地址"].ToString();
               
               
               //dtmp.AsEnumerable().Where(p => p["产品名称"].ToString() == productname && p["文件类型"].ToString() == "TVA" && p["产品架次"].ToString() == "");
           if (kk.Count() == 0)
           {
               MessageBox.Show("尚未输入该产品的TVA,请核查");
               return;
           }
           else
           {

               CATIA_task = new Thread(new ParameterizedThreadStart(check_tva));
               CATIA_task.Start(kk.ToList());
               Thread.Sleep(500);
              

           }
       }

       private void listBox2_Click(object sender, EventArgs e)
       {
           var kk = from pp in Program.prodTable.AsEnumerable()
                    where pp["名称"].ToString() == listBox2.SelectedItem.ToString().Split('_')[0]
                    select pp["名称"].ToString() +"_"+ pp["图号"].ToString();
           listBox1.SelectedItem = kk.First();
           comboBox4.SelectedIndex = -1;

           rf_filter();
            selectTVA();
        }

       private void button16_Click(object sender, EventArgs e)
       {

       }

       private void button16_Click_1(object sender, EventArgs e)
       {

           if ((listBox1.SelectedIndex != -1) && (get_datatable() != null))
           {

               string foldername = listBox1.SelectedValue.ToString();
               string prodname = foldername.Split('_')[0];
               string drawingname = foldername.Split('_')[1];



               var pathlist = (from p in get_datatable().AsEnumerable()
                               where (p["产品名称"].ToString() == prodname) && (p["文件类型"].ToString() == "Process")
                               select p["文件地址"]);
               if (pathlist.Count() != 1)
               {

                   MessageBox.Show("请确认表格中有且仅有1个该产品的Process文件！");
               }

               else
               {

                   string filepath = pathlist.First().ToString();

              pathlist = (from p in get_datatable().AsEnumerable()
                                   where (p["产品名称"].ToString() == prodname) && (p["文件类型"].ToString() == "TVA")
                                   select p["文件地址"]);

                    var pfdt = DbHelperSQL.Query("select PFname,uuidP from " + drawingname.Replace("-001", "pf")).Tables[0];
                    var pfDic = pfdt.AsEnumerable().ToDictionary(k => k["uuidP"].ToString(), v => v["PFname"].ToString());
                    TVA_Method tvaaa = new TVA_Method(pathlist.First().ToString());
                    tvaaa.ifvec = false;
                    tvaaa.coordswitch = false;
                    var tvapoints = tvaaa.TVAPointsnoVic;

                   foreach (var pp in tvapoints.Points_.Values)
                    {

                        pp.PFname = pfDic[pp.uuid];

                    }

                    TVA_locate cc = new TVA_locate();
                    cc.Show();
                    cc.showPoints = tvapoints;


            // var para =new  string [2];
            //  para[0]=(filepath);
            //  para[1]=(pathlist.First().ToString());

            //  Thread ss = new Thread(new ParameterizedThreadStart(check_product));
            //        var aa = new List<string[]>();
            //        aa.Add(para);
            //ss.Start(aa);




            // // ss.Start(para);
            //  Thread.Sleep(500);
                   // Form_method.open_process(filepath);
               }


           }
           else
           {
               MessageBox.Show("请选择一个产品！");
           }




















       }

        private void button17_Click(object sender, EventArgs e)
        {

            var Prockkk = from DataGridViewRow dd in dataGridView1.Rows
                      where Convert.ToBoolean(dd.Cells["choose"].Value) == true && dd.Cells["文件类型"].Value.ToString() == "Process"
                      select new
                      {
                          aa = dd.Cells["文件地址"].Value.ToString(),
                          bb = dd.Cells["产品名称"].Value.ToString()
                      };

           
            var TVAkkk = from DataRow dd in unidt.Rows
                         where dd["文件类型"].ToString() == "TVA"
                         select new
                         {
                             aa = dd["文件地址"].ToString(),
                             bb = dd["产品名称"].ToString()
                         };




            var kkk = from tk in TVAkkk
                      from pk in Prockkk
                      where tk.bb == pk.bb
                      select new string[2]
                      {
                         pk.aa, tk.aa
                      };

            if (checkBox7.Checked)
            {

         
                if (checkBox10.Checked)
                {
                    CATIA_task = new Thread(new ParameterizedThreadStart(check_product));
                }
                else
                { 
                      CATIA_task = new Thread(new ParameterizedThreadStart(open_process));

                }

            CATIA_task.Start(kkk.ToList());








            Thread.Sleep(500);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where Convert.ToBoolean(dd.Cells["choose"].Value) == true 
                      select dd.Cells["文件地址"].Value.ToString();

            string outputfolder = Program.InfoPath.Replace("prepare", "output") + "files\\";
            foreach (string pp in kkk)
            {
                File.Copy(pp, outputfolder + pp.Split('\\').Last(), true);


            }

            System.Diagnostics.Process.Start("explorer.exe", outputfolder);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex>-1)
            {

          
            string foldername = listBox1.SelectedValue.ToString();
       
            string drawingname = foldername.Split('_')[1];
         
           NC_TOOL.ProductData f1 = new NC_TOOL.ProductData();
                string proname;
                if (checkBox8.Checked)
                {
                    AutorivetDB.updateTVAlocation(drawingname);
                    proname = drawingname.Replace("-001", "process_backup");
                }
              else
                {
                    proname = drawingname.Replace("-001", "process");
                }



       //     string proname = drawingname.Replace("-001", "process");
                f1.inputValue = proname;
                f1.Show();
            }
            else
            {
                MessageBox.Show("请选择产品！");
            }

            //   string path = @"\\192.168.3.32\Autorivet\output\formal\" + proname;
            //if (System.IO.Directory.Exists(path))
            //{
            //    System.Diagnostics.Process.Start("explorer.exe", path);
            //}



        }

        private void button20_Click(object sender, EventArgs e)
        {
           
        }

        private void button20_Click_1(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex>=0)
            {

            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where Convert.ToBoolean(dd.Cells["choose"].Value) == true && dd.Cells["文件类型"].Value.ToString().Contains("TVA")
                      select dd.Cells["文件地址"].Value.ToString();
            TVA_Method mod = new TVA_Method(kkk.First());

            MainTools f = new MainTools();
                f.Show();
                f.inibyTVAmodel =mod;

            }
         else
            {
                MessageBox.Show("请选择产品");
                }



        }

        private void button21_Click(object sender, EventArgs e)
        {
            if ((MessageBox.Show("请确认Products已恢复初始位置？", "警告", MessageBoxButtons.OKCancel) == DialogResult.OK))
            {
           


            var Prockkk = from DataGridViewRow dd in dataGridView1.Rows
                          where Convert.ToBoolean(dd.Cells["choose"].Value) == true && dd.Cells["文件类型"].Value.ToString() == "Process"
                          select new
                          {
                              aa = dd.Cells["文件地址"].Value.ToString(),
                              bb = dd.Cells["产品名称"].Value.ToString()
                          };


            var TVAkkk = from DataRow dd in unidt.Rows
                         where dd["文件类型"].ToString() == "TVA"
                         select new
                         {
                             aa = dd["文件地址"].ToString(),
                             bb = dd["产品名称"].ToString()
                         };




            var kkk = from tk in TVAkkk
                      from pk in Prockkk
                      where tk.bb == pk.bb
                      select new string[2]
                      {
                         pk.aa, tk.aa
                      };



            process_checklist(kkk.ToList());
                //CATIA_task = new Thread(new ParameterizedThreadStart(process_checklist));
                //CATIA_task.Start(kkk.ToList());
                Thread.Sleep(500);
             
            }
        }

        private void fakeToolStripMenuItem_Click(object sender, EventArgs e)
        {
          var mapDt=  AutorivetDB.fullname_table("图号,程序编号,名称");
            var fileList = from dd in mapDt.AsEnumerable()
                           let dwg = dd["名称"].ToString()+"_"+dd["图号"].ToString()
                           let folder= Program.InfoPath + dwg + "\\NC\\"
                           select new {
                               folderName=folder,
                             fileName=  dd["程序编号"].ToString()
                           };
            string time = "1990/8/06 10:10:01";

            DateTime dt = DateTime.Parse(time);
            string targetFolder = @"\\192.168.3.32\Autorivet\output\fake\";
            foreach (var pp in fileList)
            {
              if(File.Exists(pp.fileName))
                {
                    string targetPath = targetFolder+pp.fileName;

                    localMethod.creatDir(targetFolder);
                    File.Copy(pp.folderName+pp.fileName, targetPath,true);

                    File.SetCreationTime(targetPath, dt);
                    File.SetLastWriteTime(targetPath, dt);
                    File.SetLastWriteTime(targetPath, dt);

                    Directory.SetCreationTime(targetFolder, dt);
                    Directory.SetLastWriteTime(targetFolder, dt);
                    Directory.SetLastWriteTime(targetFolder, dt);

                }
            }
            System.Diagnostics.Process.Start("explorer.exe", targetFolder);

        }

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            int k = e.ColumnIndex;
            int l = e.RowIndex;
            if (k == 6 && l >= 0)
            {
                List<historyJsonModel> p1 = jsonMethod.ReadFromStr(dataGridView1.Rows[l].Cells["历史"].Value.ToString());
                dataGridView1.Rows[l].Cells["历史"].ToolTipText = p1.JsonToString();


            }
        }

        private void button22_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            var selectedcells = from DataGridViewRow pp in dataGridView1.Rows
                                let pdname= pp.Cells["产品名称"].Value.ToString()
                                where (!pdname .Contains("半壳"))&&(!pdname.Contains("门"))  && (!pp.Cells["文件名"].Value.ToString().Contains("C013"))
                                select pp.Cells["choose"];



            if (checkBox1.Checked == true)
            {
                foreach (var dd in selectedcells)
                {
                    dd.Value = true;
                }

               
            }
            else
            {
                foreach (var dd in selectedcells)
                {
                    dd.Value =false;
                }


            }
        }
    }
}
