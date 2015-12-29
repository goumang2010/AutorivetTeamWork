using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using GoumangToolKit;
using System.Management;
using FileManagerNew;


namespace AUTORIVET_KAOHE
{
    public partial class AO : Form
    {
        public AO()
        {
            InitializeComponent();
        }
        private Microsoft.Office.Interop.Word.Application wordApp;
        private Microsoft.Office.Interop.Word.Document myAO;
        private object defaultV = System.Reflection.Missing.Value;
        //    private object documentType;
        private Dictionary<string, string> infoDic=new Dictionary<string, string>();
        public Dictionary<string,string> rncaao

        {
            get
            {
              
                return infoDic;
            }

            set
            {
                infoDic = value;
                listBox1.SelectedItem = infoDic["名称"] + "_" + infoDic["架次"];



            }
        }





        private Dictionary<string, string> fasteners = new Dictionary<string, string>();

       
    
        private void AO_Load(object sender, EventArgs e)
        {
          // 



            DataGridViewComboBoxColumn col_btn_insert = new DataGridViewComboBoxColumn();
            col_btn_insert.HeaderText = "紧固件";
            col_btn_insert.Items.Add("B0205020AD5-6S");
            col_btn_insert.Items.Add("B0205020AD5-7S");
            col_btn_insert.Items.Add("B0205020AD6-6S");
            col_btn_insert.Items.Add("B0206002AG5-3");
            col_btn_insert.Items.Add("B0206002AG5-4");
            col_btn_insert.Items.Add("B0206002AG5-5");
            col_btn_insert.Items.Add("B0206002AG6-3");
            col_btn_insert.Items.Add("B0206002AG6-4");
            col_btn_insert.Items.Add("MS20470AD5-6");
            col_btn_insert.Items.Add("MS20470AD6-6");
            col_btn_insert.Items.Add("B0205020AD5-6A");
         col_btn_insert.Width = 120;

            //col_btn_insert.v = true;
            dataGridView1.Columns.Add(col_btn_insert);


            DataGridViewTextBoxColumn col_btn_insert2 = new DataGridViewTextBoxColumn();
            col_btn_insert2.HeaderText = "数量";
            //col_btn_insert2.Text = "返修";//加上这两个就能显示
           // col_btn_insert2.UseColumnTextForButtonValue = true;

            col_btn_insert2.Width = 55;
            dataGridView1.Columns.Add(col_btn_insert2);
            listBox1.DataSource = DbHelperSQL.getlist("select concat(产品名称,'_',产品架次) from 产品流水表");
            listBox1.SelectedIndex = listBox1.Items.Count - 1;

        }

        private Document creatAAO()
        {

            List<string> prodinfo = DbHelperSQL.getlistcol("select 图号,图纸名称,下级装配号,图纸版次,AOI编号,预铆编号 from 产品列表 where 名称='" +  rncaao["名称"] + "'");
           // rncaao.Add("类型", "AAO返修");
            try
            {
                rncaao.Add("图号", prodinfo[0]);
                rncaao.Add("图纸名称", prodinfo[1]);
                rncaao.Add("下级装配号", prodinfo[2]);
                rncaao.Add("图纸版次", prodinfo[3]);
                if(comboBox1.Text=="自动钻铆")
                {
                    rncaao.Add("编号", prodinfo[4].Substring(0,8)+"--AAO");
                }
                else
                {
                    rncaao.Add("编号","C1-"+ prodinfo[5]+"--AAO");

                    rncaao["保存地址"] = Program.InfoPath + rncaao["名称"]+"_"+ rncaao["图号"]+"\\"+ rncaao["架次"]+@"\" + rncaao["图号"]+"_CHANGEAAO.doc";


                }
               
            }
            catch
            {

            }
           

            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;
            myAO = new Document();
                if (File.Exists(rncaao["保存地址"])&&(!checkBox1.Checked))
                {
                    myAO = wordApp.Documents.Open(rncaao["保存地址"]);

                }
                else
                {
                    //myAO = wordApp.Documents.Add();
              

                    myAO = wordApp.Documents.Open(@"\\192.168.3.32\Autorivet\Prepare\INFO\SAMPLE\AAO.doc");


            //下一级装配件号
            myAO.Tables[1].Cell(2, 2).Range.Text = rncaao["下级装配号"];
            string baseno = rncaao["编号"];

        int aaocount=    DbHelperSQL.getlist("select 文件编号 from Paperwork where 文件编号 like '" + baseno + "%'").Count();



        myAO.Tables[1].Cell(2, 6).Range.Text = baseno + (aaocount + 1).ToString();
         //   }


           
            //发文目的


            //标题
            if (comboBox1.Text != "更改")
            {
                string neibujushou = rncaao["内部拒收号"];
                myAO.Tables[1].Cell(2, 4).Range.Text = "配合文件RNC：" + neibujushou + "进行排故\rACCORDING TO RNC: " + neibujushou + " REPAIR";

                myAO.Tables[1].Cell(3, 1).Range.Text = "发放AAO目的 The Intent Of The Document Is: 根据编号为" + neibujushou + "的RNC对产品进行返修。 ACCORDING TO RNC: " + neibujushou;
                myAO.Tables[1].Cell(4, 1).Range.Text = "TO REWORK FOR THE PRODUCT.";
            }


            //工程图纸号
            myAO.Tables[1].Cell(5, 2).Range.Text = rncaao["图号"].Replace("-001", "");
            //版次
            string banci = rncaao["图纸版次"][1].ToString();
            myAO.Tables[1].Cell(5, 4).Range.Text = banci;
            //第一条

            myAO.Tables[1].Cell(7, 1).Range.Text = "1";
            myAO.Tables[1].Cell(7, 2).Range.Text = rncaao["图号"] + "P2";
            myAO.Tables[1].Cell(7, 4).Range.Text = "1";
            myAO.Tables[1].Cell(7, 6).Range.Text = rncaao["名称"] + rncaao["图纸名称"];
           // myAO.Tables[1].Cell(7, 10).Range.Text = rncaao["架次"];
            // myAO.Tables[1].Cell(7, 1).Range.Text = dataGridView1.Rows[0].

            //开始统计紧固件
            int fstcount = fasteners.Count();
           if(fstcount!=0)
           {

         
            int nut5 = 0;
            int nut6 = 0;
            for (int i = 0; i < fstcount; i++)
            {
                string fstname = fasteners.Keys.ElementAt(i);
                myAO.Tables[1].Cell(8 + i, 1).Range.Text = (i + 2).ToString();
                myAO.Tables[1].Cell(8 + i, 2).Range.Text = fasteners.Keys.ElementAt(i);

                myAO.Tables[1].Cell(8 + i, 4).Range.Text = fasteners[fstname];
                if (fstname.Contains("B020600"))
                {
                    myAO.Tables[1].Cell(8 + i, 6).Range.Text = "高锁 HI-LITE";
                    int hiqty = System.Convert.ToInt16(fasteners[fstname]);

                    if (fstname.Contains("AG5"))
                    {
                        nut5 = nut5 + hiqty;

                    }
                    else
                    {
                        nut6 = nut6 + hiqty;
                    }
                }
                else
                {
                    myAO.Tables[1].Cell(8 + i, 6).Range.Text = "铆钉 RIVET";
                }
            }
            fstcount = 8 + fstcount;
            if (nut5 != 0)
            {
                myAO.Tables[1].Cell(fstcount, 1).Range.Text = (fstcount - 6).ToString();
                myAO.Tables[1].Cell(fstcount, 6).Range.Text = "高锁帽 NUT";
                myAO.Tables[1].Cell(fstcount, 2).Range.Text = "B0203013-08";
                myAO.Tables[1].Cell(fstcount, 4).Range.Text = nut5.ToString();

                if (nut6 != 0)
                {
                    fstcount = fstcount + 1;
                    myAO.Tables[1].Cell(fstcount, 1).Range.Text = (fstcount - 6).ToString();
                    myAO.Tables[1].Cell(fstcount, 6).Range.Text = "高锁帽 NUT";
                    myAO.Tables[1].Cell(fstcount, 2).Range.Text = "B0203013-3";
                    myAO.Tables[1].Cell(fstcount, 4).Range.Text = nut6.ToString();
                }

            }
            else
            {
                if (nut6 != 0)
                {
                    //fstcount = fstcount + 1;
                    myAO.Tables[1].Cell(fstcount, 1).Range.Text = (fstcount - 6).ToString();
                    myAO.Tables[1].Cell(fstcount, 6).Range.Text = "高锁帽 NUT";
                    myAO.Tables[1].Cell(fstcount, 2).Range.Text = "B0203013-3";
                    myAO.Tables[1].Cell(fstcount, 4).Range.Text = nut6.ToString();
                }

            }
            //添加密封胶
            if((nut5+nut6)==0)
            {
               
            }
            else
            {
                fstcount = fstcount + 1;

            }
     





            myAO.Tables[1].Cell(fstcount, 1).Range.Text = (fstcount - 6).ToString();
            myAO.Tables[1].Cell(fstcount, 6).Range.Text = "密封剂 SEALANT";
            myAO.Tables[1].Cell(fstcount, 2).Range.Text = "MIL-PRF-81733 Type IV P/S C-24";
            myAO.Tables[1].Cell(fstcount, 4).Range.Text = "100g";
           }

            //架次
            myAO.Tables[1].Cell(15, 4).Range.Text = rncaao["架次"];
            myAO.Tables[1].Cell(15, 6).Range.Text = rncaao["图号"];
            myAO.SaveAs2(rncaao["保存地址"]);
                  
                               
                    
                    
                }
                FormMethod.scanfiledoc(rncaao["保存地址"], myAO,false);
                MessageBox.Show("生成完毕，请补全其他信息");
            return myAO;

        }

        private Document creatAOI()
        {

            if (!rncaao.Keys.Contains("图纸名称"))
            {
                rncaao.Add("图纸名称", AutorivetDB.queryno(rncaao["图号"], "图纸名称"));
            }
            


            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;
            myAO = new Document();
            if (File.Exists(rncaao["AOI保存地址"]) && (!checkBox1.Checked))
            {
                myAO = wordApp.Documents.Open(rncaao["AOI保存地址"]);

            }
            else
            {
                
                  if (rncaao["中文名称"].Contains("CS300"))
                  {
                      myAO = wordApp.Documents.Open(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "wordModel\\AOI\\AOI_CS300.docx");
                  }
                  else
                  {
                      myAO = wordApp.Documents.Open(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "wordModel\\AOI\\AOI.docx");
                  }

                      

                        myAO.SaveAs2(rncaao["AOI保存地址"]);

                }
               

         
            
                myAO.Tables[1].Cell(3,2).Range.Text = rncaao["AOI编号"];

                //替换所有编号
                wordMethod.SearchReplace(wordApp, myAO, "C1-CF510-010", rncaao["AOI编号"]);








                //标题


                myAO.Tables[1].Cell(3, 4).Range.Text = rncaao["中文名称"] + "壁板自动钻铆\rAUTO-RIVETING OF " + rncaao["图纸名称"];

                //装配图号
                    myAO.Tables[1].Cell(4, 2).Range.Text = rncaao["装配图号"];
               //替换所有装配图号

                    wordMethod.SearchReplace(wordApp, myAO, "C02311100", rncaao["装配图号"]);





            //进入零件表填写


                //第一条
            

                myAO.Tables[2].Cell(3, 1).Range.Text = "1";
        //     myAO.Tables[2].Cell(3, 1).Height=
                myAO.Tables[2].Cell(3, 2).Range.Text = rncaao["图号"] + "P1";
                myAO.Tables[2].Cell(3, 3).Range.Text = rncaao["中文名称"] +"壁板组件\r"+ rncaao["图纸名称"];
                myAO.Tables[2].Cell(3, 4).Range.Text = "A";
                myAO.Tables[2].Cell(3, 5).Range.Text = "1";
                myAO.Tables[2].Cell(3, 6).Range.Text = rncaao["装配图号"];

                // myAO.Tables[1].Cell(7, 10).Range.Text = rncaao["架次"];
                // myAO.Tables[1].Cell(7, 1).Range.Text = dataGridView1.Rows[0].

                //开始统计紧固件

                System.Data.DataTable fsttable = AutorivetDB.spfsttable(rncaao["图号"]);

                int fstcount = fsttable.Rows.Count;
                int nut5 = 0;
                int nut6 = 0;
                if (fstcount != 0)
                {


               
                    for (int i = 0; i < fstcount; i++)
                    {
                        string fstno = fsttable.Rows[i][0].ToString();
                        myAO.Tables[2].Cell(4 + i, 1).Range.Text = (i + 2).ToString();
                        myAO.Tables[2].Cell(4 + i, 2).Range.Text = fstno;
                        //string fstname = "";
                       
                        
                        if (fstno.Contains("B020600"))
                        {
                            myAO.Tables[2].Cell(4 + i, 3).Range.Text = "高锁\rHI-LITE";
                            int hiqty = System.Convert.ToInt16(fsttable.Rows[i][1].ToString());

                            if (fstno.Contains("AG5"))
                            {
                                nut5 = nut5 + hiqty;

                            }
                            else
                            {
                                nut6 = nut6 + hiqty;
                            }
                        }
                        else
                        {
                            myAO.Tables[2].Cell(4 + i, 3).Range.Text = "铆钉\rRIVET";
                        }
                        myAO.Tables[2].Cell(4 + i, 4).Range.Text = "U";
                        myAO.Tables[2].Cell(4 + i, 5).Range.Text = fsttable.Rows[i][1].ToString();
                        myAO.Tables[2].Cell(4 + i, 6).Range.Text = rncaao["装配图号"];
                    }


                   
                    fstcount = 4 + fstcount;
                    if (nut5 != 0)
                    {
                        myAO.Tables[2].Cell(fstcount, 1).Range.Text = (fstcount - 2).ToString();
                        myAO.Tables[2].Cell(fstcount, 2).Range.Text = "B0203013-08";
                        myAO.Tables[2].Cell(fstcount, 3).Range.Text = "高锁帽\rNUT";
                        myAO.Tables[2].Cell(fstcount, 4).Range.Text = "U";
                        myAO.Tables[2].Cell(fstcount, 5).Range.Text = nut5.ToString();
                        myAO.Tables[2].Cell(fstcount, 6).Range.Text = rncaao["装配图号"];
                        if (nut6 != 0)
                        {
                            fstcount = fstcount + 1;
                            myAO.Tables[2].Cell(fstcount, 1).Range.Text = (fstcount - 2).ToString();
                            myAO.Tables[2].Cell(fstcount, 2).Range.Text = "B0203013-3";
                            myAO.Tables[2].Cell(fstcount, 3).Range.Text = "高锁帽\rNUT";
                            myAO.Tables[2].Cell(fstcount, 4).Range.Text = "U";
                            myAO.Tables[2].Cell(fstcount, 5).Range.Text = nut6.ToString();
                            myAO.Tables[2].Cell(fstcount, 6).Range.Text = rncaao["装配图号"];


                           
                        }

                    }
                    else
                    {
                        if (nut6 != 0)
                        {
                            //fstcount = fstcount + 1;
                            myAO.Tables[2].Cell(fstcount, 1).Range.Text = (fstcount - 2).ToString();
                            myAO.Tables[2].Cell(fstcount, 2).Range.Text = "B0203013-3";
                            myAO.Tables[2].Cell(fstcount, 3).Range.Text = "高锁帽\rNUT";
                            myAO.Tables[2].Cell(fstcount, 4).Range.Text = "U";
                            myAO.Tables[2].Cell(fstcount, 5).Range.Text = nut6.ToString();
                            myAO.Tables[2].Cell(fstcount, 6).Range.Text = rncaao["装配图号"];
                        }

                    }
                    //添加密封胶
                    if ((nut5 + nut6) == 0)
                    {

                    }
                    else
                    {
                        fstcount = fstcount + 1;

                    }

                    myAO.Tables[2].Cell(fstcount, 1).Range.Text = (fstcount - 2).ToString();
                    myAO.Tables[2].Cell(fstcount, 3).Range.Text = "密封剂\rSEALANT";
                    myAO.Tables[2].Cell(fstcount, 2).Range.Text = "MIL-PRF-81733 Type IV";
                    myAO.Tables[2].Cell(fstcount, 4).Range.Text = "M";
                    myAO.Tables[2].Cell(fstcount,5).Range.Text = "340g";
                    myAO.Tables[2].Cell(fstcount, 6).Range.Text = rncaao["装配图号"];

                    //添加胶嘴
                    //fstcount = fstcount + 1;
                    //List<string> tipls = autorivet_op.processitemlist(rncaao["装配图号"], "胶嘴Sealant_Tip");
                    //for (int k=0;k<tipls.Count();k++)
                    //{

                    //    myAO.Tables[2].Cell(fstcount+k, 1).Range.Text = (fstcount+k - 2).ToString();
                    //    myAO.Tables[2].Cell(fstcount + k, 2).Range.Text = tipls[k];
                    //    myAO.Tables[2].Cell(fstcount + k, 3).Range.Text = "胶嘴\rSealant Tip";
                    //    myAO.Tables[2].Cell(fstcount + k, 4).Range.Text = "M";
                    //    myAO.Tables[2].Cell(fstcount + k, 5).Range.Text = "1";
                    //    myAO.Tables[2].Cell(fstcount + k, 6).Range.Text = rncaao["装配图号"];

                    //}

                    //添加下铆头

                    fstcount = fstcount+1;


                    while (!myAO.Tables[2].Cell(fstcount , 1).Range.Text.Contains("专用工具"))
                    {


                        fstcount = fstcount + 1;


                    }

                    //定位到要输入专用工具的行
                    fstcount = fstcount + 2;





                    List<string> lals = AutorivetDB.processitemlist(rncaao["装配图号"], "下铆头Lower_Anvil");
                    //最多只能放下三行
                    //用于记录行尾,指示下一行数据的索引;
                    int rear = fstcount;
                  
                        for (int k = 0; k < lals.Count(); k++)
                        {

                            myAO.Tables[2].Cell(rear , 1).Range.Text = (k + 1).ToString();
                            myAO.Tables[2].Cell(rear , 2).Range.Text = "下铆头\rLower Anvil";
                            myAO.Tables[2].Cell(rear , 3).Range.Text = "1";
                            myAO.Tables[2].Cell(rear , 4).Range.Text = lals[k];

                            rear = rear + 1;

                        }
                 //高锁限力枪
                        int toolindex = lals.Count();
                    if(nut5>0)
                    {
                        toolindex = toolindex + 1;
                        myAO.Tables[2].Cell(rear, 1).Range.Text = toolindex.ToString();
                        myAO.Tables[2].Cell(rear, 2).Range.Text = "高锁限力枪\rHI-LITE Istallation Tool";
                        myAO.Tables[2].Cell(rear, 3).Range.Text = "1";
                        myAO.Tables[2].Cell(rear, 4).Range.Text = "KTL1408A218B062";
                        toolindex += 1;
                        rear += 1;

                    }

                    if (nut6 > 0)
                    {
                        toolindex = toolindex + 1;
                        myAO.Tables[2].Cell(rear, 1).Range.Text = toolindex.ToString();
                        myAO.Tables[2].Cell(rear, 2).Range.Text = "高锁限力枪\rHI-LITE Istallation Tool";
                        myAO.Tables[2].Cell(rear, 3).Range.Text = "1";
                        myAO.Tables[2].Cell(rear, 4).Range.Text = "KTL1439A249C078";
                        rear += 1;
                    }


                    //寻找图纸列表
                    while (!myAO.Tables[2].Cell(rear, 1).Range.Text.Contains("图纸列表"))
                    {


                        rear = rear + 1;


                    }

                    rear = rear + 2;

         


                    //添加图纸

                    myAO.Tables[2].Cell(rear, 1).Range.Text = "1";
                    myAO.Tables[2].Cell(rear, 2).Range.Text = rncaao["装配图号"];
                    myAO.Tables[2].Cell(rear, 3).Range.Text = rncaao["中文名称"] + "壁板组件\r" + rncaao["图纸名称"];




                    //寻找程序列表
                    while (!myAO.Tables[2].Cell(rear, 1).Range.Text.Contains("程序列表"))
                    {


                        rear = rear + 1;


                    }

                    rear = rear + 2;

                    myAO.Tables[2].Cell(rear, 1).Range.Text = "1";
                    myAO.Tables[2].Cell(rear, 2).Range.Text = rncaao["程序编号"];

                    //替换所有程序编号
                    wordMethod.SearchReplace(wordApp, myAO, "MA10016A", rncaao["程序编号"]);


                    rear = rear + 1;



                    //寻找参考文件列表
                    while (!myAO.Tables[2].Cell(rear, 1).Range.Text.Contains("参考文件列表"))
                    {


                        rear = rear + 1;


                    }

                    rear = rear + 8;

                    myAO.Tables[2].Cell(rear, 2).Range.Text = rncaao["状态编号"];

                    //替换所有状态编号
                    wordMethod.SearchReplace(wordApp, myAO, "C1-S22100-33", rncaao["状态编号"]);



                    myAO.Tables[2].Cell(rear, 3).Range.Text = rncaao["中文名称"] + "壁板装配交付状态\rCOS OF " + rncaao["图纸名称"];



                }

                
          
                myAO.Save();
             //   wordApp.Quit();
                FormMethod.scanfiledoc(rncaao["AOI保存地址"],myAO);
                MessageBox.Show("生成完毕，请补全其他信息");


      

            return myAO;

        }
        private void button2_Click(object sender, EventArgs e)
        {
           
            creatAAO();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            fasteners = new Dictionary<string, string>();

            int datacount = dataGridView1.Rows.Count;

            for (int i = 0; i < datacount - 1;i++ )
            {
                DataGridViewRow pp = dataGridView1.Rows[i];
                fasteners.Add(pp.Cells[0].Value.ToString(), pp.Cells[1].Value.ToString());

            }

            
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private static string GetProcessUserName(int pID)
        {


            string text1 = null;


            SelectQuery query1 =
                new SelectQuery("Select * from Win32_Process WHERE processID=" + pID);
            ManagementObjectSearcher searcher1 = new ManagementObjectSearcher(query1);


            try
            {
                foreach (ManagementObject disk in searcher1.Get())
                {
                    ManagementBaseObject inPar = null;
                    ManagementBaseObject outPar = null;


                    inPar = disk.GetMethodParameters("GetOwner");


                    outPar = disk.InvokeMethod("GetOwner", inPar, null);


                    text1 = outPar["User"].ToString();
                    break;
                }
            }
            catch
            {
                text1 = "SYSTEM";
            }


            return text1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
             wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;
            myAO = new Document();
            if (File.Exists(rncaao["索引保存地址"]) && (!checkBox1.Checked))
            {
                myAO = wordApp.Documents.Open(rncaao["索引保存地址"]);

            }
                else
                {
                    //myAO = wordApp.Documents.Add();
                    
                            myAO = wordApp.Documents.Open(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "wordModel\\大纲索引"+comboBox2.Text+".doc");

                            myAO.Tables[1].Cell(2, 4).Range.Text = rncaao["装配图号"];
                            myAO.Tables[2].Cell(2, 4).Range.Text = rncaao["装配图号"];
                            myAO.Tables[1].Cell(2, 8).Range.Text = rncaao["组件号"];
                            myAO.Tables[2].Cell(2, 8).Range.Text = rncaao["组件号"];
                            myAO.Tables[1].Cell(3, 4).Range.Text = rncaao["图纸版次"];
                            myAO.Tables[2].Cell(3, 4).Range.Text = rncaao["图纸版次"];

                           myAO.SaveAs2(rncaao["保存地址"]);
                            wordMethod.SearchReplace(wordApp, myAO, "ZWH", rncaao["站位号"]);


                    }
        }

        private void killWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
       
            int k = FormMethod.killProcess("WINWORD");

            MessageBox.Show("执行成功，杀死" + k.ToString() + "个word进程");
        }




        private void listBox1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                var tt = listBox1.SelectedItem.ToString().Split('_');
                infoDic["名称"] = tt[0];
                infoDic["架次"] = tt[1];
            }
        }
    }
}
