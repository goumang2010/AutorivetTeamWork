using mysqlsolution;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using FileManagerNew;
using NC_TOOL;
using OFFICE_Method;
using System.Reflection;


namespace AUTORIVET_KAOHE
{
    public partial class couponTest : Form
    {
        Dictionary<string, string> fstenerT = DbHelperSQL.getDic("select Fasteners,Tcode from 紧固件列表");
        Dictionary<string, Control[]> controlList = new Dictionary<string, Control[]>();
        private DataTable coupondt;

        public couponTest()
        {
            InitializeComponent();
            //Fectch the coupon configuration 

            
        }



        private void button1_Click(object sender, EventArgs e)
        {

            //pointCoord test = new pointCoord("X;500");
          

            string outputfolder = Program.InfoPath + productList1.listBox1.SelectedItem.ToString() + "\\NC\\COUPON\\";

            localMethod.creatDir(outputfolder);



            foreach (var cc in controlList.Values)
            {
                List<string> sb = new List<string>();

                if (cc[1].ForeColor==Color.ForestGreen)
                {
                    //获取试片文本
                    string coupontext = cc[0].Text;
                    var cptmp = coupontext.Split('/');


                    //获取试片编号
                    var cpmarkls = from dd in coupondt.AsEnumerable()
                                 where dd["试片1"].ToString() == "C1-"+cptmp[0] && dd["试片2"].ToString() == "C1-" + cptmp[1]
                                 select new {

                                   cpno=  dd["组合编号"].ToString(),
                                     progUUID= dd["程序段编号"].ToString(),
                                     progID = dd["ID"].ToString()
                                 };
                    string cpmark;
                    if (cpmarkls.Count()>0)
                    {
                        cpmark = cpmarkls.First().cpno;
                    }
                    else
                    {
                        cpmark = "NEW";
                    }
                    

              

                    //选择“对号” 的tag
                    string mark = cc[1].Tag.ToString();
                    string folderpath = Program.InfoPath + "SAMPLE\\COUPON\\TAB\\" + mark + "\\";
            pointCoordList ptls = new pointCoordList();
        pointCoordList[] couponModel=new pointCoordList[3] {
            new pointCoordList(folderpath + mark+"_upper.tab"),
         new pointCoordList(folderpath + mark+"_lower.tab"),
         new pointCoordList(folderpath + mark+"_middle.tab")
        };
                    //判断一共有几种紧固件c[2]-c[4]

                    var fstcontrol = cc.Skip(2);
                    var fstlist = from ff in fstcontrol
                                  where ff.Text != ""
                                  select ff.Text;


                   foreach(string fstname in fstlist)
                    {

                        string fstnamemark= fstname;
                        //判断加工代码
                        string Mcode;
                        if (fstname.Contains("B020600"))
                        {
                            Mcode = "M62";

                            fstnamemark = fstname.Split('-')[0];
                        }
                        else
                        {

                            Mcode = "M60";

                        }
                        //Read the enter/out NC codes
                        string progfolder = Program.InfoPath + @"SAMPLE\COUPON\TAB\ENTER\";
                        Func<string, NCcodeList> openNCfile = delegate (string name)
                        {
                            //Type dd = typeof(NCcodeList);
                            //NCcodeList obj = (NCcodeList)dd.Assembly.CreateInstance(dd.FullName);
                            NinjectDependencyResolver dd = new NinjectDependencyResolver();
                            NCcodeList obj = new NCcodeList((IDBInfo)dd.GetService(typeof(IDBInfo)));
                            obj.ImportFromFile(progfolder + name);
                            obj.NCList.RemoveRange(0, 2);
                            obj.NCList.RemoveRange(obj.NCList.Count - 2, 2);
                            return obj;
                        };
                        string beginProg;
                        string enterProg;
                        string outProg;
                        string endProg;
                        List<string> beginNC;
                        List<string> enterNC;
                        List<string> outNC;
                        List<string> endNC;
                        if (mark.Contains("left"))
                        {
                            beginProg= "M98 P3601";
                            beginNC = openNCfile("BEGIN_LEFT").NCList;
                            enterProg = "M98 P3602";
                            enterNC= openNCfile("ENTER_LEFT").NCList;
                            outProg = "M98 P3603";
                            outNC = openNCfile("OUT_LEFT").NCList;
                            endProg = "M98 P3604";
                            endNC= openNCfile("END_LEFT").NCList;
                        }
                        else
                        {
                            beginProg = "M98 P3701";
                            beginNC = openNCfile("BEGIN_RIGHT").NCList;
                            enterProg = "M98 P3702";
                            enterNC = openNCfile("ENTER_RIGHT").NCList;
                            outProg = "M98 P3703";
                            outNC = openNCfile("OUT_RIGHT").NCList;
                            endProg = "M98 P3704";
                            endNC = openNCfile("END_RIGHT").NCList;
                        }

                       


                    

                       







                        string progID;

                        var cpmarkls2 = cpmarkls.Where(x => x.progUUID.Contains(fstnamemark));
                        if (cpmarkls2.Count() > 0)
                        {
                            progID = cpmarkls2.First().progID;
                        }
                        else
                        {
                            progID = "NEW";
                        }
                       // sb.Add("%");
                        string Ocode = "O" + productList1.listBox1.SelectedItem.ToString().Split('_')[1].Substring(3, 4) + progID.PadLeft(2, '0') + cpmark.Substring(1).PadLeft(2, '0');


                       string filename = progID+"_"+ cpmark + "_" + fstname + "_" + cptmp[0] + "_" + cptmp[1]+"_" + mark.ToUpper();

                        sb.Add("(MSG, START COUPON TEST:" + filename+")");

                        ////下铆头降到最低
                        //sb.Add("M53");
                        ////所有传感器关闭
                        //sb.Add("M51");
                        //// 取消传感器补偿清除传感器在点位模式下创建的坐标补偿
                        //sb.Add("M26");
                        ////取消视觉修正补偿清除视觉修正导致的XY坐标补偿
                        //sb.Add("M35");
                        //进入程序代码
                        //校准工装用的代码
                        sb.Add("(MSG, IF YOU ARE AT INSPECTION LOCATION THEN IGNORE NEXT STATEMENT)");
                        sb.Add("(MSG, BEGIN ENTER)");
                        sb.AddRange(beginNC);

                        sb.AddRange(enterNC);

                        sb.Add("(MSG, END ENTER)");
                        //获取换刀用的T代码
                        sb.Add("M56" + fstenerT[fstname]);


                        //更换钻头工位
                        sb.Add("M83T0");

                        //更换上铆头工位
                        sb.Add("M83T100");

                        //准备注胶循环
                        sb.Add("M83T100");

                        //试片上升
                        sb.Add(couponModel[0][0].Offset("Z;40;W;40;").ToString());


                        //处理Upper 及Lower位置
                        //取得第一个点并校准
                        sb.Add("M34T1");
                        sb.Add(couponModel[0][0].ToString());
                        sb.Add("M51");
                        sb.Add("M50T00");
                        sb.Add("M39");
                        sb.Add("M31");

                        //下铆头上升
                        sb.Add("M57T300");

                        //校准通过(该位置无紧固件)则进行铆接
                        sb.Add(Mcode);

                        for (int i = 1; i < 21; i++)
                        {
                            sb.Add("M34T1");
                            sb.Add(couponModel[0][i].ToString());
                            sb.Add("M39");
                            sb.Add("M31");
                            sb.Add(Mcode);
                        }
                        //下铆头下降
                        sb.Add("M57T4000");

                        //试片下降
                        sb.Add("(MSG, COUPON FIXTURE WILL GO UP AND MOVE TO NEXT LINE TO INSTALL FASTERNERS)");
                        sb.Add(couponModel[0][20].Offset("Z;40;W;40;").ToString());

                        //进行lower铆接
                        //添加走回工装校准点校准的脚本
                        //板子上升
                        sb.Add(couponModel[1][0].Offset("Z;40;W;40;").ToString());

                        //取得第一个点并校准
                        sb.Add("M34T1");
                        sb.Add(couponModel[1][0].ToString());
                        sb.Add("M51");
                        sb.Add("M50T00");
                        sb.Add("M39");
                        sb.Add("M31");
                       

                        //下铆头上升
                        sb.Add("M57T300");

                        //校准通过(该位置无紧固件)则进行铆接
                        sb.Add(Mcode);
                        for (int i = 1; i < 21; i++)
                        {
                            sb.Add("M34T1");
                            sb.Add(couponModel[1][i].ToString());
                            sb.Add("M39");
                            sb.Add("M31");
                            sb.Add(Mcode);
                        }


                        sb.Add("(MSG, START DRILL HOLES ON COUPONS)");
                        //开始钻孔操作
                        Mcode = "M61";
                       

                     
                        //下铆头下降
                        sb.Add("M57T4000");
                        //试片下降
                        sb.Add("(MSG, COUPON FIXTURE WILL GO UP AND MOVE TO NEXT LINE TO INSTALL FASTERNERS)");
                        sb.Add(couponModel[1][20].Offset("Z;40;W;40;").ToString());
                        sb.Add(couponModel[2][0].Offset("Z;40;W;40;").ToString());
                        //取得第一个点并校准
                        sb.Add("M34T1");
                        sb.Add(couponModel[2][0].ToString());
                        sb.Add("M51");
                        sb.Add("M50T00");
                        sb.Add("M39");
                        sb.Add("M31");
                       

                        //下铆头上升
                        sb.Add("M57T300");

                        //校准通过(该位置无紧固件)则进行加工
                        sb.Add(Mcode);
                        for (int i = 1; i < 21; i++)
                        {
                            sb.Add("M34T1");
                            sb.Add(couponModel[2][i].ToString());
                            sb.Add("M39");
                            sb.Add("M31");
                            sb.Add(Mcode);
                        }
                        //下铆头下降
                        sb.Add("M57T4000");

                        //试片降低
                        sb.Add(couponModel[2][20].Offset("Z;40;W;40;").ToString());
                        sb.Add("(MSG, END COUPON TEST:" + filename + ")");
                        sb.Add("(MSG, NOW OUT THE FIXTURE)");
                        sb.AddRange(outNC);
                        sb.Add("(MSG, IF YOU WANT TO CONTINUE COUPON TEST THEN SKIP NEXT STATEMENT)");

                        sb.AddRange(endNC);
                   //     sb.Add("%");




                        //输出文件至目录
                        int m = 0;
                        var outputNC = sb.ConvertAll(x => "N" + ((++m) * 2).ToString() + " " + x);
                        outputNC.Insert(0, "%");
                        outputNC.Insert(1, Ocode);
                        outputNC.Insert(outputNC.Count-1, "%");
                        outputNC.WriteFile(outputfolder + filename);


                   
                    }




                }

            }
            //复制进出程序至输出目录
            List<FileInfo> en = new List<FileInfo>();
            en.WalkTree(Program.InfoPath + @"SAMPLE\COUPON\TAB\ENTER\",false);
            en.copyto(outputfolder);


            System.Diagnostics.Process.Start("explorer.exe", outputfolder);
        }

        private void couponTest_Load(object sender, EventArgs e)
        {

            productList1.listBox1.Click += ListBox1_Click;


            dynamic CPNameDic;
            CPNameDic = localMethod.GetConfigValue("GetCouponName", "CouponCfg.py");




            //生成对应的字典
            for (int i=10;i<=18;i++)
            {
                string controlname = "label" + i.ToString();
                var leftLabel = this.Controls["label" + i.ToString()];
                var leftLabel2 = this.Controls["label" + (i - 9).ToString()];
                string leftindex = "l" + (i - 9).ToString();
                leftLabel.Text = CPNameDic(leftindex);
                controlList.Add(leftindex,new Control[5] { leftLabel, leftLabel2, this.Controls["comboBox" + (i - 9).ToString()], this.Controls["comboBox" + (28-i).ToString()], this.Controls["comboBox" + (55 - i).ToString()] });
                var rightLabel = this.Controls["label" + (37 - i).ToString()];
                var rightLabel2 = this.Controls["label" + (46 - i).ToString()];
                string rightindex = "r"+(i - 9).ToString();
                rightLabel.Text = CPNameDic(rightindex);
                controlList.Add(rightindex ,new Control[5] { rightLabel, rightLabel2, this.Controls["comboBox" + (46-i).ToString()], this.Controls["comboBox" + (37 - i).ToString()], this.Controls["comboBox" + (64 - i).ToString()] });
                leftLabel2.Click += new System.EventHandler(this.coupon_tweak);
                leftLabel.DoubleClick += new System.EventHandler(this.editCPName);
                rightLabel2.Click += new System.EventHandler(this.coupon_tweak);
                rightLabel.DoubleClick += new System.EventHandler(this.editCPName);


            }

       











        }

        private void LeftLabel_DoubleClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void ListBox1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            string productname = productList1.listBox1.SelectedItem.ToString();
            string prodname = productname.Split('_')[1];
            //获取试片列表

            //DbHelperSQL.ExecuteSqlTran(autorivet_op.rfcouponno(prodname));

            //coupondt= DbHelperSQL.Query("select concat('T',编号) as 组合编号,concat('C1-SKIN-',蒙皮厚度) as 试片1,concat('C1-',二层材料,'-',二层厚度) as 试片2,程序段编号 from 试片列表 where 产品图号='" + prodname + "' order by 试片1,试片2").Tables[0];
         coupondt= DbHelperSQL.Query("select concat('T',编号) as 组合编号,concat('C1-SKIN-',蒙皮厚度) as 试片1,concat('C1-',二层材料,'-',二层厚度) as 试片2,程序段编号,bb.ID from 试片列表 aa left join "+ prodname.Replace("-001","Process")+ " bb on aa.程序段编号=bb.UUID where 产品图号='" + prodname + "' order by 试片1,试片2").Tables[0];


            dataGridView2.DataSource = coupondt;

            foreach(var dd in controlList.Values)
            {
                //  dd.SetVisible(false);
                dd[1].ForeColor = Color.Black;
                dd.Skip(2).SetPropValue("Visible", false);

               
                //清除已存在的紧固件信息
          dd.Where(x => x.Name.Contains("comboBox")).SetPropValue("Text","");
                

            }


            foreach (DataRow dr in coupondt.Rows)
            {

                string querystr = (dr["试片1"].ToString() + "/" + dr["试片2"].ToString()).Replace("C1-","");

                var controlls = controlList.Values.Where(dd => dd[0].Text == querystr);
                    

                    if(controlls.Count()>0)
                {
                    var controlArray = controlls.First();
                    //  controlArray.SetVisible(true);
                    controlArray[1].ForeColor = Color.ForestGreen;
                    controlArray.Skip(2).SetPropValue("Visible", true);


                   var progname = dr["程序段编号"].ToString().Split('_');

                    if(progname.Count()>1)
                    {
                        string fstname = dr["程序段编号"].ToString().Split('_')[1];
                if(fstname.Contains("B020600"))
                        {

                            fstname = fstname + "-3";

                        }

               

                  //把紧固件填入comboBox中
                    for (int m=2;m<controlArray.Count();m++)
                {
                      if(  controlArray[m].Text=="")
                        {
                            controlArray[m].Text = fstname;
                            break;
                        }

                }

                    }


                }
                    else

                {
                    listBox1.Items.Add(dr["组合编号"].ToString() + ";" + dr["程序段编号"].ToString() +";"+ querystr);



                }



            }

       

        }



        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
           
        }

        private void coupon_tweak(object sender, EventArgs e)
        {
            var lb2 = (Label)sender;
            var selectedcontrolList = controlList.Values.Where(dd => dd[1] == lb2);
            if (selectedcontrolList.Count()>0)
            {

     
           var selectedControl= selectedcontrolList.First();

            if (lb2.ForeColor == Color.Black)
            {

                 lb2.ForeColor =Color.ForestGreen;
                selectedControl.Skip(2).SetPropValue("Visible", true);
               // selectedControl.Skip(2).SetPropValue("Text", "");
            }
            else
            {
                    lb2.ForeColor = Color.Black;
                    selectedControl.Skip(2).SetPropValue("Visible", false);
            }

            }


        }

        private void productList1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Copy all files into the CF card
            System.IO.DriveInfo[] allDrives = System.IO.DriveInfo.GetDrives();
            var targetDri = allDrives.Where(x => x.DriveType.ToString().ToUpper() != "CDROM" && x.VolumeLabel == "NC_PROGRAM");


            if (targetDri.Count() > 0)
            {
                string outputfolder = Program.InfoPath + productList1.listBox1.SelectedItem.ToString() + "\\NC\\COUPON\\";
                string newfoldername = targetDri.First().RootDirectory.FullName;
                localMethod.backupfolder(newfoldername);
                List<FileInfo> files = new List<FileInfo>();
                files.WalkTree(outputfolder, false);
                files.copyto(newfoldername);


            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(listBox2.SelectedIndex!=-1)
            {
                string product = listBox2.SelectedItem.ToString();
                string sql = "(select concat('C1-SKIN-',蒙皮厚度) as 试片,count(*) as 数量 from 试片列表 where 产品图号 like '[0]%' group by 蒙皮厚度) union (select concat('C1-', 二层材料, '-', 二层厚度), count(*) from 试片列表 where 产品图号 like '[0]%'group by 二层材料, 二层厚度)";

                switch(product)
                {
                    case "总计":
                       sql= sql.Replace("[0]", "");
                        break;
                    case "前机身":
                       sql= sql.Replace("[0]", "C023");
                        break;
                    case "中机身CS300":
                       sql= sql.Replace("[0]", "C017");
                        break;
                    case "中机身CS100":
                       sql= sql.Replace("[0]", "C013");
                        break;
                    case "舱门":
                       sql= sql.Replace("[0]", "C015");
                        break;



                }

                dataGridView1.DataSource = DbHelperSQL.Query(sql).Tables[0];

                    
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            excelMethod.SaveDataTableToExcel((DataTable)dataGridView1.DataSource);
           
        }

      


        private  void editCPName(object sender, EventArgs e)
        {
            var cl = (Label)sender;
            string oldstr = cl.Text;
            var labelIndex = (from p in controlList
                              where p.Value[0] == cl
                              select p.Key).First();

            CPNameEditor f = new CPNameEditor(labelIndex,oldstr, this);
            f.Show();
            //string newStr = localMethod.VBInputBox("不需要试片输入NONE", "输入新的编号", oldstr);
            //if (newStr != "")
            //{
            //    localMethod.UpdateConfigValue(oldstr, newStr, "CouponCfg.py");
            //}

        }
        public void UpdateLabelName(string lbi, string name)
        {

            controlList[lbi][0].Text = name;
        }

    }
}
