using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using mysqlsolution;
using FileManagerNew;
using System.IO;
using System.Management;
using System.Threading;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using CATIA_method;
using System.Reflection;

namespace AUTORIVET_KAOHE
{
    class FormMethod
    {
        public static Form GetForm(string formname)
        {

            foreach (Form Frm in System.Windows.Forms.Application.OpenForms)
            {
                if (Frm.Name == formname)
                {
                    return Frm;

                }
            }
           // Type type = Type.GetType("AUTORIVET_KAOHE."+formname);
          Assembly assembly = Assembly.GetExecutingAssembly(); // 获取当前程序集 

         var  obj = (Form)(assembly.CreateInstance("AUTORIVET_KAOHE." + formname));
            obj.Show();
            return obj;
            //Form f;
            //switch(formname)
            //{
                  


            //    case "Form3":
            //f= new Form3 ();

            //      f.Show();
            //      return f;
            //       // break;
            //    case "paperWork":


            // f= new paperWork();
            //           f.Show();
            //      return f;
            //    case "Production":
            //     f= new Production();
            //            f.Show();
            //      return f;



            //}


            //return null;
       
             
        }
        public static string GetMachineName()
        {

            try
            {
                System.Net.IPHostEntry IpEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());


                for (int i = 0; i != IpEntry.AddressList.Length; i++)
                {
                    if (IpEntry.AddressList[i].ToString().Contains("192.168"))
                    {
                        // MessageBox.Show(IpEntry.AddressList[i].ToString());

                        return IpEntry.AddressList[i].ToString() + "_" + System.Environment.UserDomainName + "_" + System.Environment.UserName;


                    }
                }


                return "unknown";


            }

            catch
            {

                return "unknown";

            }

        }
        //通知服务器
        public static void notifyServers(string msg)
        {
            var fclient = (frmClient)FormMethod.GetForm("frmClient");
            fclient.sendmsg("notify:" + Program.userID + msg);

        }


        //添加凭据


           public static void cleanbackup()
        {
            var test = new List<FileInfo>();
           
            test.WalkTree(Program.InfoPath);

            //List<FileInfo> dd = fileOP.WalkTree(Program.InfoPath);
           string backupfolder=@"\\192.168.3.32\Autorivet\Prepare\BACKUP\BACKUP_ALL\";
               localMethod.creatDir(backupfolder);

            test.Where(x => x.Name.Contains("backup")).moveto(backupfolder);
            //dd.namefilter("backup").moveto(Program.InfoPath,backupfolder);

               MessageBox.Show("成功");
               


        }
        public static void creatCredential()
        {
                        string targetName = "192.168.3.32";
            IntPtr intPtr = new IntPtr();

            bool flag = false;
            try
            {
                flag = NativeCredMan.WReadCred(targetName, AUTORIVET_KAOHE.NativeCredMan.CRED_TYPE.DOMAIN_PASSWORD, 1, out intPtr);
            }
            catch
            {
                flag = false;
            }
            if (flag)
            {

            }
            else
            {
            //    txtMsg.Text = "该凭据目前不存在";
    

            //ip地址或者网络路径 例如：TERMSRV/192.168.2.222
            string key = "192.168.3.32";
            string userName = @"A03\s50087";
            string password = "s50087";
            //用于标记凭据添加是否成功 i=0:添加成功；i=1:添加失败
            int i = 0;
            try
            {
                i = NativeCredMan.WriteCred(key, userName, password
                    , AUTORIVET_KAOHE.NativeCredMan.CRED_TYPE.DOMAIN_PASSWORD, AUTORIVET_KAOHE.NativeCredMan.CRED_PERSIST.LOCAL_MACHINE);
            }
            catch
            {
                i = 1;
            }

            if (i == 0)
            {

                
                    MessageBox.Show("成功添加凭据,请重启或稍等后再打开共享文件");

               // System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location);  //重新开启当前程序
                //Close();//关闭当前程序
            }
              
            else
                {

                    MessageBox.Show("未能成功添加凭据，这将影响正常访问共享文件，请咨询管理员");

                }
            }

            
        }




        //杀死word
      public static string GetProcessUserName(int pID)
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
        public static int  killProcess(string procName)
        {
            int k = 0;
            if (MessageBox.Show("工具会首先杀死所有的"+ procName+"进程，请先保存当前的"+ procName+"工作，然后点击确定", "警告", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
           
                foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName(procName))
                {

                    string username = GetProcessUserName(p.Id);
                    if (username == System.Environment.UserName)
                    {
                        p.Kill();
                        k = k + 1;
                    }

                }
            }

            return k;
         
        }
  

        //添加自动生成的文件

        public static void scanfiledoc(string path, Microsoft.Office.Interop.Word.Document myWordfile = null, bool closeword = true)
        {
            List<FileInfo> allfiles = new List<FileInfo>() { new FileInfo(path) };
      

            FormMethod.scanfiledoc(allfiles, myWordfile: myWordfile, closeword: closeword);
            //((Form_Interface)get_Form("paperWork")).get_datatable();
            //((Form_Interface)get_Form("paperWork")).rf_filter();

  
        }










        //备份文件





        public static string get_storefolder(string prodname, string type="")
        {
            //string prodmark =prodname+"_"+ DbHelperSQL.getlist("select 图号 from 产品列表 where 名称='" + prodname + "';").First();

            var kk = from pp in Program.prodTable.AsEnumerable()
                     where pp["名称"].ToString() == prodname
                     select pp["名称"].ToString() + "_" + pp["图号"].ToString();
            

            string prodmark = kk.First();



            string basefolder = Program.InfoPath + prodmark;
            switch (type)
            {
                case "COS":
                    return basefolder + "\\COS\\";
                  //  break;
                case "AO":
                    return basefolder + "\\AO\\";
                   // break;



                default:

                    return basefolder + "\\";
                //    break;

            }
        }
       public static void ConnectDef()
       {
          // System.Diagnostics.Process.Start("cmd.exe", @"/k Net Use");
          //// System.Diagnostics.Process.Start("cmd.exe", @"/k Net Use * \\192.168.3.32\Autorivet qwe123 /user:Dell");
          // Process proc = new Process();
          // proc.StartInfo.FileName = "cmd.exe";
          // //proc.StartInfo.UseShellExecute = false;
          // //proc.StartInfo.RedirectStandardInput = true;
          // //proc.StartInfo.RedirectStandardOutput = true;
          // //proc.StartInfo.RedirectStandardError = true;
          // //proc.StartInfo.CreateNoWindow = false;

          // proc.Start();
          // string command = @"cd\";
          // proc.StandardInput.WriteLine(command);
          // command = @"/k Net Use * \\192.168.3.32\Autorivet qwe123 /user:Dell";
          // proc.StandardInput.WriteLine(command);






         Connect("192.168.3.32","Autorivet", "Dell-PC\\Dell", "qwe123");
       }

       public static bool Connect(string remoteHost, string shareName, string userName, string passWord)
       {
           bool Flag = false;
           Process proc = new Process();
           try
           {
               proc.StartInfo.FileName = "cmd.exe";
               proc.StartInfo.UseShellExecute = false;
               proc.StartInfo.RedirectStandardInput = true;
               proc.StartInfo.RedirectStandardOutput = true;
               proc.StartInfo.RedirectStandardError = true;
               proc.StartInfo.CreateNoWindow = true;
               proc.Start();
               string dosLine = @"net use \\" + remoteHost + @"\" + shareName + " /User:" + userName + " " + passWord + " /PERSISTENT:YES";
               proc.StandardInput.WriteLine(dosLine);
               proc.StandardInput.WriteLine("exit");
               while (!proc.HasExited)
               {
                   proc.WaitForExit(1000);
               }

               string errormsg = proc.StandardError.ReadToEnd();
               proc.StandardError.Close();
               if (String.IsNullOrEmpty(errormsg))
               {
                   Flag = true;
               }
           }
           catch (Exception ex)
           {
               throw ex;
           }
           finally
           {
               //proc.Close();
               //proc.Dispose();
           }
           return Flag;
       }

        public static void open_process(object filepath)
        {
            //备份文件
           localMethod.backupfile((string)filepath);
            //使用Fastip 打开

     


           Process_Method aa = new Process_Method((string)filepath);
      // aa.iniProc();
 
        }
        public static void open_file_local(object filepath)
        {
            //备份文件
            localMethod.backupfile((string)filepath);
            System.Diagnostics.Process aa = new System.Diagnostics.Process();
            aa.StartInfo=new System.Diagnostics.ProcessStartInfo((string)filepath);
            aa.Start();
            Thread.Sleep(5);


        }
          public static void open_file(string filepath)
        {
           
            Thread t1 = new Thread(new ParameterizedThreadStart(open_file_local));
            t1.Start(filepath);
            Thread.Sleep(5);
        }

        public static void scanfiledoc(List<FileInfo> allfiles, bool rpnum = false, Microsoft.Office.Interop.Word.Document myWordfile =null,bool closeword = true)
        {
            int count = allfiles.Count;
            List<string> creatName = new List<string>();
            Microsoft.Office.Interop.Word.Application wordApp=null;

         
            for (int i = 0; i < count; i++)
            {

               // GC.Collect();


                //文件名
                string filename = allfiles[i].Name;
                //修改日期
                string revdate = allfiles[i].LastWriteTime.ToShortDateString();

                //在数据库中搜寻相关数据
                System.Data.DataTable comparefile = DbHelperSQL.Query("select 文件名,修改日期 from paperWork where 文件名='" + filename + "' and 修改日期='" + revdate + "';").Tables[0];
                if (comparefile.Rows.Count == 0)
                {





                    bool inputdata = true;

                    string[] mfprop = filename.Split('_');
                    string biaoshi = mfprop.Last().Replace(allfiles[i].Extension, "");
                    int weishu = mfprop.Count();



                    //文件类型
                    string filetype = "";

                    //文件编号
                    string filenum = "";

                    //文件名称
                    string filetitle = "";

                    //编制日期,这项之后手动填写
                    //string createdate = "";


                    //文件地址
                    string filefullname = allfiles[i].FullName.Replace("\\", "\\\\");

                    //文件夹
                    string foldername = allfiles[i].DirectoryName;
                    //对该目录进行了反转
                    List<string> folderprop = foldername.Split('\\').Reverse().ToList();
                    //关联产品
                    string productname = "";
                    //关联架次
                    string effname = "";
                    //相关文件
                    string relatedfile = "";
                    //文件格式
                   // string fileformat = "word";

                  
                    if (myWordfile == null)
                    {

                       wordApp = new Microsoft.Office.Interop.Word.Application();
                        wordApp.Visible = false;

                        myWordfile = wordApp.Documents.Open(allfiles[i].FullName);
                    }
                    else
                    {
                        wordApp = myWordfile.Application;
                    }

                    switch (biaoshi)
                    {
                        case "AAO":
                            productname = folderprop[3].Split('_')[0];
                            effname = folderprop[2];
                            relatedfile = folderprop[0];

                            filetitle = myWordfile.Tables[1].Cell(2, 4).Range.Text.Split('\r')[0];
                            filenum = myWordfile.Tables[1].Cell(2, 6).Range.Text.Replace("\a", "");
                            filetype = "RNC_AAO";
                            break;
                        case "补铆":
                            productname = folderprop[2].Split('_')[0];
                            effname = folderprop[1];
                            relatedfile = "";
                            filetitle = "对自动钻铆未完成部分进行补铆及排故";
                            filenum = myWordfile.Tables[1].Cell(2, 6).Range.Text.Replace("\a", "");

                            filetype = "补铆_AAO";
                            break;

                        case "COS":
                            productname = folderprop[1].Split('_')[0];
                            effname = "--";
                            relatedfile = "";
                            filetitle = productname + "壁板交付状态";
                            filenum = DbHelperSQL.getlist("select 状态编号 from 产品列表 where 名称='" + productname + "'").First();

                            filetype = "COS";
                            break;
                        case "AO":
                            productname = folderprop[1].Split('_')[0];
                            effname = "--";
                            relatedfile = "";
                            filetitle = "自动钻铆";
                            filenum = DbHelperSQL.getlist("select 大纲编号 from 产品列表 where 名称='" + productname + "'").First();

                            filetype = "AO";
                            break;
                        case "AOI":
                            productname = folderprop[1].Split('_')[0];
                            effname = "--";
                            relatedfile = "";
                            filetitle = "自动钻铆AOI";
                            filenum = DbHelperSQL.getlist("select AOI编号 from 产品列表 where 名称='" + productname + "'").First();

                            filetype = "AOI";
                            break;
                        case "PACR":
                            productname = folderprop[1].Split('_')[0];
                            effname = "--";
                            relatedfile = "";
                            filetitle = "自动钻铆PACR";
                            filenum = DbHelperSQL.getlist("select AOI编号 from 产品列表 where 名称='" + productname + "'").First()+"-PACR";

                            filetype = "PACR";
                            break;
                        case "INDEX":
                            productname = folderprop[1].Split('_')[0];
                            effname = "--";
                            relatedfile = DbHelperSQL.getlist("select 大纲编号 from 产品列表 where 名称='" + productname + "'").First();
                            filetitle = "装配大纲索引";
                            filenum = relatedfile.Replace("-020", "-A");

                            filetype = "INDEX";
                            break;
                        case "TS":
                            productname = folderprop[1].Split('_')[0];
                            effname = "--";
                            relatedfile = "";
                            filetitle = productname + "技术单";
                            filenum = DbHelperSQL.getlist("select 状态编号 from 产品列表 where 名称='" + productname + "'").First().Replace("S", "I");
                            filetype = "TS";
                            if (rpnum)
                            {
                                wordMethod.SearchReplace(wordApp, myWordfile, "C1-I22100-^#", filenum);

                            }
                            break;
                        case "VERI":
                            productname = folderprop[1].Split('_')[0];
                            effname = "--";
                            relatedfile = "";
                            filetitle = productname + "程序鉴定单";
                            filenum = myWordfile.Tables[1].Cell(1,5).Range.Text.Replace("\a", "");
                            filetype = "VERI";
                          
                            break;

                        default:
                            productname = folderprop[1].Split('_')[0];
                            effname = "--";
                            relatedfile = "";
                            filetitle = productname +"_"+ biaoshi;
                           // filenum = myWordfile.Tables[1].Cell(1, 5).Range.Text.Replace("\a", "");
                            filetype = "OTHER";
                            break;

                    }
                    if (closeword && myWordfile != null)
                    {
                        myWordfile.Close();
                        myWordfile = null;
                    }
                
                 

                //    GC.Collect();
                    if (inputdata)
                    {


                        StringBuilder strSqlname = new StringBuilder();
                        strSqlname.Append("INSERT INTO paperwork (");

                        strSqlname.Append("文件名,文件类型,文件编号,文件名称,修改日期,文件地址,状态,关联产品,关联架次,相关文件,位置");

                        strSqlname.Append(string.Format(") VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}') ON DUPLICATE KEY UPDATE 修改日期='{4}'", filename, filetype, filenum, filetitle, revdate, filefullname, "已签发", productname, effname, relatedfile,"工艺"));



                        creatName.Add(strSqlname.ToString());

                    }

                }
            }
            int k = DbHelperSQL.ExecuteSqlTran(creatName);
         

            if(closeword&&wordApp!=null)
            {
               // myWordfile.Save();
                
                wordApp.Quit();
                wordApp = null;
            }
           if(!closeword)
            {
                MessageBox.Show(string.Format("执行成功,增加 '{0}'条记录", k));

            }
            if (count>1)
            {

                ((FormInterface)GetForm("paperWork")).get_datatable();
                ((FormInterface)GetForm("paperWork")).rf_filter();
            }

      
            
        }
        public static void scanfileCatia(List<FileInfo> allfiles)
        {
            int count = allfiles.Count;
            List<string> creatName = new List<string>();

            for (int i = 0; i < count; i++)
            {


                //文件名
                string filename = allfiles[i].Name;
                //修改日期
                string revdate = allfiles[i].LastWriteTime.ToShortDateString();
                string ext=allfiles[i].Extension;
                //在数据库中搜寻相关数据
                System.Data.DataTable comparefile = DbHelperSQL.Query("select 文件名,修改日期 from paperWork where 文件名='" + filename + "' and 修改日期='" + revdate + "';").Tables[0];
              
              //如果没找到则要进行插入  
                if (comparefile.Rows.Count == 0)
                {



                    
         string filetype = filename.Split('.')[0].Split('_').Last();

                    if(!filetype.Contains(ext))
                    {
                        filetype = filetype + "_" + ext;
                    }
       
                



                    //文件地址
                    string filefullname = allfiles[i].FullName.Replace("\\", "\\\\");

                    //文件夹
                    string foldername = allfiles[i].DirectoryName;
                    //对该目录进行了反转
                    List<string> folderprop = foldername.Split('\\').Reverse().ToList();
                    //关联产品
                    string productname = folderprop[2].Split('_')[0];
                 
                
                   
  

                        StringBuilder strSqlname = new StringBuilder();
                        strSqlname.Append("INSERT INTO 产品数模 (");

                        strSqlname.Append("文件名,文件类型,产品名称,修改日期,文件地址");

                        strSqlname.Append(string.Format(") VALUES ('{0}','{1}','{2}','{3}','{4}') ON DUPLICATE KEY UPDATE 修改日期='{3}'", filename, filetype, productname, revdate, filefullname));



                        creatName.Add(strSqlname.ToString());

                    

                }
            }
            int k = DbHelperSQL.ExecuteSqlTran(creatName);
            MessageBox.Show(string.Format("执行成功,增加 '{0}'条记录", k));

        }






        #region 编写文件

  public static Document creatAOI(Dictionary<string, string> rncaao,bool overcover=false,bool closeword=true)
        {
            localMethod.backupfile(rncaao["AOI保存地址"]);
            if (!rncaao.Keys.Contains("图纸名称"))
            {
                
                rncaao.Add("图纸名称", AutorivetDB.queryno(rncaao["图号"], "图纸名称"));
        
            
            }
        


            var   wordApp = new Microsoft.Office.Interop.Word.Application();



         wordApp.Visible = !closeword;
        
            Document myAO;

                
          if (File.Exists(rncaao["AOI保存地址"]) && (!overcover))
            {
                myAO = wordApp.Documents.Open(rncaao["AOI保存地址"]);

            }
            else
            {
                try
                {
                    File.Copy(Program.InfoPath + "SAMPLE\\AOI\\AOI.docx", rncaao["AOI保存地址"], true);

                }
                catch
                {

                    MessageBox.Show("生成文件" + rncaao["图号"] + "AOI未成功，请关闭所有打开的word，运行kill word！");
                    return null;
                }


                myAO = wordApp.Documents.Open(rncaao["AOI保存地址"]);
                Thread.Sleep(100);
                // myAO.SaveAs2(rncaao["AOI保存地址"]);

            }




           

            //替换所有编号
            wordMethod.SearchReplace(wordApp, myAO, "[2]", rncaao["AOI编号"]);


            //上下工序

            //     myAO.Tables[1].Cell(4, 4).Range.Text =

            string prezhanwei = AutorivetDB.queryno(rncaao["图号"], "预铆编号");
            myAO.Tables[1].Cell(4, 4).Range.Text = "C1-" + prezhanwei + "-020";

            myAO.Tables[1].Cell(4, 6).Range.Text = "C1-" + prezhanwei + "-030";
            //标题


    

            myAO.Tables[1].Cell(3, 4).Range.Text =  "自动钻铆\rAUTO-RIVETING";
            //替换所有装配图号

            wordMethod.SearchReplace(wordApp, myAO, "[1]", rncaao["装配图号"]);





            //进入零件表填写


            //第一条


            myAO.Tables[2].Cell(3, 1).Range.Text = "1";
            //     myAO.Tables[2].Cell(3, 1).Height=
            myAO.Tables[2].Cell(3, 2).Range.Text = rncaao["图号"] + "P1";
            myAO.Tables[2].Cell(3, 3).Range.Text = rncaao["中文名称"] + "壁板组件\r" + rncaao["图纸名称"];
            myAO.Tables[2].Cell(3, 4).Range.Text = "A";
            myAO.Tables[2].Cell(3, 5).Range.Text = "1";
            myAO.Tables[2].Cell(3, 6).Range.Text = rncaao["装配图号"];
            if (rncaao["图号"].Contains("C017"))
            {
                myAO.Tables[1].Cell(3, 6).Range.Text = "☒CS100☑CS300";
            }
          

            // myAO.Tables[1].Cell(7, 10).Range.Text = rncaao["架次"];
            // myAO.Tables[1].Cell(7, 1).Range.Text = dataGridView1.Rows[0].

            //开始统计紧固件

           // System.Data.DataTable fsttable = autorivet_op.spfsttable(rncaao["图号"]);
      //2015.7.28改为统计加入试片耗损后的紧固件数量
            System.Data.DataTable fsttable = AutorivetDB.allqtytable(rncaao["图号"]);

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
                myAO.Tables[2].Cell(fstcount, 5).Range.Text = "340g";
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

                fstcount = fstcount + 1;


                while (!myAO.Tables[2].Cell(fstcount, 1).Range.Text.Contains("专用工具"))
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

                    myAO.Tables[2].Cell(rear, 1).Range.Text = (k + 1).ToString();
                    myAO.Tables[2].Cell(rear, 2).Range.Text = "下铆头\rLower Anvil";
                    myAO.Tables[2].Cell(rear, 3).Range.Text = "1";
                    myAO.Tables[2].Cell(rear, 4).Range.Text = lals[k];

                    rear = rear + 1;

                }
                //高锁限力枪
                int toolindex = lals.Count();
                if (nut5 > 0)
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

                //myAO.Tables[2].Cell(rear, 1).Range.Text = "1";
                //myAO.Tables[2].Cell(rear, 2).Range.Text = rncaao["程序编号"];

                //替换所有程序编号
                wordMethod.SearchReplace(wordApp, myAO, "[3]", rncaao["程序编号"]);


                rear = rear + 1;



                //寻找参考文件列表
                while (!myAO.Tables[2].Cell(rear, 1).Range.Text.Contains("参考文件列表"))
                {


                    rear = rear + 1;


                }

                rear = rear + 8;

                //判断并修改工装编号

                if(rncaao["中文名称"].Contains("下"))
                {
                    wordMethod.SearchReplace(wordApp, myAO, "C1-RIT-C01329000-001", "C1-RIT-C01333400-001");

                    wordMethod.SearchReplace(wordApp, myAO, "中中机身上半壳卡板", "中后机身下半壳卡板");

                    wordMethod.SearchReplace(wordApp,myAO, "Auto riveting framework of upper lobe MID fuselage CTR", "Auto riveting framework of lower lobe MID fuselage AFT");


                    wordMethod.SearchReplace(wordApp, myAO, "C1-RIT-C01322000-001", "C1-RIT-C01332300-001");
                    wordMethod.SearchReplace(wordApp, myAO, "中前机身上半壳卡板", "中前机身下半壳卡板");

                    wordMethod.SearchReplace(wordApp, myAO, "Auto riveting framework of upper lobe MID fuselage FWD", "Auto riveting framework of lower lobe MID fuselage FWD");


                }


           




                //2015.8.13不再需要状态编号
                //  myAO.Tables[2].Cell(rear, 2).Range.Text = rncaao["状态编号"];

                //替换所有状态编号
                //   wordMethod.SearchReplace(wordApp, myAO, "C1-S22100-33", rncaao["状态编号"]);



                //     myAO.Tables[2].Cell(rear, 3).Range.Text = rncaao["中文名称"] + "壁板装配交付状态\rCOS OF " + rncaao["图纸名称"];



            }



            myAO.Save();
            //   wordApp.Quit();
            FormMethod.scanfiledoc(rncaao["AOI保存地址"], myAO,closeword);
          //  MessageBox.Show("生成完毕，请补全其他信息");




            return myAO;

        }





  public static void creatPACR(Dictionary<string, string> rncaao, bool overcover = false,bool closeword=true)
  {

            localMethod.backupfile(rncaao["PACR保存地址"]);
      if (!rncaao.Keys.Contains("图纸名称"))
      {
          rncaao.Add("图纸名称", AutorivetDB.queryno(rncaao["图号"], "图纸名称"));
      }



    var  wordApp = new Microsoft.Office.Interop.Word.Application();
    wordApp.Visible = !closeword;
     var myAO = new Document();
      if (File.Exists(rncaao["PACR保存地址"]) && (!overcover))
      {
          myAO = wordApp.Documents.Open(rncaao["PACR保存地址"]);

      }
      else
      {

                try
                {
                    
                    File.Copy(Program.InfoPath + "SAMPLE\\AOI\\PACR.doc", rncaao["PACR保存地址"], true);
                }
                catch
                {

                    MessageBox.Show("生成文件"+rncaao["图号"]+ "PACR未成功，请关闭所有打开的word，运行kill word！");
                    return;
                }
        

          myAO = wordApp.Documents.Open(rncaao["PACR保存地址"]);


      }



            //控制记录编号
            wordMethod.SearchReplace(wordApp, myAO, "[1]", rncaao["AOI编号"] + "-PACR");
            //      myAO.Tables[1].Cell(2, 6).Range.Text = rncaao["AOI编号"] + "-PACR";
            //myAO.Tables[2].Cell(2, 6).Range.Text = rncaao["AOI编号"] + "-PACR";
            //myAO.Tables[3].Cell(2, 6).Range.Text = rncaao["AOI编号"] + "-PACR";
            //myAO.Tables[4].Cell(2, 6).Range.Text = rncaao["AOI编号"] + "-PACR";

            //控制记录名称


            myAO.Tables[1].Cell(4, 3).Range.Text = "自动钻铆\rAUTO-RIVETING";

      //装配图号
      //myAO.Tables[1].Cell(4, 2).Range.Text = rncaao["装配图号"];
      myAO.Tables[1].Cell(10, 1).Range.Text = "1";
      myAO.Tables[1].Cell(10, 2).Range.Text = rncaao["装配图号"];
      myAO.Tables[1].Cell(10, 3).Range.Text = rncaao["中文名称"] + "壁板组件\r" + rncaao["图纸名称"];
      // myAO.Tables[1].Cell(10, 5).Range.Text = rncaao["图纸版次"];
      //程序编号

      myAO.Tables[1].Cell(16, 1).Range.Text = "1";
      myAO.Tables[1].Cell(16, 2).Range.Text = rncaao["程序编号"];
            //2015.8.13不再需要状态编号
         //   myAO.Tables[1].Cell(16, 6).Range.Text = rncaao["状态编号"];
    //  myAO.Tables[1].Cell(16, 7).Range.Text = rncaao["中文名称"] + "壁板装配交付状态\rCOS OF " + rncaao["图纸名称"];



      //进入零件表填写


      //第一条


      myAO.Tables[4].Cell(5, 1).Range.Text = "1";
      //     myAO.Tables[2].Cell(3, 1).Height=
      myAO.Tables[4].Cell(5, 2).Range.Text = rncaao["图号"] ;
      myAO.Tables[4].Cell(5, 3).Range.Text = rncaao["中文名称"] + "壁板组件\r" + rncaao["图纸名称"];
      myAO.Tables[4].Cell(5, 4).Range.Text = "A";
      myAO.Tables[4].Cell(5, 5).Range.Text = "1";
  //    myAO.Tables[4].Cell(5, 6).Range.Text = rncaao["装配图号"];

      // myAO.Tables[1].Cell(7, 10).Range.Text = rncaao["架次"];
      // myAO.Tables[1].Cell(7, 1).Range.Text = dataGridView1.Rows[0].

      //开始统计紧固件

      System.Data.DataTable fsttable = AutorivetDB.allqtytable(rncaao["图号"]); ;

      int fstcount = fsttable.Rows.Count;
      int nut5 = 0;
      int nut6 = 0;
      if (fstcount != 0)
      {



          for (int i = 0; i < fstcount; i++)
          {
              string fstno = fsttable.Rows[i][0].ToString();
              myAO.Tables[4].Cell(6 + i, 1).Range.Text = (i + 2).ToString();
              myAO.Tables[4].Cell(6 + i, 2).Range.Text = fstno;
              //string fstname = "";


              if (fstno.Contains("B020600"))
              {
                  myAO.Tables[4].Cell(6 + i, 3).Range.Text = "高锁\rHI-LITE";
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
                  myAO.Tables[4].Cell(6 + i, 3).Range.Text = "铆钉\rRIVET";
              }
              myAO.Tables[4].Cell(6 + i, 4).Range.Text = "U";
              myAO.Tables[4].Cell(6 + i, 5).Range.Text = fsttable.Rows[i][1].ToString();
           //   myAO.Tables[4].Cell(6 + i, 6).Range.Text = rncaao["装配图号"];
          }



          fstcount = 6 + fstcount;
          if (nut5 != 0)
          {
              myAO.Tables[4].Cell(fstcount, 1).Range.Text = (fstcount - 2).ToString();
              myAO.Tables[4].Cell(fstcount, 2).Range.Text = "B0203013-08";
              myAO.Tables[4].Cell(fstcount, 3).Range.Text = "高锁帽\rNUT";
              myAO.Tables[4].Cell(fstcount, 4).Range.Text = "U";
              myAO.Tables[4].Cell(fstcount, 5).Range.Text = nut5.ToString();
             // myAO.Tables[4].Cell(fstcount, 6).Range.Text = rncaao["装配图号"];
              if (nut6 != 0)
              {
                  fstcount = fstcount + 1;
                  myAO.Tables[4].Cell(fstcount, 1).Range.Text = (fstcount - 2).ToString();
                  myAO.Tables[4].Cell(fstcount, 2).Range.Text = "B0203013-3";
                  myAO.Tables[4].Cell(fstcount, 3).Range.Text = "高锁帽\rNUT";
                  myAO.Tables[4].Cell(fstcount, 4).Range.Text = "U";
                  myAO.Tables[4].Cell(fstcount, 5).Range.Text = nut6.ToString();
                //  myAO.Tables[4].Cell(fstcount, 6).Range.Text = rncaao["装配图号"];



              }

          }
          else
          {
              if (nut6 != 0)
              {
                  //fstcount = fstcount + 1;
                  myAO.Tables[4].Cell(fstcount, 1).Range.Text = (fstcount - 2).ToString();
                  myAO.Tables[4].Cell(fstcount, 2).Range.Text = "B0203013-3";
                  myAO.Tables[4].Cell(fstcount, 3).Range.Text = "高锁帽\rNUT";
                  myAO.Tables[4].Cell(fstcount, 4).Range.Text = "U";
                  myAO.Tables[4].Cell(fstcount, 5).Range.Text = nut6.ToString();
                  myAO.Tables[4].Cell(fstcount, 6).Range.Text = rncaao["装配图号"];
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

          myAO.Tables[4].Cell(fstcount, 1).Range.Text = (fstcount - 2).ToString();
          myAO.Tables[4].Cell(fstcount, 3).Range.Text = "密封剂\rSEALANT";
          myAO.Tables[4].Cell(fstcount, 2).Range.Text = "MIL-PRF-81733 Type IV";
          myAO.Tables[4].Cell(fstcount, 4).Range.Text = "M";
          myAO.Tables[4].Cell(fstcount, 5).Range.Text = "340g";
          //  myAO.Tables[4].Cell(fstcount, 6).Range.Text = rncaao["装配图号"];


          fstcount = fstcount + 1;
















      }




      myAO.Save();
      FormMethod.scanfiledoc(rncaao["PACR保存地址"], myAO, closeword);
    //  MessageBox.Show("生成完毕，请补全其他信息");




      // return myAO;
  }



        public static void creatCOPYSH(Dictionary<string, string> rncaao, bool overcover = false, bool closeword = true)
        {

            localMethod.backupfile(rncaao["复制单保存地址"]);
            if (!rncaao.Keys.Contains("图纸名称"))
            {
                rncaao.Add("图纸名称", AutorivetDB.queryno(rncaao["图号"], "图纸名称"));
            }



            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = !closeword;
            var myAO = new Document();
            if (File.Exists(rncaao["复制单保存地址"]) && (!overcover))
            {
                myAO = wordApp.Documents.Open(rncaao["复制单保存地址"]);

            }
            else
            {

                try
                {

                    File.Copy(Program.InfoPath + "SAMPLE\\RULE\\REPROD_SH.doc", rncaao["复制单保存地址"], true);
                }
                catch
                {

                    MessageBox.Show("生成文件" + rncaao["图号"] + "复制单未成功，请关闭所有打开的word，运行kill word！");
                    return;
                }


                myAO = wordApp.Documents.Open(rncaao["复制单保存地址"]);


            }

            //处理程序编号
            int progno = System.Convert.ToInt32(rncaao["程序编号"].Substring(5, 2));



            //复制单编号

            myAO.Tables[1].Cell(1, 5).Range.Text ="S2B-NCC-22100-"+ progno.ToString();

            wordMethod.SearchReplace(wordApp, myAO, "[5]", rncaao["图号"]);
            wordMethod.SearchReplace(wordApp, myAO, "[6]", rncaao["中文名称"] + "壁板组件");
            wordMethod.SearchReplace(wordApp, myAO, "[8]", rncaao["程序编号"] );




            myAO.Save();
          //  Form_method.scanfiledoc(rncaao["复制单保存地址"], myAO, closeword);
            //  MessageBox.Show("生成完毕，请补全其他信息");




            // return myAO;
        }



        public static void creatVERI(Dictionary<string, string> rncaao, bool overcover = false, bool closeword = true)
        {

            localMethod.backupfile(rncaao["鉴定表保存地址"]);
            if (!rncaao.Keys.Contains("图纸名称"))
            {
                rncaao.Add("图纸名称", AutorivetDB.queryno(rncaao["图号"], "图纸名称"));
            }



            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = !closeword;
            var myAO = new Document();
            if (File.Exists(rncaao["鉴定表保存地址"]) && (!overcover))
            {
                myAO = wordApp.Documents.Open(rncaao["鉴定表保存地址"]);

            }
            else
            {

                try
                {

                    File.Copy(Program.InfoPath + "SAMPLE\\RULE\\VERI_SH.doc", rncaao["鉴定表保存地址"], true);
                }
                catch
                {

                    MessageBox.Show("生成文件" + rncaao["图号"] + "鉴定表未成功，请关闭所有打开的word，运行kill word！");
                    return;
                }


                myAO = wordApp.Documents.Open(rncaao["鉴定表保存地址"]);


            }

            //处理程序编号
            int progno = System.Convert.ToInt32(rncaao["程序编号"].Substring(5, 2));



            //鉴定表编号

            myAO.Tables[1].Cell(1, 5).Range.Text = "S2B-NCH-22100-" + progno.ToString();

            wordMethod.SearchReplace(wordApp, myAO, "[5]", rncaao["图号"]);
            wordMethod.SearchReplace(wordApp, myAO, "[6]", rncaao["中文名称"] + "壁板组件");
            wordMethod.SearchReplace(wordApp, myAO, "[7]", rncaao["图纸版次"] );
            wordMethod.SearchReplace(wordApp, myAO, "[11]", rncaao["程序编号"]);
            //开始处理试片编号
        //  DbHelperSQL.ExecuteSqlTran(autorivet_op.rfcouponno(rncaao["图号"]));

        //List< string> cplist=   DbHelperSQL.getlist("select CONCAT('T',编号) from 试片列表 where 产品图号='"+rncaao["图号"] + "' group by 编号 order by 编号;");

        //    string couponstr = rncaao["图号"].Replace("-001", cplist.First());

        //    for(int i=1;i<cplist.Count();i++)
        //    {
        //        couponstr = couponstr + "," + cplist[i];
        //    }

        //    wordMethod.SearchReplace(wordApp, myAO, "[8]", couponstr);
            myAO.Save();
            FormMethod.scanfiledoc(rncaao["鉴定表保存地址"], myAO, closeword);
            //  MessageBox.Show("生成完毕，请补全其他信息");




            // return myAO;
        }













        #endregion










    }
}
