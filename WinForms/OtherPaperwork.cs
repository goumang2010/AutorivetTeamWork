using FileManagerNew;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using mysqlsolution;
using OFFICE_Method;
using System.IO;

namespace AUTORIVET_KAOHE
{
    public partial class otherPaperwork : Form
    {
        public otherPaperwork()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog kk = new FolderBrowserDialog();
            kk.Description = "选择要扫描的文件夹";

            if (kk.ShowDialog() == DialogResult.OK)
            {
                string scanpath = kk.SelectedPath;




                List<FileInfo> allfiles=new List<FileInfo>();
                allfiles.WalkTree(scanpath);

                scanfiledoc(allfiles.extfilter("doc").namefilter("", "~$").ToList());

            }

        }
        private void otherpwrf()
        {
            dataGridView1.DataSource = DbHelperSQL.Query("select 文件名,文件编号,文件名称,文件地址,关联产品,文件格式,修改日期,工作包 from otherpaperwork ").Tables[0];
        }

        private void scanfiledoc(List<FileInfo> allfiles)
        {
          var faildoc=new System.Data.DataTable();
            //title[0] = "名称";
            //title[1] = "路径";
            faildoc.Columns.Add("名称", typeof(string));
            faildoc.Columns.Add("路径", typeof(string));

            int count = allfiles.Count;
            List<string> creatName = new List<string>();

            for (int i = 0; i < count; i++)
            {
                //文件名
                string filename = allfiles[i].Name;
                //修改日期
                string revdate = allfiles[i].LastWriteTime.ToShortDateString();
                string filetype = allfiles[i].Extension;


                string filefullname = allfiles[i].FullName.Replace("\\", "\\\\"); ;
                Document abc = wordMethod.opendoc(filefullname,false);
                string bianhao;
                string bianhaotest = wordMethod.gettext(abc, 3, 1);
                if (bianhaotest.Contains("编"))
                {

                    bianhao = wordMethod.gettext(abc, 3, 2);
                }
                else
                {
                    faildoc.Rows.Add(filename,filefullname);
                  
                    continue;

                }
                //获取装配图号
                string productname;
                string prodtest = wordMethod.gettext(abc, 4, 1);
                if (prodtest.Contains("装"))
                {
                    productname = wordMethod.gettext(abc, 4, 2);
                }
                else
                {
                    faildoc.Rows.Add(filename ,filefullname);
                    continue;
                }

                //获取名称

                string aoname="";
                string aonametest = wordMethod.gettext(abc, 3, 3);
                if (aonametest.Contains("名"))
                {
                    aoname = wordMethod.gettext(abc, 3, 4);
                }
                else
                {
                    faildoc.Rows.Add(filename , filefullname);
                    continue;
                }


                //工作包分类
                string pkg;
                if (bianhao.Contains("CR"))
                {
                    pkg = "后桶段";
                }
                else
                {
                    if (bianhao.Contains("CM"))
                    {
                        pkg = "中机身";

                    }
                    else
                    {
                        faildoc.Rows.Add(filename ,filefullname);
                        continue;
                    }
                }

                StringBuilder strSqlname = new StringBuilder();
                strSqlname.Append("INSERT INTO otherPaperwork (");

                strSqlname.Append("文件名,文件编号,文件名称,文件地址,关联产品,文件格式,修改日期,工作包");

                strSqlname.Append(string.Format(") VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}') ON DUPLICATE KEY UPDATE 文件地址='{3}'", filename, bianhao, aoname, filefullname, productname, filetype, revdate,pkg));



                creatName.Add(strSqlname.ToString());

                abc.Application.Quit();


            }

            DbHelperSQL.ExecuteSqlTran(creatName);
            otherpwrf();


             excelMethod.SaveDataTableToExcel( faildoc);
        }

        private void otherPaperwork_Load(object sender, EventArgs e)
        {
            otherpwrf();
        }

        private void button2_Click(object sender, EventArgs e)
        {
           System.Data.DataTable allfiles = DbHelperSQL.Query("select 文件编号,文件地址 from otherPaperwork").Tables[0];
           string outpufolder = @"\\192.168.3.32\Autorivet\output\INFO\files\Other\";
           foreach (DataRow pp in allfiles.Rows)
            {
                string targetpath = outpufolder + pp[0].ToString() + ".pdf";
                wordMethod.WordToPDF(pp[1].ToString(), targetpath);
            }

            
            MessageBox.Show("完成");
        }
    }
}
