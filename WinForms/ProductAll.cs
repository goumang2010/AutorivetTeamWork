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

namespace AUTORIVET_KAOHE
{
    public partial class ProductAll : Form
    {
        public ProductAll()
        {
            InitializeComponent();
        }

        private void ProductAll_Load(object sender, EventArgs e)
        {
            DataGridViewCheckBoxColumn col_btn_insert = new DataGridViewCheckBoxColumn();
            col_btn_insert.HeaderText = "操作1";
            col_btn_insert.Name = "choose";
            //col_btn_insert.Text = "编写";//加上这两个就能显示
         //   col_btn_insert.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(col_btn_insert);

            updateinfo();
            prodallrf();
        }
     
        private void updateinfo()
        {
            List<string> updatearray = new List<string>();
            updatearray.Add("update 产品列表 set AOI='',PACR='',COS='',TS='',PACR=''");

            updatearray.Add("update 产品列表 aa inner join (select 文件编号,关联产品 from paperwork where 文件类型='AO') bb on aa.名称=bb.关联产品 set aa.AO=bb.文件编号");
            updatearray.Add("update 产品列表 aa inner join (select 文件编号,关联产品 from paperwork where 文件类型='AOI') bb on aa.名称=bb.关联产品 set aa.AOI=bb.文件编号");

            updatearray.Add("update 产品列表 aa inner join (select 文件编号,关联产品 from paperwork where 文件类型='COS') bb on aa.名称=bb.关联产品 set aa.COS=bb.文件编号");
            updatearray.Add("update 产品列表 aa inner join (select 文件编号,关联产品 from paperwork where 文件类型='TS') bb on aa.名称=bb.关联产品 set aa.TS=bb.文件编号");

            updatearray.Add("update 产品列表 aa inner join (select 文件编号,关联产品 from paperwork where 文件类型='PACR') bb on aa.名称=bb.关联产品 set aa.PACR=bb.文件编号");
            updatearray.Add("update 产品列表 aa inner join (select count(文件名) as AAOqty,关联产品 from paperwork where 文件类型 like '%AAO' group by 关联产品) bb on aa.名称=bb.关联产品 set aa.AAO=bb.AAOqty");
            updatearray.Add("update 产品列表 aa inner join (select count(*) as prodqty,产品名称 from 产品流水表 group by 产品名称) bb on aa.名称=bb.产品名称 set aa.生产=bb.prodqty");
            DbHelperSQL.ExecuteSqlTran(updatearray);
        
        
        
        }
        private void prodallrf()
    {
        dataGridView1.DataSource = DbHelperSQL.Query("select 图号,名称,站位号,图纸版次,AOI,PACR,AO,COS,TS,PROGRAM,生产,AAO from 产品列表").Tables[0];

    }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int k = e.ColumnIndex;
            int l = e.RowIndex;
            DataRowView temprow = (DataRowView)dataGridView1.Rows[l].DataBoundItem;
            if(k>4)
            {
                paperWork ppwf = new paperWork();
                ppwf.Show();
                ppwf.ppworkrf(temprow[k-1].ToString(),1);
            }
            else
            {
                if(k>0)
                {
                    paperWork paperf = new paperWork();
                    paperf.Show();
                    paperf.ppworkrf(temprow["名称"].ToString());
                }
                else
                {
                    
                }
               
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excelMethod.SaveDataTableToExcel(DbHelperSQL.Query("select * from 产品列表").Tables[0]);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable output = (DataTable)dataGridView1.DataSource;
            excelMethod.SaveDataTableToExcel(output);
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

        private void OperateFiles(Action<DataGridViewRow> op)
        {

            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where Convert.ToBoolean(dd.Cells["choose"].Value) == true
                      select dd;
            foreach (DataGridViewRow temprow in kkk)
            {
                op(temprow);
            }


            MessageBox.Show("完成");
        }



        private void button3_Click(object sender, EventArgs e)
        {

            int i = 0;


     Action<DataGridViewRow> op = delegate (DataGridViewRow temprow)
                  {
                      Dictionary<string, string> aoindex = new Dictionary<string, string>();
               //组件号

               string prdname = temprow.Cells["图号"].Value.ToString();
                      aoindex.Add("组件号", prdname);
                      aoindex.Add("图号", prdname);


                      aoindex.Add("中文名称", temprow.Cells["名称"].Value.ToString());
               //装配图号
               aoindex.Add("装配图号", prdname.Replace("-001", ""));

               //装配大纲编号

               //    aoindex.Add("大纲编号", aoname);

               //AOI编号
               string aoiname = AutorivetDB.queryno(prdname, "AOI编号");
                      aoindex.Add("AOI编号", aoiname);

                      aoindex.Add("程序编号", AutorivetDB.queryno(prdname, "程序编号"));
               //状态编号
               // aoindex.Add("状态编号", autorivet_op.queryno(prdname, "状态编号"));
               //站位号
               //  string zhanwei = aoname.Split('-')[1];
               //  aoindex.Add("站位号", zhanwei);

               //图纸版次

               string dwgrev = temprow.Cells["图纸版次"].Value.ToString().Remove(0, 1);
                      aoindex.Add("图纸版次", dwgrev);
                      string foldername = Program.InfoPath + temprow.Cells["名称"].Value.ToString() + "_" + prdname + "\\AO\\";
                      localMethod.creatDir(foldername);
               //保存地址
               aoindex.Add("索引保存地址", foldername + prdname + "_" + temprow.Cells["站位号"].Value.ToString() + "_INDEX.doc");
                      aoindex.Add("AOI保存地址", foldername + prdname + "_" + temprow.Cells["站位号"].Value.ToString() + "_AOI.doc");
                      aoindex.Add("PACR保存地址", foldername + prdname + "_" + temprow.Cells["站位号"].Value.ToString() + "_PACR.doc");
                      aoindex.Add("复制单保存地址", foldername + prdname + "_" + temprow.Cells["站位号"].Value.ToString() + "_COPY.doc");
                      aoindex.Add("鉴定表保存地址", foldername + prdname + "_" + temprow.Cells["站位号"].Value.ToString() + "_VERI.doc");
                      creatFile(aoindex);
                      i = i + 1;
                  };

            OperateFiles(op);

            MessageBox.Show("完成");


        }








        private void creatFile(object Dic)
        {

            Dictionary<string, string> aoindex =( Dictionary<string, string>)Dic;
               if(checkBox2.Checked)
                {
                    FormMethod.creatAOI(aoindex, checkBox1.Checked,checkBox7.Checked);

                }

                if (checkBox4.Checked)
                {
                    FormMethod.creatPACR(aoindex, checkBox1.Checked, checkBox7.Checked);

                }
            if (checkBox5.Checked)
            {
                FormMethod.creatCOPYSH(aoindex, checkBox1.Checked, checkBox7.Checked);

            }
            if (checkBox6.Checked)
            {
                FormMethod.creatVERI(aoindex, checkBox1.Checked, checkBox7.Checked);

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            int k = FormMethod.killProcess("WINWORD");

            MessageBox.Show("执行成功，杀死" + k.ToString() + "个word进程");
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where dd.Cells["图号"].Value.ToString().Contains("C013")
                      select dd.Cells["choose"];
            if (checkBox8.Checked == true)
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

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            var kkk = from DataGridViewRow dd in dataGridView1.Rows
                      where dd.Cells["图号"].Value.ToString().Contains("C017")
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
                      where dd.Cells["图号"].Value.ToString().Contains("C023")
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

        private void button5_Click_1(object sender, EventArgs e)
        {
            var pacri_ls = DbHelperSQL.getlist("select AOI编号 from 产品列表");
            var ddd = from t in pacri_ls
                      orderby t
                      group t by t.Split('-')[1] into g
                      
                      select new
                      {
                         ACC=g.Key,
                         PACR=g
                      };
                     
            foreach (var kk in ddd)
            {
                string newfilepath = Program.InfoPath + "MATERIAL\\PACR_INDEX\\PACR_INDEX_" + kk.ACC + ".doc";

                System.IO.File.Copy(Program.InfoPath + "SAMPLE\\AOI\\PACR_INDEX.doc", Program.InfoPath+ "MATERIAL\\PACR_INDEX\\PACR_INDEX_"+kk.ACC+".doc", true);
                var page = FileManagerNew.wordMethod.opendoc(newfilepath);
                page.Tables[1].Cell(2, 4).Range.Text = kk.ACC;
                page.Tables[1].Cell(2, 8).Range.Text ="C1-"+ kk.ACC+"-PACR-NEW";
                int i = 1;
                foreach (string d in kk.PACR)
                {

                    page.Tables[1].Cell(3+i,1).Range.Text =i.ToString();
                    page.Tables[1].Cell(3 + i, 2).Range.Text = d+"-PACR";
                    page.Tables[1].Cell(3 + i, 3).Range.Text = "NEW";
                    i = i + 1;

                }

                page.Save();
                page.Application.Quit();

            }
            //pacri_ls.ForEach(x=>x.Split("")
           
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int i = 0;


            Action<DataGridViewRow> op = delegate (DataGridViewRow temprow)
            {
                string AOInum = temprow.Cells["AOI"].Value.ToString();
                if(AOInum=="")
                {
                    return;
                }
                string chnname = temprow.Cells["名称"].Value.ToString(); 

              if(checkBox11.Checked)
                {

                    //获取最新架次号

                    string sqlstr = "select 产品架次 from 产品流水表 where 产品名称='" + chnname + "' order by 产品架次 desc;";
                    var batchlst = DbHelperSQL.getlist(sqlstr);
                    string lastbatch;
                    if (batchlst.Count() == 0)
                    {
                        lastbatch = "SACI0007";
                    }
                    else
                    {
                 lastbatch = DbHelperSQL.getlist(sqlstr).First();
                    }
                       
                    int lastnum = System.Convert.ToInt32(lastbatch.Substring(4));
                    string newbatch = "SACI" + (lastnum + 1).ToString().PadLeft(4, '0');

                    //获取AOI文件
                    string AOIpath = DbHelperSQL.getlist("select 文件地址 from paperWork where 文件类型='AOI' and 关联产品='" + chnname + "';").First();


                    //填充最新架次
                  
                    var wordApp = new Microsoft.Office.Interop.Word.Application();

                   var doc= wordApp.Documents.Open(AOIpath);
                    var cover = doc.Tables[1];
                    cover.Cell(16, 5).Range.Text ="自"+ newbatch + "起";
                   // cover.Cell(4, 4).Range.Text = "N/A";
                    doc.Save();
                    doc.Close();
                    //获取PACR

                    string PACRpath = DbHelperSQL.getlist("select 文件地址 from paperWork where 文件类型='PACR' and 关联产品='" + chnname + "';").First();
                    var pacrdoc= wordApp.Documents.Open(PACRpath);

                    var pacrcover = pacrdoc.Tables[1];
                    pacrcover.Cell(4, 4).Range.Text = "自" + newbatch + "起"; ;
                    pacrdoc.Save();
                    pacrdoc.Close();


                    wordApp.Quit();

                }
            };

            OperateFiles(op);

            MessageBox.Show("完成");



        }
    }



    }

