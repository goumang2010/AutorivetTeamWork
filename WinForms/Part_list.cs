using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GoumangToolKit;
using OfficeMethod;


namespace AUTORIVET_KAOHE
{
    public partial class Part_list : Form
    {
        public Part_list()
        {
            InitializeComponent();
        }
        string unitproductname;
        private void Part_list_Load(object sender, EventArgs e)
        {
            var kk = from pp in Program.prodTable.AsEnumerable()
                    
                     select pp["图号"].ToString() + "_" + pp["名称"].ToString()+ pp["站位号"].ToString();
            listBox1.DataSource = kk.ToList();
            //更新标准件列表
           List<string> buglist= AutorivetDB.update_partlist();
            string note="以下TVA或程序规划有问题：\r\n";
            foreach (string aa in buglist)
            {
                note = note + aa + "\r\n";
            }

            MessageBox.Show(note);

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(listBox1.SelectedIndex!=-1)
            {
               string prodctname= listBox1.SelectedValue.ToString().Split('_')[0];

               label1.Text = prodctname;
               unitproductname = prodctname;

               dataGridView1.DataSource = filterData(tableMode(),new List<string> { prodctname });
               
               var proctb= AutorivetDB.spfproctable(prodctname);
                foreach (DataRow dd in proctb.Rows)
                {
                    dd["加工类型"] = Program.procDic[dd["加工类型"].ToString()];
               

                }

                dataGridView3.DataSource = proctb;
                if (dataGridView3.Columns.Count>0)
                {
                    dataGridView3.Columns[0].Width = 200;
                }
              
                if (checkBox1.Checked)
                {
                    dataGridView2.DataSource = AutorivetDB.parttable(prodctname);
                }
          



            }
            else
            {
                unitproductname = "";
            }





            }

        private void button1_Click(object sender, EventArgs e)
        {
            label1.Text = "全部产品";
            dataGridView1.DataSource = AutorivetDB.spfsttable();
            listBox1.SelectedIndex = -1;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable output = (DataTable)dataGridView1.DataSource;
      excelMethod.SaveDataTableToExcel(output);
          
        }

        private void button3_Click(object sender, EventArgs e)
        {


            showData(delegate
           {
               showHelper(unitproductname + "耗损（试片+吹钉）", AutorivetDB.overqtytable(unitproductname));
            }
               ,
               delegate
               {
                   showHelper("全部耗损（试片+吹钉）", AutorivetDB.overqtytable());
               });
                 
          



        }

        private void showHelper(string text, DataTable dt)
          {

              label1.Text = text;
              dataGridView1.DataSource = dt;
          }


        private void showData(Action t1,Action t2)
        {
            if (unitproductname != "")
            {
                t1();

            }
            else
            {
                //  whole;

                t2();
            }

        }



        private void button4_Click(object sender, EventArgs e)
        {
            label1.Text = "当前耗损（试片+吹钉）";
            dataGridView1.DataSource = AutorivetDB.overqtytable(AutorivetDB.productionname_list());
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    private DataTable filterData  (int mode, List<string> namelist)
        {
            switch (mode)
            {
                case 0:
                    //Total spf qty for one pkg
                    return AutorivetDB.spfQtyStaticTable(AutorivetDB.allqtytable, namelist);


                case 1:

                    return AutorivetDB.spfQtyStaticTable(AutorivetDB.overqtytable, namelist);

                case 2:

                    return AutorivetDB.spfQtyStaticTable(AutorivetDB.spfsttable, namelist);
                default:
                    return null;


            }


        }

   private int tableMode()
        {
            if (radioButton1.Checked == true)
            {
                return 0;
            }
            else
            {
                if (radioButton2.Checked == true)
                {
                    return 1;
                }
                else
                {
                    return 2;
                }
            }

        }


        private void listBox2_DoubleClick(object sender, EventArgs e)
        {





            if(listBox2.SelectedIndex!=-1)
            {
                switch (listBox2.SelectedItem.ToString())
                {
                    case "总计":


                        label1.Text =" 总计";


                        dataGridView1.DataSource = filterData(tableMode(), AutorivetDB.name_list());


                        break;
                    case "前机身":


                        label1.Text = "前机身";


                        dataGridView1.DataSource = filterData (tableMode(),AutorivetDB.name_list("001", "C0231"));


                        break;
                    case "中机身CS100":
                        label1.Text = "中机身CS100";
                        dataGridView1.DataSource = filterData(tableMode(), AutorivetDB.name_list("001", "C013")); 

                        break;
                    case "中机身CS300":
                        label1.Text = "中机身CS300";
                        dataGridView1.DataSource = filterData(tableMode(), AutorivetDB.name_list("001", "C017"));

                        break;




                }
            }
          
        }

        private void button5_Click(object sender, EventArgs e)
        {
           // Microsoft.Office.Interop.Excel.Worksheet abc = new Microsoft.Office.Interop.Excel.Worksheet ();
          //  int i=1;
            DataTable dt=AutorivetDB.overqtytable().Clone();
            foreach(string pp in listBox1.Items)
           
            {

                string prodctname = pp.Split('_')[0];
                dt.Rows.Add(prodctname);
                dt.Merge(AutorivetDB.overqtytable(prodctname));

            }
            excelMethod.SaveDataTableToExcel(dt);

            
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataTable output = (DataTable)dataGridView2.DataSource;
         excelMethod.SaveDataTableToExcel(output);
  
        }

        private void button10_Click(object sender, EventArgs e)
        {
          //  label1.Text = "全部产品";
            dataGridView2.DataSource = AutorivetDB.parttable();
            listBox1.SelectedIndex = -1;
        }

        private void button6_Click(object sender, EventArgs e)
        {

            if(listBox1.SelectedIndex>=0)
            {

          





    






            }






        }

        private void button7_Click(object sender, EventArgs e)
        {
            showData(
                delegate { showHelper("AOI数量",AutorivetDB.allqtytable(unitproductname)); }, delegate { showHelper("全部AOI数量", AutorivetDB.allqtytable()); }
                );
        }
    }
}
