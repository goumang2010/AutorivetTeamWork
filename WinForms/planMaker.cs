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
    public partial class planMaker : Form
    {
        DataTable datedt;
        public planMaker()
        {
            InitializeComponent();
        }

        private void planMaker_Load(object sender, EventArgs e)
        {
         //   dataGridView1.DataSource = get_datatable();
            listBox1.DataSource = AutorivetDB.fullname_list("名称");
            listBox1.SelectedIndex = -1;

            rf_gridview();
        }

        public void  rf_gridview()
        {
            datedt=DbHelperSQL.Query("select 日期,日期类型 from everyDay where TO_DAYS(NOW()) < TO_DAYS(日期);").Tables[0];
            dataGridView2.DataSource = datedt;
            label4.Text = calWorkday().ToString();


        }

        public DataTable get_datatable()
        {
            return DbHelperSQL.Query("select 日期,日期类型 from everyDay where TO_DAYS(NOW()) < TO_DAYS(日期);").Tables[0];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime dt1 = dateTimePicker1.Value;
            DateTime dt2 = dateTimePicker2.Value;
            
            TimeSpan tsdiffer = dt2.Date - dt1.Date;
            //截止日期也算一天

            int intDiffer = tsdiffer.Days+1;

            if (intDiffer>0)
            {
                

                for (int i = 0; i < intDiffer;i++ )
                {
                    DateTime dtnew=dt1.Date.AddDays(i);

                    if (dtnew.DayOfWeek == System.DayOfWeek.Sunday || dtnew.DayOfWeek == System.DayOfWeek.Saturday)
                    {

                    }

                    DbHelperSQL.ExecuteSql("insert into everyDay (日期,日期类型) value('" + dtnew.ToShortDateString() + "','" + comboBox1.Text + "') ON DUPLICATE KEY UPDATE 日期类型='"+comboBox1.Text+"';");
               
                
                }

                   
            }
            rf_gridview();
        }

        private int calWorkday()
        {
         //   DataTable dtt = (DataTable)dataGridView2.DataSource;

            var kk = from p in datedt.AsEnumerable()
                     where p["日期类型"].ToString() == "工作日"
                     select p;
            return kk.Count();



        }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime dt1 = dateTimePicker1.Value;
            DateTime dt2 = dateTimePicker2.Value;

            TimeSpan tsdiffer = dt2.Date - dt1.Date;
            //截止日期也算一天

            int intDiffer = tsdiffer.Days + 1;

            if (intDiffer > 0)
            {


                for (int i = 0; i < intDiffer; i++)
                {
                    DateTime dtnew = dt1.Date.AddDays(i);

                    if (dtnew.DayOfWeek == System.DayOfWeek.Sunday || dtnew.DayOfWeek == System.DayOfWeek.Saturday)
                    {
                        DbHelperSQL.ExecuteSql("insert ignore into everyDay (日期,日期类型) value('" + dtnew.ToShortDateString() + "','周末');");


                    }
                    else
                    {
                        DbHelperSQL.ExecuteSql("insert ignore into everyDay (日期,日期类型) value('" + dtnew.ToShortDateString() + "','工作日');");


                    }

                  
                }


            }
            rf_gridview();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox2.Items.AddRange(listBox1.Items);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            if (listBox1.SelectedIndex != -1)
            {
                foreach (var kk in listBox1.SelectedItems)
                {
                      listBox2.Items.Add(kk);
                }
               
            }
            listBox1.SelectedIndex = -1;
      
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int index = listBox2.SelectedIndex;
            if (index != -1)
            {
                //int index = listBox2.SelectedIndex;
                listBox2.Items.Remove(listBox2.SelectedItem.ToString());
                listBox2.SelectedIndex = index - 1;

            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex >0)
            {
                int index = listBox2.SelectedIndex;
                object item = listBox2.Items[index];
                listBox2.Items.RemoveAt(index);

                listBox2.Items.Insert(index - 1, item);
                listBox2.SelectedIndex = index - 1;

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex != -1)
            {
                int index = listBox2.SelectedIndex;
                object item = listBox2.Items[index];
                listBox2.Items.RemoveAt(index);
                listBox2.Items.Insert(index +1, item);
                listBox2.SelectedIndex = index+1;

            }
        }

       // public delegate DataRow getrow(DataTable abc);


        public DataRow defrow(DataTable abc)
        {
            DataRow drtmp = abc.NewRow();
            drtmp["产品架次"] = "SACI0002";
            return drtmp;
        }
        public DataRow defrow2(DataTable abc)
        {
            DataRow drtmp = abc.NewRow();
            drtmp["产品架次"] = "SACI0002";
            return drtmp;
        }

        private void button8_Click(object sender, EventArgs e)
        {


            DbHelperSQL.ExecuteSql("update everyDay set 事件='',备注='';");
            //获取最新架次
            var prodlist = listBox2.Items;
            if (prodlist.Count==0)
            {
                return;
            }
           DataTable abc=DbHelperSQL.Query("select 产品名称,产品架次 from 产品流水表").Tables[0];

           var prodlists1 = from kk in abc.AsEnumerable()
                            group kk by kk["产品名称"].ToString() into os1
                            select new
                            {
                                prodname = os1.Key,
                                jiaci = os1.Max(tk => tk["产品架次"].ToString())

                            };
    

            //获取最大架次
  Dictionary<string,int> dict = prodlists1.ToDictionary(k => k.prodname,  v => System.Convert.ToInt32( v.jiaci.Replace("SACI","")));
  DataTable dt = new DataTable();
  dt.Columns.Add("产品", typeof(string));
  dt.Columns.Add("开始日期", typeof(string));
  dt.Columns.Add("结束日期", typeof(string));
  int prodcount = prodlist.Count;


            //sql队列

  List<string> duilie = new List<string>();




       foreach(string kkk in listBox3.Items)
       {

        string[] tmparray=   kkk.Split(';');

       DateTime dt1= DateTime.Parse(tmparray[0]);

       DateTime dt2 = DateTime.Parse(tmparray[1]);

       int speed = System.Convert.ToInt16(tmparray[2]);

       //获取工作日

       var workdays = from p in datedt.AsEnumerable()
                      where p["日期类型"].ToString() == "工作日" && DateTime.Parse(p["日期"].ToString()) > dt1 && DateTime.Parse(p["日期"].ToString())  <dt2
                      select p["日期"].ToString();


       int count = workdays.Count();

       int intrecord = -1;
       for (int i = 0; i < count; i++)
       {
           string sdata = workdays.ElementAt(i).ToString();
           int prodindex = i % (prodcount * speed);
           int pdindex = (int)Math.Floor((double)prodindex / speed);


           string sprod = prodlist[pdindex].ToString();
           string jiaci;

           if (dict.Keys.Contains(sprod))
           {
               if (intrecord != pdindex)
               {
                   dict[sprod] += 1;
               }



               jiaci = "SACI" + dict[sprod].ToString().PadLeft(4, '0');


           }
           else
           {
               //CS300从第二架次开始
               jiaci = "SACI0009";
               if (intrecord != pdindex)
               {

                   dict.Add(sprod, 9);
               }
           }

           intrecord = pdindex;


           duilie.Add("update everyDay set 事件='" + sprod + "',备注='" + jiaci + "' where 日期='" + sdata + "';");

       }





       DbHelperSQL.ExecuteSqlTran(duilie);

         






            }
            excelMethod.SaveDataTableToExcel(DbHelperSQL.Query("select a.事件 as 产品,b.图号,b.站位号,a.备注 as 架次,min(a.日期) as 开始日期,max(a.日期) as 结束日期 from everyDay a left join 产品列表 b on a.事件=b.名称 group by 产品,架次 order by 开始日期").Tables[0]);




        }

        private void button9_Click(object sender, EventArgs e)
        {
            excelMethod.SaveDataTableToExcel(DbHelperSQL.Query("select a.事件 as 产品,b.图号,b.站位号,a.备注 as 架次,min(a.日期) as 开始日期,max(a.日期) as 结束日期 from everyDay a left join 产品列表 b on a.事件=b.名称 group by 产品,架次 order by 开始日期").Tables[0]);

        }

        private void button10_Click(object sender, EventArgs e)
        {
            DateTime dt1 = dateTimePicker3.Value;
            DateTime dt2 = dateTimePicker4.Value;

            
             TimeSpan tsdiffer = dt2.Date - dt1.Date;
            //截止日期也算一天

            int intDiffer = tsdiffer.Days + 1;

            if (intDiffer > 0)
            {

                listBox3.Items.Add(dt1.ToShortDateString() + ";" + dt2.ToShortDateString() + ";" + comboBox2.Text);


            }

        }

        private void button11_Click(object sender, EventArgs e)
        {
            if(listBox3.SelectedIndex>=0)
            {


                listBox3.Items.RemoveAt(listBox3.SelectedIndex);

            }
        }


    }
}
