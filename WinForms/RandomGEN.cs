using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AUTORIVET_KAOHE
{
    public partial class RandomGEN : Form
    {
        public RandomGEN()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DataTable dt = new DataTable();
            dt.Columns.Add("指标", typeof(string));
            dt.Columns.Add("1", typeof(double));
            dt.Columns.Add("2", typeof(double));
            dt.Columns.Add("3", typeof(double));
            dt.Columns.Add("4", typeof(double));
            dt.Columns.Add("5", typeof(double));
            Random ran = new Random();

            Func<object[],double,int, object[]> filldt = delegate(object[] tt, double yy,int time)
            {
                for (int i = 0; i < 5; i++)
                {

                    int rangen = ran.Next(-50, 50);
                    double dif = (double)rangen / time;
                    int dcr = yy.ToString().Length-2 ;

                  

                   tt[i+1] = Math.Round(yy + dif, dcr);


                }
                return tt;
            };


            var ddd = new object[6];
            ddd[0] = "齐平度";
            dt.Rows.Add(filldt(ddd, 0.0009, 120000));

            ddd[0] = "直径";
            dt.Rows.Add(filldt(ddd, 0.258, 7000));
            ddd[0] = "高度";
            dt.Rows.Add(filldt(ddd, 0.081, 5000));


            dataGridView1.DataSource = dt;

        }
    }
}
