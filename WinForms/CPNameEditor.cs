using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GoumangToolKit;

namespace AUTORIVET_KAOHE
{
    public partial class CPNameEditor : Form
    {
        couponTest f1;
        string oldstr;
        string lbinx;
        string ind;
        public CPNameEditor(string labelIndex,string CPname,couponTest f)
        {
            InitializeComponent();
            f1 = f;
            ind = labelIndex;
            lbinx = "\"" + labelIndex + "\":\"";
            label5.Text = "编号:" + labelIndex;
            label4.Text = "原试片:" + CPname;
            oldstr = lbinx + CPname + "\"";
            if (CPname=="NONE")
            {
                comboBox1.Text = "";
                comboBox2.Text = "";
                comboBox3.Text = "";
            }
            else
            {
                var tmpary = CPname.Split('/');
                comboBox1.Text = tmpary[0].Split('-')[1];
                var otherary= tmpary[1].Split('-');
                comboBox2.Text = otherary[0];
                comboBox3.Text = otherary[1];

            }
           
                

        }

        private void CPNameEditor_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string newLbName ="SKIN-"+ comboBox1.Text + "/" + comboBox2.Text + "-" + comboBox3.Text;
            string newStr = lbinx+ newLbName + "\"";
            //替换文件并写入
          
             localMethod.UpdateConfigValue(oldstr, newStr, "CouponCfg.py");
            f1.UpdateLabelName(ind, newLbName);
            this.Close();

           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
