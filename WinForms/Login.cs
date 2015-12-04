using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AUTORIVET_KAOHE
{
    public partial class user_key : Form
    {


        MainTeamForm ftemp;
        public user_key( MainTeamForm fcd)
        {
            InitializeComponent();
            ftemp = fcd;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            gogogo();
        }
        private void gogogo()
    {
            if (textBox1.Text == "son" || (textBox1.Text == "saci_share" && textBox2.Text == "sacia03client"))
            {
                MessageBox.Show("恭喜你已打开新世界的大门！");
                Program.ManagerActived = true;
                ftemp.ActiveManager();
                ftemp.Show();
                this.Hide();
            }
    }
        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void user_key_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyValue == 13)
            {
                gogogo();
            }
        }

        private void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                gogogo();
            }
        }
    }
}
