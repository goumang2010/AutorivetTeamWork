using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GoumangToolKit;


namespace AUTORIVET_KAOHE
{
    public partial class Task_model : Form
    {
        public Task_model()
        {
            InitializeComponent();
        }

        private void Task_model_Load(object sender, EventArgs e)
        {
            dataviewfresh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DbHelperSQL.ExecuteSql("insert into 任务模板(责任人,任务类型,任务名称,任务说明,绩效奖励) value('" + comboBox1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "','" + textBox2.Text + "'," + System.Convert.ToDouble(textBox3.Text) + ")");

            dataviewfresh();
        }
        private void dataviewfresh()
        {
            dataGridView1.DataSource = DbHelperSQL.Query("select 任务名称,责任人,任务类型,任务说明,绩效奖励 from 任务模板").Tables[0];
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
