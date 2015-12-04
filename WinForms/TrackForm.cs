using mysqlsolution;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace AUTORIVET_KAOHE
{
    public partial class Form3 : Form
    {
        const int WM_NCHITTEST = 0x0084;
        const int HTCLIENT = 0x0001;
        const int HTCAPTION = 0x0002;
       
        
                  //  [ DllImport( "user32.dll",  CharSet = CharSet.Auto ) ]
//public static extern bool SwitchToThisWindow( IntPtr hWnd,  bool fAltTab );

         
        
        
        public Form3()
        {
            InitializeComponent();
           //  SwitchToThisWindow( this.Handle, true );
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            this.Top = 200;
            this.Left = Screen.PrimaryScreen.Bounds.Width -500;
           // this.Width = 200;
           // this.Height = 200;

            rfgrid();

           // Form_method.ConnectDef();
        }

        public void rfgrid()
        {
            dataGridView1.DataSource = DbHelperSQL.Query("select 责任人,任务名称 from 任务管理 where 任务状态='重要'  order by 责任人,节点日期").Tables[0];
           
            //.Columns["任务说明"].Width = 170;
            dataGridView1.Columns["责任人"].Width = 55;
           // dataGridView1.Columns["节点日期"].Width = 55;
          //  dataGridView1.Columns["流水号"].Width = 55;
          
        }
        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            switch (m.Msg)
            {
                case WM_NCHITTEST:
                    base.WndProc(ref m);
                    if (m.Result == (IntPtr)HTCLIENT)
                        m.Result = (IntPtr)HTCAPTION;
                    break;
                default:
                    base.WndProc(ref m);
                    break;
            }
        }
    }
}
