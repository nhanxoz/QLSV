using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QLSV
{
    public partial class LoadingScreen : Form
    {
        public LoadingScreen()
        {
            InitializeComponent();
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            panel1.Width += 3;
            if(panel1.Width >=599)
            {
                timer1.Stop();
                this.Hide();
                var DangNhap = new DangNhap();
                DangNhap.Closed += (s, args) => this.Close();
                DangNhap.Show();
            }    
        }
    }
}
