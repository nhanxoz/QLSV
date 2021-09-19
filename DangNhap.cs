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
    public partial class DangNhap : Form
    {
        public static bool isLoggedIn { get; set; }
        public DangNhap()
        {
            InitializeComponent();
        }
        private void DangNhap_Load(object sender, EventArgs e)
        {
            lblError.Visible = false;
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if(txtUserName.Text.Trim().Equals("admin") && txtPassword.Text.Trim().Equals("admin"))
            {
                this.Hide();
                var GiaoDien = new GiaoDien();
                GiaoDien.Closed += (s, args) => this.Close();
                GiaoDien.Show();
                isLoggedIn = true;
            }
            else
            {
                var t = new Timer();
                t.Interval = 1000; // it will Tick in 3 seconds
                t.Tick += (s, args) =>
                {
                    lblError.Hide();
                    t.Stop();
                };
                lblError.Visible = true;
                t.Start();
            }
        }
    }
}
