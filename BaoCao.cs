using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace QLSV
{
    public partial class BaoCao : Form
    {
        public List<SinhVien> listsv = new List<SinhVien>();
        public BindingSource bds = new BindingSource();
        public BaoCao()
        {
            InitializeComponent();
        }
        private void BaoCao_Load(object sender, EventArgs e)
        {
            System.Globalization.CultureInfo vi = new System.Globalization.CultureInfo("vi-VN");
            System.Threading.Thread.CurrentThread.CurrentCulture = vi;
            string day = DateTime.Now.Day.ToString();
            string month = DateTime.Now.Month.ToString();
            string year = DateTime.Now.Year.ToString();
            day.ToString(CultureInfo.GetCultureInfo("vi-VN"));
            ReportParameterCollection rpc = new ReportParameterCollection();
            rpc.Add(new ReportParameter("StudentName", listsv[0].TenSV));
            rpc.Add(new ReportParameter("CreateDay", day));
            rpc.Add(new ReportParameter("CreateMonth", month));
            rpc.Add(new ReportParameter("CreateYear", year));
            rpc.Add(new ReportParameter("Grade1", listsv[0].diemcc.ToString()));
            rpc.Add(new ReportParameter("Grade2", listsv[0].diemgk.ToString()));
            rpc.Add(new ReportParameter("Grade3", listsv[0].diemck.ToString()));
            this.reportViewer1.LocalReport.SetParameters(rpc);
            this.reportViewer1.RefreshReport();
        }
    }
}
