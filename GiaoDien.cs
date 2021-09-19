using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QLSV
{
    public partial class GiaoDien : Form
    {
        public int index = -1;
        public GiaoDien()
        {
            InitializeComponent();
        }
        List<SinhVien> listsv = new List<SinhVien>();
        List<SinhVien> listsvr = new List<SinhVien>();
        List<SinhVien> listsvd = new List<SinhVien>();
        BindingSource bds = new BindingSource();
        BindingSource bd = new BindingSource();
        public bool CheckControl()
        {
            if (string.IsNullOrWhiteSpace(txtTenSV.Text))
            {
                MessageBox.Show("Bạn chưa nhập tên sinh viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTenSV.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtLopSV.Text))
            {
                MessageBox.Show("Bạn chưa nhập lớp cho sinh viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtLopSV.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtDiaChiSV.Text))
            {
                MessageBox.Show("Bạn chưa nhập mã sinh viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtDiaChiSV.Focus();
                return false;
            }
            return true;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            if(CheckControl())
            {
                SinhVien sv = new SinhVien();
                sv.TenSV = txtTenSV.Text;
                sv.LopSV = txtLopSV.Text.ToUpper();
                sv.DiaChiSV = txtDiaChiSV.Text;
                sv.NSSV = dateTimePickerNSSV.Value.ToString("dd/MM/yyyy");
                listsv.Add(sv);
                bd.DataSource = listsv;
                dataGridViewSV.AutoGenerateColumns = false;
                dataGridViewSV.DataSource = bd;
                foreach (DataGridViewRow row in dataGridViewSV.Rows)
                    row.Cells["Column1"].Value = (row.Index + 1).ToString();
                bd.ResetBindings(false);
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (index >= 0)
            {
                listsv[index].TenSV = txtTenSV.Text;
                listsv[index].LopSV = txtLopSV.Text;
                listsv[index].DiaChiSV = txtDiaChiSV.Text;
                listsv[index].NSSV = dateTimePickerNSSV.Value.ToString("dd/MM/yyyy");
                bd.DataSource = listsv;
                dataGridViewSV.AutoGenerateColumns = false;
                dataGridViewSV.DataSource = listsv;
                bd.ResetBindings(false);
            }
        }

        private void dataGridViewSV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            listsvd.Clear();
            index = e.RowIndex;
            if (index >=0)
            {
                txtTenSV.Text = listsv[index].TenSV;
                txtLopSV.Text = listsv[index].LopSV;
                txtDiaChiSV.Text = listsv[index].DiaChiSV;
                txtdiemcc.Text = listsv[index].diemcc.ToString();
                txtdiemgk.Text = listsv[index].diemgk.ToString();
                txtdiemck.Text = listsv[index].diemck.ToString();
                dateTimePickerNSSV.Value = DateTime.ParseExact(listsv[index].NSSV, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                listsvd.Add(listsv[index]);
                bds.DataSource = listsvd;
                dataGridViewDiem.AutoGenerateColumns = false;
                dataGridViewDiem.DataSource = bds;
                bds.ResetBindings(false);
            }
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn xoá ?", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {
                listsv.RemoveAt(index);
                bd.DataSource = listsv;
                dataGridViewSV.AutoGenerateColumns = false;
                dataGridViewSV.DataSource = bd;
                bd.ResetBindings(false);
            }
        }

        private void btnThemDiem_Click(object sender, EventArgs e)
        {
                listsvd[0].diemcc = double.Parse(txtdiemcc.Text);
                listsvd[0].diemgk = double.Parse(txtdiemgk.Text);
                listsvd[0].diemck = double.Parse(txtdiemck.Text);
                listsv[index].diemcc = listsvd[0].diemcc;
                listsv[index].diemgk = listsvd[0].diemgk;
                listsv[index].diemck = listsvd[0].diemck;
                bds.DataSource = listsvd;
                dataGridViewDiem.AutoGenerateColumns = false;
                dataGridViewDiem.DataSource = bds;
                bds.ResetBindings(false);
        }

        private void txtdiemcc_Validating(object sender, CancelEventArgs e)
        {
            if(!Regex.Match(txtdiemcc.Text, @"^(([0-9]?((\.([0-9])*)?))|(10))$").Success)
            {
                e.Cancel = true;
                txtdiemcc.Focus();
                errorProvider1.SetError(txtdiemcc, "Điểm không hợp lệ");
            }
            else
            {
                e.Cancel = false;
                errorProvider1.SetError(txtdiemcc, null);
            }
        }

        private void txtdiemgk_Validating(object sender, CancelEventArgs e)
        {
            if (!Regex.Match(txtdiemgk.Text, @"^(([0-9]?((\.([0-9])*)?))|(10))$").Success)
            {
                e.Cancel = true;
                txtdiemcc.Focus();
                errorProvider1.SetError(txtdiemgk, "Điểm không hợp lệ");
            }
            else
            {
                e.Cancel = false;
                errorProvider1.SetError(txtdiemgk, null);
            }
        }

        private void txtdiemck_Validating(object sender, CancelEventArgs e)
        {
            if (!Regex.Match(txtdiemck.Text, @"^(([0-9]?((\.([0-9])*)?))|(10))$").Success)
            {
                e.Cancel = true;
                txtdiemcc.Focus();
                errorProvider1.SetError(txtdiemck, "Điểm không hợp lệ");
            }
            else
            {
                e.Cancel = false;
                errorProvider1.SetError(txtdiemck, null);
            }
        }

        private void btnBaoCao_Click(object sender, EventArgs e)
        {
            BaoCao bc = new BaoCao();
            if (index == -1)
            {
                MessageBox.Show("Chưa chọn học viên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                bc.listsv.Add(listsv[index]);
                bc.ShowDialog();
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            if (listsv.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "XLSX (*.xlsx)|*.xlsx";
                sfd.FileName = "DSHV";
                bool fileError = false;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("Lưu file không thành công" + ex.Message);
                        }
                    }
                        if (!fileError)
                        {
                            try
                            {
                                DataTable table = new DataTable();
                                table.Columns.Add("STT");
                                table.Columns.Add("Họ và tên");
                                table.Columns.Add("Lớp");
                                table.Columns.Add("Ngày sinh");
                                table.Columns.Add("Địa chỉ");
                                table.Columns.Add("Điểm chuyên cần");
                                table.Columns.Add("Điểm giữa kỳ");
                                table.Columns.Add("Điểm cuối kỳ");
                            int stt = 1;
                                foreach (var svien in listsv)
                                {
                                    table.Rows.Add(new Object[] {stt, svien.TenSV, svien.LopSV, svien.NSSV, svien.DiaChiSV, 
                                        svien.diemcc, svien.diemgk, svien.diemck });
                                stt++;
                                }
                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Visible = true;
                                excelApp.Workbooks.Add();
                                Excel._Worksheet worksheet = excelApp.ActiveSheet;
                                worksheet.Columns.AutoFit();
                                worksheet.Cells[1][1] = "STT";
                                worksheet.Cells[2][1] = "Họ tên";
                                worksheet.Cells[3][1] = "Lớp";
                                worksheet.Cells[4][1] = "Ngày sinh";
                                worksheet.Cells[5][1] = "Địa chỉ";
                                worksheet.Cells[6][1] = "Điểm";
                                worksheet.Cells[8][1] = "";
                                worksheet.Cells[9][1] = "";
                                worksheet.Range["F1: H1"].Merge();
                                worksheet.Cells[6][2] = "Điểm chuyên cần";
                                worksheet.Cells[7][2] = "Điểm giữa kỳ";
                                worksheet.Cells[8][2] = "Điểm cuối kỳ";
                                for (var i = 0; i < table.Rows.Count; i++)
                                {
                                    for (var j = 0; j < table.Columns.Count; j++)
                                    {
                                        worksheet.Cells[i + 3, j + 1] = table.Rows[i][j];
                                    }
                                }
                                worksheet.SaveAs(sfd.FileName);
                                excelApp.Quit();
                                MessageBox.Show("Lưu thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            this.Focus();
                            }
                            catch (Exception ex)
                            {
                            MessageBox.Show("Error: "+ ex.Message);
                            }
                        }
                    }
            }
            else
            {
                MessageBox.Show("Cần có ít nhất một sinh viên để lưu", "Thông báo",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
        }
        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            listsv.Clear();
            dataGridViewSV.Rows.Clear();
            dataGridViewDiem.Rows.Clear();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open("D:\\DSHV.xlsx");
            Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets["Sheet1"];
            for (int i = 1; i < xlWorkSheet.Rows.Count; i++)
            {
                string lop = Convert.ToString(xlWorkSheet.Cells[3][i + 2].Value);
                if (lop == e.Node.Text)
                {
                    listsv.Add(new SinhVien(Convert.ToString(xlWorkSheet.Cells[2][i + 2].Text),
                    Convert.ToString(xlWorkSheet.Cells[3][i + 2].Value), Convert.ToString(xlWorkSheet.Cells[4][i + 2].Value),
                    Convert.ToString(xlWorkSheet.Cells[5][i + 2].Value), Convert.ToDouble(xlWorkSheet.Cells[6][i + 2].Value),
                    Convert.ToDouble(xlWorkSheet.Cells[7][i + 2].Value), Convert.ToDouble(xlWorkSheet.Cells[8][i + 2].Value)));
                }
                if (xlWorkSheet.Cells[1][i + 2].Value == null)
                    break;
            }
            xlApp.Quit();
            BindingSource bdr = new BindingSource();
            bdr.DataSource = listsv;
            dataGridViewSV.AutoGenerateColumns = false;
            dataGridViewSV.DataSource = bdr;
            bdr.ResetBindings(false);
        }
    }
}
