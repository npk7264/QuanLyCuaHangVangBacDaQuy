using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using QuanLyCuaHangVangBacDaQuy.Class;
using COMExcel = Microsoft.Office.Interop.Excel;

namespace QuanLyCuaHangVangBacDaQuy
{
    public partial class frmBaoCaoTonKho : Form
    {
        DataTable tblBCTK;
        public frmBaoCaoTonKho()
        {
            InitializeComponent();
        }

        private void frmBaoCaoTonKho_Load(object sender, EventArgs e)
        {
            btnIn.Enabled = false;
            LoadDataGridView();
        }

        private void LoadDataGridView()
        {
            string sql;
            sql = "select y.MaSP, TenSP, TonDau, SoLuongMua, SoLuongBan, TonCuoi, TenDVT from CTBCTK x, SANPHAM y, LOAISANPHAM z, DONVITINH t" +
                " where Thang='" + cboThang.Text + "' and Nam='" + txtNam.Text.Trim() + "'" +
                " and x.MaSP=y.MaSP and y.MaLoaiSP=z.MaLoaiSP and z.MaDVT=t.MaDVT";
            tblBCTK = Functions.GetDataToTable(sql);
            dgvBaoCaoTonKho.DataSource = tblBCTK;
            dgvBaoCaoTonKho.Columns[0].HeaderText = "Mã sản phẩm";
            dgvBaoCaoTonKho.Columns[1].HeaderText = "Tên sản phẩm";
            dgvBaoCaoTonKho.Columns[2].HeaderText = "Tồn đầu";
            dgvBaoCaoTonKho.Columns[3].HeaderText = "Số lượng mua vào";
            dgvBaoCaoTonKho.Columns[4].HeaderText = "Số lượng bán ra";
            dgvBaoCaoTonKho.Columns[5].HeaderText = "Tồn cuối";
            dgvBaoCaoTonKho.Columns[6].HeaderText = "Đơn vị tính";

            dgvBaoCaoTonKho.Columns[0].Width = 100;
            dgvBaoCaoTonKho.Columns[1].Width = 200;
            dgvBaoCaoTonKho.Columns[2].Width = 100;
            dgvBaoCaoTonKho.Columns[3].Width = 100;
            dgvBaoCaoTonKho.Columns[4].Width = 100;
            dgvBaoCaoTonKho.Columns[5].Width = 100;
            dgvBaoCaoTonKho.Columns[6].Width = 100;
            dgvBaoCaoTonKho.AllowUserToAddRows = false;
            dgvBaoCaoTonKho.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            int thang = Convert.ToInt32(cboThang.Text);
            int nam = Convert.ToInt32(txtNam.Text);

            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook; //Trong 1 chương trình Excel có nhiều Workbook
            COMExcel.Worksheet exSheet; //Trong 1 Workbook có nhiều Worksheet
            COMExcel.Range exRange;
            string sql;
            int hang = 0, cot = 0;
            DataTable tblThongtinHang;
            exBook = exApp.Workbooks.Add(COMExcel.XlWBATemplate.xlWBATWorksheet);
            exSheet = exBook.Worksheets[1];
            // Định dạng chung
            exRange = exSheet.Cells[1, 1];
            exRange.Range["A1:Z300"].Font.Name = "Times new roman"; //Font chữ
            exRange.Range["A1:B3"].Font.Size = 10;
            exRange.Range["A1:B3"].Font.Bold = true;
            exRange.Range["A1:B3"].Font.ColorIndex = 5; //Màu xanh da trời
            exRange.Range["A1:A1"].ColumnWidth = 7;
            exRange.Range["B1:B1"].ColumnWidth = 15;
            exRange.Range["A1:B1"].MergeCells = true;
            exRange.Range["A1:B1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A1:B1"].Value = "Cửa hàng đá quý";
            exRange.Range["A2:B2"].MergeCells = true;
            exRange.Range["A2:B2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:B2"].Value = "UIT - TPHCM";
            exRange.Range["A3:B3"].MergeCells = true;
            exRange.Range["A3:B3"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A3:B3"].Value = "Điện thoại: (03)88888888";
            exRange.Range["C2:G2"].Font.Size = 16;
            exRange.Range["C2:G2"].Font.Bold = true;
            exRange.Range["C2:G2"].Font.ColorIndex = 3; //Màu đỏ
            exRange.Range["C2:G2"].MergeCells = true;
            exRange.Range["C2:G2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C2:G2"].Value = "BÁO CÁO TỒN KHO";
            // Biểu diễn thông tin chung của hóa đơn bán

            exRange.Range["B6:C9"].Font.Size = 12;
            exRange.Range["B6:B6"].Value = "Tháng:";
            exRange.Range["C6:C6"].Value = thang.ToString();
            exRange.Range["C6:C6"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["B7:B7"].Value = "Năm:";
            exRange.Range["C7:C7"].Value = nam.ToString();
            exRange.Range["C7:C7"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;


            //Lấy thông tin các mặt hàng
            sql = "select y.MaSP, TenSP, TonDau, SoLuongMua, SoLuongBan, TonCuoi, TenDVT from CTBCTK x, SANPHAM y, LOAISANPHAM z, DONVITINH t" +
                " where Thang='" + cboThang.Text + "' and Nam='" + txtNam.Text.Trim() + "'" +
                " and x.MaSP=y.MaSP and y.MaLoaiSP=z.MaLoaiSP and z.MaDVT=t.MaDVT";
            tblThongtinHang = Functions.GetDataToTable(sql);
            //Tạo dòng tiêu đề bảng
            exRange.Range["A11:H11"].Font.Bold = true;
            exRange.Range["A11:H11"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["B11:H11"].ColumnWidth = 20;
            exRange.Range["A11:A11"].Value = "STT";
            exRange.Range["B11:B11"].Value = "Mã sản phẩm";
            exRange.Range["C11:C11"].Value = "Tên sản phẩm";
            exRange.Range["D11:D11"].Value = "Tồn đầu";
            exRange.Range["E11:E11"].Value = "Số lượng mua vào";
            exRange.Range["F11:F11"].Value = "Số lượng bán ra";
            exRange.Range["G11:G11"].Value = "Tồn cuối";
            exRange.Range["H11:H11"].Value = "Đơn vị tính";

            for (hang = 0; hang < tblThongtinHang.Rows.Count; hang++)
            {
                //Điền số thứ tự vào cột 1 từ dòng 12
                exSheet.Cells[1][hang + 12] = hang + 1;
                exRange = exSheet.Cells[1][hang + 12];
                exRange.HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;

                for (cot = 0; cot < tblThongtinHang.Columns.Count; cot++)
                //Điền thông tin hàng từ cột thứ 2, dòng 12
                {
                    exSheet.Cells[cot + 2][hang + 12] = tblThongtinHang.Rows[hang][cot].ToString();
                    exRange = exSheet.Cells[cot + 2][hang + 12];
                    exRange.HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
                }
            }
            exRange = exSheet.Cells[cot][hang + 14];
            exRange.Font.Bold = true;
            exRange = exSheet.Cells[cot + 1][hang + 14];
            exRange.Font.Bold = true;
            /*exRange.Value2 = tblThongtinHD.Rows[0][2].ToString();*/
            exRange = exSheet.Cells[1][hang + 15]; //Ô A1 
            exRange.Range["A1:F1"].MergeCells = true;
            exRange.Range["A1:F1"].Font.Bold = true;
            exRange.Range["A1:F1"].Font.Italic = true;
            exRange.Range["A1:F1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignRight;
            exRange = exSheet.Cells[4][hang + 17]; //Ô A1 
            exRange.Range["A1:C1"].MergeCells = true;
            exRange.Range["A1:C1"].Font.Italic = true;
            exRange.Range["A1:C1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            /*DateTime d = Convert.ToDateTime(tblThongtinHD.Rows[0][1]);*/
            exRange.Range["A1:C1"].Value = "TP HCM, ngày " + DateTime.Now.Day + " tháng " + DateTime.Now.Month + " năm " + DateTime.Now.Year;
            exRange.Range["A2:C2"].MergeCells = true;
            exRange.Range["A2:C2"].Font.Italic = true;
            exRange.Range["A2:C2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A6:C6"].MergeCells = true;
            exRange.Range["A6:C6"].Font.Italic = true;
            exRange.Range["A6:C6"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            /*exRange.Range["A6:C6"].Value = tblThongtinHD.Rows[0][6];*/
            exSheet.Name = "Hóa đơn nhập";
            exApp.Visible = true;
        }

        private void btnXuat_Click(object sender, EventArgs e)
        {
            string sql;
            if (cboThang.Text == "")
            {
                MessageBox.Show("Bạn phải nhập tháng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboThang.Focus();
                return;
            }
            if (txtNam.Text == "")
            {
                MessageBox.Show("Bạn phải nhập năm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtNam.Focus();
                return;
            }

            // Kiểm tra thời gian cần lập báo cáo tồn kho có hợp lệ không
            // Chỉ được lập báo cáo tồn kho cho một tháng khi hết tháng
            int thang = Convert.ToInt32(cboThang.Text);
            int nam = Convert.ToInt32(txtNam.Text);
            int thang_now = Convert.ToInt32(DateTime.Now.Month.ToString());
            int nam_now = Convert.ToInt32(DateTime.Now.Year.ToString());
            if ((thang >= thang_now && nam == nam_now) || (nam > nam_now))
            {
                MessageBox.Show("Thời gian không hợp lệ để lập báo cáo", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboThang.Focus();
                txtNam.Focus();
                return;
            }

            // Kiểm tra đã tháng, năm nhập vào đã tồn tại báo cáo tồn kho chưa
            int count_bctk = Convert.ToInt32(Functions.GetFieldValues("select COUNT(MaSP) from CTBCTK where Thang='" + cboThang.Text + "' and Nam='" + txtNam.Text + "'"));

            // Nếu chưa, thêm báo cáo tồn kho vào bảng CTBCTK
            if (count_bctk == 0)
            {
                sql = "SELECT MaSP FROM SANPHAM";
                DataTable tblSP = Functions.GetDataToTable(sql);
                for (int hang = 0; hang <= tblSP.Rows.Count - 1; hang++)
                {
                    int count;
                    // tính tổng sl mua trước tháng cần lập báo cáo
                    int slmua;
                    count = Convert.ToInt32(Functions.GetFieldValues("select COUNT(MaSP) from CTPMH x, PHIEUMUAHANG y where MaSP='" + tblSP.Rows[hang][0].ToString() +
                     "' and x.SoPhieu=y.SoPhieu and ((month(NgayLap)<'" + cboThang.Text + "' and year(NgayLap)='" + txtNam.Text + "')" +
                     "or year(NgayLap)<'" + txtNam.Text + "')"));
                    if (count > 0)
                        slmua = Convert.ToInt32(Functions.GetFieldValues("select SUM(SoLuong) from CTPMH x, PHIEUMUAHANG y where MaSP='" + tblSP.Rows[hang][0].ToString() +
                     "' and x.SoPhieu=y.SoPhieu and ((month(NgayLap)<'" + cboThang.Text + "' and year(NgayLap)='" + txtNam.Text + "')" +
                     "or year(NgayLap)<'" + txtNam.Text + "')"));
                    else
                        slmua = 0;
                    // tính tổng sl bán trước tháng cần lập báo cáo
                    int slban;
                    count = Convert.ToInt32(Functions.GetFieldValues("select COUNT(MaSP) from CTPBH x, PHIEUBANHANG y where MaSP='" + tblSP.Rows[hang][0].ToString() +
                     "' and x.SoPhieu=y.SoPhieu and ((month(NgayLap)<'" + cboThang.Text + "' and year(NgayLap)='" + txtNam.Text + "')" +
                     "or year(NgayLap)<'" + txtNam.Text + "')"));
                    if (count > 0)
                        slban = Convert.ToInt32(Functions.GetFieldValues("select SUM(SoLuong) from CTPBH x, PHIEUBANHANG y where MaSP='" + tblSP.Rows[hang][0].ToString() +
                     "' and x.SoPhieu=y.SoPhieu and ((month(NgayLap)<'" + cboThang.Text + "' and year(NgayLap)='" + txtNam.Text + "')" +
                     "or year(NgayLap)<'" + txtNam.Text + "')"));
                    else
                        slban = 0;
                    // tính sl tồn đầu tháng
                    int tondau = slmua - slban;

                    // tính sl mua trong tháng
                    int mua;
                    count = Convert.ToInt32(Functions.GetFieldValues("select COUNT(MaSP) from CTPMH x, PHIEUMUAHANG y where MaSP='" + tblSP.Rows[hang][0].ToString() +
                     "' and x.SoPhieu=y.SoPhieu and (month(NgayLap)='" + cboThang.Text + "' and year(NgayLap)='" + txtNam.Text + "')"));
                    if (count > 0)
                        mua = Convert.ToInt32(Functions.GetFieldValues("select SUM(SoLuong) from CTPMH x, PHIEUMUAHANG y where MaSP='" + tblSP.Rows[hang][0].ToString() +
                         "' and x.SoPhieu=y.SoPhieu and (month(NgayLap)='" + cboThang.Text + "' and year(NgayLap)='" + txtNam.Text + "')"));
                    else
                        mua = 0;

                    // tính sl bán trong tháng
                    int ban;
                    count = Convert.ToInt32(Functions.GetFieldValues("select COUNT(MaSP) from CTPBH x, PHIEUBANHANG y where MaSP='" + tblSP.Rows[hang][0].ToString() +
                     "' and x.SoPhieu=y.SoPhieu and (month(NgayLap)='" + cboThang.Text + "' and year(NgayLap)='" + txtNam.Text + "')"));
                    if (count > 0)
                        ban = Convert.ToInt32(Functions.GetFieldValues("select SUM(SoLuong) from CTPBH x, PHIEUBANHANG y where MaSP='" + tblSP.Rows[hang][0].ToString() +
                         "' and x.SoPhieu=y.SoPhieu and (month(NgayLap)='" + cboThang.Text + "' and year(NgayLap)='" + txtNam.Text + "')"));
                    else
                        ban = 0;

                    // tính sl tồn cuối tháng
                    int toncuoi = tondau + mua - ban;

                    sql = "insert into CTBCTK values('" + cboThang.Text + "', '" + txtNam.Text + "','" + tblSP.Rows[hang][0] + "','" + tondau + "','" + mua + "', '" + ban + "','" + toncuoi + "')";
                    Functions.RunSQL(sql);
                }
            }
            LoadDataGridView();
            btnIn.Enabled = true;
        }

        private void txtNam_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else
                e.Handled = true;
        }
    }
}
