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
    public partial class frmPhieuMuaHang : Form
    {
        DataTable tblCTPMH;
        public frmPhieuMuaHang()
        {
            InitializeComponent();
        }

        private void frmPhieuMuaHang_Load(object sender, EventArgs e)
        {
            btnThem.Enabled = true;
            btnLuu.Enabled = false;
            btnHuy.Enabled = false;
            btnIn.Enabled = false;
            txtTenNCC.ReadOnly = true;
            txtDiaChi.ReadOnly = true;
            txtSDT.ReadOnly = true;
            txtTenNV.ReadOnly = true;
            txtTenSP.ReadOnly = true;
            txtLoaiSP.ReadOnly = true;
            txtDVT.ReadOnly = true;
            txtThanhTien.ReadOnly = true;
            txtTongTien.ReadOnly = true;
            txtTongTien.Text = "";
            Functions.FillCombo("SELECT MaNCC, TenNCC FROM NHACUNGCAP", cboMaNCC, "MaNCC", "MaNCC");
            cboMaNCC.SelectedIndex = -1;
            Functions.FillCombo("select MaNV, TenNV from NHANVIEN", cboMaNV, "MaNV", "MaNV");
            cboMaNV.SelectedIndex = -1;
            Functions.FillCombo("SELECT MaSP, TenSP FROM SANPHAM", cboMaSP, "MaSP", "MaSP");
            cboMaSP.SelectedIndex = -1;
            //Hiển thị thông tin của một hóa đơn được gọi từ form tìm kiếm
            if (txtSoPhieu.Text != "")
            {
                LoadInfoHoaDon();
                btnHuy.Enabled = true;
                btnIn.Enabled = true;
            }
            LoadDataGridView();
        }

        private void LoadInfoHoaDon()
        {
            string str;
            str = "select NgayLap from PHIEUMUAHANG WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            dtpNgayLap.Value = DateTime.Parse(Functions.GetFieldValues(str));
            str = "SELECT MaNCC FROM PHIEUMUAHANG WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            cboMaNCC.SelectedValue = Functions.GetFieldValues(str);
            str = "SELECT MaNV FROM PHIEUMUAHANG WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            cboMaNV.SelectedValue = Functions.GetFieldValues(str);
            str = "SELECT TongTien FROM PHIEUMUAHANG WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            txtTongTien.Text = Functions.GetFieldValues(str);
        }

        private void LoadDataGridView()
        {
            string sql;
            sql = "SELECT x.MaSP, TenSP, TenLoaiSP, x.SoLuong, TenDVT, x.DonGiaMua, ThanhTien FROM CTPMH x, SANPHAM y, LOAISANPHAM z, DONVITINH t where x.SoPhieu = N'" + txtSoPhieu.Text + "'" + " and x.MaSP=y.MaSP and y.MaLoaiSP=z.MaLoaiSP and z.MaDVT=t.MaDVT";
            tblCTPMH = Functions.GetDataToTable(sql);
            dgvPhieuMuaHang.DataSource = tblCTPMH;
            dgvPhieuMuaHang.Columns[0].HeaderText = "Mã sản phẩm";
            dgvPhieuMuaHang.Columns[1].HeaderText = "Tên sản phẩm";
            dgvPhieuMuaHang.Columns[2].HeaderText = "Loại sản phẩm";
            dgvPhieuMuaHang.Columns[3].HeaderText = "Số lượng";
            dgvPhieuMuaHang.Columns[4].HeaderText = "Đơn vị tính";
            dgvPhieuMuaHang.Columns[5].HeaderText = "Đơn giá mua";
            dgvPhieuMuaHang.Columns[6].HeaderText = "Thành tiền";

            dgvPhieuMuaHang.Columns[0].Width = 100;
            dgvPhieuMuaHang.Columns[1].Width = 200;
            dgvPhieuMuaHang.Columns[2].Width = 200;
            dgvPhieuMuaHang.Columns[3].Width = 90;
            dgvPhieuMuaHang.Columns[4].Width = 100;
            dgvPhieuMuaHang.Columns[5].Width = 150;
            dgvPhieuMuaHang.Columns[6].Width = 150;
            dgvPhieuMuaHang.AllowUserToAddRows = false;
            dgvPhieuMuaHang.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void btnThem_Click_1(object sender, EventArgs e)
        {
            btnHuy.Enabled = true;
            btnLuu.Enabled = true;
            btnIn.Enabled = false;
            btnThem.Enabled = false;
            ResetValues();
            txtSoPhieu.Text = Functions.CreateKey("PMH");
            txtSoPhieu.Enabled = true;
            LoadDataGridView();
        }

        private void ResetValues()
        {
            txtSoPhieu.Text = "";
            dtpNgayLap.Value = DateTime.Now;
            cboMaNCC.SelectedIndex = -1;
            cboMaNV.SelectedIndex = -1;
            //cboMaNCC.Text = "";
            txtTongTien.Text = "0";
            cboMaSP.SelectedIndex = -1;
            //cboMaSP.Text = "";
            txtSoLuong.Text = "";
            txtDonGiaMua.Text = "";
            txtThanhTien.Text = "";
        }

        private void ResetValuesHang()
        {
            cboMaSP.SelectedIndex = -1;
            txtSoLuong.Text = "";
            txtDonGiaMua.Text = "";
            txtThanhTien.Text = "";
        }
        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cboMaNCC_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaNCC.Text=="")
            {
                txtTenNCC.Text = "";
                txtDiaChi.Text = "";
                txtSDT.Text = "";
            }
            // Khi chọn mã nhà cung cấp thì tên nhà cung cấp tự động hiện ra
            str = "select TenNCC from NHACUNGCAP where MaNCC =N'" + cboMaNCC.SelectedValue + "'";
            txtTenNCC.Text = Functions.GetFieldValues(str);
            str = "select DiaChi from NHACUNGCAP where MaNCC =N'" + cboMaNCC.SelectedValue + "'";
            txtDiaChi.Text = Functions.GetFieldValues(str);
            str = "select SDT from NHACUNGCAP where MaNCC =N'" + cboMaNCC.SelectedValue + "'";
            txtSDT.Text = Functions.GetFieldValues(str);
        }

        private void cboMaNV_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaNV.Text == "")
                txtTenNV.Text = "";
            str = "select TenNV from NHANVIEN where MaNV=N'" + cboMaNV.SelectedValue + "'";
            txtTenNV.Text = Functions.GetFieldValues(str);
        }

        private void cboMaSP_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaSP.Text=="")
            {
                txtTenSP.Text = "";
                txtLoaiSP.Text = "";
                txtDVT.Text = "";
            }
            // Khi chọn mã sản phẩm thì các thông tin về sản phẩm tự động hiện ra
            str = "select TenSP FROM SANPHAM WHERE MaSP =N'" + cboMaSP.SelectedValue + "'";
            txtTenSP.Text = Functions.GetFieldValues(str);
            str = "select TenLoaiSP FROM SANPHAM a, LOAISANPHAM b WHERE MaSP =N'" + cboMaSP.SelectedValue + "' and a.MaLoaiSP=b.MaLoaiSP";
            txtLoaiSP.Text = Functions.GetFieldValues(str);
            str = "select TenDVT FROM SANPHAM a, LOAISANPHAM b, DONVITINH c WHERE MaSP =N'" + cboMaSP.SelectedValue + "' and a.MaLoaiSP=b.MaLoaiSP and b.MaDVT=c.MaDVT";
            txtDVT.Text = Functions.GetFieldValues(str);
        }

        private void btnLuu_Click_1(object sender, EventArgs e)
        {
            string sql;
            double sl, tong, Tongmoi;
            sql = "select SoPhieu FROM PHIEUMUAHANG WHERE SoPhieu=N'" + txtSoPhieu.Text + "'";
            if (!Functions.CheckKey(sql))
            {
                if (cboMaNCC.Text.Length == 0)
                {
                    MessageBox.Show("Bạn phải nhập nhà cung cấp", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cboMaNCC.Focus();
                    return;
                }
                if (cboMaNV.Text.Length == 0)
                {
                    MessageBox.Show("Bạn phải nhập nhân viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cboMaNV.Focus();
                    return;
                }
                sql = "insert into PHIEUMUAHANG values (N'" + txtSoPhieu.Text.Trim() + "','" +
                        dtpNgayLap.Value + "', '" +
                        cboMaNCC.SelectedValue + "', '" +
                        cboMaNV.SelectedValue+"', "+
                        txtTongTien.Text + ")";
                Functions.RunSQL(sql);
            }
            // Lưu thông tin của các mặt hàng
            if (cboMaSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboMaSP.Focus();
                return;
            }
            if ((txtSoLuong.Text.Trim().Length == 0) || (txtSoLuong.Text == "0"))
            {
                MessageBox.Show("Bạn phải nhập số lượng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtSoLuong.Text = "";
                txtSoLuong.Focus();
                return;
            }
            if (txtDonGiaMua.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập đơn giá mua", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtDonGiaMua.Focus();
                return;
            }
            txtThanhTien.Text = (Convert.ToDouble(txtSoLuong.Text) * Convert.ToDouble(txtDonGiaMua.Text)).ToString();
            sql = "select MaSP FROM CTPMH WHERE MaSP=N'" + cboMaSP.SelectedValue + "' and SoPhieu = N'" + txtSoPhieu.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã sản phẩm này đã có, bạn phải nhập mã khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetValuesHang();
                cboMaSP.Focus();
                return;
            }
            sql = "insert into CTPMH values(N'" + txtSoPhieu.Text.Trim() + "',N'" + cboMaSP.SelectedValue + "'," + txtSoLuong.Text + "," + txtDonGiaMua.Text + "," + txtThanhTien.Text + ")";
            Functions.RunSQL(sql);
            LoadDataGridView();
            // Cập nhật lại số lượng của mặt hàng vào bảng SANPHAM
            sl = Convert.ToDouble(txtSoLuong.Text);
            sql = "update SANPHAM set SoLuong =SoLuong+" + sl + " where MaSP= N'" + cboMaSP.SelectedValue + "'";
            Functions.RunSQL(sql);
            // Cập nhật lại đơn giá mua của sản phẩm trong bảng SANPHAM
            sql = "update SANPHAM set DonGiaMua =" + txtDonGiaMua.Text + " where MaSP= N'" + cboMaSP.SelectedValue + "'";
            Functions.RunSQL(sql);
            // Cập nhật lại đơn giá bán của sản phẩm trong bảng SANPHAM
            double loinhuan = Convert.ToDouble(Functions.GetFieldValues("select PhanTramLoiNhuan from SANPHAM x, LOAISANPHAM y where MaSP=N'"+cboMaSP.SelectedValue+"' and x.MaLoaiSP=y.MaLoaiSP"));
            double giaban = Convert.ToDouble(txtDonGiaMua.Text) * (1 + loinhuan / 100);
            sql = "update SANPHAM set DonGiaBan =" + giaban + " where MaSP= N'" + cboMaSP.SelectedValue + "'";
            Functions.RunSQL(sql);
            // Cập nhật lại tổng tiền cho hóa đơn bán
            tong = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTien FROM PHIEUMUAHANG WHERE SoPhieu = N'" + txtSoPhieu.Text + "'"));
            Tongmoi = tong + Convert.ToDouble(txtThanhTien.Text);
            sql = "UPDATE PHIEUMUAHANG SET TongTien =" + Tongmoi + " WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            Functions.RunSQL(sql);
            txtTongTien.Text = Tongmoi.ToString();
            ResetValuesHang();
            btnHuy.Enabled = true;
            btnThem.Enabled = true;
            btnIn.Enabled = true;
        }

        private void txtSoLuong_TextChanged(object sender, EventArgs e)
        {
            //Khi thay đổi số lượng thì thực hiện tính lại thành tiền
            double tt, sl, dg;
            if (txtSoLuong.Text == "")
                sl = 0;
            else
                sl = Convert.ToDouble(txtSoLuong.Text);
            if (txtDonGiaMua.Text == "")
                dg = 0;
            else
                dg = Convert.ToDouble(txtDonGiaMua.Text);
            tt = sl * dg;
            txtThanhTien.Text = tt.ToString();
        }

        private void txtDonGiaMua_TextChanged(object sender, EventArgs e)
        {
            //Khi thay đổi giảm giá thì tính lại thành tiền
            double tt, sl, dg;
            if (txtSoLuong.Text == "")
                sl = 0;
            else
                sl = Convert.ToDouble(txtSoLuong.Text);
            if (txtDonGiaMua.Text == "")
                dg = 0;
            else
                dg = Convert.ToDouble(txtDonGiaMua.Text);
            tt = sl * dg;
            txtThanhTien.Text = tt.ToString();
        }


        private void btnTimKiem_Click_1(object sender, EventArgs e)
        {
            if (cboSoPhieu.Text == "")
            {
                MessageBox.Show("Bạn phải chọn một mã hóa đơn để tìm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboSoPhieu.Focus();
                return;
            }
            txtSoPhieu.Text = cboSoPhieu.Text;
            LoadInfoHoaDon();
            LoadDataGridView();
            btnHuy.Enabled = true;
            btnLuu.Enabled = true;
            btnIn.Enabled = true;
            cboSoPhieu.SelectedIndex = -1;
        }


        private void cboSoPhieu_DropDown(object sender, EventArgs e)
        {
            Functions.FillCombo("SELECT SoPhieu FROM PHIEUMUAHANG", cboSoPhieu, "SoPhieu", "SoPhieu");
            cboSoPhieu.SelectedIndex = -1;
        }

        private void dgvPhieuMuaHang_DoubleClick(object sender, EventArgs e)
        {
            string MaSPxoa, sql;
            Double ThanhTienxoa, SoLuongxoa, sl, slcon, tong, tongmoi;
            if (tblCTPMH.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if ((MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                //Xóa hàng và cập nhật lại số lượng hàng 
                MaSPxoa = dgvPhieuMuaHang.CurrentRow.Cells["MaSP"].Value.ToString();
                SoLuongxoa = Convert.ToDouble(dgvPhieuMuaHang.CurrentRow.Cells["SoLuong"].Value.ToString());
                ThanhTienxoa = Convert.ToDouble(dgvPhieuMuaHang.CurrentRow.Cells["ThanhTien"].Value.ToString());
                sql = "DELETE from CTPMH WHERE SoPhieu=N'" + txtSoPhieu.Text + "' AND MaSP = N'" + MaSPxoa + "'";
                Functions.RunSQL(sql);
                // Cập nhật lại số lượng cho các mặt hàng
                sl = Convert.ToDouble(Functions.GetFieldValues("SELECT SoLuong FROM SANPHAM WHERE MaSP = N'" + MaSPxoa + "'"));
                slcon = sl - SoLuongxoa;
                sql = "UPDATE SANPHAM SET SoLuong =" + slcon + " WHERE MaSP= N'" + MaSPxoa + "'";
                Functions.RunSQL(sql);
                // Cập nhật lại tổng tiền cho hóa đơn bán
                tong = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTien FROM PHIEUMUAHANG WHERE SoPhieu = N'" + txtSoPhieu.Text + "'"));
                tongmoi = tong - ThanhTienxoa;
                sql = "UPDATE PHIEUMUAHANG SET TongTien =" + tongmoi + " WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
                Functions.RunSQL(sql);
                txtTongTien.Text = tongmoi.ToString();
                LoadDataGridView();
            }
        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            double sl, slcon, slxoa;
            if (MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string sql = "SELECT MaSP,SoLuong FROM CTPMH WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
                DataTable SANPHAM = Functions.GetDataToTable(sql);
                for (int hang = 0; hang <= SANPHAM.Rows.Count - 1; hang++)
                {
                    // Cập nhật lại số lượng cho các sản phẩm
                    sl = Convert.ToDouble(Functions.GetFieldValues("SELECT SoLuong FROM SANPHAM WHERE MaSP = N'" + SANPHAM.Rows[hang][0].ToString() + "'"));
                    slxoa = Convert.ToDouble(SANPHAM.Rows[hang][1].ToString());
                    slcon = sl - slxoa;
                    sql = "UPDATE SANPHAM SET SoLuong =" + slcon + " WHERE MaSP= N'" + SANPHAM.Rows[hang][0].ToString() + "'";
                    Functions.RunSQL(sql);
                }

                //Xóa chi tiết hóa đơn
                sql = "DELETE from CTPMH WHERE SoPhieu=N'" + txtSoPhieu.Text + "'";
                Functions.RunSqlDel(sql);

                //Xóa hóa đơn
                sql = "DELETE from PHIEUMUAHANG WHERE SoPhieu=N'" + txtSoPhieu.Text + "'";
                Functions.RunSqlDel(sql);
                ResetValues();
                LoadDataGridView();
                btnHuy.Enabled = false;
                btnIn.Enabled = false;
                btnThem.Enabled = true;
            }
        }

        private void frmPhieuMuaHang_FormClosing(object sender, FormClosingEventArgs e)
        {
            ResetValues();
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook; //Trong 1 chương trình Excel có nhiều Workbook
            COMExcel.Worksheet exSheet; //Trong 1 Workbook có nhiều Worksheet
            COMExcel.Range exRange;
            string sql;
            int hang = 0, cot = 0;
            DataTable tblThongtinHD, tblThongtinHang;
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
            exRange.Range["C2:F2"].Font.Size = 16;
            exRange.Range["C2:F2"].Font.Bold = true;
            exRange.Range["C2:F2"].Font.ColorIndex = 3; //Màu đỏ
            exRange.Range["C2:F2"].MergeCells = true;
            exRange.Range["C2:F2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C2:F2"].Value = "HÓA ĐƠN MUA";
            // Biểu diễn thông tin chung của hóa đơn bán
            sql = "SELECT a.SoPhieu, a.NgayLap, a.TongTien, b.TenNCC, b.DiaChi, b.SDT, c.TenNV FROM PHIEUMUAHANG AS a, NHACUNGCAP AS b, NHANVIEN AS c WHERE a.SoPhieu = N'" + txtSoPhieu.Text + "' AND a.MaNCC = b.MaNCC AND a.MaNV = c.MaNV";
            tblThongtinHD = Functions.GetDataToTable(sql);
            exRange.Range["B6:C9"].Font.Size = 12;
            exRange.Range["B6:B6"].Value = "Mã hóa đơn:";
            exRange.Range["C6:E6"].MergeCells = true;
            exRange.Range["C6:E6"].Value = tblThongtinHD.Rows[0][0].ToString();
            exRange.Range["B7:B7"].Value = "Khách hàng:";
            exRange.Range["C7:E7"].MergeCells = true;
            exRange.Range["C7:E7"].Value = tblThongtinHD.Rows[0][3].ToString();
            exRange.Range["B8:B8"].Value = "Địa chỉ:";
            exRange.Range["C8:E8"].MergeCells = true;
            exRange.Range["C8:E8"].Value = tblThongtinHD.Rows[0][4].ToString();/////////////////////////////
            exRange.Range["B9:B9"].Value = "Số điện thoại:";
            exRange.Range["C9:E9"].MergeCells = true;
            exRange.Range["C9:E9"].Value = "'" + tblThongtinHD.Rows[0][5].ToString();
            //Lấy thông tin các mặt hàng
            sql = "SELECT b.TenSP, c.TenLoaiSP, TenDVT, a.SoLuong, a.DonGiaMua, a.ThanhTien " +
                  "FROM CTPMH AS a , SANPHAM AS b, LOAISANPHAM AS c, DONVITINH AS d WHERE a.SoPhieu = N'" +
                  txtSoPhieu.Text + "' AND a.MaSP = b.MaSP AND b.MaLoaiSP = c.MaLoaiSP AND c.MaDVT = d.MaDVT";
            tblThongtinHang = Functions.GetDataToTable(sql);
            //Tạo dòng tiêu đề bảng
            exRange.Range["A11:G11"].Font.Bold = true;
            exRange.Range["A11:G11"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A12:G12"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["B11:G11"].ColumnWidth = 20;
            exRange.Range["A11:A11"].Value = "STT";
            exRange.Range["B11:B11"].Value = "Tên sản phẩm";
            exRange.Range["C11:C11"].Value = "Loại sản phẩm";
            exRange.Range["D11:D11"].Value = "Đơn vị tính";
            exRange.Range["E11:E11"].Value = "Số lượng";
            exRange.Range["F11:F11"].Value = "Đơn giá mua";
            exRange.Range["G11:G11"].Value = "Thành tiền";

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
            exRange.HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Value2 = "Tổng tiền:";
            exRange = exSheet.Cells[cot + 1][hang + 14];
            exRange.Font.Bold = true;
            exRange.HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Value2 = tblThongtinHD.Rows[0][2].ToString();
            exRange = exSheet.Cells[1][hang + 15]; //Ô A1 
            exRange.Range["A1:F1"].MergeCells = true;
            exRange.Range["A1:F1"].Font.Bold = true;
            exRange.Range["A1:F1"].Font.Italic = true;
            exRange.Range["A1:F1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignRight;
            exRange = exSheet.Cells[4][hang + 17]; //Ô A1 
            exRange.Range["A1:C1"].MergeCells = true;
            exRange.Range["A1:C1"].Font.Italic = true;
            exRange.Range["A1:C1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            DateTime d = Convert.ToDateTime(tblThongtinHD.Rows[0][1]);
            exRange.Range["A1:C1"].Value = "TP HCM, ngày " + d.Day + " tháng " + d.Month + " năm " + d.Year;
            exRange.Range["A2:C2"].MergeCells = true;
            exRange.Range["A2:C2"].Font.Italic = true;
            exRange.Range["A2:C2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:C2"].Value = "Nhân viên bán hàng";
            exRange.Range["A6:C6"].MergeCells = true;
            exRange.Range["A6:C6"].Font.Italic = true;
            exRange.Range["A6:C6"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A6:C6"].Value = tblThongtinHD.Rows[0][6];
            exSheet.Name = "Hóa đơn nhập";
            exApp.Visible = true;
        }

        private void txtSoLuong_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void txtDonGiaMua_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else
                e.Handled = true;
        }
    }
}

