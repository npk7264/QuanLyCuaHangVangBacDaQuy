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
    public partial class frmPhieuDichVu : Form
    {
        DataTable tblCTPDV;
        public frmPhieuDichVu()
        {
            InitializeComponent();
        }

        private void frmHoaDonDichVu_Load(object sender, EventArgs e)
        {
            btnThem.Enabled = true;
            btnLuu.Enabled = false;
            btnHuy.Enabled = false;
            btnIn.Enabled = false;
            //
            txtTenKH.ReadOnly = true;
            txtDiaChi.ReadOnly = true;
            txtSDT.ReadOnly = true;
            //
            txtTenNV.ReadOnly = true;
            //
            txtLoaiDV.ReadOnly = true;
            txtDonGiaDV.ReadOnly = true;
            txtDonGiaDuocTinh.ReadOnly = true;
            txtThanhTien.ReadOnly = true;
            txtTienConLai.ReadOnly = true;
            txtTongTien.ReadOnly = true;
            txtTongTienTraTruoc.ReadOnly = true;
            txtTongTienConLai.ReadOnly = true;
            txtTongTien.Text = "0";
            txtTongTienTraTruoc.Text = "0";
            txtTongTienConLai.Text = "0";

            Functions.FillCombo("SELECT MaKH, TenKH FROM KHACHHANG", cboMaKH, "MaKH", "MaKH");
            cboMaKH.SelectedIndex = -1;
            Functions.FillCombo("SELECT MaNV, TenNV FROM NHANVIEN", cboMaNV, "MaNV", "MaNV");
            cboMaNV.SelectedIndex = -1;
            Functions.FillCombo("SELECT MaDV, LoaiDV FROM DICHVU", cboMaDV, "MaDV", "MaDV");
            cboMaDV.SelectedIndex = -1;
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
            str = "select NgayLap from PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            dtpNgayLap.Value = DateTime.Parse(Functions.GetFieldValues(str));
            str = "SELECT MaKH FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            cboMaKH.Text = Functions.GetFieldValues(str);
            str = "SELECT MaNV FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            cboMaNV.Text = Functions.GetFieldValues(str);
            str = "SELECT TongTien FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            txtTongTien.Text = Functions.GetFieldValues(str);
            str = "SELECT TongTienTraTruoc FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            txtTongTienTraTruoc.Text = Functions.GetFieldValues(str);
            str = "SELECT TongTienConLai FROM PHIEUDICHVU WHERE SoPhieu=N'" + txtSoPhieu.Text + "'";
            txtTongTienConLai.Text = Functions.GetFieldValues(str);
            str = "SELECT TinhTrang FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
            txtTinhTrangPhieuDV.Text = Functions.GetFieldValues(str);
        }

        private void LoadDataGridView()
        {
            string sql;
            sql = "SELECT x.MaDV, LoaiDV, DonGiaDV, " +
                "DonGiaDuocTinh, " +
                "SoLuong, ThanhTien, " +
                "TraTruoc, ConLai, " +
                "NgayGiao, TinhTrang " +
                "FROM CTPDV x, DICHVU y " +
                "where x.SoPhieu = N'" + txtSoPhieu.Text.Trim() + "' and x.MaDV=y.MaDV";
            tblCTPDV = Functions.GetDataToTable(sql);
            dgvPhieuDichVu.DataSource = tblCTPDV;
            dgvPhieuDichVu.Columns[0].HeaderText = "Mã dịch vụ";
            dgvPhieuDichVu.Columns[1].HeaderText = "Loại dịch vụ";
            dgvPhieuDichVu.Columns[2].HeaderText = "Đơn giá dịch vụ";
            dgvPhieuDichVu.Columns[3].HeaderText = "Đơn giá được tính";
            dgvPhieuDichVu.Columns[4].HeaderText = "Số lượng";
            dgvPhieuDichVu.Columns[5].HeaderText = "Thành tiền";
            dgvPhieuDichVu.Columns[6].HeaderText = "Trả trước";
            dgvPhieuDichVu.Columns[7].HeaderText = "Còn lại";
            dgvPhieuDichVu.Columns[8].HeaderText = "Ngày giao";
            dgvPhieuDichVu.Columns[9].HeaderText = "Tình trạng";

            dgvPhieuDichVu.Columns[0].Width = 150;
            dgvPhieuDichVu.Columns[1].Width = 200;
            dgvPhieuDichVu.Columns[2].Width = 150;
            dgvPhieuDichVu.Columns[3].Width = 150;
            dgvPhieuDichVu.Columns[4].Width = 100;
            dgvPhieuDichVu.Columns[5].Width = 150;
            dgvPhieuDichVu.Columns[6].Width = 150;
            dgvPhieuDichVu.Columns[7].Width = 150;
            dgvPhieuDichVu.Columns[8].Width = 100;
            dgvPhieuDichVu.Columns[9].Width = 100;
            dgvPhieuDichVu.AllowUserToAddRows = false;
            dgvPhieuDichVu.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnHuy.Enabled = true;
            btnLuu.Enabled = true;
            btnIn.Enabled = false;
            btnThem.Enabled = false;
            ResetValues();
            txtSoPhieu.Text = Functions.CreateKey("PDV");
            LoadDataGridView();
        }

        private void ResetValues()
        {
            txtSoPhieu.Text = "";
            dtpNgayLap.Value = DateTime.Now;
            cboMaKH.SelectedIndex = -1;
            cboMaNV.SelectedIndex = -1;
            cboMaDV.SelectedIndex = -1;
            txtChiPhiRieng.Text = "";
            txtSoLuong.Text = "";
            txtTienTraTruoc.Text = "";
            cboTinhTrangDV.SelectedIndex = -1;
            dtpNgayGiao.Value = DateTime.Now;
            txtTongTien.Text = "0";
            txtTongTienTraTruoc.Text = "0";
            txtTongTienConLai.Text = "0";
            txtTinhTrangPhieuDV.Text = "";
        }

        private void ResetValuesDichVu()
        {
            cboMaDV.SelectedIndex = -1;
            txtChiPhiRieng.Text = "";
            txtSoLuong.Text = "";
            txtTienTraTruoc.Text = "";
            cboTinhTrangDV.SelectedIndex = -1;
            dtpNgayGiao.Value = DateTime.Now;
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cboMaKH_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaKH.SelectedIndex == -1 || cboMaKH.Text == "")
            {
                txtTenKH.Text = "";
                txtDiaChi.Text = "";
                txtSDT.Text = "";
            }
            else
            {
                // Khi chọn mã khách hàng thì các thông tin khách hàng tự động hiện ra
                str = "select TenKH from KHACHHANG where MaKH =N'" + cboMaKH.SelectedValue + "'";
                txtTenKH.Text = Functions.GetFieldValues(str);
                str = "select DiaChi from KHACHHANG where MaKH =N'" + cboMaKH.SelectedValue + "'";
                txtDiaChi.Text = Functions.GetFieldValues(str);
                str = "select SDT from KHACHHANG where MaKH =N'" + cboMaKH.SelectedValue + "'";
                txtSDT.Text = Functions.GetFieldValues(str);
            }
        }

        private void cboMaNV_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaNV.SelectedIndex == -1 || cboMaNV.Text == "")
                txtTenNV.Text = "";
            else
            {
                str = "select TenNV from NHANVIEN where MaNV=N'" + cboMaNV.SelectedValue + "'";
                txtTenNV.Text = Functions.GetFieldValues(str);
            }
        }

        private void cboMaDV_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaDV.SelectedIndex == -1 || cboMaDV.Text == "")//
            {
                txtLoaiDV.Text = "";
                txtDonGiaDV.Text = "";
            }
            else
            {
                str = "select LoaiDV from DICHVU where MaDV=N'" + cboMaDV.SelectedValue + "'";
                txtLoaiDV.Text = Functions.GetFieldValues(str);
                str = "select DonGiaDV from DICHVU where MaDV=N'" + cboMaDV.SelectedValue + "'";
                txtDonGiaDV.Text = Functions.GetFieldValues(str);
            }
        }

        private void txtDonGiaDV_TextChanged(object sender, EventArgs e)
        {
            double chiphirieng, dongiadv, dongia;
            if (txtChiPhiRieng.Text == "")
                chiphirieng = 0;
            else
                chiphirieng = Convert.ToDouble(txtChiPhiRieng.Text);
            if (txtDonGiaDV.Text == "")
                dongiadv = 0;
            else
                dongiadv = Convert.ToDouble(txtDonGiaDV.Text);
            dongia = chiphirieng + dongiadv;
            txtDonGiaDuocTinh.Text = dongia.ToString();
        }

        private void txtChiPhiRieng_TextChanged(object sender, EventArgs e)
        {
            double chiphirieng, dongiadv, dongia;
            if (txtChiPhiRieng.Text == "")
                chiphirieng = 0;
            else
                chiphirieng = Convert.ToDouble(txtChiPhiRieng.Text);
            if (txtDonGiaDV.Text == "")
                dongiadv = 0;
            else
                dongiadv = Convert.ToDouble(txtDonGiaDV.Text);
            dongia = chiphirieng + dongiadv;
            txtDonGiaDuocTinh.Text = dongia.ToString();
        }

        private void txtSoLuong_TextChanged(object sender, EventArgs e)
        {
            double dongia, soluong, thanhtien;
            if (txtDonGiaDuocTinh.Text == "")
                dongia = 0;
            else
                dongia = Convert.ToDouble(txtDonGiaDuocTinh.Text);
            if (txtSoLuong.Text == "")
                soluong = 0;
            else
                soluong = Convert.ToDouble(txtSoLuong.Text);
            thanhtien = soluong * dongia;
            txtThanhTien.Text = thanhtien.ToString();
        }

        private void txtTienTraTruoc_TextChanged(object sender, EventArgs e)
        {
            double thanhtien, tratruoc, conlai;
            if (txtThanhTien.Text == "")
                thanhtien = 0;
            else
                thanhtien = Convert.ToDouble(txtThanhTien.Text);
            if (txtTienTraTruoc.Text == "")
                tratruoc = 0;
            else
                tratruoc = Convert.ToDouble(txtTienTraTruoc.Text);
            conlai = thanhtien - tratruoc;
            txtTienConLai.Text = conlai.ToString();
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            double tong, Tongmoi, tongtratruoc, Tongtratruocmoi;
            //
            sql = "select SoPhieu FROM PHIEUDICHVU WHERE SoPhieu=N'" + txtSoPhieu.Text + "'";
            if (!Functions.CheckKey(sql))
            {
                if (cboMaKH.Text.Length == 0)
                {
                    MessageBox.Show("Bạn phải nhập mã khách hàng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cboMaKH.Focus();
                    return;
                }
                if (cboMaNV.Text.Length == 0)
                {
                    MessageBox.Show("Bạn phải nhập mã nhân viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cboMaNV.Focus();
                    return;
                }
                sql = "insert into PHIEUDICHVU values (N'" + txtSoPhieu.Text.Trim() + "','" +
                        dtpNgayLap.Value + "', '" +
                        cboMaKH.SelectedValue + "', '" +
                        cboMaNV.SelectedValue + "', " +
                        txtTongTien.Text + ", " +
                        txtTongTienTraTruoc.Text + ", " +
                        txtTongTienConLai.Text + ",N'" +
                        txtTinhTrangPhieuDV.Text + "')";
                Functions.RunSQL(sql);
            }
            // Lưu thông tin các dịch vụ
            if (cboMaDV.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã dịch vụ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboMaDV.Focus();
                return;
            }
            if (txtChiPhiRieng.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập chi phí riêng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtChiPhiRieng.Focus();
                return;
            }
            if (txtSoLuong.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập số lượng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtSoLuong.Focus();
                return;
            }
            if (txtTienTraTruoc.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập số tiền trả trước", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTienTraTruoc.Focus();
                return;
            }
            if (cboTinhTrangDV.Text == "")
            {
                MessageBox.Show("Bạn phải nhập tình trạng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTienTraTruoc.Focus();
                return;
            }
            //
            sql = "select MaDV FROM CTPDV WHERE MaDV=N'" + cboMaDV.SelectedValue + "' and SoPhieu = N'" + txtSoPhieu.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã dịch vụ này đã có, bạn phải nhập mã khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetValuesDichVu();
                cboMaDV.Focus();
                return;
            }
            //
            double tratruoc = Convert.ToDouble(txtTienTraTruoc.Text);
            double thanhtien = Convert.ToDouble(txtThanhTien.Text);
            if (tratruoc < 0.5 * thanhtien || tratruoc > thanhtien)
            {
                MessageBox.Show("Tiền trả trước phải lớn hơn hoặc bằng 50% thành tiền", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTienTraTruoc.Text = "";
                txtTienTraTruoc.Focus();
                return;
            }
            if (dtpNgayGiao.Value < dtpNgayLap.Value)
            {
                MessageBox.Show("Ngày giao phải bằng hoặc sau ngày lập phiếu dịch vụ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dtpNgayGiao.Value = DateTime.Now;
                dtpNgayGiao.Focus();
                return;
            }
            //
            sql = "insert into CTPDV values(N'" + txtSoPhieu.Text.Trim() + "',N'" + cboMaDV.SelectedValue + "'," + txtChiPhiRieng.Text + ", " + txtDonGiaDuocTinh.Text + ", " + txtSoLuong.Text + "," + txtThanhTien.Text + "," + txtTienTraTruoc.Text + ", " + txtTienConLai.Text + ", '" + dtpNgayGiao.Value.Date + "', N'" + cboTinhTrangDV.Text + "')";
            Functions.RunSQL(sql);
            LoadDataGridView();

            // Cập nhật tổng
            tong = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTien FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'"));
            Tongmoi = tong + Convert.ToDouble(txtThanhTien.Text);
            //
            tongtratruoc = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTienTraTruoc FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'"));
            Tongtratruocmoi = tongtratruoc + Convert.ToDouble(txtTienTraTruoc.Text);
            //
            double Tongconlaimoi = Tongmoi - Tongtratruocmoi;
            //
            sql = "update PHIEUDICHVU set TongTien=" + Tongmoi + "where SoPhieu=N'" + txtSoPhieu.Text + "'";
            Functions.RunSQL(sql);
            sql = "update PHIEUDICHVU set TongTienTraTruoc=" + Tongtratruocmoi + "where SoPhieu=N'" + txtSoPhieu.Text + "'";
            Functions.RunSQL(sql);
            sql = "update PHIEUDICHVU set TongTienConLai=" + Tongconlaimoi + "where SoPhieu=N'" + txtSoPhieu.Text + "'";
            Functions.RunSQL(sql);

            int count = Convert.ToInt32(Functions.GetFieldValues("select count(TinhTrang) from CTPDV where TinhTrang=N'Chưa giao' and SoPhieu=N'" + txtSoPhieu.Text + "'"));
            if (count > 0)
                sql = "update PHIEUDICHVU set TinhTrang=N'Chưa hoàn thành' where SoPhieu=N'" + txtSoPhieu.Text + "'";
            else
                sql = "update PHIEUDICHVU set TinhTrang=N'Hoàn thành' where SoPhieu=N'" + txtSoPhieu.Text + "'";
            Functions.RunSQL(sql);

            string tt = Convert.ToString(Functions.GetFieldValues("select TinhTrang FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'"));
            txtTongTien.Text = Tongmoi.ToString();
            txtTongTienTraTruoc.Text = Tongtratruocmoi.ToString();
            txtTongTienConLai.Text = Tongconlaimoi.ToString();
            txtTinhTrangPhieuDV.Text = tt;

            ResetValuesDichVu();
            btnHuy.Enabled = true;
            btnThem.Enabled = true;
            btnIn.Enabled = true;
        }

        private void dgvPhieuDichVu_DoubleClick(object sender, EventArgs e)
        {
            string MaDVxoa, sql;
            Double ThanhTienxoa, TraTruocxoa, ConLaixoa, tong, tongmoi, truoc, truocmoi, conmoi;
            if (tblCTPDV.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if ((MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                //Xóa dịch vụ 
                MaDVxoa = dgvPhieuDichVu.CurrentRow.Cells["MaDV"].Value.ToString();
                ThanhTienxoa = Convert.ToDouble(dgvPhieuDichVu.CurrentRow.Cells["ThanhTien"].Value.ToString());
                TraTruocxoa = Convert.ToDouble(dgvPhieuDichVu.CurrentRow.Cells["TraTruoc"].Value.ToString());
                ConLaixoa = Convert.ToDouble(dgvPhieuDichVu.CurrentRow.Cells["ConLai"].Value.ToString());
                sql = "DELETE from CTPDV WHERE SoPhieu=N'" + txtSoPhieu.Text + "' AND MaDV = N'" + MaDVxoa + "'";
                Functions.RunSQL(sql);

                // Cập nhật lại tổng tiền cho hóa đơn bán
                tong = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTien FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'"));
                tongmoi = tong - ThanhTienxoa;
                sql = "UPDATE PHIEUDICHVU SET TongTien =" + tongmoi + " WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
                Functions.RunSQL(sql);
                txtTongTien.Text = tongmoi.ToString();

                // Cập nhật lại tổng tiền trả trước
                truoc = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTienTraTruoc FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'"));
                truocmoi = truoc - TraTruocxoa;
                sql = "UPDATE PHIEUDICHVU SET TongTienTraTruoc =" + truocmoi + " WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
                Functions.RunSQL(sql);
                txtTongTienTraTruoc.Text = truocmoi.ToString();

                // Cập nhật lại tổng tiền còn lại
                conmoi = tongmoi - truocmoi;
                sql = "UPDATE PHIEUDICHVU SET TongTienConLai =" + conmoi + " WHERE SoPhieu = N'" + txtSoPhieu.Text + "'";
                Functions.RunSQL(sql);
                txtTongTienConLai.Text = conmoi.ToString();

                // Cập nhật lại tình trạng phiếu dịch vụ
                int count = Convert.ToInt32(Functions.GetFieldValues("select count(TinhTrang) from CTPDV where TinhTrang=N'Chưa giao' and SoPhieu=N'" + txtSoPhieu.Text + "'"));
                if (count > 0)
                    sql = "update PHIEUDICHVU set TinhTrang=N'Chưa hoàn thành' where SoPhieu=N'" + txtSoPhieu.Text + "'";
                else
                    sql = "update PHIEUDICHVU set TinhTrang=N'Hoàn thành' where SoPhieu=N'" + txtSoPhieu.Text + "'";
                Functions.RunSQL(sql);
                string tt = Convert.ToString(Functions.GetFieldValues("select TinhTrang FROM PHIEUDICHVU WHERE SoPhieu = N'" + txtSoPhieu.Text + "'"));
                txtTinhTrangPhieuDV.Text = tt;
                //
                LoadDataGridView();
            }


        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            string sql;
            if (MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //Xóa chi tiết hóa đơn
                sql = "DELETE from CTPDV WHERE SoPhieu=N'" + txtSoPhieu.Text + "'";
                Functions.RunSqlDel(sql);

                //Xóa hóa đơn
                sql = "DELETE from PHIEUDICHVU WHERE SoPhieu=N'" + txtSoPhieu.Text + "'";
                Functions.RunSqlDel(sql);
                ResetValues();
                LoadDataGridView();
                btnThem.Enabled = true;
                btnHuy.Enabled = false;
                btnIn.Enabled = false;
            }
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
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
            Functions.FillCombo("SELECT SoPhieu FROM PHIEUDICHVU", cboSoPhieu, "SoPhieu", "SoPhieu");
            cboSoPhieu.SelectedIndex = -1;
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
            exRange.Range["C2:I2"].Font.Size = 16;
            exRange.Range["C2:I2"].Font.Bold = true;
            exRange.Range["C2:I2"].Font.ColorIndex = 3; //Màu đỏ
            exRange.Range["C2:I2"].MergeCells = true;
            exRange.Range["C2:I2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C2:I2"].Value = "HÓA ĐƠN DỊCH VỤ";
            // Biểu diễn thông tin chung của hóa đơn bán
            sql = "SELECT a.SoPhieu, a.NgayLap, a.TongTien, b.TenKH, b.DiaChi, b.SDT, c.TenNV FROM PHIEUDICHVU AS a, KHACHHANG AS b, NHANVIEN AS c WHERE a.SoPhieu = N'" + txtSoPhieu.Text + "' AND a.MaKH = b.MaKH AND a.MaNV = c.MaNV";
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
            exRange.Range["C8:E8"].Value = tblThongtinHD.Rows[0][4].ToString();
            exRange.Range["B9:B9"].Value = "Số điện thoại:";
            exRange.Range["C9:E9"].MergeCells = true;
            exRange.Range["C9:E9"].Value = "'" + tblThongtinHD.Rows[0][5].ToString();
            //Lấy thông tin các mặt hàng
            sql = "SELECT b.LoaiDV, a.SoLuong, b.DonGiaDV, a.ChiPhiRieng, a.DonGiaDuocTinh, a.ThanhTien, a.TraTruoc, a.ConLai, a.NgayGiao, a.TinhTrang " +
                  "FROM CTPDV AS a , DICHVU AS b WHERE a.SoPhieu = N'" +
                  txtSoPhieu.Text + "' AND a.MaDV = b.MaDV";
            tblThongtinHang = Functions.GetDataToTable(sql);
            //Tạo dòng tiêu đề bảng
            exRange.Range["A11:K11"].Font.Bold = true;
            exRange.Range["A11:K11"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A12:K12"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["B11:K11"].ColumnWidth = 20;
            exRange.Range["A11:A11"].Value = "STT";
            exRange.Range["B11:B11"].Value = "Loại dịch vụ";
            exRange.Range["C11:C11"].Value = "Số lượng";
            exRange.Range["D11:D11"].Value = "Đơn giá dịch vụ";
            exRange.Range["E11:E11"].Value = "Chi phí riêng";
            exRange.Range["F11:F11"].Value = "Đơn giá được tính";
            exRange.Range["G11:G11"].Value = "Thành tiền";
            exRange.Range["H11:H11"].Value = "Trả trước";
            exRange.Range["I11:I11"].Value = "Còn lại";
            exRange.Range["J11:J11"].Value = "Ngày giao";
            exRange.Range["K11:K11"].Value = "Tình trạng";
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
                    if (cot == 8) exSheet.Cells[cot + 2][hang + 12] = "'" + tblThongtinHang.Rows[hang][cot].ToString();
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
            exRange = exSheet.Cells[9][hang + 17]; //Ô A1 
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
            exSheet.Name = "Hóa đơn dịch vụ";
            exApp.Visible = true;
        }

        private void txtChiPhiRieng_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void txtSoLuong_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void txtTienTraTruoc_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else
                e.Handled = true;
        }
    }
}
