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

namespace QuanLyCuaHangVangBacDaQuy
{
    public partial class frmDanhMucSanPham : Form
    {
        DataTable tblSanPham;
        public frmDanhMucSanPham()
        {
            InitializeComponent();
        }

        private void frmDanhMucSanPham_Load(object sender, EventArgs e)
        {
            string sql;
            sql = "select * from LOAISANPHAM";
            txtMaSP.Enabled = false;
            btnLuu.Enabled = false;
            LoadDataGridView();
            Functions.FillCombo(sql, cboMaLoaiSP, "MaLoaiSP", "TenLoaiSP");
            cboMaLoaiSP.SelectedIndex = -1;
            ResetValues();
        }

        private void LoadDataGridView()
        {
            string sql;
            sql = "select MaSP, TenSP, TenLoaiSP, SoLuong, DonGiaMua, DonGiaBan from SANPHAM, LOAISANPHAM where SANPHAM.MaLoaiSP=LOAISANPHAM.MaLoaiSP";
            tblSanPham = Functions.GetDataToTable(sql);
            dgvSanPham.DataSource = tblSanPham;
            dgvSanPham.Columns[0].HeaderText = "Mã sản phẩm";
            dgvSanPham.Columns[1].HeaderText = "Tên sản phẩm";
            dgvSanPham.Columns[2].HeaderText = "Loại sản phẩm";
            dgvSanPham.Columns[3].HeaderText = "Số lượng";
            dgvSanPham.Columns[4].HeaderText = "Đơn giá mua";
            dgvSanPham.Columns[5].HeaderText = "Đơn giá bán";
            dgvSanPham.Columns[0].Width = 80;
            dgvSanPham.Columns[1].Width = 200;
            dgvSanPham.Columns[2].Width = 200;
            dgvSanPham.Columns[3].Width = 100;
            dgvSanPham.Columns[4].Width = 100;
            dgvSanPham.Columns[5].Width = 100;
            dgvSanPham.AllowUserToAddRows = false;
            dgvSanPham.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void ResetValues()
        {
            txtMaSP.Text = "";
            txtTenSP.Text = "";
            cboMaLoaiSP.SelectedValue = -1;
            txtSoLuong.Text = "0";
            txtDonGiaMua.Text = "0";
            txtDonGiaBan.Text = "0";
            txtSoLuong.Enabled = false;
            txtDonGiaMua.Enabled = false;
            txtDonGiaBan.Enabled = false;
        }

        private void dgvSanPham_Click(object sender, EventArgs e)
        {
            string TenLoaiSanPham;
            string sql;
            if (btnThem.Enabled == false)
            {
                MessageBox.Show("Đang ở chế độ thêm mới!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaSP.Focus();
                return;
            }
            if (tblSanPham.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            txtMaSP.Text = dgvSanPham.CurrentRow.Cells["MaSP"].Value.ToString();
            txtTenSP.Text = dgvSanPham.CurrentRow.Cells["TenSP"].Value.ToString();
            TenLoaiSanPham = dgvSanPham.CurrentRow.Cells["TenLoaiSP"].Value.ToString();
            sql = "select TenLoaiSP from LOAISANPHAM where TenLoaiSP=N'" + TenLoaiSanPham + "'";
            cboMaLoaiSP.Text = Functions.GetFieldValues(sql);
            txtSoLuong.Text = dgvSanPham.CurrentRow.Cells["SoLuong"].Value.ToString();
            txtDonGiaMua.Text = dgvSanPham.CurrentRow.Cells["DonGiaMua"].Value.ToString();
            txtDonGiaBan.Text = dgvSanPham.CurrentRow.Cells["DonGiaBan"].Value.ToString();
            btnSua.Enabled = true;
            btnBoQua.Enabled = true;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnLuu.Enabled = true;
            btnThem.Enabled = false;
            btnBoQua.Enabled = true;
            ResetValues();
            txtMaSP.Enabled = true;
            txtMaSP.Focus();
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            if (txtMaSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaSP.Focus();
                return;
            }
            if (txtTenSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenSP.Focus();
                return;
            }
            if (cboMaLoaiSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải chọn loại sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboMaLoaiSP.Focus();
                return;
            }
            if (txtSoLuong.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập số điện thoại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSoLuong.Focus();
                return;
            }
            sql = "SELECT MaSP from SANPHAM WHERE MaSP=N'" + txtMaSP.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã sản phẩm này đã có, bạn phải nhập mã khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaSP.Focus();
                txtMaSP.Text = "";
                return;
            }
            sql = "INSERT INTO SANPHAM VALUES (N'" + txtMaSP.Text.Trim() + "',N'" + txtTenSP.Text.Trim() + "','" + cboMaLoaiSP.SelectedValue.ToString() + "', " + txtSoLuong.Text.Trim() + ", " + txtDonGiaMua.Text + ", " + txtDonGiaBan.Text + ")";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnLuu.Enabled = false;
            txtMaSP.Enabled = false;
            btnBoQua.Enabled = false;
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblSanPham.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaSP.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaSP.Focus();
                return;
            }
            if (txtTenSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên hàng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTenSP.Focus();
                return;
            }
            if (cboMaLoaiSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập loại sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboMaLoaiSP.Focus();
                return;
            }
            sql = "UPDATE SANPHAM SET TenSP=N'" + txtTenSP.Text.Trim().ToString() +
                "',MaLoaiSP=N'" + cboMaLoaiSP.SelectedValue.ToString() +"' WHERE MaSP=N'" + txtMaSP.Text + "'";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnThem.Enabled = true;
            btnBoQua.Enabled = false;
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string sql;
            if ((txtMaSP.Text == "") && (txtTenSP.Text == "") && (cboMaLoaiSP.Text == ""))
            {
                MessageBox.Show("Bạn hãy nhập điều kiện tìm kiếm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            sql = "select MaSP, TenSP, TenLoaiSP, SoLuong, DonGiaMua, DonGiaBan from SANPHAM, LOAISANPHAM where 1=1 and SANPHAM.MaLoaiSP=LOAISANPHAM.MaLoaiSP";
            if (txtMaSP.Text != "")
                sql += " AND MaSP LIKE N'%" + txtMaSP.Text + "%'";
            if (txtTenSP.Text != "")
                sql += " AND TenSP LIKE N'%" + txtTenSP.Text + "%'";
            if (cboMaLoaiSP.Text != "")
                sql += " AND LOAISANPHAM.MaLoaiSP LIKE N'%" + cboMaLoaiSP.SelectedValue + "%'";
            tblSanPham = Functions.GetDataToTable(sql);
            if (tblSanPham.Rows.Count == 0)
                MessageBox.Show("Không có bản ghi thoả mãn điều kiện tìm kiếm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else MessageBox.Show("Có " + tblSanPham.Rows.Count + "  bản ghi thoả mãn điều kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dgvSanPham.DataSource = tblSanPham;
            ResetValues();
        }

        private void btnHienThi_Click(object sender, EventArgs e)
        {

            string sql;
            sql = "select MaSP, TenSP, TenLoaiSP, SoLuong, DonGiaMua, DonGiaBan from SANPHAM, LOAISANPHAM where SANPHAM.MaLoaiSP=LOAISANPHAM.MaLoaiSP";
            tblSanPham = Functions.GetDataToTable(sql);
            dgvSanPham.DataSource = tblSanPham;
        }

        private void btnBoQua_Click(object sender, EventArgs e)
        {
            ResetValues();
            btnSua.Enabled = true;
            btnThem.Enabled = true;
            btnBoQua.Enabled = false;
            btnLuu.Enabled = false;
            txtMaSP.Enabled = false;
        }
    }
}
