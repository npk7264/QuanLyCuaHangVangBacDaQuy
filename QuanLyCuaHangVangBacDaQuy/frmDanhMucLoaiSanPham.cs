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
    public partial class frmDanhMucLoaiSanPham : Form
    {
        DataTable tblLSP;
        public frmDanhMucLoaiSanPham()
        {
            InitializeComponent();
        }

        private void frmDanhMucLoaiSanPham_Load(object sender, EventArgs e)
        {
            string sql;
            sql = "select * from DONVITINH";
            txtMaLoaiSP.Enabled = false;
            btnLuu.Enabled = false;
            LoadDataGridView();
            Functions.FillCombo(sql, cboMaDVT, "MaDVT", "TenDVT");
            cboMaDVT.SelectedIndex = -1;
            ResetValues();
        }

        private void LoadDataGridView()
        {
            string sql;
            sql = "select MaLoaiSP, TenLoaiSP, TenDVT, PhanTramLoiNhuan from LOAISANPHAM x, DONVITINH y where x.MaDVT=y.MaDVT";
            tblLSP = Functions.GetDataToTable(sql);
            dgvLoaiSanPham.DataSource = tblLSP;
            dgvLoaiSanPham.Columns[0].HeaderText = "Mã loại sản phẩm";
            dgvLoaiSanPham.Columns[1].HeaderText = "Loại sản phẩm";
            dgvLoaiSanPham.Columns[2].HeaderText = "Đơn vị tính";
            dgvLoaiSanPham.Columns[3].HeaderText = "Phần trăm lợi nhuận";
            dgvLoaiSanPham.Columns[0].Width = 100;
            dgvLoaiSanPham.Columns[1].Width = 250;
            dgvLoaiSanPham.Columns[2].Width = 230;
            dgvLoaiSanPham.Columns[3].Width = 150;
            dgvLoaiSanPham.AllowUserToAddRows = false;
            dgvLoaiSanPham.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void ResetValues()
        {
            txtMaLoaiSP.Text = "";
            txtTenLoaiSP.Text = "";
            cboMaDVT.SelectedValue = -1;
            txtLoiNhuan.Text = "";
        }

        private void dgvLoaiSanPham_Click(object sender, EventArgs e)
        {
            string TenDonViTinh;
            string sql;
            if (btnThem.Enabled == false)
            {
                MessageBox.Show("Đang ở chế độ thêm mới!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaLoaiSP.Focus();
                return;
            }
            if (tblLSP.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            txtMaLoaiSP.Text = dgvLoaiSanPham.CurrentRow.Cells["MaLoaiSP"].Value.ToString();
            txtTenLoaiSP.Text = dgvLoaiSanPham.CurrentRow.Cells["TenLoaiSP"].Value.ToString();
            TenDonViTinh = dgvLoaiSanPham.CurrentRow.Cells["TenDVT"].Value.ToString();
            sql = "select TenDVT from DONVITINH where TenDVT=N'" + TenDonViTinh + "'";
            cboMaDVT.Text = Functions.GetFieldValues(sql);
            txtLoiNhuan.Text = dgvLoaiSanPham.CurrentRow.Cells["PhanTramLoiNhuan"].Value.ToString();
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
            txtMaLoaiSP.Enabled = true;
            txtMaLoaiSP.Focus();
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            if (txtMaLoaiSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã loại sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaLoaiSP.Focus();
                return;
            }
            if (txtTenLoaiSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên loại sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenLoaiSP.Focus();
                return;
            }
            if (cboMaDVT.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải chọn đơn vị tính", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboMaDVT.Focus();
                return;
            }
            if (txtLoiNhuan.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập phần trăm lợi nhuận", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtLoiNhuan.Focus();
                return;
            }
            sql = "SELECT MaLoaiSP from LOAISANPHAM WHERE MaLoaiSP=N'" + txtMaLoaiSP.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã loại sản phẩm này đã có, bạn phải nhập mã khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaLoaiSP.Focus();
                txtMaLoaiSP.Text = "";
                return;
            }
            sql = "INSERT INTO LOAISANPHAM VALUES (N'" + txtMaLoaiSP.Text.Trim() + "',N'" + txtTenLoaiSP.Text.Trim() + "','" + cboMaDVT.SelectedValue.ToString() + "', " + txtLoiNhuan.Text.Trim() + ")";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnLuu.Enabled = false;
            txtMaLoaiSP.Enabled = false;
            btnBoQua.Enabled = false;
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblLSP.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaLoaiSP.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaLoaiSP.Focus();
                return;
            }
            if (txtTenLoaiSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên loại sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTenLoaiSP.Focus();
                return;
            }
            if (cboMaDVT.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập đơn vị tính", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboMaDVT.Focus();
                return;
            }
            if (txtLoiNhuan.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập phần trăm lợi nhuận", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtLoiNhuan.Focus();
                return;
            }
            sql = "UPDATE LOAISANPHAM SET TenLoaiSP=N'" + txtTenLoaiSP.Text.Trim().ToString() +
                "',MaDVT=N'" + cboMaDVT.SelectedValue.ToString() + "',PhanTramLoiNhuan=" + txtLoiNhuan.Text + " WHERE MaLoaiSP=N'" + txtMaLoaiSP.Text + "'";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnThem.Enabled = true;
            btnBoQua.Enabled = false;
        }

        private void btnBoQua_Click(object sender, EventArgs e)
        {
            ResetValues();
            btnSua.Enabled = true;
            btnThem.Enabled = true;
            btnBoQua.Enabled = false;
            btnLuu.Enabled = false;
            txtMaLoaiSP.Enabled = false;
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtLoiNhuan_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else
                e.Handled = true;
        }
    }
}
