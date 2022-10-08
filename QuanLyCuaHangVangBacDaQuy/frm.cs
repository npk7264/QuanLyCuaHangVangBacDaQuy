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
    public partial class frmDanhMucLoaiSP : Form
    {
        DataTable tblLoaiSanPham;
        public frmDanhMucLoaiSP()
        {
            InitializeComponent();
        }

        private void frmDanhMucLoaiSP_Load(object sender, EventArgs e)
        {
            string sql;
            sql = "select * from DONVITINH";
            txtMaLoaiSP.Enabled = false;
            btnLuu.Enabled = false;
            btnBoQua.Enabled = false;
            LoadDataGridView();
            Functions.FillCombo(sql, cboTenDVT, "MaDVT", "TenDVT");
            cboTenDVT.SelectedIndex = -1;
            ResetValues();
        }

        // Phương thức nạp dữ liệu
        private void LoadDataGridView()
        {
            string sql;
            sql = "select * from LOAISANPHAM" ;
            tblLoaiSanPham = Functions.GetDataToTable(sql); //lấy dữ liệu
            dgvLoaiSP.DataSource = tblLoaiSanPham;
            dgvLoaiSP.Columns[0].HeaderText = "Mã loại sản phẩm";
            dgvLoaiSP.Columns[1].HeaderText = "Loại sản phẩm";
            dgvLoaiSP.Columns[2].HeaderText = "Đơn vị tính";
            dgvLoaiSP.Columns[3].HeaderText = "Phần trăm lợi nhuận";
            dgvLoaiSP.Columns[0].Width = 150;
            dgvLoaiSP.Columns[1].Width = 150;
            dgvLoaiSP.Columns[2].Width = 150;
            dgvLoaiSP.Columns[3].Width = 150;
            dgvLoaiSP.AllowUserToAddRows = false;
            dgvLoaiSP.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void dgvLoaiSP_Click(object sender, EventArgs e)
        {
            string MaDVT;
            string sql;
            if (btnThem.Enabled == false)
            {
                MessageBox.Show("Đang ở chế độ thêm mới!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaLoaiSP.Focus();
                return;
            }
            if (tblLoaiSanPham.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            txtMaLoaiSP.Text = dgvLoaiSP.CurrentRow.Cells["MaLoaiSP"].Value.ToString();
            txtTenLoaiSP.Text = dgvLoaiSP.CurrentRow.Cells["TenLoaiSP"].Value.ToString();
            MaDVT = dgvLoaiSP.CurrentRow.Cells["MaDVT"].Value.ToString();
            sql = "select TenDVT from DONVITINH where MaDVT=N'" + MaDVT + "'";
            cboTenDVT.Text = Functions.GetFieldValues(sql);
            txtPhanTram.Text = dgvLoaiSP.CurrentRow.Cells["PhanTramLoiNhuan"].Value.ToString();
            btnSua.Enabled = true;
            btnXoa.Enabled = true;
        }

        private void ResetValues()
        {
            txtMaLoaiSP.Text = "";
            txtTenLoaiSP.Text = "";
            cboTenDVT.Text = "";
            txtPhanTram.Text = "";
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnBoQua.Enabled = true;
            btnLuu.Enabled = true;
            btnThem.Enabled = false;
            ResetValues();
            txtMaLoaiSP.Enabled = true;
            txtMaLoaiSP.Focus();
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            if (txtMaLoaiSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập loại sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaLoaiSP.Focus();
                return;
            }
            if (txtTenLoaiSP.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên loại sản phẩm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenLoaiSP.Focus();
                return;
            }
            if (cboTenDVT.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải chọn đơn vị tính", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboTenDVT.Focus();
                return;
            }
            if (txtPhanTram.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập phần trăm lợi nhuận", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPhanTram.Focus();
                return;
            }
            sql = "select MaLoaiSP from LOAISANPHAM WHERE MaLoaiSP=N'" + txtMaLoaiSP.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã loại sản phẩm này đã có, bạn phải nhập mã khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaLoaiSP.Focus();
                txtMaLoaiSP.Text = "";
                return;
            }
            // Chèn thêm
            sql = "INSERT INTO LOAISANPHAM VALUES (N'" + txtMaLoaiSP.Text.Trim() + "',N'" + txtTenLoaiSP.Text.Trim() + "',N'" + cboTenDVT.SelectedValue.ToString() +  txtPhanTram.Text.Trim() + "')";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnXoa.Enabled = true;
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnBoQua.Enabled = false;
            btnLuu.Enabled = false;
            txtMaLoaiSP.Enabled = false;
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblLoaiSanPham.Rows.Count == 0)
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
                MessageBox.Show("Bạn phải nhập tên hàng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTenLoaiSP.Focus();
                return;
            }
            if (cboTenDVT.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập chất liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboTenDVT.Focus();
                return;
            }
            sql = "UPDATE LOAISANPHAM SET TenLoaiSP=N'" + txtTenLoaiSP.Text.Trim().ToString() +
                "',MaDVT=N'" + cboTenDVT.SelectedValue.ToString() +
                "',PhanTramLoiNhuan=" + txtPhanTram.Text.ToString() + "' WHERE MaLoaiSP=N'" + txtMaLoaiSP.Text + "'";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnBoQua.Enabled = false;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblLoaiSanPham.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaLoaiSP.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                sql = "DELETE LOAISANPHAM WHERE MaLoaiSP=N'" + txtMaLoaiSP.Text + "'";
                Functions.RunSQL(sql);
                LoadDataGridView();
                ResetValues();
            }
        }

        private void btnBoQua_Click(object sender, EventArgs e)
        {
            ResetValues();
            btnXoa.Enabled = true;
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnBoQua.Enabled = false;
            btnLuu.Enabled = false;
            txtMaLoaiSP.Enabled = false;
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
