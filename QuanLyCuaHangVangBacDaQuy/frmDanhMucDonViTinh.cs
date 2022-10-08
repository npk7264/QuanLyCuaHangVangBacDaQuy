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
    public partial class frmDanhMucDonViTinh : Form
    {
        DataTable tblDVT;
        public frmDanhMucDonViTinh()
        {
            InitializeComponent();
        }

        private void DanhMucDonViTinh_Load(object sender, EventArgs e)
        {
            txtMaDVT.Enabled = false;
            btnLuu.Enabled = false;
            LoadDataGridView();
        }

        private void LoadDataGridView()
        {
            string sql;
            sql = "select * from DONVITINH";
            tblDVT = Functions.GetDataToTable(sql); //lấy dữ liệu
            dgvDonViTinh.DataSource = tblDVT;
            dgvDonViTinh.Columns[0].HeaderText = "Mã đơn vị tính";
            dgvDonViTinh.Columns[1].HeaderText = "Tên đơn vị tính";
            dgvDonViTinh.Columns[0].Width = 200;
            dgvDonViTinh.Columns[1].Width = 350;
            dgvDonViTinh.AllowUserToAddRows = false;
            dgvDonViTinh.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void ResetValues()
        {
            txtMaDVT.Text = "";
            txtTenDVT.Text = "";
        }

        private void dgvDonViTinh_Click(object sender, EventArgs e)
        {
            if (btnThem.Enabled == false)
            {
                MessageBox.Show("Đang ở chế độ thêm mới!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaDVT.Focus();
                return;
            }
            if (tblDVT.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            txtMaDVT.Text = dgvDonViTinh.CurrentRow.Cells["MaDVT"].Value.ToString();
            txtTenDVT.Text = dgvDonViTinh.CurrentRow.Cells["TenDVT"].Value.ToString();
            btnSua.Enabled = true;
            btnBoQua.Enabled = true;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnBoQua.Enabled = true;
            btnLuu.Enabled = true;
            btnThem.Enabled = false;
            ResetValues();
            txtMaDVT.Enabled = true;
            txtMaDVT.Focus();
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblDVT.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaDVT.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtTenDVT.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên đơn vị tính", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenDVT.Focus();
                return;
            }

            sql = "UPDATE DONVITINH SET TenDVT=N'" + txtTenDVT.Text.Trim().ToString() + "' where MaDVT=N'"+txtMaDVT.Text+"'";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnBoQua.Enabled = false;
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            if (txtMaDVT.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã đơn vị tính", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaDVT.Focus();
                return;
            }
            if (txtTenDVT.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên đơn vị tính", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenDVT.Focus();
                return;
            }
            sql = "SELECT MaDVT FROM DONVITINH WHERE MaDVT =N'" + txtMaDVT.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã đơn vị tính này đã có, bạn phải nhập mã khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaDVT.Focus();
                txtMaDVT.Text = "";
                return;
            }
            sql = "INSERT INTO DONVITINH VALUES (N'" + txtMaDVT.Text.Trim() + "',N'" + txtTenDVT.Text.Trim() + "')";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnBoQua.Enabled = false;
            btnLuu.Enabled = false;
            txtMaDVT.Enabled = false;
        }

        private void btnBoQua_Click(object sender, EventArgs e)
        {
            ResetValues();
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnBoQua.Enabled = false;
            btnLuu.Enabled = false;
            txtMaDVT.Enabled = false;
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
