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
    public partial class frmDanhMucDichVu : Form
    {
        DataTable tblDV;
        public frmDanhMucDichVu()
        {
            InitializeComponent();
        }

        private void frmDanhMucDichVu_Load(object sender, EventArgs e)
        {
            txtMaDV.Enabled = false;
            btnLuu.Enabled = false;
            btnBoQua.Enabled = false;
            LoadDataGridView();
        }

        private void LoadDataGridView()
        {
            string sql;
            sql = "select*from DICHVU";
            tblDV = Functions.GetDataToTable(sql);
            dgvDichVu.DataSource = tblDV;
            dgvDichVu.Columns[0].HeaderText = "Mã dịch vụ";
            dgvDichVu.Columns[1].HeaderText = "Loại dịch vụ";
            dgvDichVu.Columns[2].HeaderText = "Đơn giá dịch vụ";
            dgvDichVu.Columns[0].Width = 100;
            dgvDichVu.Columns[1].Width = 300;
            dgvDichVu.Columns[2].Width = 220;
            dgvDichVu.AllowUserToAddRows = false;
            dgvDichVu.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void dgvDichVu_Click(object sender, EventArgs e)
        {
            if (btnThem.Enabled == false)
            {
                MessageBox.Show("Đang ở chế độ thêm mới", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaDV.Focus();
                return;
            }
            if (tblDV.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            txtMaDV.Text = dgvDichVu.CurrentRow.Cells["MaDV"].Value.ToString();
            txtLoaiDV.Text = dgvDichVu.CurrentRow.Cells["LoaiDV"].Value.ToString();
            txtDonGiaDV.Text = dgvDichVu.CurrentRow.Cells["DonGiaDV"].Value.ToString();
            btnSua.Enabled = true;
            btnBoQua.Enabled = true;
        }

        private void ResetValue()
        {
            txtMaDV.Text = "";
            txtLoaiDV.Text = "";
            txtDonGiaDV.Text = "";
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnBoQua.Enabled = true;
            btnLuu.Enabled = true;
            btnThem.Enabled = false;
            ResetValue(); //Xoá trắng các textbox
            txtMaDV.Enabled = true; //cho phép nhập mới
            txtMaDV.Focus();
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql; //Lưu lệnh sql
            if (txtMaDV.Text.Trim().Length == 0) //Nếu chưa nhập mã dịch vụ
            {
                MessageBox.Show("Bạn phải nhập mã dịch vụ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaDV.Focus();
                return;
            }
            if (txtLoaiDV.Text.Trim().Length == 0) //Nếu chưa nhập loại dịch vụ
            {
                MessageBox.Show("Bạn phải nhập loại dịch vụ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtLoaiDV.Focus();
                return;
            }
            if (txtDonGiaDV.Text.Trim().Length == 0) //Nếu chưa nhập đơn giá dịch vụ
            {
                MessageBox.Show("Bạn phải nhập đơn giá dịch vụ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtDonGiaDV.Focus();
                return;
            }
            sql = "Select MaDV From DICHVU where MaDV=N'" + txtMaDV.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã dịch vụ này đã có, bạn phải nhập mã khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaDV.Focus();
                return;
            }

            sql = "INSERT INTO DICHVU VALUES(N'" + txtMaDV.Text + "',N'" + txtLoaiDV.Text + "',N'" + txtDonGiaDV.Text + "')";
            Functions.RunSQL(sql); //Thực hiện câu lệnh sql
            LoadDataGridView(); //Nạp lại DataGridView
            ResetValue();
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnBoQua.Enabled = false;
            btnLuu.Enabled = false;
            txtMaDV.Enabled = false;
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql; //Lưu câu lệnh sql
            if (tblDV.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaDV.Text == "") //nếu chưa chọn bản ghi nào
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtLoaiDV.Text.Trim().Length == 0) //nếu chưa nhập loại dịch vụ
            {
                MessageBox.Show("Bạn chưa nhập loại dịch vụ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtDonGiaDV.Text.Trim().Length == 0) //nếu chưa nhập đơn giá dịch vụ
            {
                MessageBox.Show("Bạn chưa nhập đơn giá dịch vụ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            sql = "UPDATE DICHVU SET LoaiDV=N'" + txtLoaiDV.Text.ToString() + "',DonGiaDV = N'" + txtDonGiaDV.Text.ToString() + "' WHERE MaDV=N'" + txtMaDV.Text + "'";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValue();
            btnBoQua.Enabled = false;
        }

        private void btnBoQua_Click(object sender, EventArgs e)
        {
            ResetValue();
            btnBoQua.Enabled = false;
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnLuu.Enabled = false;
            txtMaDV.Enabled = false;
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtMaDV_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }

        private void txtLoaiDV_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }

        private void txtDonGiaDV_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else
                e.Handled = true;
        }
    }
}
