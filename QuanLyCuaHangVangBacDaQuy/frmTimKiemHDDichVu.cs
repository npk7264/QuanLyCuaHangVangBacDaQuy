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
    public partial class frmTimKiemHDDichVu : Form
    {
        DataTable tblTKHDDV;
        public frmTimKiemHDDichVu()
        {
            InitializeComponent();
        }

        private void frmTimKiemHDDichVu_Load(object sender, EventArgs e)
        {
            ResetValues();
            dgvTKHoaDon.DataSource = null;
        }

        private void ResetValues()
        {
            txtSoHD.Text = "";
            cboThang.SelectedIndex = -1;
            txtNam.Text = "";
            txtMaKH.Text = "";
            txtMaNV.Text = "";
            cboThanhToan.SelectedIndex = -1;
            cboTinhTrang.SelectedIndex = -1;
            txtSoHD.Focus();
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string sql;
            if ((txtSoHD.Text == "") && (cboThang.Text == "") && (txtNam.Text == "") &&
                (txtMaKH.Text == "") && (cboThanhToan.Text == "") &&
                (cboTinhTrang.Text == "") && (txtMaNV.Text == ""))
            {
                MessageBox.Show("Hãy nhập một điều kiện tìm kiếm!!!", "Yêu cầu ...", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            sql = "SELECT * FROM PHIEUDICHVU WHERE 1=1";
            if (txtSoHD.Text != "")
                sql = sql + " AND SoPhieu Like N'%" + txtSoHD.Text + "%'";
            if (cboThang.Text != "")
                sql = sql + " AND MONTH(NgayLap) =" + cboThang.Text;
            if (txtNam.Text != "")
                sql = sql + " AND YEAR(NgayLap) =" + txtNam.Text;
            if (txtMaKH.Text != "")
                sql = sql + " AND MaKH = N'" + txtMaKH.Text + "'";
            if (txtMaNV.Text != "")
                sql = sql + " AND MaNV = N'" + txtMaNV.Text + "'";
            if (cboThanhToan.Text != "")
            {
                if (cboThanhToan.Text == "Thanh toán xong")
                    sql = sql + " AND TongTienConLai = 0";
                else if (cboThanhToan.Text == "Còn nợ")
                    sql = sql + " AND TongTienConLai>0";
            }    
            if (cboTinhTrang.Text != "")
                sql = sql + " AND TinhTrang=N'" + cboTinhTrang.Text + "'";
            tblTKHDDV = Functions.GetDataToTable(sql);
            if (tblTKHDDV.Rows.Count == 0)
            {
                MessageBox.Show("Không có bản ghi thỏa mãn điều kiện!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("Có " + tblTKHDDV.Rows.Count + " bản ghi thỏa mãn điều kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dgvTKHoaDon.DataSource = tblTKHDDV;
            LoadDataGridView();
        }

        private void LoadDataGridView()
        {
            dgvTKHoaDon.Columns[0].HeaderText = "Số phiếu";
            dgvTKHoaDon.Columns[1].HeaderText = "Ngày lập";
            dgvTKHoaDon.Columns[2].HeaderText = "Mã khách hàng";
            dgvTKHoaDon.Columns[3].HeaderText = "Mã nhân viên";
            dgvTKHoaDon.Columns[4].HeaderText = "Tổng tiền";
            dgvTKHoaDon.Columns[5].HeaderText = "Tổng tiền trả trước";
            dgvTKHoaDon.Columns[6].HeaderText = "Tổng tiền còn lại";
            dgvTKHoaDon.Columns[7].HeaderText = "Tình trạng";
            dgvTKHoaDon.Columns[0].Width = 150;
            dgvTKHoaDon.Columns[1].Width = 150;
            dgvTKHoaDon.Columns[2].Width = 150;
            dgvTKHoaDon.Columns[3].Width = 150;
            dgvTKHoaDon.Columns[4].Width = 150;
            dgvTKHoaDon.Columns[5].Width = 150;
            dgvTKHoaDon.Columns[6].Width = 150;
            dgvTKHoaDon.Columns[7].Width = 150;
            dgvTKHoaDon.AllowUserToAddRows = false;
            dgvTKHoaDon.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void btnTimLai_Click(object sender, EventArgs e)
        {
            ResetValues();
            dgvTKHoaDon.DataSource = null;
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtTongTienConLai_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void dgvTKHoaDon_DoubleClick(object sender, EventArgs e)
        {
            string mahd;
            if (MessageBox.Show("Bạn có muốn hiển thị thông tin chi tiết?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                mahd = dgvTKHoaDon.CurrentRow.Cells["SoPhieu"].Value.ToString();
                frmPhieuDichVu frm = new frmPhieuDichVu();
                frm.txtSoPhieu.Text = mahd;
                frm.StartPosition = FormStartPosition.CenterParent;
                frm.ShowDialog();
            }
        }

        /*private void cboTinhTrang_DropDown(object sender, EventArgs e)
        {
            Functions.FillCombo("Hoàn thành, Chưa hoàn thành", cboTinhTrang, "TinhTrang", "TinhTrang");
            cboTinhTrang.SelectedIndex = -1;
        }*/

        private void txtSoHD_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }

        private void txtThang_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }

        private void txtNam_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }

        private void txtMaKH_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }

        private void txtMaNV_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }

        private void cboTinhTrang_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }

        private void txtTongTienConLai_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }


    }
}
