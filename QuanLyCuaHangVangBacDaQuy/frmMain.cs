using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QuanLyCuaHangVangBacDaQuy.Class;

namespace QuanLyCuaHangVangBacDaQuy
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            Functions.Connect();
        }

        private void mnuThoat_Click(object sender, EventArgs e)
        {
            Functions.Disconnect();
            Application.Exit();
        }

        private void mnuDichVu_Click(object sender, EventArgs e)
        {
            frmDanhMucDichVu frm = new frmDanhMucDichVu();
            frm.ShowDialog();
        }

        private void mnuSanPham_Click(object sender, EventArgs e)
        {
            frmDanhMucSanPham frm = new frmDanhMucSanPham();
            frm.ShowDialog();
        }


        private void mnuKhachHang_Click(object sender, EventArgs e)
        {
            frmDanhMucKhachHang frm = new frmDanhMucKhachHang();
            frm.ShowDialog();
        }

        private void mnuHoaDonMua_Click(object sender, EventArgs e)
        {
            frmPhieuMuaHang frm = new frmPhieuMuaHang();
            frm.ShowDialog();
        }

        private void mnuHoaDonDichVu_Click(object sender, EventArgs e)
        {
            frmPhieuDichVu frm = new frmPhieuDichVu();
            frm.ShowDialog();
        }

        private void mnuHoaDonBan_Click(object sender, EventArgs e)
        {
            frmPhieuBanHang frm = new frmPhieuBanHang();
            frm.ShowDialog();
        }

        private void mnuBaoCaoHangTon_Click(object sender, EventArgs e)
        {
            frmBaoCaoTonKho frm = new frmBaoCaoTonKho();
            frm.ShowDialog();
        }

        private void mnuNhaCungCap_Click(object sender, EventArgs e)
        {
            frmDanhMucNhaCungCap frm = new frmDanhMucNhaCungCap();
            frm.ShowDialog();
        }

        private void mnuNhanVien_Click(object sender, EventArgs e)
        {
            frmDanhMucNhanVien frm = new frmDanhMucNhanVien();
            frm.ShowDialog();
        }

        private void mnuTimKiemHoaDonDV_Click(object sender, EventArgs e)
        {
            frmTimKiemHDDichVu frm = new frmTimKiemHDDichVu();
            frm.ShowDialog();
        }

        private void mnuDonViTinh_Click(object sender, EventArgs e)
        {
            frmDanhMucDonViTinh frm = new frmDanhMucDonViTinh();
            frm.ShowDialog();
        }

        private void mnuLoaiSanPham_Click(object sender, EventArgs e)
        {
            frmDanhMucLoaiSanPham frm = new frmDanhMucLoaiSanPham();
            frm.ShowDialog();
        }
    }
}
