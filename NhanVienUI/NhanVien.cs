using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace BaiLuanDeTai7
{
    public partial class frmNhanVien : Form
    {
        public frmNhanVien()
        {
            InitializeComponent();
        }
        public void ThemDuLieuSoBan()
        {
            cbbSoBan.Items.Add("Mua Mang Đi");
            cbbSoBan.Items.Add("Bàn 1");
            cbbSoBan.Items.Add("Bàn 2");
            cbbSoBan.Items.Add("Bàn 3");
            cbbSoBan.Items.Add("Bàn 4");
            cbbSoBan.Items.Add("Bàn 5");
            cbbSoBan.Items.Add("Bàn 6");
            cbbSoBan.Items.Add("Bàn 7");
            cbbSoBan.Items.Add("Bàn 8");
            cbbSoBan.Items.Add("Bàn 9");
            cbbSoBan.Items.Add("Bàn 10");
        }
        private void bttThoatPOS_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }
        private void frmNhanVien_FormClosing(object sender, FormClosingEventArgs e)     //Nút "THOÁT POS" - tham khảo bên ngoài
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show("HÃY NHỚ XUẤT BÁO CÁO TRƯỚC KHI BẠN THOÁT POS!" + "\n" + "Bạn có muốn thoát POS không?", "ICONIC COFFEE", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    Environment.Exit(0);
                }
                else
                {
                    e.Cancel = true;
                }
            }
            else
            {
                e.Cancel = true;
            }
        }
        private void bttDangXuat_Click(object sender, EventArgs e) //Nút đăng xuất
        {
            Login oFrm = new Login(this);
            oFrm.Show();
            lbTenNguoiDung.Text = "";            
        }
        private void bttDangNhap_Click(object sender, EventArgs e)  //Nút đăng nhập
        {
            Login oFrm = new Login(this);
            oFrm.Show();
        }
        public string ValuesByProperty              //Lấy tên người dùng từ form login - tham khảo bên ngoài
        {
            get { return this.lbTenNguoiDung.Text; }
            set { this.lbTenNguoiDung.Text = value; }
        }
        private void frmNhanVien_Load(object sender, EventArgs e)
        {
            ThemDuLieuSoBan();
            tabControlMain.Enabled = false;
            bttDangXuat.Enabled = false;
            lbTenNguoiDung.Text = "";
            lbXinChao.Text = "HÃY ĐĂNG NHẬP ĐỂ CÓ THỂ SỬ DỤNG";
        }
        private void lbTenNguoiDung_TextChanged(object sender, EventArgs e)     //Demo tính năng đăng nhập qua các tài khoản có phân quyền
        {
            if (lbTenNguoiDung.Text == "Võ Ngọc Lê Duy")
            {
                UIBlock_TabNhanVien.Visible = false;
                bttDangXuat.Enabled = true;
                bttDangNhap.Enabled = false;
                tabControlMain.Enabled = true;
                lbXinChao.Text = "Xin chào";
                tabPageAdmin.Enabled = true;        // Cho phép dùng các tính năng của quản lý
                lbAdminWarning_TabBaoCao.Text = "";
                UIBlock_TabBaoCao.Visible = false;

            }
            else if (lbTenNguoiDung.Text != "")
            {
                UIBlock_TabNhanVien.Visible = false;
                bttDangXuat.Enabled = true;
                bttDangNhap.Enabled = false;
                tabControlMain.Enabled = true;
                lbXinChao.Text = "Xin chào";
                tabPageAdmin.Enabled = false;       // Không cho phép dùng các tính năng của quản lý
                lbAdminWarning_TabBaoCao.Text = "⚠️ CHỈ QUẢN LÝ MỚI CÓ THỂ SỬ DỤNG CÁC TÍNH NĂNG NÀY";
                UIBlock_TabBaoCao.Visible = true;
            }
            else
            {
                bttDangXuat.Enabled = false;
                bttDangNhap.Enabled = true;
                tabControlMain.Enabled = false;
                lbXinChao.Text = "HÃY ĐĂNG NHẬP ĐỂ CÓ THỂ SỬ DỤNG";
                UIBlock_TabBaoCao.Visible = true;
                UIBlock_TabNhanVien.Visible = true;
            }
        }

        public double Tongtientatcasanpham()    //Tính tổng bill
        {
            int sum = 0;
            int i = 0;
            for (i = 0; i < listViewBill.Items.Count; i++)
            {
                sum = sum + Convert.ToInt32(listViewBill.Items[i].SubItems[2].Text);
            }
            return sum;
        }

        private void TongTien()             //Định dạng textbox tổng tiền để đễ đọc hơn
        {
            if (listViewBill.Items.Count > 0)
            {
                txtTongTien.Text = ((String.Format("{0:n0}", Tongtientatcasanpham())));
            }
            else
            {
                txtTongTien.Text = "";
            }
        }

        private void bttXoa_Click(object sender, EventArgs e)       //Xóa 1 items trong bill
        {
            listViewBill.Items.Clear();
            txtTongTien.Text = "";
            txtSanPham_SuaBill.Text = "";
            txtSoLuong_SuaBill.Text = "";
            txtDONGIA.Text = "";
            txtThanhTien_SuaBill.Text = "";
            txtSoLuong_SuaBill.Enabled = false;
            bttSua_SuaBill.Enabled = false;
            bttHuy_SuaBill.Enabled = false;
            bttXoaSP_SuaBill.Enabled = false;
            cbbSoBan.Text = null;
        }

        private void bttCaPheSua_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Cà Phê Sữa";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttCaPheDen_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Cà Phê Đen";
            int SOLUONG = 1;
            int THANHTIEN = 20000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttSinhToDau_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Sinh Tố Dâu";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttSinhToDuaHau_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Sinh Tố Dưa Hấu";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttCapuchino_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Capuchino";
            int SOLUONG = 1;
            int THANHTIEN = 20000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttLatte_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Latte";
            int SOLUONG = 1;
            int THANHTIEN = 20000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttMachiato_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Machiato";
            int SOLUONG = 1;
            int THANHTIEN = 20000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttSinhToMangCau_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Sinh Tố Mãng Cầu";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttSinhToBo_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Sinh Tố Bơ";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttXoai_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Sinh Tố Xoài";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttTraSuaTruyenThong_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Trà Sữa Truyền Thống";
            int SOLUONG = 1;
            int THANHTIEN = 30000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttMatcha_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Trà Sữa Matcha";
            int SOLUONG = 1;
            int THANHTIEN = 35000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttTraSuaSocola_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Trà Sữa Socola";
            int SOLUONG = 1;
            int THANHTIEN = 35000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttTraSuaOreoCream_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Trà Sữa Oreo Cream";
            int SOLUONG = 1;
            int THANHTIEN = 35000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttTraSuaThaiDo_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Trà Sữa Thái Đỏ";
            int SOLUONG = 1;
            int THANHTIEN = 35000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttTraSuaThaiXanh_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Trà Sữa Thái Xanh";
            int SOLUONG = 1;
            int THANHTIEN = 35000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttTraDao_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Trà Đào";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttTraChanh_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Trà Chanh";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttTraTac_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Trà Tắc";
            int SOLUONG = 1;
            int THANHTIEN = 20000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttNuocEpTao_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Nước Ép Táo";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttNuocEpOi_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Nước Ép Ổi";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttNuocEpDua_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Nước Ép Dứa";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttNuocEpCam_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Nước Ép Cam";
            int SOLUONG = 1;
            int THANHTIEN = 20000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttSinhToDua_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Sinh Tố Dừa";
            int SOLUONG = 1;
            int THANHTIEN = 25000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttHatHuongDuong_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Hạt Hướng Dương";
            int SOLUONG = 1;
            int THANHTIEN = 4000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttTiramisu_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Bánh Tiramisu";
            int SOLUONG = 1;
            int THANHTIEN = 48000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void bttCrepeTranChau_Click(object sender, EventArgs e)
        {
            string SANPHAM = "Bánh Crepe Trân Châu";
            int SOLUONG = 1;
            int THANHTIEN = 54000;

            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    int SOLUONGSAN = Convert.ToInt32(listViewBill.Items[i].SubItems[1].Text);
                    SOLUONGSAN = SOLUONGSAN + 1;
                    THANHTIEN = THANHTIEN * SOLUONGSAN;
                    listViewBill.Items[i].SubItems[1].Text = SOLUONGSAN.ToString();
                    listViewBill.Items[i].SubItems[2].Text = THANHTIEN.ToString();
                    TongTien();
                    return;
                }
            }
            ListViewItem itemCaPheDen = new System.Windows.Forms.ListViewItem(new string[]
                { SANPHAM, SOLUONG.ToString(), THANHTIEN.ToString() });
            listViewBill.Items.AddRange(new System.Windows.Forms.ListViewItem[] { itemCaPheDen });
            TongTien();
        }

        private void listViewBill_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)      //Thao tác chọn sản phẩm để Sửa hoặc Xóa
        {
            txtSoLuong_SuaBill.Enabled = true;
            bttSua_SuaBill.Enabled = true;
            bttHuy_SuaBill.Enabled = true;
            bttXoaSP_SuaBill.Enabled = true;
            var item = e.Item;
            txtSanPham_SuaBill.Text = item.SubItems[0].Text;
            txtSoLuong_SuaBill.Text = item.SubItems[1].Text;
            txtThanhTien_SuaBill.Text = item.SubItems[2].Text;
            int TINHDONGIA = Convert.ToInt32(item.SubItems[2].Text) / Convert.ToInt32(txtSoLuong_SuaBill.Text);
            txtDONGIA.Text = Convert.ToString(TINHDONGIA);
        }

        private void bttSua_SuaBill_Click(object sender, EventArgs e)               //Thực hiện thao tác sửa số lượng sản phẩm
        {
            string SANPHAM = txtSanPham_SuaBill.Text;
            int tongDong = listViewBill.Items.Count;
            for (int i = 0; i < tongDong; i++)
            {
                if (listViewBill.Items[i].Text == SANPHAM)
                {
                    listViewBill.Items[i].SubItems[1].Text = txtSoLuong_SuaBill.Text;
                    listViewBill.Items[i].SubItems[2].Text = txtThanhTien_SuaBill.Text;
                    TongTien();
                    txtSanPham_SuaBill.Text = "";
                    txtSoLuong_SuaBill.Text = "";
                    txtDONGIA.Text = "";
                    txtThanhTien_SuaBill.Text = "";
                    txtSoLuong_SuaBill.Enabled = false;
                    bttSua_SuaBill.Enabled = false;
                    bttHuy_SuaBill.Enabled = false;
                    bttXoaSP_SuaBill.Enabled = false;
                    return;
                }
            }
        }

        private void txtSoLuong_SuaBill_TextChanged_1(object sender, EventArgs e)           //Fix bug textbox Số lượng, cập nhật giá tiền sau khi sửa số lượng theo thời gian thực
        {
            if (txtSoLuong_SuaBill.Text == "")
            {
                txtSoLuong_SuaBill.Text = "0";
            }
            else if (txtDONGIA.Text != "")
            {
                string DONGIA = txtDONGIA.Text;
                int SOLUONG = Convert.ToInt32(txtSoLuong_SuaBill.Text);
                int THANHTIEN = Convert.ToInt32(DONGIA) * (SOLUONG);
                txtThanhTien_SuaBill.Text = Convert.ToString(THANHTIEN);
            }
        }

        private void txtSoLuong_SuaBill_KeyPress_1(object sender, KeyPressEventArgs e)      //Textbox Số lượng chỉ được nhập các kí tự số - tham khảo bên ngoài
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void bttHuy_SuaBill_Click(object sender, EventArgs e)       //Hủy thao tác sửa bill
        {
            txtSanPham_SuaBill.Text = "";
            txtSoLuong_SuaBill.Text = "";
            txtDONGIA.Text = "";
            txtThanhTien_SuaBill.Text = "";
            txtSoLuong_SuaBill.Enabled = false;
            bttSua_SuaBill.Enabled = false;
            bttHuy_SuaBill.Enabled = false;
            bttXoaSP_SuaBill.Enabled = false;
        }
        public int TongDoanhThuBaoCao()             //Tính tổng doanh thu tại trang Báo Cáo (tính dựa trên tổng sp của listViewBaoCao)
        {
            int sum = 0;
            int i = 0;
            for (i = 0; i < listViewBaoCao.Items.Count; i++)
            {
                sum = sum + Convert.ToInt32(listViewBaoCao.Items[i].SubItems[2].Text);
            }
            return sum;
        }

        private void txtTongDoanhThuCuaBaoCao()     //Định dạng số để dễ đọc
        {
            if (listViewBaoCao.Items.Count > 0)
            {
                txtTongDoanhThu_BaoCao.Text = ((String.Format("{0:n0}", TongDoanhThuBaoCao())));
            }
            else
            {
                txtTongDoanhThu_BaoCao.Text = "";
            }
        }
        private void bttThanhToan_Click(object sender, EventArgs e) //Thực hiện các tính năng của thanh toán, lưu bill vào "BÁO CÁO", refresh đơn để nhập đơn mới
        {
            if (cbbSoBan.Text == "")
            {
                MessageBox.Show("Hãy chọn số bàn!", "ICONIC COFFEE");
            }
            else
            {
                dtpkHours.Value = DateTime.Now;
                dtpkDate.Value = DateTime.Now;
                listViewBaoCao.Groups.Add(new ListViewGroup(dtpkHours.Text + " " + dtpkDate.Text + " | " + cbbSoBan.Text + " | Người bán: " + lbTenNguoiDung.Text, HorizontalAlignment.Center));
                int TongBanBaoCao = listViewBaoCao.Groups.Count;
                int TongSanPhamBaoCao = listViewBaoCao.Items.Count;

                ListViewItem itemClone;                                                     //Chuyển bill hiện tại vào lvBaoCao - tham khảo bên ngoài
                ListView.ListViewItemCollection coll = listViewBill.Items;
                foreach (ListViewItem item in coll)
                {
                    itemClone = item.Clone() as ListViewItem;
                    listViewBill.Items.Remove(item);
                    listViewBaoCao.Items.Add(itemClone);

                }

                for (int i = TongSanPhamBaoCao; i < listViewBaoCao.Items.Count; i++)            //Nhóm lại từng bill để quản lý dễ dàng hơn
                {
                    listViewBaoCao.Items[i].Group = listViewBaoCao.Groups[TongBanBaoCao - 1];
                }

                TongDoanhThuBaoCao();
                txtTongDoanhThuCuaBaoCao();
                if (txtTongDoanhThu_BaoCao.Text != "0")
                {
                    int LoiNhuan = 0;
                    int DoanhThu = TongDoanhThuBaoCao();
                    LoiNhuan = DoanhThu / 2;
                    txtTongLoiNhuan_BaoCao.Text = ((String.Format("{0:n0}", LoiNhuan)));
                }

                txtTongTien.Text = "";
                txtSanPham_SuaBill.Text = "";
                txtSoLuong_SuaBill.Text = "";
                txtDONGIA.Text = "";
                txtThanhTien_SuaBill.Text = "";
                txtSoLuong_SuaBill.Enabled = false;
                bttSua_SuaBill.Enabled = false;
                bttHuy_SuaBill.Enabled = false;
                bttXoaSP_SuaBill.Enabled = false;
                cbbSoBan.Text = null;
            }
        }

        private void bttXoaSP_Click(object sender, EventArgs e) //Xóa 1 sản phẩm trước khi ấn nút thanh toán
        {
            string TenSanPham = txtSanPham_SuaBill.Text.Trim();
            foreach (ListViewItem it in listViewBill.Items)
            {
                if (it.SubItems[0].Text == TenSanPham)
                {
                    it.Remove();
                    MessageBox.Show("Đã xóa " + TenSanPham);
                    TongTien();
                    txtSanPham_SuaBill.Text = "";
                    txtSoLuong_SuaBill.Text = "";
                    txtDONGIA.Text = "";
                    txtThanhTien_SuaBill.Text = "";
                    txtSoLuong_SuaBill.Enabled = false;
                    bttSua_SuaBill.Enabled = false;
                    bttHuy_SuaBill.Enabled = false;
                    bttXoaSP_SuaBill.Enabled = false;
                }
            }
        }

        private void bttXuatBaoCao_BaoCao_Click(object sender, EventArgs e) // Xuất báo cáo sang file excel - tham khảo bên ngoài
        {
            int nam, thang, ngay;
            string[] textngaythang = dtpkDate.Text.Split('/');
            nam = Convert.ToInt32(textngaythang[2]);
            thang = Convert.ToInt32(textngaythang[1]);
            ngay = Convert.ToInt32(textngaythang[0]);
            string formatNgayThang = ngay + "-" + thang + "-" + nam;
            using (SaveFileDialog sfd = new SaveFileDialog() { FileName = "Báo cáo " + formatNgayThang, Filter = "Excel Workbook|*.xls", ValidateNames = true })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    dtpkDate.Text = DateTime.Now.ToString();
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    Workbook wb = app.Workbooks.Add(XlSheetType.xlWorksheet);
                    Worksheet ws = (Worksheet)app.ActiveSheet;
                    app.Visible = false;
                    ws.Cells[1, 1] = "SẢN PHẨM";
                    ws.Cells[1, 2] = "SỐ LƯỢNG";
                    ws.Cells[1, 3] = "THÀNH TIỀN";
                    ws.Cells[3, 5] = "DOANH THU:";
                    ws.Cells[3, 6] = txtTongDoanhThu_BaoCao.Text + " đồng";
                    ws.Cells[5, 5] = "LỢI NHUẬN:";
                    ws.Cells[5, 6] = txtTongLoiNhuan_BaoCao.Text + " đồng";
                    ws.Cells[7, 5] = "THỜI GIAN:";
                    ws.Cells[7, 6] = dtpkDate.Text;
                    int i = 2;
                    foreach (ListViewItem item in listViewBaoCao.Items)
                    {
                        ws.Cells[i, 1] = item.SubItems[0].Text;
                        ws.Cells[i, 2] = item.SubItems[1].Text;
                        ws.Cells[i, 3] = item.SubItems[2].Text;
                        i++;
                    }
                    wb.SaveAs(sfd.FileName, XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, (XlSaveAsAccessMode)XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                    app.Quit();
                    MessageBox.Show("ĐÃ XUẤT BÁO CÁO!", "THÀNH CÔNG!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void listViewBaoCao_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e) //Xóa sản phẩm sau khi ấn nút thanh toán ở menu Báo cáo
        {
            if (cbXoaSanPham.Checked == true)
            {
                var item = e.Item;
                item.Remove();
                TongDoanhThuBaoCao();
                txtTongDoanhThuCuaBaoCao();
                if (txtTongDoanhThu_BaoCao.Text != "0")
                {
                    int LoiNhuan = 0;
                    int DoanhThu = TongDoanhThuBaoCao();
                    LoiNhuan = DoanhThu / 2;
                    txtTongLoiNhuan_BaoCao.Text = ((String.Format("{0:n0}", LoiNhuan)));
                }
                return;
            }
        }
    }
}