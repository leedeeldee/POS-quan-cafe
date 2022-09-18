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

namespace BaiLuanDeTai7
{
    public partial class Login : Form
    {

        public Login()
        {
            InitializeComponent();
        }
        frmNhanVien CalledFrom;
        public Login(frmNhanVien viaParameter) : this()
        {
            CalledFrom = viaParameter;  
        }
        private void LoginInfo()
        {
            string TenDangNhap = txtTenDangNhap.Text;
            string MatKhau = txtMatKhau.Text;
            if (TenDangNhap == "VNLD" &&
                MatKhau == "123")           // Võ Ngọc Lê Duy - 215120125
            {
                CalledFrom.ValuesByProperty = "Võ Ngọc Lê Duy";
                string GettingBack = CalledFrom.ValuesByProperty;
                this.Close();
            }
            else if (TenDangNhap == "LBK" &&
                     MatKhau == "123")      // Lê Bá Kông - 215120184
            {
                CalledFrom.ValuesByProperty = "Lê Bá Kông";
                string GettingBack = CalledFrom.ValuesByProperty;
                this.Close();
            }
            else if (TenDangNhap == "NVTT" &&
                     MatKhau == "123")      // Nguyễn Văn Thanh Thuyên - 215122379
            {
                CalledFrom.ValuesByProperty = "Nguyễn Văn Thanh Thuyên";
                string GettingBack = CalledFrom.ValuesByProperty;
                this.Close();
            }
            else if (TenDangNhap == "TTT" &&
                     MatKhau == "123")      // Trần Thủy Tiên - 215120326
            {
                CalledFrom.ValuesByProperty = "Trần Thủy Tiên";
                string GettingBack = CalledFrom.ValuesByProperty;
                this.Close();
            }
            else
            {
                MessageBox.Show("Lỗi! Hãy xem lại thông tin đăng nhập!");
            }
        }

        private void bttLogin_Click(object sender, EventArgs e)
        {
            LoginInfo();
        }

        private void bttCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
// DEMO TÍNH NĂNG ĐĂNG NHẬP, CHÚNG EM LÀM THEO CÁCH NÀY VÌ CHÚNG EM CHƯA HỌC CƠ SỎ DỮ LIỆU