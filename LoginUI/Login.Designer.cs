namespace BaiLuanDeTai7
{
    partial class Login
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.txtTenDangNhap = new MaterialSkin.Controls.MaterialSingleLineTextField();
            this.txtMatKhau = new MaterialSkin.Controls.MaterialSingleLineTextField();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.bttLogin = new System.Windows.Forms.Button();
            this.bttCancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(148)))), ((int)(((byte)(121)))));
            this.pictureBox1.Image = global::BaiLuanDeTai7.Properties.Resources.Logo_xóa_nền;
            this.pictureBox1.Location = new System.Drawing.Point(-3, -2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(648, 723);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // txtTenDangNhap
            // 
            this.txtTenDangNhap.BackColor = System.Drawing.Color.White;
            this.txtTenDangNhap.Depth = 0;
            this.txtTenDangNhap.Font = new System.Drawing.Font("Roboto", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTenDangNhap.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(148)))), ((int)(((byte)(121)))));
            this.txtTenDangNhap.Hint = "";
            this.txtTenDangNhap.Location = new System.Drawing.Point(814, 290);
            this.txtTenDangNhap.MaxLength = 32767;
            this.txtTenDangNhap.MouseState = MaterialSkin.MouseState.HOVER;
            this.txtTenDangNhap.Name = "txtTenDangNhap";
            this.txtTenDangNhap.PasswordChar = '\0';
            this.txtTenDangNhap.SelectedText = "";
            this.txtTenDangNhap.SelectionLength = 0;
            this.txtTenDangNhap.SelectionStart = 0;
            this.txtTenDangNhap.Size = new System.Drawing.Size(252, 23);
            this.txtTenDangNhap.TabIndex = 1;
            this.txtTenDangNhap.TabStop = false;
            this.txtTenDangNhap.UseSystemPasswordChar = false;
            // 
            // txtMatKhau
            // 
            this.txtMatKhau.Depth = 0;
            this.txtMatKhau.Font = new System.Drawing.Font("Roboto", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMatKhau.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(148)))), ((int)(((byte)(121)))));
            this.txtMatKhau.Hint = "";
            this.txtMatKhau.Location = new System.Drawing.Point(814, 363);
            this.txtMatKhau.MaxLength = 32767;
            this.txtMatKhau.MouseState = MaterialSkin.MouseState.HOVER;
            this.txtMatKhau.Name = "txtMatKhau";
            this.txtMatKhau.PasswordChar = '●';
            this.txtMatKhau.SelectedText = "";
            this.txtMatKhau.SelectionLength = 0;
            this.txtMatKhau.SelectionStart = 0;
            this.txtMatKhau.Size = new System.Drawing.Size(252, 23);
            this.txtMatKhau.TabIndex = 2;
            this.txtMatKhau.TabStop = false;
            this.txtMatKhau.UseSystemPasswordChar = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(148)))), ((int)(((byte)(121)))));
            this.label1.Location = new System.Drawing.Point(809, 262);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 19);
            this.label1.TabIndex = 5;
            this.label1.Text = "TÊN ĐĂNG NHẬP";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(148)))), ((int)(((byte)(121)))));
            this.label2.Location = new System.Drawing.Point(809, 335);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(90, 19);
            this.label2.TabIndex = 5;
            this.label2.Text = "MẬT KHẨU";
            // 
            // bttLogin
            // 
            this.bttLogin.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(148)))), ((int)(((byte)(121)))));
            this.bttLogin.Font = new System.Drawing.Font("Roboto", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttLogin.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.bttLogin.Location = new System.Drawing.Point(798, 404);
            this.bttLogin.Name = "bttLogin";
            this.bttLogin.Size = new System.Drawing.Size(160, 48);
            this.bttLogin.TabIndex = 6;
            this.bttLogin.Text = "ĐĂNG NHẬP";
            this.bttLogin.UseVisualStyleBackColor = false;
            this.bttLogin.Click += new System.EventHandler(this.bttLogin_Click);
            // 
            // bttCancel
            // 
            this.bttCancel.Font = new System.Drawing.Font("Roboto", 15.75F, System.Drawing.FontStyle.Bold);
            this.bttCancel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(148)))), ((int)(((byte)(121)))));
            this.bttCancel.Location = new System.Drawing.Point(979, 404);
            this.bttCancel.Name = "bttCancel";
            this.bttCancel.Size = new System.Drawing.Size(160, 48);
            this.bttCancel.TabIndex = 7;
            this.bttCancel.Text = "HỦY";
            this.bttCancel.UseVisualStyleBackColor = true;
            this.bttCancel.Click += new System.EventHandler(this.bttCancel_Click);
            // 
            // Login
            // 
            this.AcceptButton = this.bttLogin;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1285, 718);
            this.Controls.Add(this.bttCancel);
            this.Controls.Add(this.bttLogin);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtMatKhau);
            this.Controls.Add(this.txtTenDangNhap);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.Name = "Login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ICONIC COFFEE - POS Login";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.PictureBox pictureBox1;
        private MaterialSkin.Controls.MaterialSingleLineTextField txtTenDangNhap;
        private MaterialSkin.Controls.MaterialSingleLineTextField txtMatKhau;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button bttLogin;
        private System.Windows.Forms.Button bttCancel;
    }
}

