using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace LTEOPT
{
    public partial class LoginForm : DevExpress.XtraEditors.XtraForm
    {
        public string manufacturer = "huawei";

        public LoginForm()
        {
            InitializeComponent();

            this.Text = "LTE调度参数优化工具";
        }

        //登录
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (comboBoxEdit1.Text != "cszhangm" || textEdit1.Text != "123456")
            {
                XtraMessageBox.Show("用户名或密码错误！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (comboBoxEdit2.Text == "中兴")
            {
                manufacturer = "zte";
            }
            else if (comboBoxEdit2.Text == "阿朗")
            {
                manufacturer = "allu";
            }
            //else
            //{
            //    manufacturer = "huawei";
            //}

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        //退出，取消
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        }
    }
}