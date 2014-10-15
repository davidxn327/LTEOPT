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
        }

        //登录
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (comboBoxEdit1.Text == "中兴")
            {
                manufacturer = "zte";
            }
            else if (comboBoxEdit1.Text == "朗讯")
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