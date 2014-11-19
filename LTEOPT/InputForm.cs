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
    public partial class InputForm : DevExpress.XtraEditors.XtraForm
    {
        public InputForm()
        {
            InitializeComponent();
        }

        public string enode;

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            enode = textEdit1.Text;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        }
    }
}