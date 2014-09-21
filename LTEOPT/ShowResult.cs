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
    public partial class ShowResult : DevExpress.XtraEditors.XtraForm
    {
        

        public ShowResult()
        {
            InitializeComponent();
            
        }

        DataSet ds;
        public DataSet DataSource
        {
            get { return ds; }
            set { ds = value; }
        }

        public void SetTable(DataTable dt)
        {
            gridControl1.DataSource = dt;
        }
    }
}