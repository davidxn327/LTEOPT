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
            set { ds = value; ConfigDS(ds); }
        }

        void ConfigDS(DataSet ds)
        {
            if(ds==null || ds.Tables.Count==0)
                return;
            foreach (DataTable item in ds.Tables)
            {
                repositoryItemComboBox1.Items.Add(item.TableName);
            }
            string firstTableName = ds.Tables[0].TableName;
            barEditItem1.EditValue = firstTableName;
            SetTable(ds.Tables[0]);
        }

        public void SetTable(DataTable dt)
        {
            gridControl1.DataSource = dt;
        }

        private void repositoryItemComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string tableName = ((DevExpress.XtraEditors.ComboBoxEdit)sender).SelectedItem.ToString();
            if (string.IsNullOrEmpty(tableName))
                return;
            DataTable dt = ds.Tables[tableName];
            if (dt != null)
                SetTable(dt);
        }
    }
}