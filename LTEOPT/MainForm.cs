using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Aspose.Cells;

namespace LTEOPT
{
    public partial class MainForm : DevExpress.XtraEditors.XtraForm
    {
        public MainForm()
        {
            InitializeComponent();
        }

        DataSet baseDataSet;
        DataSet providerDataSet;

        DataSet specDataSet;
        DataSet switchDataSet;
        DataSet rateDataSet;

        DataSet ExcelToDataSet(string excelfile)
        {
            DataSet ds = new DataSet();

            try
            {
                Workbook book = new Workbook();
                book.Open(excelfile);

                foreach (Worksheet sheet in book.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int rowCount = cells.MaxDataRow + 1;
                    int cellCount = cells.MaxDataColumn + 1;

                    if (rowCount > 1)
                    {
                        DataTable dt = new DataTable(sheet.Name);

                        for (int i = 0; i < cellCount; i++)
                        {
                            dt.Columns.Add(cells[0, i].StringValue);
                        }
                        cells.ExportDataTable(dt, 1, 0, rowCount - 1, false, true);

                        //cells.ExportDataTable(dt, 0, 0, rowCount, true, true);

                        ds.Tables.Add(dt);


                    }
                    Application.DoEvents();

                }
            }
            catch (Exception e)
            {
                XtraMessageBox.Show(e.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            return ds;
        }

        //全网数据导入
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog()==DialogResult.OK)
            {
                baseDataSet = ExcelToDataSet(ofd.FileName);
            }
        }

        //查看基础数据
        private void simpleButton4_Click(object sender, EventArgs e)
        {
            if (baseDataSet == null)
            {
                XtraMessageBox.Show("请先导入全网基础数据！",
                    "警告",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return;
            }

            ShowResult sr = new ShowResult();
            sr.SetTable(baseDataSet.Tables["MmeAccess"]);
            sr.ShowDialog();
        }

        //全网数据检查
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            foreach (DataTable baseTbl in baseDataSet.Tables)
            {
                DataTable table = providerDataSet.Tables[baseTbl.TableName];
                if (table != null)
                {
                    DataRow[] rows = table.Select("");

                    DataTable dt = table.Clone();
                    ds.Tables.Add(dt);
                    for (int i = 0; i < rows.Length; i++)
                    {
                        dt.ImportRow(rows[i]);
                    }
                }

            }

            ShowResult sr = new ShowResult();
            sr.DataSource = ds;
            sr.ShowDialog();
        }

        //单个eNodes参数检查
        private void simpleButton3_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            foreach (DataTable baseTbl in baseDataSet.Tables)
            {
                DataTable table = providerDataSet.Tables[baseTbl.TableName];
                if (table != null)
                {
                    DataRow[] rows = table.Select("");

                    DataTable dt = table.Clone();
                    ds.Tables.Add(dt);
                    for (int i = 0; i < rows.Length; i++)
                    {
                        dt.ImportRow(rows[i]);
                    }
                }

            }

            ShowResult sr = new ShowResult();
            sr.DataSource = ds;
            sr.Text = "";
            sr.ShowDialog();
        }

        //特殊场景导入
        private void simpleButton5_Click(object sender, EventArgs e)
        {

        }

        //特殊场景基础数据查看
        private void simpleButton6_Click(object sender, EventArgs e)
        {

        }

        //切换优先场景
        private void simpleButton7_Click(object sender, EventArgs e)
        {

        }

        //速率优先场景
        private void simpleButton8_Click(object sender, EventArgs e)
        {

        }



    }
}