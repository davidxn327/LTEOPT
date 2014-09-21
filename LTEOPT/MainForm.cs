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

            if (!System.IO.Directory.Exists(dataDir))
            {
                System.IO.Directory.CreateDirectory(dataDir);
            }

            baseDataSet = new DataSet();
            if (System.IO.File.Exists(baseFile))
                baseDataSet.ReadXml(baseFile, XmlReadMode.Auto);

            specDataSet = new DataSet();
            if (System.IO.File.Exists(specFile))
                specDataSet.ReadXml(specFile, XmlReadMode.Auto);
        }

        static string dataDir = Application.StartupPath + "/data/";
        static string baseFile = dataDir + "base.dat";
        static string specFile = dataDir + "spec.dat";

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

        private bool CheckAll(DataRow dataRow, DataRow baseRow)
        {
            foreach (DataColumn col in baseRow.Table.Columns)
            {
                if (dataRow[col.ColumnName] != null)
                {
                    // TODO:
                    if (baseRow[col.ColumnName] != dataRow[col.ColumnName])
                        return false;
                }
            }

            return true;
        }

        private bool CheckENodes(DataRow dataRow, DataRow baseRow, string enodes)
        {
            if(dataRow["xxxxxxxxxxx"].ToString() != enodes)
                return true;//enodes不一样就不用比了
            foreach (DataColumn col in baseRow.Table.Columns)
            {
                if (dataRow[col.ColumnName] != null)
                {
                    // TODO:
                    if (baseRow[col.ColumnName] != dataRow[col.ColumnName])
                        return false;
                }
            }

            return true;
        }

        //全网数据导入
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog()==DialogResult.OK)
            {
                baseDataSet = ExcelToDataSet(ofd.FileName);
            }
            baseDataSet.WriteXml(baseFile, XmlWriteMode.WriteSchema);
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
            sr.DataSource = baseDataSet;
            sr.ShowDialog();
        }

        //导入全网数据
        private void simpleButton9_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                providerDataSet = ExcelToDataSet(ofd.FileName);
            }
        }

        //全网数据检查
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (baseDataSet == null || providerDataSet == null)
            {
                XtraMessageBox.Show("请导入全网数据！",
                     "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataSet ds = new DataSet();
            foreach (DataTable baseTbl in baseDataSet.Tables)
            {
                if (baseTbl.Rows.Count < 1)
                    continue;
                DataRow baseRow = baseTbl.Rows[0];

                DataTable cmpTable = providerDataSet.Tables[baseTbl.TableName];
                if (cmpTable != null)
                {
                    DataTable dt = cmpTable.Clone();
                    ds.Tables.Add(dt);
                    for (int i = 0; i < cmpTable.Rows.Count; i++)
                    {
                        if (!CheckAll(cmpTable.Rows[i], baseRow))
                            dt.ImportRow(cmpTable.Rows[i]);
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
            if (string.IsNullOrEmpty(comboBoxEdit1.Text))
            {
                XtraMessageBox.Show("请输入eNodes！",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string enodes = comboBoxEdit1.Text;

            if (baseDataSet == null || providerDataSet == null)
            {
                XtraMessageBox.Show("请导入全网数据！",
                     "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataSet ds = new DataSet();
            foreach (DataTable baseTbl in baseDataSet.Tables)
            {
                if (baseTbl.Rows.Count < 1)
                    continue;
                DataRow baseRow = baseTbl.Rows[0];

                DataTable cmpTable = providerDataSet.Tables[baseTbl.TableName];
                if (cmpTable != null)
                {
                    DataTable dt = cmpTable.Clone();
                    ds.Tables.Add(dt);
                    for (int i = 0; i < cmpTable.Rows.Count; i++)
                    {
                        if (!CheckENodes(cmpTable.Rows[i], baseRow, enodes))
                            dt.ImportRow(cmpTable.Rows[i]);
                    }
                }

            }

            ShowResult sr = new ShowResult();
            sr.DataSource = ds;
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