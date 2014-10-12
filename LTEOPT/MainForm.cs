using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.Skins;
using DevExpress.LookAndFeel;
using DevExpress.UserSkins;
using DevExpress.XtraEditors;
using Aspose.Cells;
using DevExpress.XtraTab;
using DevExpress.XtraGrid;


namespace WindowsApplication1
{
    public partial class MainForm : XtraForm
    {
        string data_path = Application.StartupPath + "/data/";

        DataSet qwBase;//全网基准参数
        DataSet qwDataSet;//全网数据

        DataSet qhBase;//切换优先基准
        DataSet qhDataSet;//切换优先数据

        DataSet slBase;//速率优先基准
        DataSet slDataSet;//速率优先数据

        public MainForm()
        {
            InitializeComponent();

            //ImportBase();

            InitTab();

            qwBase = new DataSet();
            qwBase.ReadXml(data_path + "huawei_qwBase.dat");

            qhBase = new DataSet();
            qhBase.ReadXml(data_path + "huawei_qhBase.dat");

        }
        void InitTab()
        {
            pages = new Dictionary<string, XtraTabPage>();
            pages.Add("all", xtraTabPage1);
            pages.Add("handoff", xtraTabPage2);
            pages.Add("rate", xtraTabPage3);

            grids = new Dictionary<string, GridControl>();
            grids.Add("all", gridControl);
            grids.Add("handoff", gridControl1);
            grids.Add("rate", gridControl2);

            combos = new Dictionary<string, ComboBoxEdit>();
            combos.Add("all", comboBoxEdit1);
            combos.Add("handoff", comboBoxEdit2);
            combos.Add("rate", comboBoxEdit3);
        }

        Dictionary<string, XtraTabPage> pages;
        Dictionary<string, GridControl> grids;
        //Dictionary<string, DataGridView> views;
        Dictionary<string, ComboBoxEdit> combos;
        //切换Tab页面
        void SwitchPage(string page)
        {
            // TODO：
            xtraTabControl1.SelectedTabPage = pages[page];
        }

        //tab页面显示数据
        void ShowDataSet(string page, DataSet ds)
        {
            if (ds == null || ds.Tables.Count == 0)
                return;

            //下拉框填充表名
            combos[page].Properties.Items.Clear();
            foreach (DataTable item in ds.Tables)
            {
                combos[page].Properties.Items.Add(item.TableName);
            }
            string firstTableName = ds.Tables[0].TableName;

            //绑定数据源
            grids[page].Tag = ds;

            //grids[page].DataSource = null;
            //grids[page].RefreshDataSource();

            (grids[page].MainView as DevExpress.XtraGrid.Views.Grid.GridView).Columns.Clear(); 
            grids[page].DataSource = ds.Tables[0];
            grids[page].MainView.RefreshData();

            //切换页面
            xtraTabControl1.SelectedTabPage = pages[page];
        }

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

        void ImportBase()
        {
            string man = "huawei";

            qwBase = ExcelToDataSet(data_path + "LTE参数_20140702-基准.xlsx");
            qwBase.WriteXml(data_path +man+ "_qwBase.dat", XmlWriteMode.WriteSchema);

            qhBase = ExcelToDataSet(data_path + "异厂家切换相关参数及功能开关-ALU-基准.xlsx");
            qhBase.WriteXml(data_path + man + "_qhBase.dat", XmlWriteMode.WriteSchema);

            //slBase = ExcelToDataSet(data_path +"速率基准");
            //slBase.WriteXml(data_path + man + "_slBase.dat", XmlWriteMode.WriteSchema);
        }

        //检查全部数据
        private bool CheckAll(DataRow dataRow, DataRow baseRow)
        {
            bool flag = true;
            foreach (DataColumn col in baseRow.Table.Columns)
            {
                if ( col.ColumnName!="ENBEquipment" && dataRow[col.ColumnName] != null)
                {
                    // TODO: 1.n选1  2.区间  3.整体匹配     1,3可以组合；区间内没有分号
                    if (baseRow[col.ColumnName] != dataRow[col.ColumnName])
                    {
                        dataRow.RowError += col.ColumnName+";";
                        flag = false;
                    }
                }
            }

            return flag;
        }

        private bool CheckENodes(DataRow dataRow, DataRow baseRow, string enodes)
        {
            if (dataRow["xxxxxxxxxxx"].ToString() != enodes)
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

        //导入全网数据
        private void inboxItem_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string xls = ofd.FileName;
                qwDataSet = ExcelToDataSet(xls);

                //把结果显示在tab页中
                ShowDataSet("all", qwDataSet);
            }
        }

        //全网参数核查
        private void outboxItem_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            if (qwDataSet == null || qwBase ==null)
            {
                XtraMessageBox.Show("请导入全网数据！",
                     "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataSet ds = new DataSet();
            foreach (DataTable baseTbl in qwBase.Tables)
            {
                if (baseTbl.Rows.Count < 1)
                    continue;
                DataRow baseRow = baseTbl.Rows[0];

                DataTable cmpTable = qwDataSet.Tables[baseTbl.TableName];
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

            //显示核查结果
            ShowDataSet("all", ds);

        }

        //单个eNodeB参数核查
        private void navBarItem5_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            // 获取需要核查的eNodeB
            if (string.IsNullOrEmpty(comboBoxEdit1.Text))
            {
                XtraMessageBox.Show("请输入eNodes！",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string enodes = comboBoxEdit1.Text;


            if (qwBase == null || qwDataSet == null)
            {
                XtraMessageBox.Show("请导入全网数据！",
                     "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataSet ds = new DataSet();
            foreach (DataTable baseTbl in qwBase.Tables)
            {
                if (baseTbl.Rows.Count < 1)
                    continue;
                DataRow baseRow = baseTbl.Rows[0];

                DataTable cmpTable = qwDataSet.Tables[baseTbl.TableName];
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

            //显示核查结果
            ShowDataSet("all", ds);
        }

        //导入切换优先数据
        private void navBarItem1_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string xls = ofd.FileName;
                qhDataSet = ExcelToDataSet(xls);

                //把结果显示在tab页中
                ShowDataSet("handoff", qhDataSet);
            }
        }

        //切换优先核查
        private void navBarItem2_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            if (qhDataSet == null || qhBase == null)
            {
                XtraMessageBox.Show("请导入切换优先场景数据！",
                     "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataSet ds = new DataSet();
            foreach (DataTable baseTbl in qhBase.Tables)
            {
                if (baseTbl.Rows.Count < 1)
                    continue;
                DataRow baseRow = baseTbl.Rows[0];

                DataTable cmpTable = qhDataSet.Tables[baseTbl.TableName];
                if (cmpTable != null)
                {
                    DataTable dt = baseTbl.Clone();
                    ds.Tables.Add(dt);
                    for (int i = 0; i < cmpTable.Rows.Count; i++)
                    {
                        if (!CheckAll(cmpTable.Rows[i], baseRow))
                            dt.ImportRow(cmpTable.Rows[i]);
                    }
                }

            }

            //显示核查结果
            ShowDataSet("handoff", ds);
        }

        //导入速率优先数据
        private void navBarItem3_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string xls = ofd.FileName;
                slDataSet = ExcelToDataSet(xls);

                //把结果显示在tab页中
                ShowDataSet("rate", qhDataSet);
            }
        }

        //速率优先核查
        private void navBarItem4_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            if (slDataSet == null || slBase == null)
            {
                XtraMessageBox.Show("请导入速率优先场景数据！",
                     "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataSet ds = new DataSet();
            foreach (DataTable baseTbl in slBase.Tables)
            {
                if (baseTbl.Rows.Count < 1)
                    continue;
                DataRow baseRow = baseTbl.Rows[0];

                DataTable cmpTable = slDataSet.Tables[baseTbl.TableName];
                if (cmpTable != null)
                {
                    DataTable dt = baseTbl.Clone();
                    ds.Tables.Add(dt);
                    for (int i = 0; i < cmpTable.Rows.Count; i++)
                    {
                        if (!CheckAll(cmpTable.Rows[i], baseRow))
                            dt.ImportRow(cmpTable.Rows[i]);
                    }
                }

            }

            //显示核查结果
            ShowDataSet("rate", ds);
        }

        //全网列表
        private void comboBoxEdit1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cn = comboBoxEdit1.Text;
            DataSet ds = grids["all"].Tag as DataSet;

            (grids["all"].MainView as DevExpress.XtraGrid.Views.Grid.GridView).Columns.Clear(); 
            grids["all"].DataSource = ds.Tables[cn];
            grids["all"].RefreshDataSource();
            grids["all"].MainView.RefreshData();
        }

    }
}