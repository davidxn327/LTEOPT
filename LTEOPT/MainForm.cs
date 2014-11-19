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


namespace LTEOPT
{
    public partial class MainForm : XtraForm
    {
        public string manufacturer = "huawei";
        string data_path = Application.StartupPath + "/data/";

        DataSet qwBase;//全网基准参数
        DataSet qwDataSet;//全网数据

        DataSet qhBase;//切换优先基准
        DataSet qhDataSet;//切换优先数据

        DataSet slBase;//速率优先基准
        DataSet slDataSet;//速率优先数据

        DataSet rlBase;//容量优先基准
        DataSet rlDataSet;//容量优先数据

        public MainForm()
        {
            InitializeComponent();
        }

        public MainForm(string man)
        {
            InitializeComponent();

            manufacturer = man;
            //ImportBase();

            InitTab();

            ImportBase(man);

        }

        void InitTab()
        {
            pages = new Dictionary<string, XtraTabPage>();
            pages.Add("all", xtraTabPage1);
            pages.Add("handoff", xtraTabPage2);
            pages.Add("rate", xtraTabPage3);
            pages.Add("capacity", xtraTabPage4);

            //grids = new Dictionary<string, GridControl>();
            //grids.Add("all", gridControl);
            //grids.Add("handoff", gridControl1);
            //grids.Add("rate", gridControl2);

            views = new Dictionary<string, DevExpress.XtraGrid.Views.Grid.GridView>();
            views.Add("all", gridView1);
            views.Add("handoff", gridView2);
            views.Add("rate", gridView3);
            views.Add("capacity", gridView4);

            combos = new Dictionary<string, ComboBoxEdit>();
            combos.Add("all", comboBoxEdit1);
            combos.Add("handoff", comboBoxEdit2);
            combos.Add("rate", comboBoxEdit3);
            combos.Add("capacity", comboBoxEdit4);

            emptyitems = new Dictionary<string, DevExpress.XtraLayout.EmptySpaceItem>();
            emptyitems.Add("all", emptySpaceItem1);
            emptyitems.Add("handoff", emptySpaceItem2);
            emptyitems.Add("rate", emptySpaceItem3);
            emptyitems.Add("capacity", emptySpaceItem4);
        }

        Dictionary<string, XtraTabPage> pages;
        //Dictionary<string, GridControl> grids;
        Dictionary<string, DevExpress.XtraGrid.Views.Grid.GridView> views;
        Dictionary<string, ComboBoxEdit> combos;
        Dictionary<string, DevExpress.XtraLayout.EmptySpaceItem> emptyitems;

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

            //绑定数据源
            views[page].Tag = ds;
            views[page].Columns.Clear();
            views[page].GridControl.DataSource = ds.Tables[0];
            views[page].RefreshData();

            //显示统计结果
            emptyitems[page].Text = ds.DataSetName;

            //下拉框填充表名
            combos[page].Properties.Items.Clear();
            foreach (DataTable item in ds.Tables)
            {
                combos[page].Properties.Items.Add(item.TableName);
            }
            //string firstTableName = ds.Tables[0].TableName;
            combos[page].SelectedIndex = 0;

            //切换页面
            xtraTabControl1.SelectedTabPage = pages[page];
        }

        void ChangeTable(string page, string tablename)
        {
            DataSet ds = views[page].Tag as DataSet;

            views[page].Columns.Clear();
            views[page].GridControl.DataSource = ds.Tables[tablename];
            views[page].RefreshData();
        }

        DataSet ExcelToDataSet(string excelfile)
        {
            DataSet ds = new DataSet("   ");

            try
            {
                Workbook book = new Workbook();
                book.Open(excelfile);

                int titleRowIndex = 0;
                int firstRowIndex = 1;

                if (is_import_base)
                {
                    if (manufacturer == "huawei")
                    {
                        titleRowIndex = 1;
                        firstRowIndex = 2;
                    }
                    //else { }//默认值
                }
                else
                {
                    if (manufacturer == "huawei")
                    {
                        titleRowIndex = 1;
                        firstRowIndex = 2;
                    }
                    else if (manufacturer == "zte")
                    {
                        titleRowIndex = 0;
                        firstRowIndex = 5;
                    }
                    //else { } //阿朗，默认值
                }

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
                            string colname = cells[titleRowIndex, i].StringValue;
                            if (string.IsNullOrEmpty(colname))
                            {
                                cellCount = i;
                                break;
                            }
                            else
                            {
                                dt.Columns.Add(cells[titleRowIndex, i].StringValue);
                            }
                        }
                        cells.ExportDataTable(dt, firstRowIndex, 0, rowCount - 1, false, true);

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

        bool is_import_base = false;
        void ImportBase(string man)
        {
            is_import_base = true;
            //string man = "huawei";

            string qwfile = "data/";//全网数据
            string tsfile = "data/";//特殊场景

            navBarGroup1.Visible = false;
            navBarGroup2.Visible = false;
            navBarGroup3.Visible = false;
            xtraTabPage2.PageVisible = false;
            xtraTabPage3.PageVisible = false;
            xtraTabPage4.PageVisible = false;

            //每个厂家对应一个特殊场景：切换优先-阿朗；容量优先-中兴；速率优先-华为
            if (man == "huawei")
            {
                qwfile += "huawei_qw.xls";
                tsfile += "huawei_sl.xls";

                qwBase = ExcelToDataSet(qwfile);
                slBase = ExcelToDataSet(tsfile);

                navBarGroup2.Visible = true;
                xtraTabPage3.PageVisible = true;
            }
            else if (man == "zte")
            {
                qwfile += "zte_qw.xlsm";
                tsfile += "zte_rl.xlsm";

                qwBase = ExcelToDataSet(qwfile);
                rlBase = ExcelToDataSet(tsfile);

                navBarGroup3.Visible = true;
                xtraTabPage4.PageVisible = true;
            }
            else //if (man == "allu")
            {
                qwfile += "allu_qw.xlsx";
                tsfile += "allu_qh.xlsx";

                qwBase = ExcelToDataSet(qwfile);
                qhBase = ExcelToDataSet(tsfile);

                navBarGroup1.Visible = true;
                xtraTabPage2.PageVisible = true;
            }

            is_import_base = false;
        }

        //检查DataSet
        DataSet CheckDataSet(DataSet base_ds, DataSet cmp_ds, string single = null)
        {
            DataSet ds = new DataSet();
            int totalRow = 0;
            int totalCell = 0;
            int errRow = 0;
            int errCell = 0;
            foreach (DataTable baseTbl in base_ds.Tables)
            {
                if (baseTbl.Rows.Count < 1)
                    continue;

                DataRow baseRow = baseTbl.Rows[0];
                DataTable cmpTable = cmp_ds.Tables[baseTbl.TableName];
                if (cmpTable != null)
                {
                    DataTable dt = baseTbl.Copy();
                    
                    for (int i = 0; i < cmpTable.Rows.Count; i++)
                    {
                        int errcnt;
                        if (single != null)
                        {
                            errcnt = CheckENodes(cmpTable.Rows[i], baseRow, single);
                        }
                        else
                        {
                            errcnt = CheckAll(cmpTable.Rows[i], baseRow);
                        }

                        if (errcnt > 0)
                        {
                            dt.ImportRow(cmpTable.Rows[i]);
                            errRow++;
                            errCell += errcnt;
                        }

                    }

                    if (dt.Rows.Count>1)
                    {
                        ds.Tables.Add(dt);
                    }
                }

                totalRow += cmpTable.Rows.Count;
                totalCell += totalRow * baseTbl.Columns.Count;
            }
            ds.DataSetName = string.Format("共检查{0}行，有{1}行不匹配；共{2}个字段，有{3}个不匹配。", totalRow, errRow, totalCell, errCell);

            return ds;
        }

        //检查全部数据
        private int CheckAll(DataRow dataRow, DataRow baseRow)
        {
            
            int err = 0;
            foreach (DataColumn col in baseRow.Table.Columns)
            {
                bool flag = false;//表示字段是否有正确
                if (col.ColumnName != "ENBEquipment" && dataRow[col.ColumnName] != null)
                {
                    // TODO: 1.n选1  2.区间  3.整体匹配     1,3可以组合；区间内没有分号
                    string baseStr = baseRow[col.ColumnName].ToString();
                    string cmpStr = dataRow[col.ColumnName].ToString();
                    if (baseStr == cmpStr)
                    {
                        //完全一样就不用比了
                        continue;
                    }
                    if (baseStr.Contains("]"))
                    {
                        string pattern = @"^\[(\d)[,，](\d)\]$";
                        System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(pattern);
                        var match = reg.Match(baseStr);
                        if (match.Success)
                        {
                            double n1 = double.Parse(match.Groups[1].Value);
                            double n2 = double.Parse(match.Groups[2].Value);
                            double n = double.Parse(cmpStr);
                            if (n >= n1 && n <= n2)
                            {
                                flag = true;
                            }
                            else
                            {
                                err++;
                                dataRow.RowError += col.ColumnName + ";";
                            }
                        }
                    }
                    else
                    {
                        string[] options = baseStr.Split(';', '；');
                        for (int i = 0; i < options.Length; i++)
                        {
                            string option = options[i];

                            if ((option == cmpStr)
                                || (option == "空白" && cmpStr == "")
                                || (option == "任意数"))
                            {
                                flag = true;
                                break;
                            }
                            
                        }
                        if (flag == false)
                        {
                            err++;
                            dataRow.RowError += col.ColumnName + ";";
                        }
                    }

                }
            }

            return err;
        }

        // -1：编号不一致；  0：相同；  >0：不相同的个数
        private int CheckENodes(DataRow dataRow, DataRow baseRow, string enodes)
        {
            if (dataRow["ENBEquipment"].ToString() != enodes)
                return -1;//enodes不一样就不用比了

            return CheckAll(dataRow, baseRow);
        }

        //导入全网数据
        private void inboxItem_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel（*.xls;*.xlsx;*.xlsm）|*.xls;*.xlsx;*.xlsm";
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

            DataSet ds = CheckDataSet(qwBase, qwDataSet);

            //显示核查结果
            ShowDataSet("all", ds);

        }

        //单个eNodeB参数核查
        private void navBarItem5_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            // 获取需要核查的eNodeB
            InputForm inputform = new InputForm();
            if (inputform.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            string enodes = comboBoxEdit1.Text;

            if (string.IsNullOrEmpty(enodes))
            {
                XtraMessageBox.Show("请输入eNodes！",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (qwBase == null || qwDataSet == null)
            {
                XtraMessageBox.Show("请导入全网数据！",
                     "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataSet ds = CheckDataSet(qwBase, qwDataSet, enodes);

            //显示核查结果
            ShowDataSet("all", ds);
        }

        //导入切换优先数据
        private void navBarItem1_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel（*.xls;*.xlsx;*.xlsm）|*.xls;*.xlsx;*.xlsm";
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

            DataSet ds = CheckDataSet(qhBase, qhDataSet);

            //显示核查结果
            ShowDataSet("handoff", ds);
        }

        //导入速率优先数据
        private void navBarItem3_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel（*.xls;*.xlsx;*.xlsm）|*.xls;*.xlsx;*.xlsm";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string xls = ofd.FileName;
                slDataSet = ExcelToDataSet(xls);

                //把结果显示在tab页中
                ShowDataSet("rate", slDataSet);
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

            DataSet ds = CheckDataSet(slBase, slDataSet);

            //显示核查结果
            ShowDataSet("rate", ds);
        }

        //导入容量优先数据
        private void navBarItem6_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel（*.xls;*.xlsx;*.xlsm）|*.xls;*.xlsx;*.xlsm";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string xls = ofd.FileName;
                rlDataSet = ExcelToDataSet(xls);

                //把结果显示在tab页中
                ShowDataSet("capacity", rlDataSet);
            }
        }

        //容量优先核查
        private void navBarItem7_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            if (rlDataSet == null || rlBase == null)
            {
                XtraMessageBox.Show("请导入容量优先场景数据！",
                     "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataSet ds = CheckDataSet(rlBase, rlDataSet);

            //显示核查结果
            ShowDataSet("capacity", ds);
        }

        //全网列表
        private void comboBoxEdit1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cn = comboBoxEdit1.Text;
            ChangeTable("all", cn);

            //DataSet ds = views["all"].Tag as DataSet;
            //views["all"].Columns.Clear();
            //views["all"].GridControl.DataSource = ds.Tables[cn];
            //views["all"].RefreshData();
        }

        //切换优先列表
        private void comboBoxEdit2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cn = comboBoxEdit2.Text;
            ChangeTable("handoff", cn);
        }

        //速率优先列表
        private void comboBoxEdit3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cn = comboBoxEdit3.Text;
            ChangeTable("rate", cn);
        }

        //容量优先列表
        private void comboBoxEdit4_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cn = comboBoxEdit4.Text;
            ChangeTable("capacity", cn);
        }

        private void gridView1_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            //第一行  
            if (e.RowHandle == 0)
            {
                e.Appearance.BackColor = Color.DeepSkyBlue;
                //e.Appearance.BackColor2 = Color.LightCyan;
            }
            else
            {
                var gv = ((DevExpress.XtraGrid.Views.Base.ColumnView)(sender));
                DataRow dr = gv.GetDataRow(e.RowHandle);
                if (dr.RowError.Contains(e.Column.FieldName))//
                {
                    e.Appearance.BackColor = Color.Red;
                }
            }
            ////单元格  
            //if (e.RowHandle == 0 && e.Column.ColumnHandle == 0)
            //{
            //    e.Appearance.BackColor = Color.DeepSkyBlue;
            //    e.Appearance.BackColor2 = Color.LightCyan;
            //} 
        }

        void ExportToExcel(string page)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel（*.xls）|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                views[page].ExportToXls(sfd.FileName);
                XtraMessageBox.Show("导出成功！");
            }
        }

        void ExportAllToExcel(string page)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DataSet ds = views[page].Tag as DataSet;
                GridControl gc = new GridControl();
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    DataTable dt = ds.Tables[i];
                    string file = fbd.SelectedPath + "/" + dt.TableName + ".xls";

                    gc.DataSource = dt;
                    gc.ExportToXls(file);

                }
                XtraMessageBox.Show("导出成功！");
            }
        }

        //基本参数导出Excel
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            ExportToExcel("all");
        }

        //基本参数导出所有
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            ExportAllToExcel("all");
        }

        //切换优先导出Excel
        private void simpleButton5_Click(object sender, EventArgs e)
        {
            ExportToExcel("handoff");
        }

        //切换优先导出所有
        private void simpleButton6_Click(object sender, EventArgs e)
        {
            ExportAllToExcel("handoff");
        }

        //速率优先导出Excel
        private void simpleButton3_Click(object sender, EventArgs e)
        {
            ExportToExcel("rate");
        }

        //速率优先导出所有
        private void simpleButton4_Click(object sender, EventArgs e)
        {
            ExportAllToExcel("rate");
        }

        //容量优先导出Excel
        private void simpleButton7_Click(object sender, EventArgs e)
        {
            ExportToExcel("capacity");
        }

        //容量优先导出所有
        private void simpleButton8_Click(object sender, EventArgs e)
        {
            ExportAllToExcel("capacity");
        }



       

    }
}