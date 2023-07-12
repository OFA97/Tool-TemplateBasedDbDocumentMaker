using Dapper;
using DbDocumentMaker.Models;
using DbDocumentMaker.Utility;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using NPOI.OpenXmlFormats.Dml.Diagram;

namespace DbDocumentMaker
{
    public partial class FrmMain : Form
    {
        // Fields
        private DbManager _dbManager;


        // Constructors
        public FrmMain()
        {
            InitializeComponent();

            // init worker
            _dbManager = new DbManager(Config.GetInstance().Content.CurrentConnection.Str);

        }


        // Methods
        private void ShowTables(List<Table> dbTables)
        {
            // reset ui
            clbTables.Items.Clear();
            dgvColumns.DataSource = null;

            // find selected table names for current connection in config
            var checkedTableNames = new List<string>();
            var connName = Config.GetInstance().Content.CurrentConnectionName;
            var docTablePackages = Config.GetInstance().Content.DocTablePackages;
            if (docTablePackages.ContainsKey(connName))
            {
                checkedTableNames = docTablePackages[connName];
            }

            // display table list
            clbTables.Items.Clear();
            foreach (var table in dbTables)
            {
                var isChecked = checkedTableNames.Contains(table.TableName);
                clbTables.Items.Add(table, isChecked);
            }
        }

        private void ShowTableColumns(string tableName)
        {
            dgvColumns.DataSource =
                _dbManager.DbTables.Where(t => t.TableName == tableName).First()
                        .Columns
                        .Select(c => new
                        {
                            Index = c.No,
                            ColumnName = c.ColumnName,
                            Type = c.FullDataType,
                            Nullable = c.IsNullable,
                            PK = c.IsPrimaryKey,
                            FK = c.IsForeignKey,
                            FKReference = c.FkReferencedInfo,
                            Identity = c.IsIdentity,
                            Default = c.Default,
                            Description = c.Description,
                            MS_Description = c.MS_Description
                            //$@"=""IF not exists(SELECT * FROM ::fn_listextendedproperty (NULL, 'user', 'dbo', 'table', '{tableName}', 'column', '{c.ColumnName}'))
                            //BEGIN  
                            // exec sp_addextendedproperty 'MS_Description', '""&J{c.No + 5}&""', 'user', 'dbo', 'table', '{tableName}', 'column', '{c.ColumnName}'
                            //END  
                            //ELSE
                            //BEGIN  
                            // exec sp_updateextendedproperty 'MS_Description', '""&J{c.No + 5}&""', 'user', 'dbo', 'table', '{tableName}', 'column', '{c.ColumnName}'
                            //END"""
                        }).ToList();

            dgvColumns.AutoResizeColumns();
        }

        private List<string> GetCheckedTableNames()
        {
            var checkedTables = new List<string>();
            foreach (var table in clbTables.CheckedItems.Cast<Table>())
            {
                checkedTables.Add(table.TableName);
            }

            return checkedTables;
        }


        // Events
        private void btnReload_Click(object sender, EventArgs e)
        {
            _dbManager.LoadTables();
            ShowTables(_dbManager.DbTables);
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {

            var checkedTableNames = GetCheckedTableNames();
            if (checkedTableNames.Count > 0)
            {
                try
                {
                    // generate db document
                    string templatePath = Config.GetInstance().Content.CurrentDocTemplatePath;
                    string outputPath = Config.GetInstance().Content.OutputDocLocation + Guid.NewGuid().ToString() + ".xlsx";
                    _dbManager.GenerateDocument(checkedTableNames, templatePath, outputPath);

                    // open it
                    Process.Start(outputPath);
                }
                catch (Exception ex)
                {
                    MsgBoxHelper.Error("Failure: " + ex.Message);
                }


            }
            else
            { MsgBoxHelper.Warning("Please select table!"); }


        }

        private void clbTables_SelectedValueChanged(object sender, EventArgs e)
        {
            var tableName = string.Empty;
            if (clbTables.SelectedIndex > -1)
            {
                tableName = (clbTables.Items[clbTables.SelectedIndex] as Table).TableName;
                ShowTableColumns(tableName);
            }
        }

        private void btnCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < clbTables.Items.Count; i++)
                clbTables.SetItemChecked(i, true);
        }

        private void btnUnCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < clbTables.Items.Count; i++)
                clbTables.SetItemChecked(i, false);
        }

        private void btnRememberCheckedTables_Click(object sender, EventArgs e)
        {
            var checkedTables = GetCheckedTableNames();
            if (checkedTables.Count > 0)
            {
                try
                {
                    // save to config
                    var connName = Config.GetInstance().Content.CurrentConnectionName;
                    Config.GetInstance().Content.DocTablePackages[connName] = checkedTables;
                    Config.GetInstance().SaveConfig();

                    MsgBoxHelper.Done();
                }
                catch (Exception ex)
                {
                    MsgBoxHelper.Error("Failure: " + ex.Message);
                }
            }
            else
            { MsgBoxHelper.Warning("Please select table!"); }
        }

        private void btnSetting_Click(object sender, EventArgs e)
        {
            using (var frmSetting = new FrmSetting())
            {
                frmSetting.ShowDialog();
                if (frmSetting.DialogResult == DialogResult.OK)
                {
                    if (frmSetting.IsCurrentConnectionChanged)
                    {
                        // change connect string
                        _dbManager.LoadTables(Config.GetInstance().Content.CurrentConnection.Str);
                        ShowTables(_dbManager.DbTables);
                    }

                }
            }
        }

        private void btnFillBackServer_Click(object sender, EventArgs e)
        {
            ///step1.讀取xlsx
            ///step2.撈現有的資料表,欄位,表名稱對應[sheet],欄位名對應[]
            ///

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog.Title = "選擇 Excel 檔案";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                ConfigContent cc = Config.GetInstance().Content;

                if (File.Exists(Path.Combine(cc.OutputDocLocation, Export.TableFile)))
                {
                    File.Delete(Path.Combine(cc.OutputDocLocation, Export.TableFile));
                }

                if (File.Exists(Path.Combine(cc.OutputDocLocation, Export.ColumnFile)))
                {
                    File.Delete(Path.Combine(cc.OutputDocLocation, Export.ColumnFile));
                }

                NpoiHelper.ReadExcelFile(filePath);
            }
        }
    }
}
