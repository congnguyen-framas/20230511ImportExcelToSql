using DevExpress.Spreadsheet;
using DevExpress.XtraPrinting.Export;
using DevExpress.XtraBars.Docking2010.Views.Tabbed;
using DevExpress.XtraEditors;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraWaitForm;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Dapper;

namespace ImportExcelToSQL
{
    public partial class frmMain : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        string _conString = "ed3YbBgz3fEdyTkRahthFY5ktQmH2er+ubV7i40QDz2+hAazJukJ2OH+oPgSxn5U/zuDnzqJhjX8Hf4Ll2BNVKIZOFtreteKzvCMIyaoCsM/7GMn4dF5LjLPdPYzP1LfBjbKUEXI6H4ywgTNkXMaC5ggAbdvWL90BlikTavI3oeqrMY08r8wM7Wvttn/kY+l";
        public frmMain()
        {
            InitializeComponent();

            Load += FrmMain_Load;
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            _barButtonItemImport.ItemClick += _barButtonItemImport_ItemClick;
            _barButtonItemUpdate.ItemClick += _barButtonItemUpdate_ItemClick;
        }

        private void _barButtonItemUpdate_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void _barButtonItemImport_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel File|*.xlsx";
                ofd.Title = "Import Excel File";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    SplashScreenManager.ShowForm(this, typeof(WaitForm), true, true, false);
                    SplashScreenManager.Default.SetWaitFormCaption("Vui lòng chờ trong giây lát");
                    SplashScreenManager.Default.SetWaitFormDescription("Loading...");
                    List<ExcelModel> coreData = new List<ExcelModel>();

                    using (Workbook wb = new Workbook())
                    {
                        wb.LoadDocument(ofd.FileName);

                        if (wb.Worksheets.Count > 0)
                        {
                            Worksheet ws = wb.Worksheets[0];

                            //Get ra số hàng và cột có data
                            //index bắt đầu từ 1
                            var _usedRange = ws.GetUsedRange();
                            int _rowUsed = _usedRange.RowCount;
                            int _colUsed = _usedRange.ColumnCount;

                            for (int i = 1; i < _rowUsed + 1; i++)
                            {
                                Row _row = ws.Rows[i];

                                if (_row[$"A{i}"].Value.TextValue == "6199322102-*-D172")
                                {
                                    var a = 100;
                                }

                                if (!string.IsNullOrEmpty(_row[$"A{i}"].Value.TextValue))
                                {
                                    coreData.Add(new ExcelModel()
                                    {
                                        C000 = _row[$"A{i}"].Value.TextValue,
                                        U003 = _row[$"B{i}"].Value.TextValue,
                                    }); ;
                                }
                            }

                            using (var connection = new SqlConnection(EncodeMD5.DecryptString(_conString, "ITFramasBDVN")))
                            {
                                int res = 0;
                                foreach (var item in coreData)
                                {
                                    if (connection.Execute($"Update [VNT86]..[t026] SET U003 = '{item.U003}' WHERE c000 = '{item.C000}'") > 0)
                                    {
                                        res += 1;
                                    }
                                }

                                MessageBox.Show($"Total row affected: {res}.");
                            }
                        }
                        else
                        {
                            XtraMessageBox.Show($"File temple is empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                SplashScreenManager.CloseForm();
            }
        }
    }
}