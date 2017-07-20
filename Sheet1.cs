using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using DataAccess;
using VSTOCommon;

namespace HospitalTargetCheckSFA
{
    public partial class Sheet1
    {
        const string START_COLUMN = "B";
        const string END_COLUMN = "B";
        const int TITLE_ROW_COUNT = 10;
        string DATA_ROW = START_COLUMN + "{0}:" + END_COLUMN + "{1}";
        private string[] columns = { "RegionCode" };
        public int DataCount
        {
            get
            {
                return Convert.ToInt32(ExcelUtils.GetPropertyValue("SHEET1_DATACOUNT", this));
            }
            set
            {
                ExcelUtils.AddProperty("SHEET1_DATACOUNT", value.ToString(), this);
            }
        }
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet1_Startup);
            this.Shutdown += new System.EventHandler(Sheet1_Shutdown);
        }

        #endregion

        #region DashBoard

        public void Init(DataTable dt)
        {
            ExcelUtils.UnLockSheet(this);

            LoadData(dt);

            //SetFormula();

            ResetFormat();

            //ExcelUtils.LockSheet(this);
        }

        public void LoadData(DataTable dt)
        {
            //DataTable mydt = dt.DefaultView.ToTable(false, columns);
            string startCell, endCell;

            DataCount = dt.Rows.Count;

            //数据填充
            if (DataCount > 0)
            {
                startCell = START_COLUMN + (TITLE_ROW_COUNT + 1).ToString();
                endCell = END_COLUMN + (TITLE_ROW_COUNT + DataCount).ToString();
                VSTOCommon.ExcelUtils.PrintData(this, dt, startCell, endCell);
            }

            List<string> ls = new List<string>();
            //string lastYear = (TNTPlanCheck.Year - 1).ToString();
            //string q1 = TNTPlanCheck.HY1 == 1 ? "Q1" : "Q3";
            //string q2 = TNTPlanCheck.HY1 == 1 ? "Q2" : "Q4";
            //ls.Add(lastYear);
            //ls.Add(q1);
            //ls.Add(q2);
            //VSTOCommon.ExcelUtils.PrintData(this, ls, "G50", "G52");
            ls.Add(Globals.ThisWorkbook.TNTPlanCheck.Year.ToString());
            ls.Add(Globals.ThisWorkbook.TNTPlanCheck.HY1.ToString());
            ls.Add(Globals.ThisWorkbook.TNTPlanCheck.HY2.ToString());
            ls.Add(Globals.ThisWorkbook.TNTPlanCheck.Q1.ToString());
            ls.Add(Globals.ThisWorkbook.TNTPlanCheck.Q2.ToString());
            ls.Add(Globals.ThisWorkbook.TNTPlanCheck.Q3.ToString());
            ls.Add(Globals.ThisWorkbook.TNTPlanCheck.Q4.ToString());
            VSTOCommon.ExcelUtils.PrintData(this, ls, "G1", "G7");
        }

        public void ResetFormat()
        {
            Excel.Range rg;
            //格式
            rg = this.Range["B11","B30"];
            VSTOCommon.ExcelUtils.SetCellsColor_ReadOnely(rg, true);
        }

        public void ClearData()
        {
            try
            {
                ExcelUtils.UnLockSheet(this);
                if (DataCount > 0)
                {
                    Excel.Range rg;
                    rg = this.Range[string.Format(DATA_ROW, TITLE_ROW_COUNT + 1, TITLE_ROW_COUNT + DataCount)];
                    rg.Clear();
                    VSTOCommon.ExcelUtils.SetCellsColor_ReadOnely(rg, true);
                }
                ExcelUtils.LockSheet(this);
            }
            catch (Exception ex)
            {
                LogHelper.WriteError("", ex);
            }
        }

        #endregion
    }
}
