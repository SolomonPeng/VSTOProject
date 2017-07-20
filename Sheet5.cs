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
    public partial class Sheet5
    {
        const string START_COLUMN = "A";
        const string END_COLUMN = "H";
        const int TITLE_ROW_COUNT = 1;
        string DATA_ROW = START_COLUMN + "{0}:" + END_COLUMN + "{1}";
        private string[] columns = { "RegionCode", "ProductCode" };
        public int DataCount
        {
            get
            {
                return Convert.ToInt32(ExcelUtils.GetPropertyValue("SHEET5_DATACOUNT", this));
            }
            set
            {
                ExcelUtils.AddProperty("SHEET5_DATACOUNT", value.ToString(), this);
            }
        }
        private void Sheet5_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet5_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet5_Startup);
            this.Shutdown += new System.EventHandler(Sheet5_Shutdown);
        }

        #endregion

        #region 大区指标汇总

        public void Init(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                return;
            }
            DataCount = dt.Rows.Count;
            ExcelUtils.UnLockSheet(this);

            LoadData(dt);

            SetFormula();

            ResetFormat();

            ExcelUtils.LockSheet(this);
        }

        public void LoadData(DataTable dt)
        {
            //DataTable mydt = dt.DefaultView.ToTable(false, columns);
            string startCell, endCell;

            //数据填充
            if (DataCount > 0)
            {
                startCell = "A" + (TITLE_ROW_COUNT + 1).ToString();
                endCell = "B" + (TITLE_ROW_COUNT + DataCount).ToString();
                VSTOCommon.ExcelUtils.PrintData(this, dt, startCell, endCell);
            }
        }

        public void SetFormula()
        {
            try
            {
                Excel.Range rg;
                rg = this.Range["C" + (TITLE_ROW_COUNT + 1).ToString(), "C" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=SUMIFS(大区指标调整明细!$V:$V,大区指标调整明细!$A:$A,$A{0},大区指标调整明细!$J:$J,$B{0})", TITLE_ROW_COUNT + 1);
                rg = this.Range["D" + (TITLE_ROW_COUNT + 1).ToString(), "D" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=SUMIFS(大区指标调整明细!$X:$X,大区指标调整明细!$A:$A,$A{0},大区指标调整明细!$J:$J,$B{0})", TITLE_ROW_COUNT + 1);
                rg = this.Range["E" + (TITLE_ROW_COUNT + 1).ToString(), "E" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=SUMIFS(数据整合明细!$R:$R,数据整合明细!$A:$A,$A{0},数据整合明细!$J:$J,$B{0})", TITLE_ROW_COUNT + 1);
                rg = this.Range["F" + (TITLE_ROW_COUNT + 1).ToString(), "F" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=SUMIFS(数据整合明细!$S:$S,数据整合明细!$A:$A,$A{0},数据整合明细!$J:$J,$B{0})", TITLE_ROW_COUNT + 1);
                rg = this.Range["G" + (TITLE_ROW_COUNT + 1).ToString(), "G" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=E{0}-C{0}", TITLE_ROW_COUNT + 1);
                rg = this.Range["H" + (TITLE_ROW_COUNT + 1).ToString(), "H" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=F{0}-D{0}", TITLE_ROW_COUNT + 1);
            }
            catch (Exception ex)
            {
                LogHelper.WriteError("", ex);
            }
        }

        public void ResetFormat()
        {
            Excel.Range rg;
            //格式
            rg = this.Range[string.Format(DATA_ROW, TITLE_ROW_COUNT + 1, TITLE_ROW_COUNT + DataCount)];
            VSTOCommon.ExcelUtils.SetCellsColor_ReadOnely(rg, true);

            rg = this.Range["C" + (TITLE_ROW_COUNT + 1).ToString(), "F" + (TITLE_ROW_COUNT + DataCount).ToString()];
            rg.NumberFormat = "#,##0";

            rg = this.Range["A" + (TITLE_ROW_COUNT + 1).ToString(), "D" + (TITLE_ROW_COUNT + DataCount).ToString()];
            rg.Interior.Color = System.Drawing.Color.Yellow;

            rg = this.Range["G" + (TITLE_ROW_COUNT + 1).ToString(), "H" + (TITLE_ROW_COUNT + DataCount).ToString()];
            rg.NumberFormat = "#,##0;[Red](#,##0)";

            if (Globals.ThisWorkbook.TNTPlanCheck.Q1 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q3 == 0)
            {
                this.Range["C1"].EntireColumn.Hidden = true;
                this.Range["E1"].EntireColumn.Hidden = true;
                this.Range["G1"].EntireColumn.Hidden = true;
            }
            else
            {
                this.Range["C1"].EntireColumn.Hidden = false;
                this.Range["E1"].EntireColumn.Hidden = false;
                this.Range["G1"].EntireColumn.Hidden = false;
            }

            if (Globals.ThisWorkbook.TNTPlanCheck.Q2 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q4 == 0)
            {
                this.Range["D1"].EntireColumn.Hidden = true;
                this.Range["F1"].EntireColumn.Hidden = true;
                this.Range["H1"].EntireColumn.Hidden = true;
            }
            else
            {
                this.Range["D1"].EntireColumn.Hidden = false;
                this.Range["F1"].EntireColumn.Hidden = false;
                this.Range["H1"].EntireColumn.Hidden = false;
            }

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
                }
                ExcelUtils.LockSheet(this);
            }
            catch (Exception ex)
            {
                LogHelper.WriteError("", ex);
            }
        }

        public void ClearFilter()
        {
            try
            {
                if (!this.AutoFilter.FilterMode || DataCount == 0)
                {
                    return;
                }

                ExcelUtils.UnLockSheet(this);
                for (int i = 1; i < columns.Length + 1; i++)
                {
                    this.Range[string.Format(DATA_ROW, TITLE_ROW_COUNT + 1, TITLE_ROW_COUNT + DataCount)].AutoFilter(i, System.Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, System.Type.Missing, true);
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
