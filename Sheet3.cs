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
    public partial class Sheet3
    {
        const string START_COLUMN = "A";
        const string END_COLUMN = "T";
        const int TITLE_ROW_COUNT = 1;
        string DATA_ROW = START_COLUMN + "{0}:" + END_COLUMN + "{1}";
        private string[] columns = { "RegionCode","DistrictName","DMNumber","DMName","HospitalCode","HospitalName","Province","City",
                                     "HospitalLevel","ProductCode","PositionCode","UserName","ProductLineName",
                                     "PreQ1","PreQ2","PreQ3","PreQ4","SumQ1","SumQ2"};
        public int DataCount
        {
            get
            {
                return Convert.ToInt32(ExcelUtils.GetPropertyValue("SHEET3_DATACOUNT", this));
            }
            set
            {
                ExcelUtils.AddProperty("SHEET3_DATACOUNT", value.ToString(), this);
            }
        }
        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet3_Startup);
            this.Shutdown += new System.EventHandler(Sheet3_Shutdown);
        }

        #endregion

        #region 数据整合明细

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
            DataTable mydt = dt.DefaultView.ToTable(false, columns);
            string startCell, endCell;

            //数据填充
            if (DataCount > 0)
            {
                startCell = "A" + (TITLE_ROW_COUNT + 1).ToString();
                endCell = "S" + (TITLE_ROW_COUNT + DataCount).ToString();
                VSTOCommon.ExcelUtils.PrintData(this, mydt, startCell, endCell);
            }
        }

        public void SetFormula()
        {
            try
            {
                Excel.Range rg;

                rg = this.Range["T" + (TITLE_ROW_COUNT + 1).ToString(), "T" + (TITLE_ROW_COUNT + DataCount).ToString()];
                if (Globals.ThisWorkbook.TNTPlanCheck.Q1 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q3 == 0)
                {
                    rg.Formula = string.Format("=S{0}", TITLE_ROW_COUNT + 1);
                }
                else if (Globals.ThisWorkbook.TNTPlanCheck.Q2 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q4 == 0)
                {
                    rg.Formula = string.Format("=R{0}", TITLE_ROW_COUNT + 1);
                }
                else
                {
                    rg.Formula = string.Format("=R{0}+S{0}", TITLE_ROW_COUNT + 1);
                }
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

            rg = this.Range["N" + (TITLE_ROW_COUNT + 1).ToString(), "T" + (TITLE_ROW_COUNT + DataCount).ToString()];
            rg.NumberFormat = "#,##0";

            rg = this.Range["R" + (TITLE_ROW_COUNT + 1).ToString(), "S" + (TITLE_ROW_COUNT + DataCount).ToString()];
            rg.Interior.Color = System.Drawing.Color.Yellow;

            if (Globals.ThisWorkbook.TNTPlanCheck.Q1 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q3 == 0)
            {
                this.Range["R1"].EntireColumn.Hidden = true;
            }
            else
            {
                this.Range["R1"].EntireColumn.Hidden = false;
            }

            if (Globals.ThisWorkbook.TNTPlanCheck.Q2 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q4 == 0)
            {
                this.Range["S1"].EntireColumn.Hidden = true;
            }
            else
            {
                this.Range["S1"].EntireColumn.Hidden = false;
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
