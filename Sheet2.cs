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
using DataAccess.Model;
using VSTOCommon;

namespace HospitalTargetCheckSFA
{
    public partial class Sheet2
    {
        const string START_COLUMN = "A";
        const string END_COLUMN = "AB";
        const int TITLE_ROW_COUNT = 1;
        string DATA_ROW = START_COLUMN + "{0}:" + END_COLUMN + "{1}";
        private string[] columns = { "RegionCode","DistrictName","DMNumber","DMName","HospitalCode","HospitalName","Province",
                                     "City","HospitalLevel","ProductCode","PositionCode","UserName","ProductLineName",
                                     "PreQ1","PreQ2","PreQ3","PreQ4","OriginalQ1","OriginalQ2","HalfYearTarget","RevisedQ1",
		                             "FinalQ1","RevisedQ2","FinalQ2","FinalTarget","Verification","PLineCorrect","SuNoWeight" };
        public int DataCount
        {
            get
            {
                return Convert.ToInt32(ExcelUtils.GetPropertyValue("SHEET2_DATACOUNT", this));
            }
            set
            {
                ExcelUtils.AddProperty("SHEET2_DATACOUNT", value.ToString(), this);
            }
        }
        private void Sheet2_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet2_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet2_Startup);
            this.Shutdown += new System.EventHandler(Sheet2_Shutdown);
        }

        #endregion

        #region 大区指标调整明细

        public void Init(DataTable dt)
        {
            if(dt == null || dt.Rows.Count == 0)
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
                startCell = START_COLUMN + (TITLE_ROW_COUNT + 1).ToString();
                endCell = END_COLUMN + (TITLE_ROW_COUNT + DataCount).ToString();
                VSTOCommon.ExcelUtils.PrintData(this, mydt, startCell, endCell);
            }
            
        }

        public void SetFormula()
        {
            try
            {
                Excel.Range rg;

                rg = this.Range["T" + (TITLE_ROW_COUNT + 1).ToString(), "T" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=R{0}+S{0}", TITLE_ROW_COUNT + 1);

                rg = this.Range["V" + (TITLE_ROW_COUNT + 1).ToString(), "V" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=R{0}+U{0}", TITLE_ROW_COUNT + 1);

                rg = this.Range["X" + (TITLE_ROW_COUNT + 1).ToString(), "X" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=S{0}+W{0}", TITLE_ROW_COUNT + 1);

                rg = this.Range["Y" + (TITLE_ROW_COUNT + 1).ToString(), "Y" + (TITLE_ROW_COUNT + DataCount).ToString()];
                if (Globals.ThisWorkbook.TNTPlanCheck.Q1 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q3 == 0)
                {
                    rg.Formula = string.Format("=X{0}", TITLE_ROW_COUNT + 1);
                }
                else if (Globals.ThisWorkbook.TNTPlanCheck.Q2 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q4 == 0)
                {
                    rg.Formula = string.Format("=V{0}", TITLE_ROW_COUNT + 1);
                }
                else
                {
                    rg.Formula = string.Format("=V{0}+X{0}", TITLE_ROW_COUNT + 1);
                }

                rg = this.Range["Z" + (TITLE_ROW_COUNT + 1).ToString(), "Z" + (TITLE_ROW_COUNT + DataCount).ToString()];
                rg.Formula = string.Format("=IF(OR(AND(DashBoard!$J$2<>\"\",V{0}=0),AND(DashBoard!$J$3<>\"\",X{0}=0)),\"医院对应的岗位指标为0，这家医院的销量将无法计入该岗位\",\"\")", TITLE_ROW_COUNT + 1);

                // 设置条件公式.
                Excel.FormatCondition cond = rg.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlNotEqual, "\"\"");
                cond.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                rg = this.Range["AA" + (TITLE_ROW_COUNT + 1).ToString(), "AA" + (TITLE_ROW_COUNT + DataCount).ToString()];
                cond = rg.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "FALSE");
                cond.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                rg = this.Range["AB" + (TITLE_ROW_COUNT + 1).ToString(), "AB" + (TITLE_ROW_COUNT + DataCount).ToString()];
                cond = rg.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "TRUE");
                cond.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
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

            rg = this.Range["N" + (TITLE_ROW_COUNT + 1).ToString(), "Y" + (TITLE_ROW_COUNT + DataCount).ToString()];
            rg.NumberFormat = "#,##0";

            rg = this.Range["AA" + (TITLE_ROW_COUNT + 1).ToString(), "AB" + (TITLE_ROW_COUNT + DataCount).ToString()];
            rg.Interior.Color = System.Drawing.Color.Yellow;

            if (Globals.ThisWorkbook.TNTPlanCheck.Q1 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q3 == 0)
            {
                this.Range["R1"].EntireColumn.Hidden = true;
                this.Range["U1"].EntireColumn.Hidden = true;
                this.Range["V1"].EntireColumn.Hidden = true;
            }
            else
            {
                this.Range["R1"].EntireColumn.Hidden = false;
                this.Range["U1"].EntireColumn.Hidden = false;
                this.Range["V1"].EntireColumn.Hidden = false;
            }

            if (Globals.ThisWorkbook.TNTPlanCheck.Q2 == 0 && Globals.ThisWorkbook.TNTPlanCheck.Q4 == 0)
            {
                this.Range["S1"].EntireColumn.Hidden = true;
                this.Range["W1"].EntireColumn.Hidden = true;
                this.Range["X1"].EntireColumn.Hidden = true;
            }
            else
            {
                this.Range["S1"].EntireColumn.Hidden = false;
                this.Range["W1"].EntireColumn.Hidden = false;
                this.Range["X1"].EntireColumn.Hidden = false;
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
