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
using DataAccess.BLL.Vsto;
using DataAccess.Model;
using DataAccess;
using VSTOCommon;
using TNTVSTO;

namespace HospitalTargetCheckSFA
{
    public partial class ThisWorkbook
    {
        public static readonly string Version = "1.0";
        public static readonly string ExcelName = "HospitalTargetSFA";
        private const string CUSTOM_PROPERTY_IS_FIRST_LAUNCH = "TNT_IS_FIRST_LAUNCH";
        private const string CUSTOM_PROPERTY_USER_ID = "TNT_USER_ID";
        private const string CUSTOM_PROPERTY_USER_NAME = "TNT_USER_NAME";
        private const string CUSTOM_PROPERTY_ROLE_CODE = "TNT_ROLE_CODE";
        private const string CUSTOM_PROPERTY_POSITION_CODE = "TNT_POSITION_CODE";
        private const string CUSTOM_PROPERTY_IS_LOGIN_AS = "TNT_IS_LOGIN_AS";
        private const string CUSTOM_PROPERTY_PRODUCT_CODE = "TNT_PRODUCT_CODE";
        private const string CUSTOM_PROPERTY_SU_CODE = "TNT_SU_CODE";
        public UserInfo LoginUser;
        public TNTPlan TNTPlanCheck;
        private TargetCheckHospitalBll bll = new TargetCheckHospitalBll();
        public Dictionary<string, string> dicSu;
        //public string PositionCode { get; set; }
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            dicSu = bll.GetSUList();
            Init();
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Login code

        public bool IsFirstLaunch()
        {
            return !ExcelUtils.IsPropertyExisted(CUSTOM_PROPERTY_IS_FIRST_LAUNCH, this);
        }


        public void ClearProperties()
        {
            ExcelUtils.DeleteProperty(CUSTOM_PROPERTY_IS_FIRST_LAUNCH, this);
            ExcelUtils.DeleteProperty(CUSTOM_PROPERTY_USER_ID, this);
            ExcelUtils.DeleteProperty(CUSTOM_PROPERTY_USER_NAME, this);
            ExcelUtils.DeleteProperty(CUSTOM_PROPERTY_ROLE_CODE, this);
            ExcelUtils.DeleteProperty(CUSTOM_PROPERTY_POSITION_CODE, this);
            ExcelUtils.DeleteProperty(CUSTOM_PROPERTY_IS_LOGIN_AS, this);
            ExcelUtils.DeleteProperty(CUSTOM_PROPERTY_PRODUCT_CODE, this);
            ExcelUtils.DeleteProperty(CUSTOM_PROPERTY_SU_CODE, this);
        }


        #endregion

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

        private void Init()
        {
            if (IsFirstLaunch())
            {
                formLogin login = new formLogin(DataAccess.Model.TNTPlan.TNTPlanType.HospitalTarget, ExcelName, Version);
                login.ShowDialog();

                if (login.DialogResult == DialogResult.OK)
                {
                    LoginUser = login.LoginUser;

                    ExcelUtils.AddProperty(CUSTOM_PROPERTY_USER_ID, LoginUser.UserId.ToString(), this);
                    ExcelUtils.AddProperty(CUSTOM_PROPERTY_USER_NAME, LoginUser.UserName, this);
                    ExcelUtils.AddProperty(CUSTOM_PROPERTY_ROLE_CODE, LoginUser.RoleCode.ToString(), this);
                    //AddProperty(CUSTOM_PROPERTY_POSITION_CODE, LoginUser.PositionCode);
                    ExcelUtils.AddProperty(CUSTOM_PROPERTY_IS_LOGIN_AS, LoginUser.IsLoginAs.ToString(), this);
                    //AddProperty(CUSTOM_PROPERTY_SU_CODE, LoginUser.PositionCode);
                    ExcelUtils.AddProperty(CUSTOM_PROPERTY_IS_FIRST_LAUNCH, "NO", this);

                    if (!CheckTNTPlan())
                    {
                        SelectSU();
                    }
                    bll.AddDBLog(LoginUser.PositionCode, LoginUser.UserId, "SFA指标核查", "登录");
                    Globals.Ribbons.RibbonTnt.SetRibbonButtonEnable(true);
                    Globals.Ribbons.RibbonTnt.InitInfor();

                    LoadAllData();
                }
                else
                {
                    //doesn't login
                    Globals.Ribbons.RibbonTnt.SetRibbonButtonEnable(false);
                    Globals.Ribbons.RibbonTnt.InitInfor();
                }
            }
            else
            {
                //not the first launch, read from cache
                Guid userId = Guid.Parse(ExcelUtils.GetPropertyValue(CUSTOM_PROPERTY_USER_ID, this));
                String userName = ExcelUtils.GetPropertyValue(CUSTOM_PROPERTY_USER_NAME, this);
                Int32 roleCode = Int32.Parse(ExcelUtils.GetPropertyValue(CUSTOM_PROPERTY_ROLE_CODE, this));
                String positionCode = ExcelUtils.GetPropertyValue(CUSTOM_PROPERTY_POSITION_CODE, this);
                bool isLoginAs = Boolean.Parse(ExcelUtils.GetPropertyValue(CUSTOM_PROPERTY_IS_LOGIN_AS, this));
                //String selectedProductCode = GetPropertyValue(CUSTOM_PROPERTY_PRODUCT_CODE);

                LoginUser = new UserInfo
                {
                    UserId = userId,
                    UserName = userName,
                    RoleCode = roleCode,
                    PositionCode = positionCode,
                    IsLoginAs = isLoginAs
                };

                //SelectSU();

                Globals.Ribbons.RibbonTnt.InitInfor();
                //Globals.Ribbons.RibbonTnt.SetVisble();
                Globals.Ribbons.RibbonTnt.SetRibbonButtonEnable(true);
            }


            //
            //Globals.Ribbons.RibbonTnt.SetRibbonButtonEnable(true);
        }

        public bool CheckTNTPlan()
        {
            bool bresult = false;
            TNTPlanCheck = bll.GetTNTPlanSFACheck(LoginUser.SUID);
            if (TNTPlanCheck == null)
            {
                MessageBox.Show("无此SU指标数据");
                return false;
            }
            if (TNTPlanCheck.State != 16)
            {
                if (MessageBox.Show("此SU未完成调整,确定查看吗？", "未开始您的环节", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    bresult = true;
                }
                else
                {
                    bresult = false;
                }
            }
            else
            {
                bresult = true;
            }
            return bresult;
        }

        public void SelectSU()
        {
            formSelectPosition form = new formSelectPosition(dicSu.Keys.ToList(), "请选择SU：");
            form.Text = "选择登录SU";
            do
            {
                form.ShowDialog();
                LoginUser.SUID = Guid.Parse(dicSu[form.SelectedPosition]);
                LoginUser.PositionCode = form.SelectedPosition;
                //TNTPlanCheck = bll.GetTNTPlanSFACheck(LoginUser.SUID);
            } while (!CheckTNTPlan());
        }



        /// <summary>
        /// 获取各个表需要的数据
        /// </summary>
        public void LoadAllData()
        {
            formProcessing form = new formProcessing();
            try
            {
                this.Application.ScreenUpdating = false;
                bll.AddDBLog(LoginUser.PositionCode, LoginUser.UserId, "SFA指标核查", "读取");
                
                form.Show();
                ClearAllData();
                ClearFilter();
                form.SetProcessing(10);
                DataSet ds = bll.ProcessHospitalTargetCheckSFA(LoginUser.SUID);
                if (ds.Tables[0] != null && ds.Tables[0].Rows.Count > 0)
                {
                    form.SetProcessing(30);
                    Globals.Sheet1.Init(ds.Tables[0]);
                    form.SetProcessing(50);
                    Globals.Sheet2.Init(ds.Tables[1]);
                    form.SetProcessing(70);
                    Globals.Sheet3.Init(ds.Tables[2]);
                    form.SetProcessing(80);
                    Globals.Sheet4.Init(ds.Tables[3]);
                    form.SetProcessing(90);
                    Globals.Sheet5.Init(ds.Tables[4]);
                }
                else
                {
                    MessageBox.Show("无数据");
                }

                form.SetProcessing(100);
                form.Close();
                this.Application.ScreenUpdating = true;
                this.Save();
            }
            catch(Exception ex)
            {
                form.Close();
                LogHelper.WriteError("",ex);
            }
        }

        public void ClearAllData()
        {
            try
            {
                Globals.Sheet1.ClearData();
                Globals.Sheet2.ClearData();
                Globals.Sheet3.ClearData();
                Globals.Sheet4.ClearData();
                Globals.Sheet5.ClearData();
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
                Globals.Sheet2.ClearFilter();
                Globals.Sheet3.ClearFilter();
                Globals.Sheet4.ClearFilter();
                Globals.Sheet5.ClearFilter();
            }
            catch (Exception ex)
            {
                LogHelper.WriteError("", ex);
            }
        }

        public void Logout()
        {
            try
            {
                ClearAllData();
                ClearFilter();
                this.LoginUser = new UserInfo();
                //更新Ribbon
                //Globals.Ribbons.RibbonTnt.SetVisble();
                ClearProperties();
                this.Save();
                bll.AddDBLog(LoginUser.PositionCode, LoginUser.UserId, "SFA指标核查", "注销");
                Init();
            }
            catch (Exception ex)
            {
                LogHelper.WriteError("", ex);
            }
        }

        public void GetSubmitStatus()
        {
            formViewStatus form = new formViewStatus();
            if (form.LoadData(LoginUser.RoleCode, LoginUser.PositionCode, 2))
            {
                form.ShowDialog();
            }
        }
    }
}
