using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using DataAccess.Model;
using DataAccess.BLL.Vsto;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Windows.Forms;
using TNTVSTO;

namespace HospitalTargetCheckSFA
{
    public partial class RibbonTnt
    {
        private void RibbonTnt_Load(object sender, RibbonUIEventArgs e)
        {
            //设置基本信息
            InitInfor();
            //根据角色控制可用按钮
            SetVisble();

        }
        /// <summary>
        /// 根据角色控制可用按钮
        /// </summary>
        public void SetVisble()
        {
            char enter = (char)13;
            btnCheckData.Label += enter;
            btnDownload.Label += enter;
            btnCheckStatus.Label += enter;

            if (Globals.ThisWorkbook.LoginUser == null)
                return;

            switch (Globals.ThisWorkbook.LoginUser.RoleCode)
            {
                case (int)RoleCode.SFA:

                    break;

                case (int)RoleCode.SULeader:

                    break;

                case (int)RoleCode.RM:

                    break;

                case (int)RoleCode.RME:

                    break;

                case (int)RoleCode.DM:

                    break;
                     
            }
        }
        /// <summary>
        /// 设置基本信息
        /// </summary>
        public void InitInfor()
        {
            if (Globals.ThisWorkbook == null || Globals.ThisWorkbook.IsFirstLaunch())
            {
                this.lblUserName.Label = "未登录";
                this.lblPosition.Label = "未登录";
                this.lblRegion.Label = "未登录";
                this.lblUpdateTime.Label = "未登录";
                this.lblUpdateUser.Label = "未登录";
                return;
            }
            //姓名
            lblUserName.Label = Globals.ThisWorkbook.LoginUser.UserName;
            //职位
            lblPosition.Label = Globals.ThisWorkbook.LoginUser.RoleName;
            //大区
            if (Globals.ThisWorkbook.LoginUser.RoleCode != (int)RoleCode.SFA)
            {
                lblRegion.Label = Globals.ThisWorkbook.LoginUser.RMCode;
            }
            else
            {
                lblRegion.Label = Globals.ThisWorkbook.LoginUser.PositionCode;
            }
            //流程状态
            lblStatus.Label = "[SFA指标核查]";
            //更新时间
            lblUpdateTime.Label =  "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "]";
            //更新者
            lblUpdateUser.Label = Globals.ThisWorkbook.LoginUser.UserName;           

        }
        /// <summary>
        ///  校验数据  
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheckData_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisWorkbook.CheckAllData();
        }                
        /// <summary>
        /// 查看提交状态
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheckStatus_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.GetSubmitStatus();
        }

        /// <summary>
        /// 更新下载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDownload_Click(object sender, RibbonControlEventArgs e)
        {
            CommonBLL checkbll = new CommonBLL();
            bool re = checkbll.CheckCanLogonForVersion(ThisWorkbook.ExcelName, ThisWorkbook.Version);
            if (!re)
            {
                MessageBox.Show("请当前使用的客户端版本过低,请升级客户端版本");
                return;
            }
            //Dictionary<string, string> dicSu = new LoginBll().GetAllSuList();
            //formSelectPosition form = new formSelectPosition(dicSu.Keys.ToList(), "请选择SU：");
            //form.Text = "选择登录SU";
            //form.ShowDialog();
            //Globals.ThisWorkbook.LoginUser.SUID = Guid.Parse(dicSu[form.SelectedPosition]);
            //Globals.ThisWorkbook.LoginUser.PositionCode = form.SelectedPosition;
            
            Globals.ThisWorkbook.LoadAllData();
        }
        /// <summary>
        /// 数据提交
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSubmit_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisWorkbook.SubmitData();
        }

        public void SetRibbonButtonEnable(bool enable)
        {
            btnDownload.Enabled = enable;
            btnCheckData.Enabled = enable;
            btnCheckStatus.Enabled = enable;
            //btnLogout.Enabled = enable;
        }

        private void btnLogout_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Logout();
        }

        private void btnChangeSU_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.SelectSU();
            InitInfor();
            Globals.ThisWorkbook.LoadAllData();
        }

    }

}
