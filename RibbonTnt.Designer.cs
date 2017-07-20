namespace HospitalTargetCheckSFA
{
    partial class RibbonTnt : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonTnt()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonTnt));
            this.tabTnt = this.Factory.CreateRibbonTab();
            this.grpUserInfo = this.Factory.CreateRibbonGroup();
            this.btnUserName = this.Factory.CreateRibbonButton();
            this.btnRegion = this.Factory.CreateRibbonButton();
            this.btnPosition = this.Factory.CreateRibbonButton();
            this.lblUserName = this.Factory.CreateRibbonLabel();
            this.lblRegion = this.Factory.CreateRibbonLabel();
            this.lblPosition = this.Factory.CreateRibbonLabel();
            this.grpDetail = this.Factory.CreateRibbonGroup();
            this.btnStatus = this.Factory.CreateRibbonButton();
            this.btnUpdateTime = this.Factory.CreateRibbonButton();
            this.btnUpdateUser = this.Factory.CreateRibbonButton();
            this.lblStatus = this.Factory.CreateRibbonLabel();
            this.lblUpdateTime = this.Factory.CreateRibbonLabel();
            this.lblUpdateUser = this.Factory.CreateRibbonLabel();
            this.grpHC = this.Factory.CreateRibbonGroup();
            this.btnDownload = this.Factory.CreateRibbonButton();
            this.btnCheckData = this.Factory.CreateRibbonButton();
            this.btnCheckStatus = this.Factory.CreateRibbonButton();
            this.btnLogout = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnChangeSU = this.Factory.CreateRibbonButton();
            this.tabTnt.SuspendLayout();
            this.grpUserInfo.SuspendLayout();
            this.grpDetail.SuspendLayout();
            this.grpHC.SuspendLayout();
            this.group2.SuspendLayout();
            // 
            // tabTnt
            // 
            this.tabTnt.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabTnt.Groups.Add(this.grpUserInfo);
            this.tabTnt.Groups.Add(this.grpDetail);
            this.tabTnt.Groups.Add(this.group2);
            this.tabTnt.Groups.Add(this.grpHC);
            this.tabTnt.Label = "TNT";
            this.tabTnt.Name = "tabTnt";
            // 
            // grpUserInfo
            // 
            this.grpUserInfo.Items.Add(this.btnUserName);
            this.grpUserInfo.Items.Add(this.btnRegion);
            this.grpUserInfo.Items.Add(this.btnPosition);
            this.grpUserInfo.Items.Add(this.lblUserName);
            this.grpUserInfo.Items.Add(this.lblRegion);
            this.grpUserInfo.Items.Add(this.lblPosition);
            this.grpUserInfo.Name = "grpUserInfo";
            // 
            // btnUserName
            // 
            this.btnUserName.Image = ((System.Drawing.Image)(resources.GetObject("btnUserName.Image")));
            this.btnUserName.Label = "当前用户：";
            this.btnUserName.Name = "btnUserName";
            this.btnUserName.ShowImage = true;
            // 
            // btnRegion
            // 
            this.btnRegion.Image = ((System.Drawing.Image)(resources.GetObject("btnRegion.Image")));
            this.btnRegion.Label = "所在大区：";
            this.btnRegion.Name = "btnRegion";
            this.btnRegion.ShowImage = true;
            // 
            // btnPosition
            // 
            this.btnPosition.Image = ((System.Drawing.Image)(resources.GetObject("btnPosition.Image")));
            this.btnPosition.Label = "职位：";
            this.btnPosition.Name = "btnPosition";
            this.btnPosition.ShowImage = true;
            // 
            // lblUserName
            // 
            this.lblUserName.Label = " ";
            this.lblUserName.Name = "lblUserName";
            // 
            // lblRegion
            // 
            this.lblRegion.Label = " ";
            this.lblRegion.Name = "lblRegion";
            // 
            // lblPosition
            // 
            this.lblPosition.Label = " ";
            this.lblPosition.Name = "lblPosition";
            // 
            // grpDetail
            // 
            this.grpDetail.Items.Add(this.btnStatus);
            this.grpDetail.Items.Add(this.btnUpdateTime);
            this.grpDetail.Items.Add(this.btnUpdateUser);
            this.grpDetail.Items.Add(this.lblStatus);
            this.grpDetail.Items.Add(this.lblUpdateTime);
            this.grpDetail.Items.Add(this.lblUpdateUser);
            this.grpDetail.Name = "grpDetail";
            // 
            // btnStatus
            // 
            this.btnStatus.Image = ((System.Drawing.Image)(resources.GetObject("btnStatus.Image")));
            this.btnStatus.Label = "所处流程：";
            this.btnStatus.Name = "btnStatus";
            this.btnStatus.ShowImage = true;
            // 
            // btnUpdateTime
            // 
            this.btnUpdateTime.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdateTime.Image")));
            this.btnUpdateTime.Label = "上次更新时间：";
            this.btnUpdateTime.Name = "btnUpdateTime";
            this.btnUpdateTime.ShowImage = true;
            // 
            // btnUpdateUser
            // 
            this.btnUpdateUser.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdateUser.Image")));
            this.btnUpdateUser.Label = "最近操作者：";
            this.btnUpdateUser.Name = "btnUpdateUser";
            this.btnUpdateUser.ShowImage = true;
            // 
            // lblStatus
            // 
            this.lblStatus.Label = " ";
            this.lblStatus.Name = "lblStatus";
            // 
            // lblUpdateTime
            // 
            this.lblUpdateTime.Label = " ";
            this.lblUpdateTime.Name = "lblUpdateTime";
            // 
            // lblUpdateUser
            // 
            this.lblUpdateUser.Label = " ";
            this.lblUpdateUser.Name = "lblUpdateUser";
            // 
            // grpHC
            // 
            this.grpHC.Items.Add(this.btnDownload);
            this.grpHC.Items.Add(this.btnCheckData);
            this.grpHC.Items.Add(this.btnCheckStatus);
            this.grpHC.Items.Add(this.btnLogout);
            this.grpHC.Name = "grpHC";
            // 
            // btnDownload
            // 
            this.btnDownload.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDownload.Image = ((System.Drawing.Image)(resources.GetObject("btnDownload.Image")));
            this.btnDownload.Label = "更新数据";
            this.btnDownload.Name = "btnDownload";
            this.btnDownload.ShowImage = true;
            this.btnDownload.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDownload_Click);
            // 
            // btnCheckData
            // 
            this.btnCheckData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCheckData.Image = ((System.Drawing.Image)(resources.GetObject("btnCheckData.Image")));
            this.btnCheckData.Label = " 校验数据  ";
            this.btnCheckData.Name = "btnCheckData";
            this.btnCheckData.ShowImage = true;
            this.btnCheckData.Visible = false;
            this.btnCheckData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCheckData_Click);
            // 
            // btnCheckStatus
            // 
            this.btnCheckStatus.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCheckStatus.Image = ((System.Drawing.Image)(resources.GetObject("btnCheckStatus.Image")));
            this.btnCheckStatus.Label = "查看提交状态";
            this.btnCheckStatus.Name = "btnCheckStatus";
            this.btnCheckStatus.ShowImage = true;
            this.btnCheckStatus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCheckStatus_Click);
            // 
            // btnLogout
            // 
            this.btnLogout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLogout.Image = ((System.Drawing.Image)(resources.GetObject("btnLogout.Image")));
            this.btnLogout.Label = "Logout注销";
            this.btnLogout.Name = "btnLogout";
            this.btnLogout.ShowImage = true;
            this.btnLogout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLogout_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnChangeSU);
            this.group2.Name = "group2";
            // 
            // btnChangeSU
            // 
            this.btnChangeSU.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnChangeSU.Image = ((System.Drawing.Image)(resources.GetObject("btnChangeSU.Image")));
            this.btnChangeSU.Label = "更换SU";
            this.btnChangeSU.Name = "btnChangeSU";
            this.btnChangeSU.ShowImage = true;
            this.btnChangeSU.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeSU_Click);
            // 
            // RibbonTnt
            // 
            this.Name = "RibbonTnt";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabTnt);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonTnt_Load);
            this.tabTnt.ResumeLayout(false);
            this.tabTnt.PerformLayout();
            this.grpUserInfo.ResumeLayout(false);
            this.grpUserInfo.PerformLayout();
            this.grpDetail.ResumeLayout(false);
            this.grpDetail.PerformLayout();
            this.grpHC.ResumeLayout(false);
            this.grpHC.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabTnt;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpUserInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblUserName;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblRegion;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblPosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpHC;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDetail;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblUpdateTime;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblUpdateUser;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUserName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRegion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateTime;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateUser;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDownload;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLogout;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeSU;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonTnt RibbonTnt
        {
            get { return this.GetRibbon<RibbonTnt>(); }
        }
    }
}
