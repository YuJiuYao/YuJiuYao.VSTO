namespace VSTO.Word
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.customTab = this.Factory.CreateRibbonTab();
            this.groupDev = this.Factory.CreateRibbonGroup();
            this.btnGroupDevRowOne = this.Factory.CreateRibbonButtonGroup();
            this.buttonShowMessageBox = this.Factory.CreateRibbonButton();
            this.buttonUser = this.Factory.CreateRibbonButton();
            this.btnGroupDevRowTwo = this.Factory.CreateRibbonButtonGroup();
            this.btnOpenCusPane = this.Factory.CreateRibbonButton();
            this.btnOpenPopupBox = this.Factory.CreateRibbonButton();
            this.groupFunny = this.Factory.CreateRibbonGroup();
            this.customTab.SuspendLayout();
            this.groupDev.SuspendLayout();
            this.btnGroupDevRowOne.SuspendLayout();
            this.btnGroupDevRowTwo.SuspendLayout();
            this.SuspendLayout();
            // 
            // customTab
            // 
            this.customTab.Groups.Add(this.groupDev);
            this.customTab.Groups.Add(this.groupFunny);
            this.customTab.Label = "YuJiuYao工具箱";
            this.customTab.Name = "customTab";
            // 
            // groupDev
            // 
            this.groupDev.Items.Add(this.btnGroupDevRowOne);
            this.groupDev.Items.Add(this.btnGroupDevRowTwo);
            this.groupDev.Label = "开发者测试";
            this.groupDev.Name = "groupDev";
            // 
            // btnGroupDevRowOne
            // 
            this.btnGroupDevRowOne.Items.Add(this.buttonShowMessageBox);
            this.btnGroupDevRowOne.Items.Add(this.buttonUser);
            this.btnGroupDevRowOne.Name = "btnGroupDevRowOne";
            // 
            // buttonShowMessageBox
            // 
            this.buttonShowMessageBox.Label = "测试";
            this.buttonShowMessageBox.Name = "buttonShowMessageBox";
            this.buttonShowMessageBox.OfficeImageId = "UpdateFolder";
            this.buttonShowMessageBox.ShowImage = true;
            this.buttonShowMessageBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonShowMessageBox_Click);
            // 
            // buttonUser
            // 
            this.buttonUser.Label = "作者";
            this.buttonUser.Name = "buttonUser";
            this.buttonUser.OfficeImageId = "UpdateFolder";
            this.buttonUser.ShowImage = true;
            this.buttonUser.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonUser_Click);
            // 
            // btnGroupDevRowTwo
            // 
            this.btnGroupDevRowTwo.Items.Add(this.btnOpenCusPane);
            this.btnGroupDevRowTwo.Items.Add(this.btnOpenPopupBox);
            this.btnGroupDevRowTwo.Name = "btnGroupDevRowTwo";
            // 
            // btnOpenCusPane
            // 
            this.btnOpenCusPane.Label = "侧边弹框";
            this.btnOpenCusPane.Name = "btnOpenCusPane";
            this.btnOpenCusPane.OfficeImageId = "UpdateFolder";
            this.btnOpenCusPane.ShowImage = true;
            this.btnOpenCusPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenCusPane_Click);
            // 
            // btnOpenPopupBox
            // 
            this.btnOpenPopupBox.Label = "浮动弹框";
            this.btnOpenPopupBox.Name = "btnOpenPopupBox";
            this.btnOpenPopupBox.OfficeImageId = "UpdateFolder";
            this.btnOpenPopupBox.ShowImage = true;
            this.btnOpenPopupBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenPopupBox_Click);
            // 
            // groupFunny
            // 
            this.groupFunny.Label = "娱乐";
            this.groupFunny.Name = "groupFunny";
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.customTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.customTab.ResumeLayout(false);
            this.customTab.PerformLayout();
            this.groupDev.ResumeLayout(false);
            this.groupDev.PerformLayout();
            this.btnGroupDevRowOne.ResumeLayout(false);
            this.btnGroupDevRowOne.PerformLayout();
            this.btnGroupDevRowTwo.ResumeLayout(false);
            this.btnGroupDevRowTwo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab customTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDev;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFunny;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup btnGroupDevRowOne;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup btnGroupDevRowTwo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonShowMessageBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUser;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenCusPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenPopupBox;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon Ribbon1
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
