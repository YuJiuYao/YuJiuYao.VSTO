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
            this.btnGroup = this.Factory.CreateRibbonButtonGroup();
            this.buttonShowMessageBox = this.Factory.CreateRibbonButton();
            this.buttonNext = this.Factory.CreateRibbonButton();
            this.groupFunny = this.Factory.CreateRibbonGroup();
            this.customTab.SuspendLayout();
            this.groupDev.SuspendLayout();
            this.btnGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // customTab
            // 
            this.customTab.Groups.Add(this.groupDev);
            this.customTab.Groups.Add(this.groupFunny);
            this.customTab.Label = "YuJiuYao";
            this.customTab.Name = "customTab";
            // 
            // groupDev
            // 
            this.groupDev.Items.Add(this.btnGroup);
            this.groupDev.Label = "开发者测试";
            this.groupDev.Name = "groupDev";
            // 
            // btnGroup
            // 
            this.btnGroup.Items.Add(this.buttonShowMessageBox);
            this.btnGroup.Items.Add(this.buttonNext);
            this.btnGroup.Name = "btnGroup";
            // 
            // buttonShowMessageBox
            // 
            this.buttonShowMessageBox.Label = "测试";
            this.buttonShowMessageBox.Name = "buttonShowMessageBox";
            this.buttonShowMessageBox.ShowImage = true;
            this.buttonShowMessageBox.OfficeImageId = "UpdateFolder";
            this.buttonShowMessageBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonShowMessageBox_Click);
            // 
            // buttonNext
            // 
            this.buttonNext.Label = "下一个";
            this.buttonNext.Name = "buttonNext";
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
            this.btnGroup.ResumeLayout(false);
            this.btnGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab customTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDev;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFunny;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup btnGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonShowMessageBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNext;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon Ribbon1
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
