namespace VSTO.Forms.Forms
{
    partial class YzControl
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.webFormView = new VSTO.Forms.BaseForms.WebFormView();
            this.SuspendLayout();
            // 
            // webFormView1
            // 
            this.webFormView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webFormView.JsApi = null;
            this.webFormView.Location = new System.Drawing.Point(0, 0);
            this.webFormView.Name = "webFormView1";
            this.webFormView.Size = new System.Drawing.Size(150, 150);
            this.webFormView.TabIndex = 0;
            this.webFormView.TaskPaneVisibleEv = null;
            this.webFormView.Url = null;
            this.webFormView.UserAgent = false;
            // 
            // DxControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.webFormView);
            this.Name = "DxControl";
            this.ResumeLayout(false);
        }

        #endregion

        private VSTO.Forms.BaseForms.WebFormView webFormView;
    }
}
