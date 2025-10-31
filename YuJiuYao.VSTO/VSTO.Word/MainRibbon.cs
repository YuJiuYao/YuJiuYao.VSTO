using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using VSTO.Common;
using VSTO.Forms.Common;
using VSTO.Forms.Forms;
using VSTO.Word.Common;

namespace VSTO.Word
{
    public partial class MainRibbon
    {
        private Microsoft.Office.Interop.Word.Application _application;

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _application = Globals.ThisAddIn.Application;
        }

        private void ShowMessageBox_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(_application.ActiveDocument.ToString());
        }

        private void User_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show($@"作者：{_application.UserName}");
        }

        private void OpenCusPane_Click(object sender, RibbonControlEventArgs e)
        {
            // 创建或指向本地页面
            var htmlFileDir = Path.Combine(ParametersBaseHelper.BaseDirectoryPath, "HTML", "test.html");
            var fileUrl = new Uri(htmlFileDir).AbsoluteUri; // file:/// URL
            CustomTaskPaneHelper.OpenCustomTaskPane(fileUrl, ((RibbonButton)sender).Label);
        }

        private void OpenPopupBox_Click(object sender, RibbonControlEventArgs e)
        {
            var popupBoxForm = new PopupBoxForm("https://www.youtube.com/", new ClientPortal(new Forms.BaseForms.WebFormView()))
            {
                Dock = DockStyle.None,
                Icon = Properties.Resources.popupBoxFormIco,
                Text = @"浮动弹框",
                Size = new System.Drawing.Size(Convert.ToInt32(ParametersBaseHelper.WindowWidth * 0.75),
                    Convert.ToInt32(ParametersBaseHelper.WindowHeight * 0.75))
            };
            popupBoxForm.ShowDialog();
        }
    }
}
