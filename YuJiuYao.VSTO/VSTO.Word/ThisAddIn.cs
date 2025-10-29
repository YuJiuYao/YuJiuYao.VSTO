using System;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace VSTO.Word
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            #region Right-Click Menu
            try
            {
                InitRightClickMenu();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("InitRightClickMenu error: " + ex);
                // Do not rethrow to avoid Word disabling the add-in
            }
            #endregion
        }

        private void InitRightClickMenu()
        {
            try
            {
                var commandBars = Globals.ThisAddIn.Application.CommandBars;
                CommandBar commandBar = null;
                foreach (CommandBar bar in commandBars)
                {
                    if (string.Equals(bar.Name, "Text", StringComparison.OrdinalIgnoreCase))
                    {
                        commandBar = bar;
                        break;
                    }
                }

                if (commandBar == null)
                {
                    // Context menu not found in this Word version/state
                    return;
                }

                // 自定义按钮
                if (commandBar.Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true) is CommandBarButton commandBarButton)
                {
                    commandBarButton.Tag = "右键按钮";
                    commandBarButton.Caption = "右键按钮";
                    commandBarButton.DescriptionText = "右键按钮";
                    //faceId 查找地址 http://www.kebabshopblues.co.uk/2007/01/04/visual-studio-2005-tools-for-office-commandbarbutton-faceid-property/
                    commandBarButton.FaceId =852;
                    commandBarButton.Click += CommandBarButton_Click;
                }

                // 自定义组合
                if (commandBar.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true) is CommandBarPopup commandBarPopup)
                {
                    commandBarPopup.Tag = "右键组";
                    commandBarPopup.Caption = "右键组";
                    commandBarPopup.DescriptionText = "右键组";

                    if (commandBarPopup.Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true) is CommandBarButton groupChildButton)
                    {
                        groupChildButton.Tag = "右键组子按钮";
                        groupChildButton.Caption = "右键组子按钮";
                        groupChildButton.DescriptionText = "右键组子按钮";
                        groupChildButton.FaceId =852;
                        groupChildButton.Click += CommandBarButton_Click;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("InitRightClickMenu inner error: " + ex);
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        private void CommandBarButton_Click(CommandBarButton ctrl, ref bool cancelDefault)
        {
            MessageBox.Show(ctrl.Tag);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
