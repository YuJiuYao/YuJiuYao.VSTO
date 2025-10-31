using System.Collections.Generic;
using Microsoft.Office.Tools;
using VSTO.Forms.Forms;

namespace VSTO.Word.Common
{
    internal class CustomTaskPaneHelper
    {
        public static void OpenCustomTaskPane(string url, string title)
        {
            var docId = Globals.ThisAddIn.Application.ActiveDocument.DocID;
            if (ParametersHelper.CustomTaskPanes.ContainsKey(docId))
            {
                var customTaskPane = ParametersHelper.CustomTaskPanes[docId].Find(c => c.Control.Tag.ToString() == title);
                if (customTaskPane != null)
                {
                    if (customTaskPane.Visible)
                        return;

                    ParametersHelper.CustomTaskPanes[docId].ForEach(c => c.Visible = false);
                    customTaskPane.Visible = true;
                }
                else
                {
                    ParametersHelper.CustomTaskPanes[docId].Add(GetCustomTaskPane(url, title));
                }
            }
            else
            {
                ParametersHelper.CustomTaskPanes.Add(docId, new List<CustomTaskPane>() { GetCustomTaskPane(url, title) });
            }
        }

        private static CustomTaskPane GetCustomTaskPane(string url, string title)
        {
            var view = new YzControl(url, title);
            ParametersHelper.CustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(view, title);
            ParametersHelper.CustomTaskPane.Control.Tag = title;
            ParametersHelper.CustomTaskPane.Width = 300;
            ParametersHelper.CustomTaskPane.Visible = true;
            ParametersHelper.CustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            return ParametersHelper.CustomTaskPane;
        }

        public static void RemoveCustomTaskPane()
        {
            var docId = Globals.ThisAddIn.Application.ActiveDocument.DocID;
            ParametersHelper.CustomTaskPanes[docId].ForEach(c =>
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(c);
            });
            ParametersHelper.CustomTaskPanes[docId].Clear();
        }
    }
}
