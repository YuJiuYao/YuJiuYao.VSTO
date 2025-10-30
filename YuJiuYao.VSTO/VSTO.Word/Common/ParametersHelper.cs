using System.Collections.Generic;
using Microsoft.Office.Tools;
using VSTO.Common;

namespace VSTO.Word.Common
{
    internal class ParametersHelper : ParametersBaseHelper
    {
        public static CustomTaskPane _customTaskPane = null;

        /// <summary>
        /// 当前项目中每个文档中左侧栏集合
        /// </summary>
        public static Dictionary<int, List<CustomTaskPane>> CustomTaskPanes = new Dictionary<int, List<CustomTaskPane>>();
    }
}
