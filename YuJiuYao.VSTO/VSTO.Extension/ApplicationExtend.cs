using System;
using System.Collections.Generic;
using System.Reflection;

namespace VSTO.Extension
{
    /// <summary>
    /// Word Application 扩展方法（无 Office 引用，使用反射后期绑定）
    /// 适合在不直接引用 Microsoft.Office.Interop.Word.dll 的情况下动态操作 Word。
    /// 所有与 Word 的交互均通过反射完成，避免强依赖 Office PIA。
    /// </summary>
    public static class ApplicationExtend
    {
        // ------------------------
        // Word 常量（与 Interop 枚举等价）
        // ------------------------

        private const int WdNumberOfPagesInDocument = 4; // 对应 WdInformation.wdNumberOfPagesInDocument
        private const int WdStatisticPages = 2;          // 对应 WdStatistic.wdStatisticPages
        private const int WdGoToPage = 1;                // 对应 WdGoToItem.wdGoToPage
        private const int WdGoToAbsolute = 1;            // 对应 WdGoToDirection.wdGoToAbsolute

        /// <summary>
        /// 获取当前活动文档的总页数。
        /// </summary>
        /// <param name="app">Word Application 对象（通过 VSTO 的 Globals.ThisAddIn.Application 传入）</param>
        /// <returns>页数整数</returns>
        public static int GetActiveDocumentPageCount(this object app)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            // 获取 Word.Application 类型对象
            var appType = app.GetType();

            // 通过反射获取 Application.Selection 属性（当前选区）
            var selection = appType.InvokeMember("Selection", BindingFlags.GetProperty, null, app, null);
            if (selection == null) throw new InvalidOperationException("无法获取 Selection");

            // 获取选区对象类型
            var selType = selection.GetType();

            // 读取选区的 Information 属性，传入 wdNumberOfPagesInDocument（4），得到页数
            var pagesObj = selType.InvokeMember("Information", BindingFlags.GetProperty, null, selection,
                new object[] { WdNumberOfPagesInDocument });

            // 转换为整数返回
            return Convert.ToInt32(pagesObj);
        }

        /// <summary>
        /// 获取当前文档中所有“空白页”的页码列表。
        /// 空白页：指该页的文本范围（Range.Text）为空或仅包含空白字符。
        /// </summary>
        public static List<int> GetSpacePages(this object app)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));
            var result = new List<int>();

            try
            {
                var appType = app.GetType();

                // 1. 获取当前活动文档 ActiveDocument
                var doc = appType.InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, app, null);
                if (doc == null) return result;
                var docType = doc.GetType();

                // 2. 调用 ComputeStatistics(wdStatisticPages) 获取总页数
                var allPagesObj = docType.InvokeMember("ComputeStatistics", BindingFlags.InvokeMethod, null, doc,
                    new object[] { WdStatisticPages });
                var allPages = Convert.ToInt32(allPagesObj);

                // 3. 获取文档内容对象 Content，用于确定文档末尾位置
                var endRange = docType.InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                var endRangeType = endRange.GetType();

                // 获取文档结束位置（End 属性）
                var endPosObj = endRangeType.InvokeMember("End", BindingFlags.GetProperty, null, endRange, null);
                var endPos = Convert.ToInt32(endPosObj);

                // 4. 获取当前 Selection 对象（用于跳页 GoTo）
                var selection = appType.InvokeMember("Selection", BindingFlags.GetProperty, null, app, null);
                var selType = selection.GetType();

                // 5. 遍历每一页，检查其文字内容是否为空
                for (var i = 1; i <= allPages; i++)
                {
                    // 跳转到第 i 页
                    var range = selType.InvokeMember("GoTo", BindingFlags.InvokeMethod, null, selection,
                        new[] { WdGoToPage, WdGoToAbsolute, i, Type.Missing });
                    var rangeType = range.GetType();

                    // 设置当前页的结束位置
                    if (i == allPages)
                    {
                        // 最后一页：结束位置是整个文档末尾
                        rangeType.InvokeMember("End", BindingFlags.SetProperty, null, range, new object[] { endPos });
                    }
                    else
                    {
                        // 非最后一页：结束位置是下一页开始前一个字符
                        var nextRange = selType.InvokeMember("GoTo", BindingFlags.InvokeMethod, null, selection,
                            new[] { WdGoToPage, WdGoToAbsolute, i + 1, Type.Missing });
                        var nextRangeType = nextRange.GetType();
                        var nextStartObj =
                            nextRangeType.InvokeMember("Start", BindingFlags.GetProperty, null, nextRange, null);
                        var nextStart = Convert.ToInt32(nextStartObj);

                        rangeType.InvokeMember("End", BindingFlags.SetProperty, null, range,
                            new object[] { nextStart - 1 });
                    }

                    // 读取该页文本内容
                    var textObj = rangeType.InvokeMember("Text", BindingFlags.GetProperty, null, range, null);
                    var text = textObj?.ToString();

                    // 如果该页无实质内容，则记录页码
                    if (string.IsNullOrWhiteSpace(text?.Trim()))
                        result.Add(i);
                }
            }
            catch
            {
                // 捕获并忽略异常，返回当前结果
            }

            return result;
        }

        /// <summary>
        /// 将 Word 光标定位到指定页。
        /// </summary>
        /// <param name="app"></param>
        /// <param name="page">目标页码</param>
        public static void SpacePagePosition(this object app, int page)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            var appType = app.GetType();
            var doc = appType.InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, app, null);
            var docType = doc.GetType();

            // 调用 Document.GoTo(wdGoToPage, wdGoToAbsolute, page)
            docType.InvokeMember("GoTo", BindingFlags.InvokeMethod, null, doc,
                new[] { WdGoToPage, WdGoToAbsolute, page, Type.Missing });
        }

        /// <summary>
        /// 删除指定页码的空白页。
        /// </summary>
        /// <param name="app"></param>
        /// <param name="pages">需要删除的页码列表</param>
        public static void RemoveSpacePages(this object app, List<int> pages)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));
            if (pages == null || pages.Count == 0) return;

            var appType = app.GetType();
            var doc = appType.InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, app, null);
            var docType = doc.GetType();

            // 1. 获取文档总页数
            var allPagesObj = docType.InvokeMember("ComputeStatistics", BindingFlags.InvokeMethod, null, doc,
                new object[] { WdStatisticPages });
            var allPages = Convert.ToInt32(allPagesObj);

            // 2. 获取整个文档范围对象 Content（删除页时用于末尾对齐）
            var allRange = docType.InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);

            // 按页码倒序删除，避免页数变化导致错位
            pages.Sort();
            pages.Reverse();

            try
            {
                var selection = appType.InvokeMember("Selection", BindingFlags.GetProperty, null, app, null);
                var selType = selection.GetType();

                var allRangeType = allRange.GetType();
                var allEndObj = allRangeType.InvokeMember("End", BindingFlags.GetProperty, null, allRange, null);
                var allEnd = Convert.ToInt32(allEndObj);

                foreach (var page in pages)
                {
                    // 跳转到当前页
                    var range = selType.InvokeMember("GoTo", BindingFlags.InvokeMethod, null, selection,
                        new[] { WdGoToPage, WdGoToAbsolute, page, Type.Missing });
                    var rangeType = range.GetType();

                    // 设置该页的结束位置
                    if (page == allPages)
                    {
                        rangeType.InvokeMember("End", BindingFlags.SetProperty, null, range, new object[] { allEnd });
                    }
                    else
                    {
                        // 获取下一页开始位置
                        var nextPage = selType.InvokeMember("GoTo", BindingFlags.InvokeMethod, null, selection,
                            new[] { WdGoToPage, WdGoToAbsolute, page + 1, Type.Missing });
                        var nextPageType = nextPage.GetType();
                        var nextStartObj =
                            nextPageType.InvokeMember("Start", BindingFlags.GetProperty, null, nextPage, null);
                        var nextStart = Convert.ToInt32(nextStartObj);
                        rangeType.InvokeMember("End", BindingFlags.SetProperty, null, range,
                            new object[] { nextStart - 1 });
                    }

                    // 删除该页内容
                    rangeType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, range, null);
                }
            }
            catch
            {
                // 忽略异常，确保不会影响其他页
            }
        }
    }
}
