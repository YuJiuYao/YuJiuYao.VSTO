using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Newtonsoft.Json;
using VSTO.Extension;
using VSTO.Forms.BaseForms;
using VSTO.Models;

namespace VSTO.Forms.Common
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class ClientPortal
    {
        private readonly WebFormView _webFormView;

        /// <summary>
        /// 构造函数，如果webFormView没用的话，建议删除
        /// </summary>
        /// <param name="webFormView"></param>
        public ClientPortal(WebFormView webFormView)
        {
            _webFormView = webFormView;
        }

        /// <summary>
        /// 和前端交互的方法
        /// </summary>
        /// <param name="action"></param>
        /// <param name="data"></param>
        /// <param name="evalJs"></param>
        /// <returns></returns>
        public string ExecuteJs(string action, string data, string evalJs = "")
        {
            try
            {
                switch (action)
                {
                    case "GetDate":
                        return JsonConvert.SerializeObject(ApiResult<DateTime>.Success(DateTime.Now));
                    case "InsertText":
                        InsertTextAtSelection(data ?? string.Empty);
                        if (!string.IsNullOrWhiteSpace(evalJs))
                        {
                            // 回调前端
                            _webFormView?.WebView2?.CoreWebView2?.ExecuteScriptAsync(evalJs);
                        }
                        return JsonConvert.SerializeObject(ApiResult<string>.Success("ok"));
                    default:
                        return JsonConvert.SerializeObject(ApiResult<string>.Fail("未知操作"));
                }
            }
            catch (Exception ex)
            {
                return JsonConvert.SerializeObject(ApiResult<string>.Fail(ex.Message));
            }
        }

        private void InsertTextAtSelection(string text)
        {
            //通过反射进行 COM 后期绑定，避免添加 Office 引用或 dynamic依赖
            var app = GetWordApp();
            if (app == null) throw new InvalidOperationException("未找到Word进程");

            var appType = app.GetType();
            // 获取 Selection 对象
            var selection = appType.InvokeMember("Selection", BindingFlags.GetProperty, null, app, null);
            if (selection == null) throw new InvalidOperationException("无法获取选择对象");

            var selType = selection.GetType();
            // 调用 TypeText 方法
            selType.InvokeMember("TypeText", BindingFlags.InvokeMethod, null, selection, new object[] { text });

            MessageBox.Show($@"当前页数：{app.GetActiveDocumentPageCount()}");
            app.SpacePagePosition(app.GetActiveDocumentPageCount() - 5);
        }

        private object GetWordApp()
        {
            try
            {
                return Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                return null;
            }
        }

    }
}
