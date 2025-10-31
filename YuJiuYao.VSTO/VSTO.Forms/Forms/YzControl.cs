using System;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using VSTO.Common;
using VSTO.Forms.Common;

namespace VSTO.Forms.Forms
{
    public partial class YzControl : UserControl
    {
        /// <summary>
        /// 当前模块下的浏览器控件
        /// </summary>
        public static WebView2 WebView2;

        /// <summary>
        /// UserControl 实例化时传递url已经前端参数，不可在load里面加载无效
        /// </summary>
        /// <param name="url">前端url路由</param>
        /// <param name="param">url携带的参数</param>
        /// <param name="isAgent"></param>
        public YzControl(string url = "", string param = "", bool isAgent = false)
        {
            InitializeComponent();
            try
            {
                webFormView.InitForm(ParametersBaseHelper.DicCacheInMemo, ParametersBaseHelper.DicWebShot, new ClientPortal(webFormView), ParametersBaseHelper.IsDebugModel ? 1 : 0);
                if (url.StartsWith("http"))
                {
                    webFormView.Url = url;
                }
                else
                {
                    webFormView.Url = ParametersBaseHelper.FrontUrl + url;
                }
                WebView2 = webFormView.WebView2;
            }
            catch (Exception ex)
            {
                // LogHelper.Instance.Error(ex);
            }
        }
    }
}
