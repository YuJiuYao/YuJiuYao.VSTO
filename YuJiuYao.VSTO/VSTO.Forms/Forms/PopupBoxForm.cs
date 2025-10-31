using System;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using VSTO.Forms.BaseForms;

namespace VSTO.Forms.Forms
{
    public partial class PopupBoxForm : Form
    {
        private readonly string _url;

        private readonly object _jsApi;

        private static UserControlView _userControlView;


        public PopupBoxForm(string url, object jsApi)
        {
            _url = url;
            _jsApi = jsApi;
            Init();
            InitializeComponent();
        }

        /// <summary>
        /// 窗体初始化
        /// </summary>
        private void Init()
        {
            _userControlView = new UserControlView(_url, _jsApi)
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_userControlView);
        }

        /// <summary>
        /// 刷新
        /// </summary>
        /// <param name="methods"></param>
        public static void Reload(string methods)
        {
            if (_userControlView.WebView2.CoreWebView2 != null)
            {
                _userControlView.WebView2.Reload();
            }
        }

        /// <summary>
        /// 回调前端js
        /// </summary>
        /// <param name="methods"></param>
        public static void EvenJavaScript(string methods)
        {
            _userControlView.WebView2.CoreWebView2?.ExecuteScriptAsync(methods);
        }

        /// <summary>
        /// 打印当前界面为pdf
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="headerTitle"></param>
        public static async void PrintPdf(string filepath, string headerTitle)
        {
            try
            {
                if (_userControlView.WebView2.CoreWebView2 == null) return;
                var env = await CoreWebView2Environment.CreateAsync(null, AppDomain.CurrentDomain.BaseDirectory + "Temp\\WebUserData\\");

                var coreWebView2PrintSettings = env.CreatePrintSettings();

                coreWebView2PrintSettings.HeaderTitle = headerTitle;
                coreWebView2PrintSettings.MarginBottom = 0.6;
                coreWebView2PrintSettings.MarginTop = 0.6;
                coreWebView2PrintSettings.ShouldPrintBackgrounds = true;
                coreWebView2PrintSettings.ShouldPrintHeaderAndFooter = true;

                await _userControlView.WebView2.CoreWebView2.PrintToPdfAsync(filepath, coreWebView2PrintSettings);
            }
            catch (Exception)
            {
                // TODO 处理异常
            }
        }

    }
}
