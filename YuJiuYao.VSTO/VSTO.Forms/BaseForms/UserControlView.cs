using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using VSTO.Common;
using VSTO.Forms.Common;

namespace VSTO.Forms.BaseForms
{
    public partial class UserControlView : UserControl
    {
        /// <summary>
        /// 页面跳转URL
        /// </summary>
        private readonly string _url;

        /// <summary>
        /// 1.自建一个类，该类注解：[ClassInterface(ClassInterfaceType.AutoDual)][ComVisible(true)]，然后在该类中写个方法，比如：Func(string action,string data,string evalJs)
        /// 2.则js在调用的时候，使用：window.chrome.webview.hostObjects.sync.CallClient.GetData(msg,'66',"funcX")
        /// </summary>
        private readonly object _jsApi;

        /// <summary>
        /// 返回当前控件的webview2对象(对于回调前端js方法)
        /// </summary>
        public WebView2 WebView2 => webView;

        public UserControlView(string url, object jsApi)
        {
            _url = url;
            _jsApi = jsApi;
            InitializeComponent();
        }


        /// <summary>
        /// 控件加载的时候初始化webview2
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void UserControlView_Load(object sender, EventArgs e)
        {
            try
            {
                if (Parent != null)
                    //销毁多线程webview2，避免占用过多的内容
                    //副作用就是速度不够快吧，也可能是官方的BUG，目前为止
                    Parent.Disposed += Parent_Disposed;

                /*
                 * 通过CoreWebView2Environment设置浏览器的环境变量。可设置在父控件的Load事件中，注意使用async异步（要求）。
                 * 务必注意，不能在设置属性Source后再设置环境变量。参见：https://stackoverflow.com/questions/62470733/set-cache-directory-for-webview2
                 * 变量1：browserExecutableFolder：可设置为null，默认会在执行的dll文件的同目录下创建一个runtimes文件夹，不用理会，当然也可以设置到其它地方。
                 * 变量2：userDataFolder（用户数据）：默认两个位置：
                 * 对于打包的 Windows 应用商店应用，默认用户文件夹是ApplicationData\LocalFolder包文件夹中的子文件夹。
                 * 对于现有的桌面应用程序，默认用户数据文件夹是exe+.WebView2.。比如VSTO则会在C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE.WebView2
                 * 参见：https://docs.microsoft.com/en-us/microsoft-edge/webview2/concepts/user-data-folder
                 * 而Office的安装目录具有特殊权限，我们的exe无权限，所以此处更改。
                 * 变量3：option，可不设置。
                 */
                //webview2在设计模式（设计器）下阻止初始化，问题出在CoreWebView2Environment.CreateAsync上，这是官方dll的bug，官方已收到反馈，正在处理。
                //参见：https://github.com/MicrosoftEdge/WebView2Feedback/issues/1497
                //参见：https://githubhot.com/repo/MicrosoftEdge/WebView2Feedback/issues/2046 （1个月前说下周出预发行版修复，但安装后仍未修复）。
                //此处使用DesignMode判断暂时规避，可以正常使用。
                await InitializeCoreWebView2Async();

                //Navigate to URI by setting Source property
                try
                {
                    webView.Source = new Uri(_url);

                    //设置前端访问客户端接口
                    webView.CoreWebView2.AddHostObjectToScript("CallClient", _jsApi);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($@"您的网络存在问题，请检查网络状况或重新启动Word!{ex.Message}");
                }

                if (ParametersBaseHelper.IsDebugModel)
                {
                    webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
                    webView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = true;
                }

#if DEBUG
                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
                webView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = true;
#endif
            }
            catch (Exception)
            {
                // TODO 处理异常
            }
        }

        /// <summary>
        /// 销毁webview2，释放内存，这个主要是Form的(带考虑是否需要)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Parent_Disposed(object sender, EventArgs e)
        {
            try
            {
                if (!webView.Disposing && !webView.IsDisposed) webView.Dispose();
            }
            catch (InvalidOperationException)
            {
                //无需记录，多窗口关闭过快
            }
        }

        /// <summary>
        /// 实现网站url加载的必须事件
        /// </summary>
        /// <returns></returns>
        private async Task InitializeCoreWebView2Async()
        {
            //Specify options regarding the coreView2 initialization process
            var webView2Environment =
                await CoreWebView2Environment.CreateAsync(userDataFolder: ParametersHelper.WebUserDataDir);

            //CoreWebView2 creation
            await webView.EnsureCoreWebView2Async(webView2Environment);
        }
    }
}