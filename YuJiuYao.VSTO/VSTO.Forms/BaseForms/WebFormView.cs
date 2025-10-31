using System;
using System.Collections.Concurrent;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;

namespace VSTO.Forms.BaseForms
{
    public partial class WebFormView : UserControl
    {
        public string Url { get; set; }

        public bool UserAgent { get; set; }

        /// <summary>
        /// 给JS提供的接口。比如：new DomInterop()
        /// <para>不传递，则不提供接口</para>
        /// <para>自建一个类，该类注解：[ClassInterface(ClassInterfaceType.AutoDual)][ComVisible(true)]，然后在该类中写个方法，比如：Func(string action,string data,string evalJs)</para>
        /// <para>则js在调用的时候，使用：window.chrome.webview.hostObjects.sync.CallClient.Func(msg,'66',"funcX")</para>
        /// </summary>
        public object JsApi { get; set; }

        /// <summary>
        /// 返回控件的webview2
        /// </summary>
        public Microsoft.Web.WebView2.WinForms.WebView2 WebView2 => webView;

        /// <summary>
        /// 传递自定义任务窗格的Visible事件，以便于释放WebView2及截图
        /// <para>可以不用传递</para>
        /// </summary>
        public EventHandler TaskPaneVisibleEv { get; set; }

        /// <summary>
        /// webview2的用户数据文件目录，cookie、web缓存。软件运行目录\Temp\WebUserData
        /// </summary>
        private static string _webUserDataDir = AppDomain.CurrentDomain.BaseDirectory + "Temp\\WebUserData\\";

        /// <summary>
        /// web静态文件缓存地址，手动管理。静态文件根据版本号管理。软件运行目录\Cache\Web\域名\文件名.v0
        /// </summary>
        private static string _webCacheDir = AppDomain.CurrentDomain.BaseDirectory + "Cache\\Web\\";

        /// <summary>
        /// 缓存文件缓存到内存中，一个项目大概十来M静态文件，完全可以缓存到内存重
        /// <para>key：文件物理路径。value：文件流</para>
        /// </summary>
        private ConcurrentDictionary<string, MemoryStream> _dicCacheInMemo;

        /// <summary>
        /// 内存缓存web截图，用户打开窗口无感
        /// </summary>
        private ConcurrentDictionary<string, Bitmap> _dicWebShot;

        /// <summary>
        /// 是否开启Debug模式
        /// </summary>
        private int _isDebugModel;

        /// <summary>
        /// 隐藏于UserAgent中的版本号
        /// </summary>
        private string _version;


        public WebFormView()
        {
            InitializeComponent();
        }

        public void InitForm(ConcurrentDictionary<string, MemoryStream> dicCacheInMemo,
            ConcurrentDictionary<string, Bitmap> dicWebShot, object jsApi, int isDebugModel, string version = "8.9.2")
        {
            _dicCacheInMemo = dicCacheInMemo;
            _dicWebShot = dicWebShot;
            JsApi = jsApi;
            TaskPaneVisibleEv += TaskPaneVisibleChange;
            _isDebugModel = isDebugModel;
            _version = version;
        }

        private async void WebFormView_Load(object sender, EventArgs e)
        {
            //销毁进程
            if (Parent != null)
                //销毁多线程webview2，避免占用过多的内容
                //副作用就是速度不够快吧，也可能是官方的BUG，目前为止
                Parent.Disposed += Parent_Disposed;
            //Debug.WriteLine("控件加载");
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
            if (DesignMode) return;
            if (!string.IsNullOrEmpty(Url)) GetWebShot(Url); //放在初始化之前，因为初始化也会白屏一段时间
            var env = await CoreWebView2Environment.CreateAsync(null, _webUserDataDir, null);
            //异步设置使变量生效。
            try
            {
                await webView.EnsureCoreWebView2Async(env);
            }
            catch //页面未加载关闭窗口，会报：System.Runtime.InteropServices.COMException:“已中止操作 (异常来自 HRESULT:0x80004004 (E_ABORT))”
            {
                return;
            }

            //必须在初始化完成后加载
            if (!string.IsNullOrEmpty(Url))
                try
                {
                    webView.Source = new Uri(Url, UriKind.Absolute);
                }
                catch
                {
                    MessageBox.Show(@"您的网络存在问题，请检查网络状况或重新启动Word!");
                }

            //是否开启DEBUG模式
            if (_isDebugModel == 1)
            {
                //开启DEBUG模式
                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
                webView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = true;
            }

            #region 打包的时候需要注释

#if DEBUG
            webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            webView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = true;
#endif

            #endregion

            //屏幕小的时候，采用移动端模式
            if (UserAgent)
                webView.CoreWebView2.Settings.UserAgent =
                    $"Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Bider/{_version} Version/13.0.3 Mobile/15E148 Safari/604.1";
        }

        /// <summary>
        /// 浏览器禁用内置功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webView_NavigationStarting(object sender, CoreWebView2NavigationStartingEventArgs e)
        {
            try
            {
                if (webView.CoreWebView2 != null)
                {
                    //禁用网页右键功能
                    webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;

                    //禁用开发工具，设为true，可通过F12打开开发者调试功能
                    webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                    //禁用网页的放大缩小功能
                    webView.CoreWebView2.Settings.IsZoomControlEnabled = true;

                    //禁用网页上加载页面的进度条功能
                    webView.CoreWebView2.Settings.IsStatusBarEnabled = true;
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// 获取该控件的顶级控件，比如Form窗体。如果不在Form窗体，则返回该控件自身
        /// </summary>
        /// <returns></returns>
        public Control GetTopControl()
        {
            Control topCtl = this;
            while (topCtl.Parent != null) topCtl = topCtl.Parent;
            return topCtl;
        }

        /// <summary>
        /// 跳转到URL
        /// <para>同时注意：如果没有初始化，则会自动初始化（比如默认用户数据文件位置为默认位置，后续不可更改），所以，必须先初始化</para>
        /// <para>而初始化设置，已经在用户控件加载时设置。所以，此处变量的设置在：先加载后设置。</para>
        /// </summary>
        /// <param name="url"></param>
        public void ToUrl(string url)
        {
            Url = url;
            webView.CoreWebView2.Navigate(url);
        }

        /// <summary>
        /// 销毁webview2，释放内存，这个主要是CustomTaskPane
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TaskPaneVisibleChange(object sender, EventArgs e)
        {
            try
            {
                var senderPropertyInfo = sender.GetType().GetProperty("Visible");
                if (senderPropertyInfo == null) return;
                var isVisible = (bool)senderPropertyInfo.GetValue(sender);
                if (isVisible) return;
                if (!webView.Disposing && !webView.IsDisposed) webView.Dispose();
                //var pane = (Microsoft.Office.Tools.CustomTaskPaneImpl)sender;
            }
            catch
            {
                // ignored
            }
        }

        /// <summary>
        /// 放置遮挡web快照的图片
        /// </summary>
        /// <param name="formName">窗口名，一般传递url地址来定义窗口名</param>
        private void GetWebShot(string formName)
        {
            if (string.IsNullOrEmpty(formName))
            {
                pictureBox.Hide();
                return;
            }

            if (_dicWebShot.TryGetValue(formName, out var value))
                pictureBox.Image = value;
            else
                pictureBox.Hide(); //首次加载，隐藏遮挡框
        }

        /// <summary>
        /// 销毁webview2，释放内存，这个主要是Form的
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
    }
}