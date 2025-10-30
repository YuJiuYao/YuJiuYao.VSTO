using System;
using System.Collections.Concurrent;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace VSTO.Common
{
    public abstract class ParametersBaseHelper
    {
        /// <summary>
        /// 缓存文件缓存到内存中，一个项目大概十来M静态文件，完全可以缓存到内存重
        /// <para>key：文件物理路径。value：文件流</para>
        /// </summary>
        public static ConcurrentDictionary<string, MemoryStream> DicCacheInMemo = new ConcurrentDictionary<string, MemoryStream>();
        /// <summary>
        /// 内存缓存web截图，用户打开窗口无感
        /// </summary>
        public static ConcurrentDictionary<string, Bitmap> DicWebShot = new ConcurrentDictionary<string, Bitmap>();
        /// <summary>
        /// Webview2支持的最低版本
        /// <para>16.0.13127.20082</para>
        /// </summary>
        public static readonly Version LowestVersion = new Version("16.0.13127.20082");

        /// <summary>
        /// 电脑宽度
        /// </summary>
        public static readonly int WindowWidth = Screen.PrimaryScreen.Bounds.Width;

        /// <summary>
        /// 电脑高度
        /// </summary>
        public static readonly int WindowHeight = Screen.PrimaryScreen.Bounds.Height;

        /// <summary>
        /// 项目运行的目录
        /// </summary>
        protected static readonly string BaseDirectoryPath = AppDomain.CurrentDomain.BaseDirectory;

        /// <summary>
        /// 用户配置文件路径
        /// </summary>
        public static readonly string UserSettingIniPath = BaseDirectoryPath + "UserSetting.ini";

        /// <summary>
        /// 前端地址
        /// </summary>
        public static string FrontUrl = null;

        /// <summary>
        /// 后端地址
        /// </summary>
        public static string BackendUrl = null;

        /// <summary>
        /// 是否开启Debug模式，1开启，0关闭
        /// </summary>
        public static bool IsDebugModel = true;
    }
}