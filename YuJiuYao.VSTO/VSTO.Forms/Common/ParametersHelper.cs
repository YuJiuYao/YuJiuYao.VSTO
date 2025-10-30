using VSTO.Common;

namespace VSTO.Forms.Common
{
    internal class ParametersHelper : ParametersBaseHelper
    {
        /// <summary>
        /// 用户数据文件目录，cookie、web缓存。软件运行目录\Temp\WebUserData
        /// </summary>
        public static string WebUserDataDir = BaseDirectoryPath + "Temp\\WebUserData\\";

        /// <summary>
        ///  web静态文件缓存地址，手动管理。静态文件根据版本号管理。软件运行目录\Cache\Web\域名\文件名.v0
        /// </summary>
        public static string WebCacheDir = BaseDirectoryPath + "Cache\\Web\\";

        /// <summary>
        /// 用户下载软件文件所在位置（升级，下载之类的，用完切记删除掉）
        /// </summary>
        public static string DownLoadFilePath = BaseDirectoryPath + "Temp\\DownLoad\\";
    }
}
