using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VSTO.Common
{
    public abstract class FuncBaseHelper
    {
        /// <summary>
        /// 数据写入配置文件INI
        /// </summary>
        /// <param name="section">写入的节点</param>
        /// <param name="key">写入的键</param>
        /// <param name="val">写入值</param>
        /// <returns></returns>
        public static bool SetIniValue(string section, string key, string val)
        {
            var i = WritePrivateProfileString(section, key, val, ParametersBaseHelper.UserSettingIniPath);
            return i > 0;
        }

        /// <summary>
        /// 获取INI文件
        /// </summary>
        /// <param name="section">写入的节点</param>
        /// <param name="key">写入的键</param>
        /// <returns>为空表示失败或者没有数据</returns>
        public static string GetIniValue(string section, string key)
        {
            var strB = new StringBuilder(500);
            //第三个参数表示没有值是返回的数据，此处没有值返回空“”
            GetPrivateProfileString(section, key, "", strB, 500, ParametersBaseHelper.UserSettingIniPath);
            return strB.ToString();
        }

        /// <summary>
        /// 根据文件地址删除文件
        /// </summary>
        /// <param name="filePath"></param>
        public static void DeleteFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch
            {
                // ignored
            }
        }

        /// <summary>
        /// 删除目录下所有文件
        /// </summary>
        /// <param name="directoryInfo"></param>
        public static void DeleteAllFile(DirectoryInfo directoryInfo)
        {
            try
            {
                //directoryInfo.Delete(true);不能这样直接删除文件夹，wps占用
                foreach (var item in directoryInfo.GetFiles())
                {
                    try
                    {
                        item.Delete();
                    }
                    catch
                    {
                        // ignored
                    }
                }
                foreach (var item in directoryInfo.GetDirectories())
                {
                    DeleteAllFile(item);
                }
            }
            catch
            {
                // ignored
            }
        }

        /// <summary>
        /// 根据文件的扩展名，判断是否需要存储的文件
        /// </summary>
        /// <param name="fileName">带扩展名的文件名</param>
        /// <param name="contentType">mime类型，有的文件没扩展名，但mime类型符合缓存逻辑</param>
        /// <param name="uri">URL地址，用于判断是否包含“v=”的参数，以便于进一步判断是否进行缓存</param>
        /// <returns></returns>
        public static bool NeedSaveByExt(string fileName, string contentType, Uri uri)
        {

            var ext = Path.GetExtension(fileName).ToLower();
            switch (ext)
            {
                case ".js":
                case ".css":
                case ".jpg":
                case ".jpeg":
                case ".bmp":
                case ".png":
                case ".gif":
                case ".ico":
                case ".svg":
                case ".ttf":
                case ".woff":
                case ".woff2":
                case ".eot":
                case ".mp3":
                case ".mp4":
                    return true;
                default:
                    //判断返回的媒体类型，有的时候返回的没扩展名，但是需要缓存的对象，比如c#动态打包的css文件，形如：https://officeweb365.com/Content/css?v=gV4hLx2wMNPtRlvNaYaW5R_tzyrnJr15qo4QyApcF4E1
                    //其类型为：content-type:text/css; charset=utf-8
                    //有一些业务图片，比如：https://bb.daxiit.net/compilation/file/img?fid=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJyZXNvdXJjZSI6
                    //它本身的Content-Type为image/jpg，但是是业务里的图片，业务里的图片可能很多，我们就不缓存了。交给浏览器自己处理吧
                    //这里总体约定：1、扩展名未知，2、mime类型需要缓存但是不包含v=标识，则认为不需要缓存
                    if (string.IsNullOrEmpty(contentType))
                    {
                        return false;
                    }
                    var uriHasV = UriHasV(uri.Query);
                    if (!uriHasV)
                    {
                        return false;//不包含v=的参数，则不缓存了
                    }
                    var mime = contentType.Split(';')[0];
                    switch (mime)
                    {
                        case "text/css":
                        case "application/x-javascript":
                        case "audio/mp3":
                        case "video/mpeg4":
                            return true;
                        default:
                            return mime.StartsWith("image") || mime.StartsWith("application/font");
                    }
            }
        }

        /// <summary>
        /// 判断URL地址中是否包含“v=”的参数，以便于进一步判断是否进行缓存
        /// </summary>
        /// <param name="queryStr">url的查询字符串</param>
        /// <returns></returns>
        public static bool UriHasV(string queryStr)
        {
            if (string.IsNullOrEmpty(queryStr))
            {
                return false;
            }
            var queryArr = queryStr.Split('&');
            return queryArr.Any(val => val.StartsWith("v="));
        }

        /// <summary>
        /// 从URI中获取请求的文件名，如果文件名太长，则返回空字符串。
        /// </summary>
        /// <param name="url">URL请求地址</param>
        /// <param name="hasVersion">是否包含版本号，若包含，则文件名后面增加 .v0</param>
        /// <returns></returns>
        protected static string GetFileName(string url, bool hasVersion = true)
        {
            try
            {
                var uri = new Uri(url, UriKind.Absolute);
                return GetFileName(uri, hasVersion);
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 从URI中获取请求的文件名，如果文件名太长，则返回空字符串。
        /// </summary>
        /// <param name="uri">URL请求地址</param>
        /// <param name="hasVersion">是否包含版本号，若包含，则文件名后面增加 .v0</param>
        /// <returns></returns>
        private static string GetFileName(Uri uri, bool hasVersion = true)
        {
            //获取文件名
            var fileName = Path.GetFileName(uri.AbsolutePath);
            //有的文件名太长。比如一些svg请求。指定的路径或文件名太长，或者两者都太长。完全限定文件名必须少于 260 个字符，并且目录名必须少于 248 个字符。
            if (string.IsNullOrEmpty(fileName) || fileName.Length > 128)
            {
                return string.Empty;
            }
            if (!hasVersion)
            {
                return fileName;
            }
            var queryStr = uri.Query.ToLower();
            if (queryStr.Contains("v="))
            {
                var pos0 = queryStr.IndexOf("v=", StringComparison.Ordinal);
                var pos1 = queryStr.IndexOf("&", pos0 + 1, StringComparison.Ordinal);
                var ver = pos1 > -1 ? queryStr.Substring(pos0 + 2, pos1 - pos0 - 2) : queryStr.Substring(pos0 + 2);
                return fileName + ".v" + ver;
            }
            else
            {
                return fileName + ".v0";
            }
        }

        /// <summary>
        /// 获取http响应头content-type类型
        /// </summary>
        /// <param name="fileExt">文件扩展名</param>
        /// <returns></returns>
        public static string GetContentType(string fileExt)
        {
            var ext = fileExt.StartsWith(".") ? fileExt.Substring(1) : fileExt;
            switch (ext.ToLower())
            {
                case "js":
                    return "content-type: application/x-javascript";
                case "css":
                    return "content-type: text/css";
                case "jpg":
                case "jpeg":
                case "bmp":
                case "png":
                case "gif":
                    return "content-type: image/" + ext;
                case "ico":
                    return "image/x-icon";
                case "svg":
                    return "content-type: text/xml";
                case "ttf":
                case "woff":
                case "woff2":
                case "eot":
                    return "content-type: application/font-" + ext;
                case "mp3":
                    return "content-type: audio/mp3";
                case "mp4":
                    return "content-type: video/mpeg4";
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// 获得活动窗口的句柄
        /// </summary>
        /// <returns></returns>
        [DllImport("user32.dll", EntryPoint = "GetActiveWindow")]
        public static extern int GetActiveWindow();

        /// <summary>
        /// 写入INI文件
        /// </summary>
        /// <param name="section">节点名称[如[TypeName]]</param>
        /// <param name="key">键</param>
        /// <param name="val">值</param>
        /// <param name="filepath">文件路径</param>
        /// <returns>Long，非零表示成功，零表示失败。会设置GetLastError</returns>
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);

        /// <summary>
        /// 读取INI文件
        /// </summary>
        /// <param name="section">节点名称</param>
        /// <param name="key">键</param>
        /// <param name="defVal">值</param>
        /// <param name="retVal">stringBuilder对象</param>
        /// <param name="size">字节大小</param>
        /// <param name="filepath">文件路径</param>
        /// <returns></returns>
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string defVal, StringBuilder retVal, int size, string filepath);
    }
}
