using System;
using System.Threading.Tasks;
using VSTO.Common;

namespace VSTO.Forms.Common
{
    internal class FuncHelper : FuncBaseHelper
    {
        /// <summary>
        /// 公用无参委托
        /// </summary>
        internal delegate void YzDelegate();

        /// <summary>
        /// 延迟执行函数
        /// </summary>
        /// <param name="yzDelegate">要执行的方法</param>
        /// <param name="delayMillSec">延迟x秒后执行</param>
        internal static void DelayExe(YzDelegate yzDelegate, int delayMillSec)
        {
            Task.Factory.StartNew(() =>
            {
                System.Threading.Thread.Sleep(delayMillSec);
                yzDelegate?.Invoke();
            });
        }

        /// <summary>
        /// 根据URL获取本地缓存文件路径
        /// </summary>
        /// <param name="webCacheDir"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static string GetCacheFilePath(string webCacheDir, string url)
        {
            try
            {
                var uri = new Uri(url, UriKind.Absolute);
                return webCacheDir + uri.Host + "\\" + GetFileName(url);
            }
            catch (Exception)
            {
                // LogHelper.Instance.Error(ex);
            }
            return string.Empty;
        }
    }
}
