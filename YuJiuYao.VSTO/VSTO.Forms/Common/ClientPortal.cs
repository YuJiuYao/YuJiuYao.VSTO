using System;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
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
        /// 构造函数，如果ctl没用的话，建议删除
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
        public string GetData(string action, string data, string evalJs = "")
        {
            string result;
            switch (action)
            {
                case "GetDate":
                    result = JsonConvert.SerializeObject(ApiResult<DateTime>.Success(DateTime.Now));
                    break;
                default:
                    result = JsonConvert.SerializeObject(ApiResult<string>.Fail("获取失败"));
                    break;
            }
            return result;
        }

    }
}
