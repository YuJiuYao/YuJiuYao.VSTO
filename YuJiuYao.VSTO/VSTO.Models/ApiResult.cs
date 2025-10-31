namespace VSTO.Models
{
    /// <summary>
    /// 表示api调用的结果。包括状态、提示信息、数据
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ApiResult<T>
    {
        /// <summary>
        /// 状态
        /// </summary>
        public bool Status { get; set; }

        /// <summary>
        /// 提示信息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 数据
        /// </summary>
        public T Data { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="status"></param>
        /// <param name="message"></param>
        /// <param name="data"></param>
        private ApiResult(bool status, string message, T data)
        {
            Status = status;
            Message = message;
            Data = data;
        }

        /// <summary>
        /// 获取成功的结果
        /// </summary>
        /// <param name="data">数据</param>
        /// <param name="message">数据获取成功提示信息。默认为 获取成功</param>
        /// <returns></returns>
        public static ApiResult<T> Success(T data, string message = "获取成功")
        {
            return new ApiResult<T>(true, message, data);
        }

        /// <summary>
        /// 获取失败的结果
        /// </summary>
        /// <param name="message">失败信息</param>
        /// <returns></returns>
        public static ApiResult<T> Fail(string message)
        {
            return new ApiResult<T>(false, message, default);
        }
    }
}