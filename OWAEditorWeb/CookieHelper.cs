using System;
using System.Web;

namespace OWAEditorWeb
{
    /// <summary>
    /// 对Cookie操作的封装
    /// </summary>
    public class CookieHelper
    {
        private HttpContext context;		//获取HTTP请求的信息
        private HttpCookie acookie;			//cookie类
        private int operateState = 0;

        public CookieHelper()
        {
            context = HttpContext.Current;
        }

        /// <summary>
        /// 判断客户端是否支持使用Cookie
        /// </summary>
        /// <returns>返回的是布尔型变量,用来判断客户端是否支持Cookie</returns>
        public bool Estop()
        {
            //判断客户端是否支持Cookies
            if (context.Request.Browser.Cookies)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 创建一个有过期时间的cookie
        /// </summary>
        /// <param name="cookieName">cookie名</param>
        /// <param name="cookieValue">cookie值</param>
        /// <param name="time">cookie的过期时间</param>
        /// <returns>返回operateState代表操作成功
        ///	返回operateState + 3代表客户端不支持Cookie</returns>
        public int CreateCookie(string cookieName, string cookieValue, DateTime time)
        {
            //判断客户端是否支持Cookie
            if (Estop() == false)
            {
                return operateState + 3;
            }
            acookie = new HttpCookie(cookieName, cookieValue);
            acookie.Expires = time;

            context.Response.SetCookie(acookie);

            //context.Response.SetCookie(acookie);
            return 0;
        }

        /// <summary>
        /// 添加一个无日期的cookie
        /// </summary>
        /// <param name="cookieName">cookie名</param>
        /// <param name="cookieValue">cookie值</param>
        /// <returns>返回operateState代表操作成功
        ///	返回operateState + 3代表客户端不支持Cookie</returns>
        public int CreateCookie(string cookieName, string cookieValue)
        {
            //判断客户端是否支持Cookie
            if (Estop() == false)
            {
                return operateState + 3;
            }

            acookie = new HttpCookie(cookieName, cookieValue);

            context.Response.SetCookie(acookie);

            return 0;
        }

        /// <summary>
        ///  判断指定的cookie是否存在
        /// </summary>
        /// <param name="cookieName">cookie名</param>
        /// <param name="being">输出型判断是否存在的布尔值</param>
        /// <returns>返回operateState代表操作成功
        ///	返回operateState + 3代表客户端不支持Cookie</returns>
        public int Exists(string cookieName, out bool being)
        {
            being = false;
            //判断客户端是否支持Cookie
            if (Estop() == false)
            {
                return operateState + 3;
            }

            acookie = context.Request.Cookies.Get(cookieName);
            if (acookie != null && acookie.Value != "")
                being = true;
            else
                being = false;
            return 0;
        }

        /// <summary>
        /// 获取指定的cookie的值
        /// </summary>
        /// <param name="cookieName">cookie名</param>
        /// <param name="cookieValue">输出型的cookie的值</param>
        /// <returns>返回operateState代表操作成功,返回operateState + 2代表指定的cookie不存在
        ///	返回operateState + 3代表客户端不支持Cookie</returns>
        public int GetCookieValue(string cookieName, out string cookieValue)
        {
            cookieValue = null;
            //判断客户端是否支持Cookie
            if (Estop() == false)
            {
                return operateState + 3;
            }

            bool being;
            //调用本类的判断是否有值的函数来判断指定cookie是否有值
            Exists(cookieName, out being);

            //如果cookie存在就获取cookie的值,否则返回错误编码2
            if (being == true)
            {
                cookieValue = context.Request.Cookies.Get(cookieName).Value;
            }
            else
            {
                cookieValue = null;
                return operateState + 2;
            }
            return 0;
        }

        /// <summary>
        /// 删除一个指定的Cookie
        /// </summary>
        /// <param name="cookieName">cookie名</param>
        /// <returns>返回operateState代表操作成功
        ///	返回operateState + 3代表客户端不支持Cookie</returns>
        public int DeleteCookie(string cookieName)
        {
            //判断客户端是否支持Cookie
            if (Estop() == false)
            {
                return operateState + 3;
            }
            CreateCookie(cookieName, "");
            return 0;
        }

        /// <summary>
        /// 获得所有的cookie名
        /// </summary>
        /// <param name="cookieName">输出型参数用户获得所有cookie,添加时可以给cookie名起有规律的,取的时候就可以判断分开不同项目之间的cookie</param>
        /// <returns></returns>
        public int GetAllCookieName(out string[] cookieName)
        {
            HttpCookieCollection allCookie = context.Request.Cookies;
            System.Collections.IEnumerator e = allCookie.GetEnumerator();
            cookieName = new string[allCookie.Count];
            for (int i = 0; i < allCookie.Count; i++)
            {
                e.MoveNext();
                if (e.Current.ToString() != "ASP.NET_SessionId")
                {
                    cookieName[i] = e.Current.ToString();
                }
            }
            return operateState;
        }
    }
}