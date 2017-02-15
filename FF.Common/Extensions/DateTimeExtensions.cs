using System; 

namespace FF.Common.Extensions {
	/// <summary>
	/// Date time extension methods
	/// </summary>
	public static class DateTimeExtensions {

        /// <summary>
        /// 将DateTime时间换成中文
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="dateTime">时间</param>
        /// <returns>System.String.</returns>
        /// <example>
        /// 2016-12-21 12:12:21.012 → 1月前
        /// 2015-12-21 12:12:21.012 → 1年前
        /// </example>
        public static string ToChsStr(this DateTime dateTime)
        {
            var ts = DateTime.Now - dateTime;
            if ((int)ts.TotalDays >= 365)
                return (int)ts.TotalDays / 365 + "年前";
            if ((int)ts.TotalDays >= 30 && ts.TotalDays <= 365)
                return (int)ts.TotalDays / 30 + "月前";
            if ((int)ts.TotalDays == 1)
                return "昨天";
            if ((int)ts.TotalDays == 2)
                return "前天";
            if ((int)ts.TotalDays >= 3 && ts.TotalDays <= 30)
                return (int)ts.TotalDays + "天前";
            if ((int)ts.TotalDays != 0) return dateTime.ToString("yyyy年MM月dd日");
            if ((int)ts.TotalHours != 0)
                return (int)ts.TotalHours + "小时前";
            if ((int)ts.TotalMinutes == 0)
                return "刚刚";
            return (int)ts.TotalMinutes + "分钟前";
        }

        /// <summary>
        /// Truncate datetime, only keep seconds
        /// </summary>
        /// <param name="time">The time</param>
        /// <returns></returns>
        public static DateTime Truncate(this DateTime time) {
			return time.AddTicks(-(time.Ticks % TimeSpan.TicksPerSecond));
		}

		/// <summary>
		/// Return unix style timestamp
		/// Return a minus value if the time early than 1970-1-1
		/// The given time will be converted to UTC time first
		/// </summary>
		/// <param name="time">The time</param>
		/// <returns></returns>
		public static TimeSpan ToTimestamp(this DateTime time) {
			return time.ToUniversalTime() - new DateTime(1970, 1, 1);
		}
	}
}
