using System;
using System.Text.RegularExpressions;

namespace FF.Common.Extensions {
	/// <summary>
	/// Decimal extension methods
	/// </summary>
	public static class DecimalExtensions {
		/// <summary>
		/// Remove excess 0s after decimal
		/// Eg: giving 12.3000 will return 12.3
		/// See: http://stackoverflow.com/questions/4298719/parse-decimal-and-filter-extra-0-on-the-right
		/// </summary>
		public static decimal Normalize(this decimal value) {
			return value / 1.000000000000000000000000000000000m;
		}

        /// <summary>
        /// Decimal 转换为大写汉字
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public static string ConvertToChinese(this Decimal number)
        {
            var s = number.ToString("#L#E#D#C#K#E#D#C#J#E#D#C#I#E#D#C#H#E#D#C#G#E#D#C#F#E#D#C#.0B0A");
            var d = Regex.Replace(s, @"((?<=-|^)[^1-9]*)|((?'z'0)[0A-E]*((?=[1-9])|(?'-z'(?=[F-L\.]|$))))|((?'b'[F-L])(?'z'0)[0A-L]*((?=[1-9])|(?'-z'(?=[\.]|$))))", "${b}${z}");
            var r = Regex.Replace(d, ".", m => "负元空零壹贰叁肆伍陆柒捌玖空空空空空空空分角拾佰仟万亿兆京垓秭穰"[m.Value[0] - '-'].ToString());

            return r;
        }
    }
}
