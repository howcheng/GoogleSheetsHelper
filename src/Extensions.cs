using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsHelper
{
	public static class Extensions
	{
		/// <summary>
		/// Converts a <see cref="System.Drawing.Color">System.Drawing.Color</see> object to a Google <see cref="Color"/> object
		/// </summary>
		/// <param name="color"></param>
		/// <returns></returns>
		public static Color ToGoogleColor(this System.Drawing.Color color) => new Color
		{
			Alpha = color.A / 255f,
			Blue = color.B / 255f,
			Green = color.G / 255f,
			Red = color.R / 255f
		};

		public static System.Drawing.Color ToSystemColor(this Color color)
			=> System.Drawing.Color.FromArgb(color.Alpha.FloatToByte(), color.Red.FloatToByte(), color.Green.FloatToByte(), color.Blue.FloatToByte());

		private static byte FloatToByte(this float? f) => (byte)Math.Floor((f ?? 1) * 255f);

		/// <summary>
		/// Compares a <see cref="Color"/> object with a <see cref="System.Drawing.Color">System.Drawing.Color</see> object to see if they 
		/// represent the same color
		/// </summary>
		/// <param name="color1"></param>
		/// <param name="color2"></param>
		/// <returns></returns>
		public static bool GoogleColorEquals(this Color color1, System.Drawing.Color color2)
			=> GoogleColorEquals(color1, color2.ToGoogleColor());

		/// <summary>
		/// Compares two <see cref="Color"/> objects to see if they represent the same color
		/// </summary>
		/// <param name="color1"></param>
		/// <param name="color2"></param>
		/// <returns></returns>
		public static bool GoogleColorEquals(this Color color1, Color color2)
			=> (color1.Alpha ?? 1) == (color2.Alpha ?? 1) // when retrieving from the sheet, the value may be null instead of the default (1 for Alpha, 0 for RGB)
				&& (color1.Blue ?? 0) == (color2.Blue ?? 0)
				&& (color1.Green ?? 0) == (color2.Green ?? 0)
				&& (color1.Red ?? 0) == (color2.Red ?? 0);

		public static string ToCamelCase(this string s) => $"{s.Substring(0, 1).ToLower()}{s.Substring(1)}";

		public static void AddRange<T>(this IList<T> list, IEnumerable<T> items)
		{
			foreach (T item in items)
			{
				list.Add(item);
			}
		}
	}
}
