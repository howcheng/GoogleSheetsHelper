using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsHelper
{
	public static class Extensions
	{
		public static Color ToGoogleColor(this System.Drawing.Color color)
		{
			return new Color
			{
				Alpha = color.A / 255f,
				Blue = color.B / 255f,
				Green = color.G / 255f,
				Red = color.R / 255f
			};
		}
	}
}
