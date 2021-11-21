using System;
using System.Collections.Generic;
using System.Text;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsHelper
{
	public static class CellDataExtensions
	{
		public static void SetBackgroundColor(this CellData cell, int red, int green, int blue)
			=> cell.SetBackgroundColor(System.Drawing.Color.FromArgb(red, green, blue));

		public static void SetBackgroundColor(this CellData cell, System.Drawing.Color color)
		{
			CellFormat format = GetOrCreateCellFormat(cell);
			format.BackgroundColor = color.ToGoogleColor();
		}

		public static void SetBoldText(this CellData cell, bool useBold)
		{
			CellFormat cellFormat = GetOrCreateCellFormat(cell);
			TextFormat textFormat = GetOrCreateTextFormat(cellFormat);
			textFormat.Bold = useBold;
		}

		public static void SetForegroundColor(this CellData cell, int red, int green, int blue)
			=> cell.SetForegroundColor(System.Drawing.Color.FromArgb(red, green, blue));

		public static void SetForegroundColor(this CellData cell, System.Drawing.Color color)
		{
			CellFormat cellFormat = GetOrCreateCellFormat(cell);
			TextFormat textFormat = GetOrCreateTextFormat(cellFormat);
			textFormat.ForegroundColor = color.ToGoogleColor();
		}

		public static void SetHorizontalAlignment(this CellData cell, HorizontalAlignment alignment)
		{
			CellFormat cellFormat = GetOrCreateCellFormat(cell);
			cellFormat.HorizontalAlignment = alignment.ToString();
		}

		/// <summary>
		/// Sets the number format for a cell (also works for dates), see https://developers.google.com/sheets/api/guides/formats
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="formatString"></param>
		public static void SetNumberFormat(this CellData cell, string formatString)
		{
			CellFormat cellFormat = GetOrCreateCellFormat(cell);
			NumberFormat numberFormat = GetOrCreateNumberFormat(cellFormat);
			numberFormat.Pattern = formatString;
		}

		private static CellFormat GetOrCreateCellFormat(CellData cell)
		{
			if (cell.UserEnteredFormat == null)
				cell.UserEnteredFormat = new CellFormat();
			return cell.UserEnteredFormat;
		}

		private static TextFormat GetOrCreateTextFormat(CellFormat cellFormat)
		{
			if (cellFormat.TextFormat == null)
				cellFormat.TextFormat = new TextFormat();
			return cellFormat.TextFormat;
		}

		private static NumberFormat GetOrCreateNumberFormat(CellFormat cellFormat)
		{
			if (cellFormat.NumberFormat == null)
				cellFormat.NumberFormat = new NumberFormat { Type = "number" };
			return cellFormat.NumberFormat;
		}
	}
}
