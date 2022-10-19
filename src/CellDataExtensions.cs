using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsHelper
{
	/// <summary>
	/// Extension methods for <see cref="CellData"/> objects
	/// </summary>
	public static class CellDataExtensions
	{
		/// <summary>
		/// Sets the background color of a cell
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="red">Integer value of the red value (1 to 255)</param>
		/// <param name="green">Integer value of the green value (1 to 255)</param>
		/// <param name="blue">Integer value of the blue value (1 to 255)</param>
		public static CellData SetBackgroundColor(this CellData cell, int red, int green, int blue)
			=> cell.SetBackgroundColor(System.Drawing.Color.FromArgb(red, green, blue));

		/// <summary>
		/// Sets the background color of a cell
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="color"></param>
		public static CellData SetBackgroundColor(this CellData cell, System.Drawing.Color color)
			=> cell.SetBackgroundColor(color.ToGoogleColor());

		/// <summary>
		/// Sets the background color of a cell
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="color"></param>
		public static CellData SetBackgroundColor(this CellData cell, Color color)
		{
			CellFormat format = GetOrCreateCellFormat(cell);
			format.BackgroundColor = color;
			return cell;
		}

		/// <summary>
		/// Sets the text formatting to be bold or not
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="useBold"></param>
		public static CellData SetBoldText(this CellData cell, bool useBold)
		{
			CellFormat cellFormat = GetOrCreateCellFormat(cell);
			TextFormat textFormat = GetOrCreateTextFormat(cellFormat);
			textFormat.Bold = useBold;
			return cell;
		}

		/// <summary>
		/// Sets the foreground (text) color of a cell
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="red">Integer value of the red value (1 to 255)</param>
		/// <param name="green">Integer value of the green value (1 to 255)</param>
		/// <param name="blue">Integer value of the blue value (1 to 255)</param>
		public static CellData SetForegroundColor(this CellData cell, int red, int green, int blue)
			=> cell.SetForegroundColor(System.Drawing.Color.FromArgb(red, green, blue));

		/// <summary>
		/// Sets the foreground (text) color of a cell
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="color"></param>
		public static CellData SetForegroundColor(this CellData cell, System.Drawing.Color color)
			=> cell.SetBackgroundColor(color.ToGoogleColor());

		/// <summary>
		/// Sets the foreground (text) color of a cell
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="color"></param>
		public static CellData SetForegroundColor(this CellData cell, Color color)
		{
			CellFormat cellFormat = GetOrCreateCellFormat(cell);
			TextFormat textFormat = GetOrCreateTextFormat(cellFormat);
			textFormat.ForegroundColor = color;
			return cell;
		}

		/// <summary>
		/// Sets the horizontal alignment (justification) of a cell
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="alignment"></param>
		public static CellData SetHorizontalAlignment(this CellData cell, HorizontalAlignment alignment)
		{
			CellFormat cellFormat = GetOrCreateCellFormat(cell);
			cellFormat.HorizontalAlignment = alignment.ToString();
			return cell;
		}

		/// <summary>
		/// Sets the number format for a cell (also works for dates), see https://developers.google.com/sheets/api/guides/formats
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="formatString"></param>
		public static CellData SetNumberFormat(this CellData cell, string formatString)
		{
			CellFormat cellFormat = GetOrCreateCellFormat(cell);
			NumberFormat numberFormat = GetOrCreateNumberFormat(cellFormat);
			numberFormat.Pattern = formatString;
			return cell;
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
