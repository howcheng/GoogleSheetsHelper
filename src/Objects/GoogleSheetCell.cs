using System;
using System.Drawing;

/// <summary>
/// Represents a single cell in a Google Sheet
/// </summary>
namespace GoogleSheetsHelper
{
	public class GoogleSheetCell
	{
		/// <summary>Gets or sets the string value of the cell</summary>
		public string StringValue { get; set; }

		/// <summary>Gets or sets the numeric value of a cell</summary>
		public double? NumberValue { get; set; }

		/// <summary>Gets or sets the Boolean value of a cell</summary>
		public bool? BoolValue { get; set; }

		/// <summary>Gets or sets the date/time value of a cell</summary>
		public DateTime? DateTimeValue { get; set; }

		/// <summary>Gets or sets the formula value of a cell</summary>
		public string FormulaValue { get; set; } 

		/// <summary>Gets or sets the number format of a cell: see https://developers.google.com/sheets/api/guides/formats </summary>
		public string NumberFormat { get; set; }

		/// <summary>Gets or sets the date/time format of a cell: see https://developers.google.com/sheets/api/guides/formats </summary>
		public string DateTimeFormat { get; set; }

		/// <summary>Flag indicating that the cell contents are in boldface</summary>
		public bool? Bold { get; set; }

		/// <summary>Gets or sets the background color of a cell</summary>
		public Color? BackgroundColor { get; set; }

		/// <summary>Gets or sets the foreground (text) color of a cell</summary>
		public Color? ForegroundColor { get; set; }

		/// <summary>Get or sets the horizontal alignment of the cell contents/summary>
		public HorizontalAlignment? HorizontalAlignment { get; set; }

		public GoogleSheetCell()
		{
		}

		public GoogleSheetCell(string value)
		{
			StringValue = value;
		}

		public GoogleSheetCell(double? value)
		{
			NumberValue = value;
		}

		public GoogleSheetCell(bool? value)
		{
			BoolValue = value;
		}

		public GoogleSheetCell(DateTime? value)
		{
			DateTimeValue = value;
		}

		public static GoogleSheetCell Create(object value)
		{
			if (value == null) return null;

			switch (value)
			{
				case string s:
					return new GoogleSheetCell(s);
				case int i:
					return new GoogleSheetCell(i);
				case double d:
					return new GoogleSheetCell(d);
				case bool b:
					return new GoogleSheetCell(b);
				case DateTime dt:
					return new GoogleSheetCell(dt);
			}
			return new GoogleSheetCell(value.ToString());
		}

		public object CellValue
		{
			get => StringValue ?? FormulaValue ?? NumberValue ?? (object)BoolValue ?? DateTimeValue;
		}
	}
}
