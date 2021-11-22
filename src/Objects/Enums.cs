using System;

namespace GoogleSheetsHelper
{
	public enum HorizontalAlignment
	{
		Left,
		Right,
		Center
	}

	[Flags]
	public enum CellRangeOptions
	{
		None = 0,
		/// <summary>
		/// Indicates that the column should be fixed, i.e., $A
		/// </summary>
		FixColumn = 1,
		/// <summary>
		/// Indicates that the row number should be fixed, i.e. $10
		/// </summary>
		FixRow = 2,
	}

	public enum RepeatDirection
	{
		Vertical,
		Horizontal
	}
}
