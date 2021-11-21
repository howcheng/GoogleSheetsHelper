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
		FixColumn = 1,
		FixRow = 2,
	}

	public enum RepeatDirection
	{
		Vertical,
		Horizontal
	}
}
