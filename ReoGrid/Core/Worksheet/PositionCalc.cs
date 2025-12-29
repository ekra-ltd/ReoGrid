/*****************************************************************************
 * 
 * ReoGrid - .NET Spreadsheet Control
 * 
 * https://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Author: Jingwood <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2025 Jingwood <jingwood at unvell.com>
 * Copyright (c) 2012-2025 UNVELL Inc. All rights reserved.
 * 
 ****************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

#if WINFORM || ANDROID
using RGFloat = System.Single;

#elif WPF
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;

#endif // WPF

using unvell.ReoGrid.Data;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid
{
	partial class Worksheet
	{
		#region Header
		internal int FindColIndexMiddle(RGFloat x)
		{
			return ArrayHelper.QuickFind(0, this.cols.Count, (i) =>
			{
				var col = this.cols[i];

				if (x > col.Left + col.InnerWidth / 2)
					return 1;

				if (i > 0)
				{
					var prevCol = this.cols[i - 1];

					if (x < prevCol.Left + prevCol.InnerWidth / 2)
					{
						return -1;
					}
				}

				return 0;
			});
		}

		internal int FindRowIndexMiddle(RGFloat x)
		{
			return ArrayHelper.QuickFind(0, this.rows.Count, (i) =>
			{
				var row = this.rows[i];

				if (x > row.Top + row.InnerHeight / 2)
					return 1;

				if (i > 0)
				{
					var prevCol = this.rows[i - 1];

					if (x < prevCol.Top + prevCol.InnerHeight / 2)
					{
						return -1;
					}
				}

				return 0;
			});
		}

        internal enum ColumnDirection
        {
            Left,
            Right,
        }

        internal class FindColumnByPositionResult
        {
            public bool IsInline { get; set; }
            public int Column { get; set; }
            public ColumnDirection Direction { get; set; }
        }

        internal bool FindColumnByPosition(RGFloat x, out int col)
        {
            var result = FindColumnByPositionWithResult(x);
            col = result.Column;
            return result.IsInline;
        }

        // TODO: need performance improvement
        internal FindColumnByPositionResult FindColumnByPositionWithResult(RGFloat x)
        {
            var result = new FindColumnByPositionResult { IsInline = true, Column = -1, Direction = ColumnDirection.Left, };

            RGFloat scaleThumb = 2 / renderScaleFactor;
            RGFloat scaleThumbHiddenColumn = 5 / renderScaleFactor;

            for (int i = 0; i < this.cols.Count; i++)
            {
                if (!cols[i].IsVisible)
                {
                    if (x <= cols[i].Right - scaleThumbHiddenColumn)
                    {
                        result.IsInline = false;
                        result.Column = i;
                        break;
                    }
                    else if (x <= cols[i].Right + scaleThumbHiddenColumn)
                    {
                        result.Column = i;
                        for (int j = i; j < cols.Count; j++)
                        {
                            if (cols[j].IsVisible) break;
                            result.Column = j;
                        }
                        result.Direction = ColumnDirection.Right;
                        break;
                    }
                }
                else
                {
                    if (x <= this.cols[i].Right - scaleThumb)
                    {
                        result.Column = i;
                        result.IsInline = false;
                        // Если следующий столбец - скрытый, то следует проверить его
                        if (i + 1 < cols.Count && !cols[i + 1].IsVisible)
                        {
                            if (x <= this.cols[i].Right - scaleThumbHiddenColumn)
                                result.IsInline = false;
                            else
                                result.IsInline = true;
                        }
                        break;
                    }
                    else if (x <= this.cols[i].Right + scaleThumb)
                    {
                        result.Column = i;
                        break;
                    }
                }
            }
            return result;
        }

        // TODO: need performance improvement
        internal FindRowByPositionResult FindRowByPositionWithResult(RGFloat y)
        {
            var result = new FindRowByPositionResult { IsInline = true, Row = -1, Direction = RowDirection.Top, };

            RGFloat scaleThumb = 2 / this.renderScaleFactor;
            RGFloat scaleThumbHiddenRow = 5 / this.renderScaleFactor;

            for (int i = 0; i < rows.Count; i++)
            {
                if (!rows[i].IsVisible)
                {
                    if (y <= rows[i].Bottom - scaleThumbHiddenRow)
                    {
                        result.IsInline = false;
                        result.Row = i;
                        break;
                    }
                    else if (y <= rows[i].Bottom + scaleThumbHiddenRow)
                    {
                        result.Row = i;
                        for (int j = i; j < rows.Count; j++)
                        {
                            if (rows[j].IsVisible) break;
                            result.Row = j;
                        }
                        result.Direction = RowDirection.Buttom;
                        break;
                    }
                }
                else
                {
                    if (y <= rows[i].Bottom - scaleThumb)
                    {
                        result.IsInline = false;
                        result.Row = i;
                        // Если следующая строка - скрытая, то следует проверить его
                        if (i + 1 < cols.Count && !cols[i + 1].IsVisible)
                        {
                            if (y <= rows[i].Bottom - scaleThumbHiddenRow)
                                result.IsInline = false;
                            else
                                result.IsInline = true;
                        }
                        break;
                    }
                    else if (y <= rows[i].Bottom + scaleThumb)
                    {
                        result.Row = i;
                        break;
                    }
                }
            }
            return result;
        }

        internal enum RowDirection
        {
            Top,
            Buttom,
        }


        internal class FindRowByPositionResult
        {
            public bool IsInline { get; set; }
            public int Row { get; set; }
            public RowDirection Direction { get; set; }
        }

        internal bool FindRowByPosition(RGFloat y, out int row)
        {
            var result = FindRowByPositionWithResult(y);
            row = result.Row;
            return result.IsInline;
        }
        #endregion // Header

        internal Rectangle GetRangeBounds(int row, int col, int rows, int cols)
		{
			return GetRangePhysicsBounds(new RangePosition(row, col, rows, cols));
		}
		internal Rectangle GetRangeBounds(CellPosition startPos, CellPosition endPos)
		{
			return GetRangePhysicsBounds(new RangePosition(startPos, endPos));
		}
		
		/// <summary>
		/// Get physics rectangle bounds from specified range position.
		/// Be careful that this is different from the rectangle bounds displayed on screen,
		/// the actual bound positions displayed on screen are transformed and scaled 
		/// in order to scroll, zoom and freeze into different viewports.
		/// </summary>
		/// <param name="range">The range position to get bounds</param>
		/// <returns>Rectangle bounds from specified range position</returns>
		public Rectangle GetRangePhysicsBounds(RangePosition range)
		{
			RangePosition fixedRange = FixRange(range);

			var rowHead = rows[fixedRange.Row];
			var colHead = cols[fixedRange.Col];
			var toRowHead = rows[fixedRange.EndRow];
			var toColHead = cols[fixedRange.EndCol];

			int width = toColHead.Right - colHead.Left;
			int height = toRowHead.Bottom - rowHead.Top;

			return new Rectangle(colHead.Left, rowHead.Top, width + 1, height + 1);
		}

		/// <summary>
		/// Get physics position from specified cell position.
		/// Be careful that this is different from the rectangle bounds displayed on screen,
		/// the actual bound positions displayed on the screen are transformed and scaled 
		/// in order to scroll, zoom and freeze into different viewports.
		/// </summary>
		/// <param name="row">Zero-based index of row</param>
		/// <param name="col">Zero-based index of column</param>
		/// <returns>Point position of specified cell position in pixel.</returns>
		public Point GetCellPhysicsPosition(int row, int col)
		{
			if (row < 0 || row >= this.rows.Count
				|| col < 0 || col >= this.cols.Count)
			{
				throw new ArgumentException("row or col invalid");
			}

			var rowHeader = this.rows[row].Top;
			var colHeader = this.cols[col].Left;

			return new Point(colHeader, rowHeader);
		}

		internal Rectangle GetScaledRangeBounds(RangePosition range)
		{
			var rowHead = rows[range.Row];
			var colHead = cols[range.Col];
			var toRowHead = rows[range.EndRow];
			var toColHead = cols[range.EndCol];

			RGFloat width = (toColHead.Right - colHead.Left) * this.renderScaleFactor;
			RGFloat height = (toRowHead.Bottom - rowHead.Top) * this.renderScaleFactor;

			return new Rectangle(colHead.Left * this.renderScaleFactor, rowHead.Top * this.renderScaleFactor, width, height);
		}


		internal Rectangle GetCellBounds(CellPosition pos)
		{
			return GetCellBounds(pos.Row, pos.Col);
		}

		internal Rectangle GetCellBounds(int row, int col)
		{
			if (cells[row, col] == null)
			{
				return GetCellRectFromHeader(row, col);
			}
			else if (cells[row, col].MergeStartPos != CellPosition.Empty)
			{
				Cell cell = GetCell(cells[row, col].MergeStartPos);
				return cell?.Bounds ?? GetCellRectFromHeader(row, col);
			}
			else
				return cells[row, col].Bounds;
		}

		private Rectangle GetCellRectFromHeader(int row, int col)
		{
			return new Rectangle(cols[col].Left, rows[row].Top, cols[col].InnerWidth + 1, rows[row].InnerHeight + 1);
		}	
	}
}
