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
 * Author: Jing Lu <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2021 Jing Lu <jingwood at unvell.com>
 * Copyright (c) 2012-2016 unvell.com, all rights reserved.
 * 
 ****************************************************************************/

#if FORMULA

using System;
using System.Collections.Generic;
using System.Linq;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Core.SerialFill;
using System.Collections.Generic;
using System.Linq;
using unvell.ReoGrid.Formula;

namespace unvell.ReoGrid
{
	partial class Worksheet
	{
		/// <summary>
		/// Auto fill specified serial in range.
		/// </summary>
		/// <param name="fromAddressOrName">Range to read filling rules.</param>
		/// <param name="toAddressOrName">Range to be filled.</param>
		public void AutoFillSerial(string fromAddressOrName, string toAddressOrName)
		{
			RangePosition fromRange, toRange;

#region fromRange
            if (this.TryGetNamedRange(fromAddressOrName, out var fromNRange))
			{
				fromRange = fromNRange.Position;
			}
			else if (RangePosition.IsValidAddress(fromAddressOrName))
			{
				fromRange = new RangePosition(fromAddressOrName);
			}
			else
			{
				throw new InvalidAddressException(fromAddressOrName);
			}
#endregion // fromRange

#region toRange
            if (this.TryGetNamedRange(toAddressOrName, out var toNRange))
			{
				toRange = toNRange.Position;
			}
			else if (RangePosition.IsValidAddress(toAddressOrName))
			{
				toRange = new RangePosition(toAddressOrName);
			}
			else
			{
				throw new InvalidAddressException(toAddressOrName);
			}
#endregion // toRange

			this.AutoFillSerial(fromRange, toRange);
		}

		/// <summary>
		/// Auto fill specified serial in range.
		/// </summary>
		/// <param name="fromRange">Range to read filling rules.</param>
		/// <param name="toRange">Range to be filled.</param>
		public void AutoFillSerial(RangePosition fromRange, RangePosition toRange)
		{
			fromRange = this.FixRange(fromRange);
			toRange = this.FixRange(toRange);

            #region Arguments Check
            if (fromRange.IntersectWith(toRange))
            {
                throw new ArgumentException("fromRange and toRange cannot being intersected.");
            }

			if (toRange != CheckMergedRange(toRange))
			{
				throw new ArgumentException("cannot change a part of merged range.");
			}
#endregion // Arguments Check

            List<CellPosition> fromCells, toCells;

            if (CheckRangeReadonly(toRange))
            {
                throw new RangeContainsReadonlyCellsException(toRange);
            }

			if (fromRange.Col == toRange.Col && fromRange.Cols == toRange.Cols)
            {
                // for (var c = toRange.Col; c <= toRange.EndCol; c++)
                // {
                //     fromCells = GetColumnCellPositionsFromRange(fromRange, c);
                //     toCells = GetColumnCellPositionsFromRange(toRange, c);
                //     AutoFillSerialCells(fromCells, toCells);
                // }
                try
                {
                    BeforeSerialFill?.Invoke(this, new Events.RangeSerialFillEventArgs(fromRange, toRange));
                }
                catch
                {
                    // ignored
                }
                ExecuteFillByAction(fromRange, toRange, ExecuteVerticalFill);
                try
                {
                    AfterSerialFill?.Invoke(this, new Events.RangeSerialFillEventArgs(fromRange, toRange));
                }
                catch
                {
                    // ignored
                }
            }
            else if (fromRange.Row == toRange.Row && fromRange.Rows == toRange.Rows)
            {
                //for (var r = toRange.Row; r <= toRange.EndRow; r++)
                //{
                //    fromCells = GetRowCellPositionsFromRange(fromRange, r);
                //    toCells = GetRowCellPositionsFromRange(toRange, r);
                //    AutoFillSerialCells(fromCells, toCells);
                //}
                try
                {
                    BeforeSerialFill?.Invoke(this, new Events.RangeSerialFillEventArgs(fromRange, toRange));
                }
                catch
                {
                    // ignored
                }
                ExecuteFillByAction(fromRange, toRange, ExecuteHorizontalFill);
                try
                {
                    AfterSerialFill?.Invoke(this, new Events.RangeSerialFillEventArgs(fromRange, toRange));
                }
                catch
                {
                    // ignored
                }
            }
            else
				throw new InvalidOperationException("The fromRange and toRange must be having same number of rows or same number of columns.");
		}

        private void ExecuteFillByAction(RangePosition fromRange, RangePosition toRange, Action<RangePosition, RangePosition> fillExecutor)
        {
            var before = GetPartialGrid(toRange);                     // старые значения
            fillExecutor?.Invoke(fromRange, toRange);                 // выполняем действи
            var after = GetPartialGrid(toRange);                      // новые значения
            SetPartialGrid(toRange, before);                          // говорим что так и было
            DoAction(new SetPartialGridAction(toRange, after, true)); // и делаем действие
        }
        
        private void ExecuteVerticalFill(RangePosition fromRange, RangePosition toRange)
        {
            for (int c = toRange.Col; c <= toRange.EndCol; c++)
            {
                List<object> fromData = new List<object>();
                for (int r = fromRange.Row; r <= fromRange.EndRow; r++)
                {
                    fromData.Add(cells[r, c]?.Data);
                }
                var filler = SerialFillerBase.GetSerialFiller(fromData.ToArray());


                    #region Up to Down
                for (int toRow = toRange.Row, index = 0; toRow < toRange.EndRow + 1; index++)
                {
                    Cell toCell = cells[toRow, c];

                    if (toCell != null && toCell.Rowspan < 0)
                    {
                        toRow++;
                        continue;
                    }

                    CellPosition fromPos = new CellPosition(fromRange.Row + (index % fromRange.Rows), c);

                    Cell fromCell = cells[fromPos.Row, fromPos.Col];

                    if (fromCell == null || fromCell.Rowspan <= 0)
                    {
                        this[toRow, c] = null;
                        toRow++;
                        continue;
                    }

                    if (fromCell != null && !string.IsNullOrEmpty(fromCell.InnerFormula))
                    {
                        #region Fill Formula
                        FormulaRefactor.Reuse(this, fromPos, new RangePosition(toRow, c, 1, 1));
                        #endregion // Fill Formula
                    }
                    else
                    {
                        #region Fill Number
                        this[toRow, c] = filler.GetSerialValue(toRow - fromRange.Row);
                        #endregion // Fill Number
                    }

                    toRow += Math.Max(fromCell.Rowspan, toCell?.Rowspan ?? 1);
                }
                #endregion // Up to Down
            }
        }
            
        /// <summary>
        /// Метод заполнения большего диапазона значений на основе меньшего диапазона
        /// </summary>
        /// <param name="fromRange"></param>
        /// <param name="toRange"></param>
        private void ExecuteHorizontalFill(RangePosition fromRange, RangePosition toRange)
        {
            for (int r = toRange.Row; r <= toRange.EndRow; r++)
            {
                List<object> fromData = new List<object>();
                for (int c = fromRange.Col; c<= fromRange.EndCol; c++)
                {
                    fromData.Add(cells[r, c]?.Data);
                }
                var filler = SerialFillerBase.GetSerialFiller(fromData.ToArray());


                #region Left to Right
                for (int toCol = toRange.Col, index = 0; toCol < toRange.EndCol + 1; index++)
                {
                    Cell toCell = cells[r, toCol];

                    if (toCell != null && toCell.Colspan < 0)
                    {
                        toCol++;
                        continue;
                    }

                    CellPosition fromPos = new CellPosition(r, fromRange.Col + (index % fromRange.Cols));

                    Cell fromCell = cells[fromPos.Row, fromPos.Col];

                    if (fromCell == null || fromCell.Colspan <= 0)
                    {
                        this[r, toCol] = null;
                        toCol++;
                        continue;
                    }

                    if (fromCell != null && !string.IsNullOrEmpty(fromCell.InnerFormula))
                    {
                        #region Fill Formula
                        FormulaRefactor.Reuse(this, fromPos, new RangePosition(r, toCol, 1, 1));
                        #endregion // Fill Formula
                    }
                    else
                    {
                        #region Fill Number
                        this[r, toCol] = filler.GetSerialValue(toCol - fromRange.Col);
                        #endregion // Fill Number
                    }

                    toCol += Math.Max(fromCell.Colspan, toCell?.Colspan ?? 1);
                }
                #endregion // Left to Right
            }
        }

        private List<CellPosition> GetColumnCellPositionsFromRange(RangePosition fromRange, int columnIndex)
        {
            var result = new List<CellPosition>();
            for (int rowIndex = fromRange.Row; rowIndex < fromRange.EndRow + 1; rowIndex++)
            {
                var cellPosition = new CellPosition(rowIndex, columnIndex);
                AddCellIfValid(cellPosition, result);
            }

            return result;
        }

        private List<CellPosition> GetRowCellPositionsFromRange(RangePosition fromRange, int rowIndex)
        {
            var result = new List<CellPosition>();
            for (int columnIndex = fromRange.Col; columnIndex < fromRange.EndCol + 1; columnIndex++)
            {
                var cellPosition = new CellPosition(rowIndex, columnIndex);
                AddCellIfValid(cellPosition, result);
        }

            return result;
        }

        private void AddCellIfValid(CellPosition cellPosition, List<CellPosition> result)
        {
            var cell = Cells[cellPosition];

            // Exclude merged cells
            if (cell != null && !cell.IsValidCell)
            {
                return;
    }

            result.Add(cellPosition);
}

        private void AutoFillSerialCells(List<CellPosition> fromCells, List<CellPosition> toCells)
        {
            if (!fromCells.Any() || !toCells.Any())
            {
                return;
            }

            var autoFillSequenceInput = fromCells
                .Select(cellPosition => this[cellPosition])
                .ToList();
            var autoFillSequence = new AutoFillSequence(autoFillSequenceInput);
            var autoFillExtrapolatedValues = autoFillSequence.Extrapolate(toCells.Count);

            for (var toCellIndex = 0; toCellIndex < toCells.Count; toCellIndex++)
            {
                var fromCellIndex = toCellIndex % fromCells.Count;
                var fromCellPosition = fromCells[fromCellIndex];
                var fromCell = Cells[fromCellPosition];
                var toCellPosition = toCells[toCellIndex];
                var toCell = Cells[toCellPosition];

                if (!string.IsNullOrEmpty(fromCell?.InnerFormula))
                {
                    FormulaRefactor.Reuse(this, fromCellPosition, new RangePosition(toCellPosition));
                }
                else
                {
                    toCell.Data = autoFillExtrapolatedValues[toCellIndex];
                }
            }
        }
    }
}

#endif // FORMULA
