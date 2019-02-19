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

using System;
using System.Linq;
using System.Text;
using System.Threading;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Main;
using unvell.ReoGrid.Interaction;

#if WINFORM
using DataObject = System.Windows.Forms.DataObject;
using Clipboard = System.Windows.Forms.Clipboard;
#elif WPF
using DataObject = System.Windows.DataObject;
using Clipboard = System.Windows.Clipboard;
#endif // WINFORM

#if EX_SCRIPT
using unvell.ReoScript;
using unvell.ReoGrid.Script;
#endif // EX_SCRIPT

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        private static readonly string ClipBoardDataFormatIdentify = "{CB3BE3D1-2BF9-4fa6-9B35-374F6A0412CE}";

        private RangePosition currentCopingRange = RangePosition.Empty;

		public string StringifyRange(string addressOrName)
		{
			if (RangePosition.IsValidAddress(addressOrName))
			{
				return this.StringifyRange(new RangePosition(addressOrName));
			}
			else if (this.registeredNamedRanges.TryGetValue(addressOrName, out var namedRange))
			{
				return this.StringifyRange(namedRange);
			}
			else
				throw new InvalidAddressException(addressOrName);
		}

        /// <summary>
        /// Convert all data from specified range to a tabbed string.
        /// </summary>
        /// <param name="range">The range to be converted.</param>
        /// <returns>Tabbed string contains all data converted from specified range.</returns>
        public string StringifyRange(RangePosition range)
        {
			int erow = range.EndRow;
			int ecol = range.EndCol;

            // copy plain text
            StringBuilder sb = new StringBuilder();

            bool isFirst = true;
			for (int r = range.Row; r <= erow; r++)
            {
                if (isFirst) isFirst = false; else sb.Append('\n');

                bool isFirst2 = true;
				for (int c = range.Col; c <= ecol; c++)
                {
                    if (isFirst2) isFirst2 = false; else sb.Append('\t');

                    var cell = this.cells[r, c];
                    if (cell != null)
                    {
                        var text = cell.DisplayText;

                        if (!string.IsNullOrEmpty(text))
                        {
                            if (text.Contains('\n'))
                            {
                                text = string.Format("\"{0}\"", text);
                            }

                            sb.Append(text);
                        }
                    }
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Paste data from tabbed string into worksheet.
        /// </summary>
		/// <param name="address">Start cell position to be filled.</param>
		/// <param name="content">Data to be pasted.</param>
        /// <returns>Range position that indicates the actually filled range.</returns>
		public RangePosition PasteFromString(string address, string content)
        {
			if (!CellPosition.IsValidAddress(address))
            {
				throw new InvalidAddressException(address);
			}

			return this.PasteFromString(new CellPosition(address), content);
		}

		/// <summary>
		/// Paste data from tabbed string into worksheet.
		/// </summary>
		/// <param name="startPos">Start position to fill data.</param>
		/// <param name="content">Tabbed string to be pasted.</param>
		/// <returns>Range position that indicates the actually filled range.</returns>
		public RangePosition PasteFromString(CellPosition startPos, string content)
		{
			//int rows = 0, cols = 0;

			//string[] lines = content.Split(new string[] { "\r\n" }, StringSplitOptions.None);
			//for (int r = 0; r < lines.Length; r++)
			//{
			//	string line = lines[r];
			//	if (line.EndsWith("\n")) line = line.Substring(0, line.Length - 1);
			//	//line = line.Trim();

			//	if (line.Length > 0)
			//	{
			//		string[] tabs = line.Split('\t');
			//		cols = Math.Max(cols, tabs.Length);

			//		for (int c = 0; c < tabs.Length; c++)
			//		{
			//			int toRow = startPos.Row + r;
			//			int toCol = startPos.Col + c;

			//			if (!this.IsValidCell(toRow, toCol))
			//			{
			//				throw new RangeIntersectionException(new RangePosition(toRow, toCol, 1, 1));
			//			}

			//			string text = tabs[c];

			//			if (text.StartsWith("\"") && text.EndsWith("\""))
			//			{
			//				text = text.Substring(1, text.Length - 2);
			//			}

			//			SetCellData(toRow, toCol, text);
			//		}

			//		rows++;
			//	}
			//}

			object[,] parsedData = RGUtility.ParseTabbedString(content);

			int rows = parsedData.GetLength(0);
			int cols = parsedData.GetLength(1);

			var range = new RangePosition(startPos.Row, startPos.Col, rows, cols);

			this.SetRangeData(range, parsedData);

			return range;
        }

        #region Copy

        /// <summary>
        /// Copy data and put into Clipboard.
        /// </summary>
        public bool Copy()
        {
            if (IsEditing)
            {
                this.controlAdapter.EditControlCopy();
            }
            else
            {
                this.controlAdapter.ChangeCursor(CursorStyle.Busy);

                try
                {
                    if (BeforeCopy != null)
                    {
                        var evtArg = new BeforeRangeOperationEventArgs(selectionRange);
                        BeforeCopy(this, evtArg);
                        if (evtArg.IsCancelled)
                        {
                            return false;
                        }
                    }

#if EX_SCRIPT
                    var scriptReturn = RaiseScriptEvent("oncopy");
                    if (scriptReturn != null && !ScriptRunningMachine.GetBoolValue(scriptReturn))
                    {
                        return false;
                    }
#endif // EX_SCRIPT

                    // highlight current copy range
                    currentCopingRange = selectionRange;

#if WINFORM || WPF
                    DataObject data = new DataObject();
                    data.SetData(ClipBoardDataFormatIdentify,
                        GetPartialGrid(currentCopingRange, PartialGridCopyFlag.All, ExPartialGridCopyFlag.None, true));

                    string text = StringifyRange(currentCopingRange);
                    if (!string.IsNullOrEmpty(text)) data.SetText(text);

                    // set object data into clipboard
                    Clipboard.SetDataObject(data);
#endif // WINFORM || WPF

                    if (AfterCopy != null)
                    {
                        AfterCopy(this, new RangeEventArgs(this.selectionRange));
                    }
                }
                catch (Exception ex)
                {
                    this.NotifyExceptionHappen(ex);
                    return false;
                }
                finally
                {
                    this.controlAdapter.ChangeCursor(CursorStyle.PlatformDefault);
                }

            }

            return true;
        }

        #endregion // Copy

        #region Paste

        private class PasteValues
        {
            public PartialGrid PartialGrid { get; set; }= null;
            public string ClipboardText { get; set; }= null;
        }
        
        /// <summary>
        /// Copy data from Clipboard and put on grid.
        /// 
        /// Currently ReoGrid supports the following types of source from the clipboard.
        ///  - Data from another ReoGrid instance
        ///  - Plain/Unicode Text from any Windows Applications
        ///  - Tabbed Plain/Unicode Data from Excel or similar applications
        /// 
        /// When data copied from another ReoGrid instance, and the destination range 
        /// is bigger than the source, ReoGrid will try to repeat putting data to fill 
        /// the destination range entirely.
        /// 
        /// Todo: Copy border and cell style from Excel.
        /// </summary>
        public bool Paste()
        {
            if (IsEditing)
            {
                this.controlAdapter.EditControlPaste();
            }
            else
            {
                if (!IsPateActionCanExecute()) return false;
                try
                {
                    PreparePasteAction();
                    if (!PasteAction(GetPasteValues())) return false;
                }
                catch (Exception ex)
                {
                    NotifyExceptionHappen(ex);
                }
                finally
                {
                    FinallyPasteAction();
                }
                NotifyPasteAction();
            }

            return true;
        }

        private bool IsPateActionCanExecute()
        {
            if (HasSettings(WorksheetSettings.Edit_Readonly) || selectionRange.IsEmpty)
            {
                return false;
            }
            return true;
        }

        private void PreparePasteAction()
        {
            controlAdapter.ChangeCursor(CursorStyle.Busy);
        }

        private bool PasteAction(PasteValues pasteValues)
        {
            if (pasteValues.PartialGrid != null)
            {
                if (!PastlePartialGrid(pasteValues.PartialGrid))
                {
                    return false;
                }
            }
            else if (!string.IsNullOrEmpty(pasteValues.ClipboardText))
            {
                if (!PastePlainText(pasteValues.ClipboardText) )
                {
                    return false;
                }
            }
            return true;
        }
        
        private bool PastlePartialGrid(PartialGrid partialGrid)
        {
            var pastleData =  CreatePastlePartialGridTargetRange(this, partialGrid);
            if (IsSomethingWrongInPastlePartialGrid(pastleData.DstRangePosition)) return false;
            DoAction(new SetPartialGridAction(pastleData.DstRangePosition, pastleData.SourcePartialGrid, pastleData.ForceUnmerge));
            return true;
        }

        private void FinallyPasteAction()
        {
            controlAdapter.ChangeCursor(CursorStyle.Selection);
            RequestInvalidate();
        }

        private void NotifyPasteAction()
        {
            AfterPaste?.Invoke(this, new RangeEventArgs(selectionRange));
        }
        
        private class PastleRangesData
        {
            public RangePosition DstRangePosition { get; set; }

            public PartialGrid SourcePartialGrid { get; set; }

            public bool ForceUnmerge { get; set; }
        }

        private static PastleRangesData CreatePastlePartialGridTargetRange(Worksheet dstWorksheet, PartialGrid sourcePartialGrid)
        {
            var dstPosition = dstWorksheet.selectionRange;
            bool forceUnmerge = true;

            var startRow = dstPosition.Row;
            var startCol = dstPosition.Col;

            var rows = sourcePartialGrid.Rows;
            var cols = sourcePartialGrid.Columns;

            // Если источник кратно (несколько раз целиком) помещается в диапазон - то используется весь диапазон
            if (dstPosition.Rows % sourcePartialGrid.Rows == 0)
            {
                rows = dstPosition.Rows;
            }
            if (dstPosition.Cols % sourcePartialGrid.Columns == 0)
            {
                cols = dstPosition.Cols;
            }
            
            // Если источник это одна ячейка *(возможно объединенная) и диапазон вставки - это тоже одна ячейка*(возможно объединенная)
            // то копируем только одну ячейку
            // GetPartialGrid(currentCopingRange, PartialGridCopyFlag.All, ExPartialGridCopyFlag.None, true)
            var dstPartialgrid = dstWorksheet.GetPartialGrid(dstPosition);
            if (
                IsSingleMergedCell(dstPartialgrid) &&
                (IsSingleCell(sourcePartialGrid) || IsSingleMergedCell(sourcePartialGrid)))
            {
                rows = 1;
                cols = 1;
                forceUnmerge = false;
            }
            return new PastleRangesData
            {
                DstRangePosition = new RangePosition(startRow, startCol, rows, cols),
                SourcePartialGrid = sourcePartialGrid,
                ForceUnmerge = forceUnmerge
            };
        
            // return new RangePosition(startRow, startCol, rows, cols);
        }

        private static bool IsSingleCell(PartialGrid partialGrid)
        {
            var result = false;
            var cell = partialGrid?.Cells[0, 0];
            if (cell != null)
            {
                if (!cell.IsMergedCell)
                {
                    if (partialGrid.Rows == 1 && partialGrid.Columns == 1)
                    {
                        result = true;
                    }
                }
            }
            return result;
        }

        private static bool IsSingleMergedCell(PartialGrid partialGrid)
        {
            var result = false;
            var cell = partialGrid?.Cells[0, 0];
            if (cell != null)
            {
                
                if (cell.IsMergedCell)
                {
                    var cols = cell.MergeEndPos.Col - cell.MergeStartPos.Col + 1;
                    var rows = cell.MergeEndPos.Row - cell.MergeStartPos.Row + 1;

                    if (rows == partialGrid.Rows && cols == partialGrid.Columns)
                        result = true;
                }
            }
            return result;
        }

        private bool IsSomethingWrongInPastlePartialGrid(RangePosition targetRange)
        {
            if (IsPastePartialGridExternCodeCanceled(targetRange)) return true;
            if (IsPateTargetoutOfrange(targetRange)) return true;
            if (IsPastePartialGridCheckReadOnly(targetRange)) return true;
            return false;
        }
        
         private static PasteValues GetPasteValues()
        {
            int tryings = 1;

            do
            {
                try
                {
                    var result = new PasteValues
                    {
                        PartialGrid = null,
                        ClipboardText = null,
                    };

#if WINFORM || WPF
                    if (Clipboard.GetDataObject() is DataObject data)
                    {
                        result.PartialGrid = data.GetData(ClipBoardDataFormatIdentify) as PartialGrid;
                        if (data.ContainsText())
                        {
                            result.ClipboardText = data.GetText();
                        }
                    }
#elif ANDROID

#endif // WINFORM || WPF
                    return result;
                }
                catch
                {
                    if (--tryings <= 0) throw;
                    Thread.Sleep(100);
                }
            } while (true);

        }

        private bool PastePlainText(string clipboardText)
        {
            var arrayData = RGUtility.ParseTabbedString(clipboardText);
            var targetRange = CreatePastePlaintTextTargteRange(arrayData);
            if (!RaiseBeforePasteEvent(targetRange))
            {
                return false;
            }

            var actionSupportedControl = controlAdapter?.ControlInstance as IActionControl;
            actionSupportedControl?.DoAction(this, new SetRangeDataAction(targetRange, arrayData));
            return true;
        }

        private RangePosition CreatePastePlaintTextTargteRange(object[,] arrayData)
        {
            var maxRows = Math.Max(selectionRange.Rows, arrayData.GetLength(0));
            var maxCols = Math.Max(selectionRange.Cols, arrayData.GetLength(1));

            var targetRange = new RangePosition(selectionRange.Row, selectionRange.Col, maxRows, maxCols);
            return targetRange;
        }





        private bool IsPastePartialGridExternCodeCanceled(RangePosition targetRange)
        {
            if (!RaiseBeforePasteEvent(targetRange))
            {
                return true;
            }
            return false;
        }

        private bool IsPateTargetoutOfrange(RangePosition targetRange)
        {
            if (targetRange.EndRow >= rows.Count || targetRange.EndCol >= cols.Count)
            {
                // TODO: paste range overflow
                // need to notify user-code to handle this 
                return true;
            }
            return false;
        }

        private bool IsPastePartialGridCheckReadOnly(RangePosition targetRange)
        {
            // check whether the range to be pasted contains readonly cell
            if (CheckRangeReadonly(targetRange))
            {
                NotifyExceptionHappen(new OperationOnReadonlyCellException("specified range contains readonly cell"));
                return true;
            }
            return false;
        }
        
        private bool CheckIntersectedMergeRangeInPartialGridError(
            PartialGrid partialGrid, 
            int rowRepeat, 
            int colRepeat,
            int startRow, 
            int startCol)
        {
            bool isError = false;
            if (partialGrid.Cells != null)
            {
                try
                {
                    #region Check repeated intersected ranges

                    for (var rr = 0; rr < rowRepeat; rr++)
                    {
                        for (var cc = 0; cc < colRepeat; cc++)
                        {
                            var rrClosure = rr;
                            var ccClosure = cc;
                            partialGrid.Cells.Iterate((row, col, cell) =>
                            {
                                if (cell.IsMergedCell)
                                {
                                    for (var r = startRow; r < cell.MergeEndPos.Row - cell.InternalRow + startRow + 1; r++)
                                    {
                                        for (var c = startCol;
                                            c < cell.MergeEndPos.Col - cell.InternalCol + startCol + 1;
                                            c++)
                                        {
                                            var tr = r + rrClosure * partialGrid.Rows;
                                            var tc = c + ccClosure * partialGrid.Columns;

                                            var existedCell = cells[tr, tc];

                                            if (existedCell != null)
                                            {
                                                if (
                                                    // cell is a part of merged cell
                                                    (existedCell.Rowspan == 0 && existedCell.Colspan == 0)
                                                    // cell is merged cell
                                                    || existedCell.IsMergedCell)
                                                {
                                                    throw new RangeIntersectionException(selectionRange);
                                                }
                                                // cell is readonly
                                                else if (existedCell.IsReadOnly)
                                                {
                                                    throw new CellDataReadonlyException(cell.InternalPos);
                                                }
                                            }
                                        }
                                    }
                                }
                                return Math.Min(cell.Colspan, (short) 1);
                            });
                        }
                    }

                    #endregion // Check repeated intersected ranges
                }
                catch (Exception ex)
                {
                    isError = true;
                    // raise event to notify user-code there is error happened during paste operation
                    OnPasteError?.Invoke(this, new RangeOperationErrorEventArgs(selectionRange, ex));
                }
            }
            return isError;
        }

        private bool RaiseBeforePasteEvent(RangePosition range)
        {
            if (BeforePaste != null)
            {
                var evtArg = new BeforeRangeOperationEventArgs(range);
                BeforePaste(this, evtArg);
                if (evtArg.IsCancelled)
                {
                    return false;
                }
            }

#if EX_SCRIPT
            object scriptReturn = RaiseScriptEvent("onpaste", new RSRangeObject(this, range));
            if (scriptReturn != null && !ScriptRunningMachine.GetBoolValue(scriptReturn))
            {
                return false;
            }
#endif // EX_SCRIPT

            return true;
        }

        #endregion // Paste

        #region Cut
        /// <summary>
        /// Copy any remove anything from selected range into Clipboard.
        /// </summary>
        public bool Cut()
        {
            if (IsEditing)
            {
                this.controlAdapter.EditControlCut();
            }
            else
            {
                if (!Copy()) return false;

                if (BeforeCut != null)
                {
                    var evtArg = new BeforeRangeOperationEventArgs(this.selectionRange);

                    BeforeCut(this, evtArg);

                    if (evtArg.IsCancelled)
                    {
                        return false;
                    }
                }

#if EX_SCRIPT
                object scriptReturn = RaiseScriptEvent("oncut");
                if (scriptReturn != null && !ScriptRunningMachine.GetBoolValue(scriptReturn))
                {
                    return false;
                }
#endif

                if (!HasSettings(WorksheetSettings.Edit_Readonly))
                {
                    this.DeleteRangeData(currentCopingRange);
                    this.RemoveRangeStyles(currentCopingRange, PlainStyleFlag.All);
                    this.RemoveRangeBorders(currentCopingRange, BorderPositions.All);
                }

                if (AfterCut != null)
                {
                    AfterCut(this, new RangeEventArgs(this.selectionRange));
                }
            }

            return true;
        }
        #endregion // Cut

        //private void CheckCanPaste()
        //{
        //	// TODO
        //}

        //void ClipboardMonitor_ClipboardChanged(object sender, ClipboardChangedEventArgs e)
        //{
        //	CheckCanPaste();
        //}

        #region Checks

        /// <summary>
        /// Determine whether the selected range can be copied.
        /// </summary>
        /// <returns>True if the selected range can be copied.</returns>
        public bool CanCopy()
        {
            //TODO
            return true;
        }

        /// <summary>
        /// Determine whether the selected range can be cutted.
        /// </summary>
        /// <returns>True if the selected range can be cutted.</returns>
        public bool CanCut()
        {
            //TODO
            return true;
        }

        /// <summary>
        /// Determine whether the data contained in Clipboard can be pasted into grid control.
        /// </summary>
        /// <returns>True if the data contained in Clipboard can be pasted</returns>
        public bool CanPaste()
        {
            //TODO
            return true;
        }

        #endregion // Checks

        #region Events

        /// <summary>
        /// Before a range will be pasted from Clipboard
        /// </summary>
        public event EventHandler<BeforeRangeOperationEventArgs> BeforePaste;

        /// <summary>
        /// When a range has been pasted into grid
        /// </summary>
        public event EventHandler<RangeEventArgs> AfterPaste;

        /// <summary>
        /// When an error happened during perform paste
        /// </summary>
        [Obsolete("use ReoGridControl.ErrorHappened instead")]
        public event EventHandler<RangeOperationErrorEventArgs> OnPasteError;

        /// <summary>
        /// Before a range to be copied into Clipboard
        /// </summary>
        public event EventHandler<BeforeRangeOperationEventArgs> BeforeCopy;

        /// <summary>
        /// When a range has been copied into Clipboard
        /// </summary>
        public event EventHandler<RangeEventArgs> AfterCopy;

        /// <summary>
        /// Before a range to be moved into Clipboard
        /// </summary>
        public event EventHandler<BeforeRangeOperationEventArgs> BeforeCut;

        /// <summary>
        /// After a range to be moved into Clipboard
        /// </summary>
        public event EventHandler<RangeEventArgs> AfterCut;

        /// <summary>
        /// Возникает до последовательного заполнения диапазона ячеек на основе другого диапазона
        /// </summary>
        public event EventHandler<RangeSerialFillEventArgs> BeforeSerialFill;

        /// <summary>
        /// Возникает до последовательного заполнения диапазона ячеек на основе другого диапазона
        /// </summary>
        public event EventHandler<RangeSerialFillEventArgs> AfterSerialFill;
        #endregion // Events
    }
}

#if WINFORM || WPF

#endif // WINFORM || WPF