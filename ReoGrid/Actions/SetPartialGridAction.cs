/*****************************************************************************
 * 
 * ReoGrid - Opensource .NET Spreadsheet Control
 * 
 * https://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Thank you to all contributors!
 * 
 * (c) 2012-2020 Jingwood, unvell.com <jingwood at unvell.com>
 * 
 ****************************************************************************/

using System.Diagnostics;

namespace unvell.ReoGrid.Actions
{
	/// <summary>
	/// Action to set partial grid.
	/// </summary>
	public class SetPartialGridAction : WorksheetReusableAction
	{
		private PartialGrid data;
		private PartialGrid backupData;
		private readonly bool _forceUnmerge;

		/// <summary>
		/// Create action to set partial grid.
		/// </summary>
		/// <param name="range">target range to set partial grid.</param>
		/// <param name="data">partial grid to be set.</param>
		/// <param name="forceUnmerge">Указывает, что для диапазона <see cref="range"/>нужно принудительно вызвать
		/// <seealso cref="Worksheet.UnmergeRange(int,int,int,int)"/></param>
		public SetPartialGridAction(RangePosition range, PartialGrid data, bool forceUnmerge)
			: base(range)
		{
			this.data = data;
			_forceUnmerge = forceUnmerge;
		}

		public override WorksheetReusableAction Clone(RangePosition range)
		{
			return new SetPartialGridAction(range, data, _forceUnmerge);
		}

		/// <summary>
		/// Do action to set partial grid.
		/// </summary>
		public override void Do()
		{
			backupData = Worksheet.GetPartialGrid(base.Range, PartialGridCopyFlag.All, ExPartialGridCopyFlag.BorderOutsideOwner);
			Debug.Assert(backupData != null);
			//#5439-43 Сбрасываем Merge всех ячеек
			if (_forceUnmerge)
			{
				// Выполнеяется Unmerge для существующих на рабочем листе ячеек
				Worksheet.UnmergeRange(Range, new SkipCellUnmergeBehavior());
				Range = Worksheet.SetPartialGridRepeatly(Range, data);
			}
			else
			{
				Worksheet.SetPartialGrid(Range, data, PartialGridCopyFlag.CellData, ExPartialGridCopyFlag.BorderOutsideOwner);
			}
			Worksheet.TryAddConditionalFormats();
			Worksheet.RecalcConditionalFormats();
		}

		/// <summary>
		/// Undo action to restore setting partial grid.
		/// </summary>
		public override void Undo()
		{
			Debug.Assert(backupData != null);
			//#5439-43 Сбрасываем Merge всех ячеек
			if (_forceUnmerge)
				Worksheet.UnmergeRange(Range, new SkipCellUnmergeBehavior());
			base.Worksheet.SetPartialGrid(Range, backupData, PartialGridCopyFlag.All, ExPartialGridCopyFlag.BorderOutsideOwner);
		}

		/// <summary>
		/// Get friendly name of this action.
		/// </summary>
		/// <returns>Friendly name of this action.</returns>
		public override string GetName()
		{
			return "Set Partial Grid";
		}
	}
}
