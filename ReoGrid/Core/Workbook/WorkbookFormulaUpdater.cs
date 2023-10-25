using System;
using System.Collections.Generic;
using System.Linq;
using unvell.ReoGrid.Formula;

namespace unvell.ReoGrid
{
	/// <summary>
	/// Класс, который вызывает перерасчет формул для всех ячеек во всех листах книги.
	/// Понадобился при решении задачи #11982, так как reogrid при импорте книги на перерасчитывает формулы.
	/// </summary>
	internal static class WorkbookFormulaUpdater
	{
		/// <summary>
		/// Список ячеек, для которых значение формулы перерасчитано
		/// </summary>
		private static readonly HashSet<Cell> ProcessedCells = new();
		
		/// <summary>
		/// Контроль за чересчур вложенными формулами, а также зацикливанием
		/// </summary>
		private static uint _updateCellFormulaCallGuard = 0;
	
		/// <summary>
		/// Метод, который пробегается по всем ячейкам всех листов и вызывает перерасчет формулы.
		/// </summary>
		public static void Update(IWorkbook workbook)
		{
			var workSheets = workbook?.Worksheets;
			if (workSheets == null) return;
			foreach (var worksheet in workSheets)
			{
				worksheet?.IterateCells(worksheet.UsedRange, true, (rowIndex, colIndex, cell) =>
				{
					UpdateCellFormula(cell, worksheet);
					return true;
				});
			}
			ProcessedCells.Clear();
		}

		/// <summary>
		/// Рекурсивное вычисление формулы ячейки (с вычислением формул для всех ячеек, входящих в текущую формулу)
		/// </summary>
		private static void UpdateCellFormula(Cell cell, Worksheet worksheet)
		{
			++_updateCellFormulaCallGuard;
			try
			{
				if (_updateCellFormulaCallGuard <= 50)
				{
					if (cell == null || !cell.HasFormula || cell.FormulaTree == null || ProcessedCells.Contains(cell))
						return;
					var stCellNodes = GetFormulaCellReferences(cell.FormulaTree);
					stCellNodes.Where(c => c != null && !ProcessedCells.Contains(c)).ToList()
							.ForEach(c => UpdateCellFormula(c, worksheet));
				}
				worksheet.RecalcCell(cell);
				ProcessedCells.Add(cell);
			}
			finally
			{
				--_updateCellFormulaCallGuard;
			}
		}

		/// <summary>
		/// Получение списка всех ячеек, входящих в текущую формулу
		/// </summary>
		private static IEnumerable<Cell> GetFormulaCellReferences(STNode formula)
		{
			var result = new List<Cell>();
			var queue = new Queue<STNode>();
			queue.Enqueue(formula);
			while (queue.Any())
			{
				var node = queue.Dequeue();
				if (node.Children == null) continue;
				foreach (var subNode in node.Children)
				{
					switch (subNode.Type)
					{
						case STNodeType.CELL:
						{
							if (subNode is STCellNode cellNode)
							{
								var cell = cellNode.Worksheet?.GetCellOrNull(cellNode.Position.Row, cellNode.Position.Col);
								if (cell != null)
									result.Add(cell);
							}
							break;
						}
						case STNodeType.RANGE:
						{
							if (subNode is STRangeNode cellNode)
							{
								cellNode.Worksheet?.IterateCells(cellNode.Range, true, (rowIndex, colIndex, cell) =>
								{
									if (cell != null)
										result.Add(cell);
									return true;
								});
							}
							break;
						}
						case STNodeType.IDENTIFIER:
						{
							if (subNode is STIdentifierNode cellNode)
							{
								var name = cellNode.Identifier;
								if (name != null && cellNode.Worksheet != null && cellNode.Worksheet.TryGetNamedRange(name, out var range))
								{
									if (range != null)
									{
										cellNode.Worksheet?.IterateCells(range.Position, true, (rowIndex, colIndex, cell) =>
										{
											if (cell != null)
												result.Add(cell);
											return true;
										});
									}
								}
							}
							break;
						}
					}
					queue.Enqueue(subNode);
				}
			}
			return result;
		}
	}
}