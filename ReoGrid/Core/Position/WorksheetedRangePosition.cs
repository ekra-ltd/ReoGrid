using System;
using System.Diagnostics;
using System.Linq;
using unvell.ReoGrid.Formula;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid
{
    /// <summary>
    /// Класс описания позиции на конкретном листе (worksheet)
    /// </summary>
    /// <remarks>
    /// Изначально reogrid был написан с предположением что лист будет только один и при разных сравнениях не учитывалось 
    /// то что позиции могут быть на разных листах. На данный класс следует переходить в тех местах где используются 
    /// позиции на разных листах (практически везде)
    /// </remarks>
    public class WorksheetedRangePosition
    {
        #region Конструктор

        public WorksheetedRangePosition(Worksheet worksheet, string address)
            : this(worksheet, new RangePosition(address))
        {
        }

        public WorksheetedRangePosition(Worksheet worksheet, RangePosition position)
        {
            Worksheet = worksheet;
            Position = position;
        }

        #endregion

        #region Методы

        /// <summary>
        /// Преобразовывает в формулу Excel для экспорта графиков
        /// </summary>
        /// <returns></returns>
        public string ToExcelFormula()
            => FormulaExtension.ConcatAddress(Worksheet, Position.ToAbsoluteAddress());


        /// <summary>
        /// Получает данные в i-ом столбце или строке
        /// В случае если строк и столбцов больше 1, то проход выполняется сначала по столбцам первой строки,
        /// далее по по второй и так далее
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public string GetValue(uint index)
        {
            string result = null;
            var cellPosition = GetCellPosition(index);
            if (cellPosition.HasValue)
            {
                result = Worksheet.Cells[cellPosition.Value].DisplayText;
            }
            return result;
        }

        /// <summary>
        /// Возвращает количество ячеек в данной позиции
        /// </summary>
        /// <returns></returns>
        public uint CellsCountInRange()
            => (uint) (Position.Cols * Position.Rows);

        internal static WorksheetedRangePosition TryCreateFromFormula(Worksheet worksheet, string address)
        {
            WorksheetedRangePosition result = null;
            try
            {
                if (FormulaUtility.SplitAddress(address, out string worksheetName, out string s))
                {
                    if (string.IsNullOrEmpty(worksheetName))
                    {
                        result = new WorksheetedRangePosition(worksheet, s);
                    }
                    else if (worksheet.Name == worksheetName)
                    {
                        result = new WorksheetedRangePosition(worksheet, s);
                    }
                    else
                    {
                        var ws = worksheet.Workbook?.GetWorksheetByName(worksheetName);
                        if (ws != null)
                        {
                            result = new WorksheetedRangePosition(worksheet, s);
                        }
                    }
                }
            }
            catch(Exception e)
            {
                Debug.Fail($"Ожидается что исключение не будет сгенерированно:{e.Message}");
                // ignore
            }
            return result;
        }

        internal WorksheetedCellPosition GetWorksheetedCellPosition(uint index)
        {
            WorksheetedCellPosition result = null;
            var cellPosition = GetCellPosition(index);
            if (cellPosition.HasValue)
            {
                result = new WorksheetedCellPosition(Worksheet, cellPosition.Value);
            }
            return result;
        }

        private CellPosition? GetCellPosition(uint index)
        {
            var cols = Position.Cols;
            var rows = Position.Rows;
            var maxIndexValue = cols * rows;
            if (index >= maxIndexValue)
            {
                Debug.Fail($"Ожидается что параметр {nameof(index)} будет меньше {maxIndexValue}. Принято:{index}. (Колонок:{cols}, Строк:{rows}, Лист:\"{Worksheet.Name}\", Позиция:{Position.ToAddress()})");
                return null;
            }
            var colOffset = (int)(index % cols); // Номер колонки в ряды
            var rowOffset = (int)(index / cols); // номер ряда
            return new CellPosition(Position.Row + rowOffset, Position.Col + colOffset);
        }

        internal static WorksheetedRangePosition CombineSequentalCellPostions(System.Collections.Generic.IEnumerable<WorksheetedCellPosition> sequence)
        {
            if (sequence is null)
            {
                throw new ArgumentNullException(nameof(sequence));
            }
            var list = sequence.ToList();
            if(!list.Any())
            {
                throw new CombineSequentalCellPostionsException(CombineSequentalCellPostionsExceptionCode.SequinceIsEmpty);
            }
            ReoGrid.Worksheet ws = null;
            foreach (var cell in list)
            {
                if (ws is null)
                {
                    ws = cell.Worksheet;
                }
                else if (ws != cell.Worksheet)
                {
                    throw new CombineSequentalCellPostionsException(CombineSequentalCellPostionsExceptionCode.DifferentWorkshetsInSequence);
                }
                // else - all correct
            }
            var columns = list.Select(c => c.Position.Col).ToList();
            var rows = list.Select(c => c.Position.Row).ToList();
            var sameColumn = columns.All(c => c == columns[0]);
            var sameRow = rows.All(r => r == rows[0]);
            if (!sameRow && !sameColumn)
            {
                throw new CombineSequentalCellPostionsException(CombineSequentalCellPostionsExceptionCode.NotSameRowOrColumn);
            }
            var s = sameRow ? columns : rows;
            for (var i = 1; i < s.Count; i++)
            {
                if (s[i] != s[i - 1] + 1)
                    throw new CombineSequentalCellPostionsException(CombineSequentalCellPostionsExceptionCode.ErrorInSequence);
            }

            return new WorksheetedRangePosition(ws, new RangePosition(rows[0], columns[0], sameRow ? 1 : s.Count, sameRow ? s.Count : 1));
        }
        #endregion

        #region Свойства

        /// <summary>
        /// Лист на котором располагается диапазон
        /// </summary>
        public Worksheet Worksheet { get; }

        /// <summary>
        /// Позиция диапазона
        /// </summary>
        public RangePosition Position { get; }

        #endregion
    }
}
