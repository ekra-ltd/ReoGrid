using System;
using System.Collections.Generic;
using System.Diagnostics;
using unvell.ReoGrid.Formula;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid { 

    /// <summary>
    /// Класс описания позиции ячейки на конкретном листе (worksheet)
    /// </summary>
    /// <remarks>
    /// Изначально reogrid был написан с предположением что лист будет только один и при разных сравнениях не учитывалось 
    /// то что позиции могут быть на разных листах. На данный класс следует переходить в тех местах где используются 
    /// позиции на разных листах (практически везде)
    /// </remarks>
    public class WorksheetedCellPosition
    {
        #region Конструктор

        public WorksheetedCellPosition(Worksheet worksheet, string address)
            : this(worksheet, new CellPosition(address))
        {
        }

        public WorksheetedCellPosition(Worksheet worksheet, CellPosition position)
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

        public string GetCellText()
        {
            if (Worksheet is null) return String.Empty;
            if (Position.IsEmpty) return String.Empty;
            return Worksheet.GetCellText(Position);
        }
        ///// <summary>
        ///// Получает данные в i-ом столбце или строке
        ///// В случае если строк и столбцов больше 1, то проход выполняется сначала по столбцам первой строки,
        ///// далее по по второй и так далее
        ///// </summary>
        ///// <param name="index"></param>
        ///// <returns></returns>
        //public string GetValue(uint index)
        //{
        //    var cols = Position.Cols;
        //    var rows = Position.Rows;
        //    var maxIndexValue = cols * rows;
        //    if (index >= maxIndexValue)
        //    {
        //        Debug.Fail($"Ожидается что параметр {nameof(index)} будет меньше {maxIndexValue}. Принято:{index}. (Колонок:{cols}, Строк:{rows}, Лист:\"{Worksheet.Name}\", Позиция:{Position.ToAddress()})");
        //        return null;
        //    }
        //    var colOffset = (int)(index % cols); // Номер колонки в ряды
        //    var rowOffset = (int)(index / cols); // номер ряда
        //    return Worksheet.Cells[Position.Row + rowOffset, Position.Col + colOffset].DisplayText;
        //}

        ///// <summary>
        ///// Возвращает количество ячеек в данной позиции
        ///// </summary>
        ///// <returns></returns>
        //public uint CellsCountInRange()
        //    => (uint)(Position.Cols * Position.Rows);

        /// <summary>
        /// Пытается создать позицию на основе формулы
        /// </summary>
        /// <param name="worksheet">любой лист книги (книга требуется для проверки адреса)</param>
        /// <param name="address">адрес: пример: Sheet2!$A$3, 'Drawing & Charts'!$A$3</param>
        /// <returns></returns>
        internal static WorksheetedCellPosition TryCreateFromFormula(Worksheet worksheet, string address)
        {
            WorksheetedCellPosition result = null;
            try
            {
                if (FormulaUtility.SplitAddress(address, out string worksheetName, out string cellAddress))
                {
                    if (string.IsNullOrEmpty(worksheetName))
                    {
                        result = new WorksheetedCellPosition(worksheet, cellAddress);
                    }
                    else if (worksheet.Name == worksheetName)
                    {
                        result = new WorksheetedCellPosition(worksheet, cellAddress);
                    }
                    else
                    {
                        var ws = worksheet.Workbook?.GetWorksheetByName(worksheetName);
                        result = new WorksheetedCellPosition(ws ?? worksheet, cellAddress);
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

       

        #endregion

        #region Свойства

        /// <summary>
        /// Лист на котором располагается диапазон
        /// </summary>
        public Worksheet Worksheet { get; }

        /// <summary>
        /// Позиция диапазона
        /// </summary>
        public CellPosition Position { get; }

        #endregion
    }
}
