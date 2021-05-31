namespace unvell.ReoGrid
{
    /// <summary>
    /// Действие по разделению ячейки на листе документа по 
    /// </summary>
    public interface IUnmergeRangeBehavior
    {
        /// <summary>
        /// Выполняет действие по разделению ячейки 
        /// </summary>
        /// <param name="worksheet">Лист, на котором распологается ячейка</param>
        /// <param name="row">Номер строки, на которой распологается ячейка <see cref="Cell"/> </param>
        /// <param name="column">Номер столбца на котором распологается ячейка. Возвращается колонка где распологается
        /// следующая за данной ячейка <see cref="Cell"/> </param>
        void UnmergeCellInRange(Worksheet worksheet, ref int row, ref int column);
    }
}