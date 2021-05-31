namespace unvell.ReoGrid
{
    /// <summary>
    /// Выполняет Unmerge над ячекой, если ячейка является объединенной.
    /// В случае, если ячека не существует - то создается новая ячейка
    /// </summary>
    public class CreateCellUnmergeBehavior: IUnmergeRangeBehavior
    {
        void IUnmergeRangeBehavior.UnmergeCellInRange(Worksheet worksheet, ref int row, ref int column)
        {
            Cell cell = worksheet.CreateAndGetCell(row, column);

            if (cell.Colspan > 1 || cell.Rowspan > 1)
            {
                worksheet.UnmergeCell(cell);
                column += cell.Colspan;
            }
        }
    }
}