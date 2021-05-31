namespace unvell.ReoGrid
{
    /// <summary>
    /// Выполняет Unmerge над ячекой, если ячейка является объединенной.
    /// В случае, если ячека не существует - то ячейка пропускается
    /// </summary>
    public class SkipCellUnmergeBehavior: IUnmergeRangeBehavior
    {
        void IUnmergeRangeBehavior.UnmergeCellInRange(Worksheet worksheet, ref int row, ref int column)
        {
            Cell cell = worksheet.GetCellOrNull(row, column);

            if (cell is null) return;
            
            if (cell.Colspan > 1 || cell.Rowspan > 1)
            {
                worksheet.UnmergeCell(cell);
                column += cell.Colspan;
            }
        }
    }
}