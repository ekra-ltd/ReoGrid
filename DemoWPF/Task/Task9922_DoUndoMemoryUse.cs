namespace unvell.ReoGrid.WPFDemo.Task
{
    /// <summary>
    /// Класс для тестирования задачи #9922
    /// Создает область 10000 * 1000, при копировании и вставки которыоый наблюдается повышенное выделение памяти
    /// </summary>
    public class Task9922_DoUndoMemoryUse: ITaskExample
    {
        public void Apply(ReoGridControl grid)
        {
            var worksheet = grid.NewWorksheet("#9922");
            
            worksheet.SetCols(1000);
            worksheet.SetRows(10000);
        }
    }
}