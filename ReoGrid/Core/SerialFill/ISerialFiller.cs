namespace unvell.ReoGrid.Core.SerialFill
{
    /// <summary>
    /// Интерфейс отвечающий за последовательное заполнение данных
    /// </summary>
    internal interface ISerialFiller
    {
        /// <summary>
        /// Указывает может ли данный экземпляр объекта создать новый объект последовательности
        /// </summary>
        /// <param name="toIndex">номер элемента в последовательности</param>
        /// <returns></returns>
        bool CanGetValue(int toIndex);

        /// <summary>
        /// Создает новый элемент в последовательности
        /// </summary>
        /// <param name="toIndex">Номер элемента в последовательности</param>
        /// <returns></returns>
        object GetSerialValue(int toIndex);
    }
}
