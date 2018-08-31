using System;

namespace unvell.ReoGrid.Core.SerialFill
{
    abstract class SerialFillerBase : ISerialFiller
    {
        protected object[] Data;

        protected SerialFillerBase(object[] data)
        {
            Data = new object[data.Length];
            Array.Copy(data, Data, data.Length);
        }

        public bool CanGetValue(int toIndex)
            => CanGetValueInternal(toIndex);

        public object GetSerialValue(int toIndex)
            => GetSerialValueInternal(toIndex);
        

        protected abstract bool CanGetValueInternal(int toIndex);

        protected abstract object GetSerialValueInternal(int toIndex);

        protected class Position
        {
            public int SequenceNumber { get; set; }
            public int ElementNumber { get; set; }
        }

        protected Position GetPosition(int index)
            => new Position
            {
                SequenceNumber = index / Data.Length ,
                ElementNumber = index % Data.Length
            };

        public static ISerialFiller GetSerialFiller(object[] data)
            => new MultiSerialFiller(
                new Type[] {
                    typeof(NumberSerialFiller),
                    typeof(CopySerialFiller),
                    typeof(IncrementNumberFiller),
                    typeof(NullSerialFiller) },
                data);

       //  /// <summary>
       //  /// Указывает что такой объект может ввходить в данный Filler
       //  /// </summary>
       //  /// <param name="obj"></param>
       //  /// <returns></returns>
       //  protected abstract bool AllowObject(object obj);
       // 
       // /// <summary>
       // /// Указывает что такой объект может ввходить в данный Filler
       // /// </summary>
       // /// <param name="obj"></param>
       // /// <returns></returns>
       // protected abstract bool AllowObjects(object[] obj);
    }
}
