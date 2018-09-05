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
        {
            if (index >= 0)
            {
                return new Position
                {
                    SequenceNumber = index / Data.Length,
                    ElementNumber = index % Data.Length
                };
            }
            else
            {
                var el = GetPositiveElementIndex(index, Data.Length);
                var seq = (index - (Data.Length - 1));
                seq = seq / Data.Length;

                return new Position
                {
                    SequenceNumber = seq,
                    ElementNumber = el,
                };
            }
        }

        public static ISerialFiller GetSerialFiller(object[] data) => new MultiSerialFiller(data);

        protected int GetPositiveElementIndex(int index, int length)
        {
            var result = (index % length);
            if (result < 0)
            {
                result += length;
            }
            return result;
        }
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
