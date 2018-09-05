using System;
using System.Collections.Generic;
using System.Linq;

namespace unvell.ReoGrid.Core.SerialFill
{
    /// <summary>
    /// Заполнитель, содержащий в себе другие заполнители и использующий их по мере необходимости
    /// </summary>
    class MultiSerialFiller: SerialFillerBase
    {
        private List<SerialFillerBase> _fillers = new List<SerialFillerBase>();
        private List<FillerListDescription> _extendedFillers = null;

        public MultiSerialFiller( object[] data) : base(data)
        {
            Type[] fillerTypes = new Type[] { typeof(NumberSerialFiller), typeof(CopySerialFiller), typeof(IncrementNumberFiller), typeof(NullSerialFiller) };
            Type[] justCopy = new Type[] { typeof(CopySerialFiller) };

            for (int i = 0; i < fillerTypes.Length; i++)
            {
                _fillers.Add(Activator.CreateInstance(fillerTypes[i], new object[] { data }) as SerialFillerBase);
            }
            try
            {
                if (data.Length == 1)
                {
                    // Пользователи просят чтобы при растяжении области с одним элементом копировался этот один элемент
                    // даже если это число. Ранее число инкрементировалось
                    _extendedFillers = TrimToFillers(justCopy, data);
                }
                else
                {
                    _extendedFillers = TrimToFillers(fillerTypes, data);
                }
            }
            catch
            {
                // ignored
            }
        }

        protected override bool CanGetValueInternal(int toIndex)
            => _extendedFillers == null ?
            _fillers.Any(i => i.CanGetValue(toIndex)) :
            CanGetValueEx(toIndex);

        protected override object GetSerialValueInternal(int toIndex)
            => _extendedFillers == null ?
            _fillers.FirstOrDefault(i => i.CanGetValue(toIndex))?.GetSerialValue(toIndex):
            GetSerialValueEx(toIndex);

        private class FillerListDescription
        {
            public SerialFillerBase Filler { get; set; }

            public int Start { get; set; }

            public int Length { get; set; }
        }

        /// <summary>
        /// Функция на основе входной области создает списк заполнителей для каждого участка области
        /// </summary>
        /// <param name="fillerTypes"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        private static List<FillerListDescription> TrimToFillers(Type[] fillerTypes, object[] data)
        {
            var fillers = new List<SerialFillerBase>();
            List<Tuple<AllowedObjectsAttribute, Type>> allowedFillers = new List<Tuple<AllowedObjectsAttribute, Type>>();
            List<FillerListDescription> result = new List<FillerListDescription>();
            foreach (var type in fillerTypes)
            {
                var attribute = type.GetCustomAttributes(typeof(AllowedObjectsAttribute), true).FirstOrDefault() as AllowedObjectsAttribute;
                if (attribute != null)
                {
                    allowedFillers.Add(new Tuple<AllowedObjectsAttribute, Type>(attribute, type));
                }
            }
            var list = new List<object>();
            Tuple<AllowedObjectsAttribute, Type> lastAllowed = null;
            for(int i = 0; i < data.Length; i++)
            {
                list.Add(data[i]);                      // Текущий набор для входных данных для filler-а
                #region Проверяем может ли какой то Filler принять этот массив как входные данные
                var allower = GetAllower(allowedFillers, list);
                #endregion

                #region Если такого Filler-а нет
                if (null == allower)
                {
                    if (lastAllowed != null)               // Если на предыдущем шаге существовал такой Filler, который мог принять эти данные
                    {
                        var lastObj = list[list.Count - 1];
                        list.RemoveAt(list.Count - 1);
                        result.Add(
                            new FillerListDescription
                            {
                                Filler = Activator.CreateInstance(lastAllowed.Item2, new object[] { list.ToArray() }) as SerialFillerBase,
                                Start = i - list.Count,
                                Length = list.Count,
                            });
                        list.Clear();
                        list.Add(lastObj);
                        lastAllowed = GetAllower(allowedFillers, list);
                        if(lastAllowed == null)
                        {
                            // что то сломалось в алгоритме
                            return null;
                        }
                    }
                    else
                    {
                        // что то сломалось в алгоритме
                        return null;
                    }
                }
                else
                {
                    lastAllowed = allower;
                }
                #endregion
            }
            if (lastAllowed != null)
                result.Add(new FillerListDescription
                {
                    Filler = Activator.CreateInstance(lastAllowed.Item2, new object[] { list.ToArray() }) as SerialFillerBase,
                    Start = data.Length - list.Count,
                    Length = list.Count,
                });
            else
            {
                // что то сломалось
                return null;
            }
            return result;
        }

        private static Tuple<AllowedObjectsAttribute, Type> GetAllower(List<Tuple<AllowedObjectsAttribute, Type>> allowedFillers, List<object> list)
        {
            foreach (var checker in allowedFillers)
            {
                if (true == checker.Item1?.IsAllow(list.ToArray()))
                {
                    return checker;
                }
            }
            return null;
        }

        private object GetSerialValueEx(int index)
        {
            var pos = GetPosition(index);
            if(_extendedFillers != null)
            {
                foreach(var eFiller in _extendedFillers)
                {
                    if(pos.ElementNumber >= eFiller.Start && pos.ElementNumber < eFiller.Start + eFiller.Length)
                    {
                        return eFiller.Filler.GetSerialValue(pos.SequenceNumber * eFiller.Length + pos.ElementNumber - eFiller.Start);
                    }
                }
            }
            return null;
        }

        private bool CanGetValueEx(int index)
        {
            var pos = GetPosition(index);
            if (_extendedFillers != null)
            {
                foreach (var eFiller in _extendedFillers)
                {
                    if (pos.ElementNumber >= eFiller.Start && pos.ElementNumber < eFiller.Start + eFiller.Length)
                    {
                        return eFiller.Filler.CanGetValue(pos.SequenceNumber * eFiller.Length + pos.ElementNumber - eFiller.Start);
                    }
                }
            }
            return false;
        }
    }
}
