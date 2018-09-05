using System.Collections.Generic;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid.Core.SerialFill
{
    [NumberListAllowedObjects]
    class NumberSerialFiller: SerialFillerBase
    {
        private double k = 0;

        private double c = 0;

        private int Numbers = 0;

        public NumberSerialFiller(object[] fromData)
                : base(fromData)
            {
            // мнк
            k = 0;
            c = 0;
            int n = Data.Length;
            double sum1 = 0;
            double sum2 = 0;
            double sum3 = 0;
            double sum4 = 0;
            if (Data.Length > 0)
            {
                var x = new List<double>();
                var y = new List<double>();
                for (int i = 0; i < n; i++)
                {
                    double d;
                    if (CellUtility.TryGetNumberData(Data[i], out d))
                    {
                        x.Add(x.Count);
                        y.Add(d);
                        Numbers = x.Count;
                    }
                }
                n = x.Count;
                if (n >= 2)
                {
                    for (int i = 0; i < x.Count; i++)
                    {
                        sum1 += x[i] * y[i];
                        sum2 += y[i];
                        sum3 += x[i] * x[i];
                        sum4 += x[i];
                    }

                    k = (n * sum1 - sum4 * sum2) / (n * sum3 - sum4 * sum4);
                    c = (sum2 - k * sum4) / n;
                }
                else if (n == 1)
                {
                    k = 0;
                    c = y[0];
                }
            }
            else if (n == 1)
            {
                k = 0;
                double d;
                if (CellUtility.TryGetNumberData(Data[0], out d))
                    c = d;
            }
        }

        protected override object GetSerialValueInternal(int toIndex)
        {
            try
            {
                // if (Data.Length > 0)
                // {
                //     double d;
                //     // var data = Data[toIndex % Data.Length];
                //     // if (CellUtility.TryGetNumberData(data, out d))
                //     // {
                        return k * toIndex + c;
                 //   // }
                 //   // return data;
                //}
            }
            catch
            {
                // ignored
            }
            return null;
        }

        protected override bool CanGetValueInternal(int toIndex)
        {
            double d;
            var data = Data[GetPositiveElementIndex(toIndex, Data.Length)];
            return CellUtility.TryGetNumberData(data, out d);
        }
    }
}
