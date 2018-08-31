using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid.Core.SerialFill
{
    [SingleNumberAllowedObjects]
    class IncrementNumberFiller : SerialFillerBase
    {
        double value;
        public IncrementNumberFiller(object[] data) : base(data) {
            value = 0;
            if (Data.Length > 0)
            {
                double d;
                if (CellUtility.TryGetNumberData(Data[0], out d))
                {
                    value = d;
                }
            }
        }

        protected override bool CanGetValueInternal(int toIndex)
        {
            return true;
        }

        protected override object GetSerialValueInternal(int toIndex)
        {
            return value + toIndex;
        }
    }
}
