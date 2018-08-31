using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace unvell.ReoGrid.Core.SerialFill
{
    [DateTimeListAllowedObjects]
    class NullSerialFiller : SerialFillerBase
    {
        public NullSerialFiller(object[] data) : base(data) { }
        protected override bool CanGetValueInternal(int toIndex)
        {
            return true;
        }

        protected override object GetSerialValueInternal(int toIndex)
        {
            return null;
        }
    }
}
