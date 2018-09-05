namespace unvell.ReoGrid.Core.SerialFill
{
    [StringListAllowedObjectsAttribute]
    class CopySerialFiller : SerialFillerBase
    {
        public CopySerialFiller(object[] data)
                : base(data) { }

        protected override bool CanGetValueInternal(int toIndex) => true;

        protected override object GetSerialValueInternal(int toIndex)
        {
            if (Data.Length > 0)
                return Data[GetPositiveElementIndex(toIndex, Data.Length)];
            return null;
        }
    }
}
