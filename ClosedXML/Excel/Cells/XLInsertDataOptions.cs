namespace ClosedXML.Excel
{
    public class XLInsertDataOptions
    {
        public IXLNumberFormat DateOnlyFormat { get; set; }
        public IXLNumberFormat DateTimeFormat { get; set; }
        public IXLNumberFormat NumericFormat { get; set; }
        public IXLNumberFormat TimespanFormat { get; set; }
    }
}
