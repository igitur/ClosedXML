using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotSourceCalculatedField
    {
        String Formula { get; set; }
        String Name { get; }
    }
}
