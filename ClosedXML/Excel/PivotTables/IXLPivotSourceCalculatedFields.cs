using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotSourceCalculatedFields : IEnumerable<IXLPivotSourceCalculatedField>
    {
        IXLPivotSourceCalculatedField Add(String name, String formula);

        void Clear();

        Boolean Contains(String name);

        IXLPivotSourceCalculatedField Get(String name);

        void Remove(String name);

        Boolean TryGetCalculatedField(String name, out IXLPivotSourceCalculatedField calculatedField);
    }
}
