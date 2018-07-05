// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotValues : IXLPivotValues
    {
        private readonly IXLPivotTable _pivotTable;
        private readonly Dictionary<String, IXLPivotValue> _pivotValues = new Dictionary<string, IXLPivotValue>(StringComparer.OrdinalIgnoreCase);

        internal XLPivotValues(IXLPivotTable pivotTable)
        {
            this._pivotTable = pivotTable;
        }

        public IXLPivotValue Add(String sourceName)
        {
            return Add(sourceName, sourceName);
        }

        public IXLPivotValue Add(String sourceName, String customName)
        {
            if (sourceName != XLConstants.PivotTableValuesSentinalLabel && !this._pivotTable.Source.SourceRangeFields.Contains(sourceName))
                throw new ArgumentOutOfRangeException(nameof(sourceName), String.Format("The column '{0}' does not appear in the source range.", sourceName));

            var pivotValue = new XLPivotValue(sourceName) { CustomName = customName };
            _pivotValues.Add(customName, pivotValue);

            if (_pivotValues.Count > 1 && this._pivotTable.ColumnLabels.All(cl => cl.SourceName != XLConstants.PivotTableValuesSentinalLabel) && this._pivotTable.RowLabels.All(rl => rl.SourceName != XLConstants.PivotTableValuesSentinalLabel))
                _pivotTable.ColumnLabels.Add(XLConstants.PivotTableValuesSentinalLabel);

            return pivotValue;
        }

        public void Clear()
        {
            _pivotValues.Clear();
        }

        public Boolean Contains(String customName)
        {
            return _pivotValues.ContainsKey(customName);
        }

        public Boolean Contains(IXLPivotValue pivotValue)
        {
            return _pivotValues.ContainsKey(pivotValue.SourceName);
        }

        public Boolean ContainsSourceField(string sourceName)
        {
            return _pivotValues.Values.Select(v => v.SourceName).Contains(sourceName, StringComparer.OrdinalIgnoreCase);
        }

        public IXLPivotValue Get(String customName)
        {
            return _pivotValues[customName];
        }

        public IXLPivotValue Get(Int32 index)
        {
            return _pivotValues.Values.ElementAt(index);
        }

        public IEnumerator<IXLPivotValue> GetEnumerator()
        {
            return _pivotValues.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Int32 IndexOf(String customName)
        {
            if (!TryGetValue(customName, out IXLPivotValue pivotValue))
                throw new ArgumentException("Invalid field name.", nameof(customName));

            var selectedItem = _pivotValues
                .Select((item, index) => new { Item = item, Position = index })
                .First(i => i.Item.Key.Equals(customName, StringComparison.OrdinalIgnoreCase));

            return selectedItem.Position;
        }

        public Int32 IndexOf(IXLPivotValue pivotValue)
        {
            return IndexOf(pivotValue.SourceName);
        }

        public void Remove(String customName)
        {
            _pivotValues.Remove(customName);
        }

        public Boolean TryGetValue(string customName, out IXLPivotValue pivotValue)
        {
            return _pivotValues.TryGetValue(customName, out pivotValue);
        }
    }
}
