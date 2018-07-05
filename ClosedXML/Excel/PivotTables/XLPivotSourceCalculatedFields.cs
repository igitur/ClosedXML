// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPivotSourceCalculatedFields : IXLPivotSourceCalculatedFields
    {
        private readonly Dictionary<String, IXLPivotSourceCalculatedField> _calculatedFields = new Dictionary<String, IXLPivotSourceCalculatedField>(StringComparer.OrdinalIgnoreCase);
        private readonly IXLPivotSource _pivotSource;

        internal XLPivotSourceCalculatedFields(IXLPivotSource pivotSource)
        {
            this._pivotSource = pivotSource;
        }

        public IXLPivotSourceCalculatedField Add(String name, String formula)
        {
            if (_calculatedFields.Keys.Contains(name) || this._pivotSource.CachedFields.Keys.Contains(name, StringComparer.OrdinalIgnoreCase))
                throw new ArgumentException(nameof(name), String.Format("The name '{0}' is already in use by another pivot field.", name));

            var calculatedField = new XLPivotSourceCalculatedField(name, formula);
            _calculatedFields.Add(name, calculatedField);

            return calculatedField;
        }

        public void Clear()
        {
            _calculatedFields.Clear();
        }

        public Boolean Contains(String name)
        {
            return _calculatedFields.ContainsKey(name);
        }

        public IXLPivotSourceCalculatedField Get(String name)
        {
            return _calculatedFields[name];
        }

        public IEnumerator<IXLPivotSourceCalculatedField> GetEnumerator()
        {
            return _calculatedFields.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Remove(String name)
        {
            _calculatedFields.Remove(name);
        }

        public Boolean TryGetCalculatedField(String name, out IXLPivotSourceCalculatedField calculatedField)
        {
            return _calculatedFields.TryGetValue(name, out calculatedField);
        }
    }
}
