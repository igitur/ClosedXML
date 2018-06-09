using ClosedXML.Attributes;
using FastMember;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{Name}")]
    internal class XLTable : XLRange, IXLTable
    {
        #region Private fields

        private string _name;
        internal bool _showTotalsRow;
        internal HashSet<String> _uniqueNames;

        #endregion Private fields

        #region Constructor

        /// <summary>
        /// The direct contructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
        /// </summary>
        public XLTable(XLRangeParameters xlRangeParameters)
            : base(xlRangeParameters)
        {
            InitializeValues(false);
        }

        #endregion Constructor

        public override XLRangeType RangeType
        {
            get { return XLRangeType.Table; }
        }

        private IXLRangeAddress _lastRangeAddress;
        private Dictionary<String, IXLTableField> _fieldNames = null;

        public Dictionary<String, IXLTableField> FieldNames
        {
            get
            {
                if (_fieldNames != null && _lastRangeAddress != null && _lastRangeAddress.Equals(RangeAddress))
                    return _fieldNames;

                if (_fieldNames == null)
                {
                    _fieldNames = new Dictionary<String, IXLTableField>();
                    _lastRangeAddress = RangeAddress;
                    HeadersRow();
                }
                else
                {
                    HeadersRow(false);
                }

                RescanFieldNames();

                _lastRangeAddress = RangeAddress;

                return _fieldNames;
            }
        }

        private void RescanFieldNames()
        {
            if (ShowHeaderRow)
            {
                var detectedFieldNames = new Dictionary<String, IXLTableField>();
                var headersRow = HeadersRow(false);
                Int32 cellPos = 0;
                foreach (var cell in headersRow.Cells())
                {
                    var name = cell.GetString();
                    if (_fieldNames.ContainsKey(name) && _fieldNames[name].Column.ColumnNumber() == cell.Address.ColumnNumber)
                    {
                        (_fieldNames[name] as XLTableField).Index = cellPos;
                        detectedFieldNames.Add(name, _fieldNames[name]);
                        cellPos++;
                        continue;
                    }

                    if (String.IsNullOrWhiteSpace(name))
                    {
                        name = GetUniqueName("Column", cellPos + 1, true);
                        cell.SetValue(name);
                        cell.DataType = XLDataType.Text;
                    }
                    if (_fieldNames.ContainsKey(name))
                        throw new ArgumentException("The header row contains more than one field name '" + name + "'.");

                    _fieldNames.Add(name, new XLTableField(this, name) { Index = cellPos++ });
                    detectedFieldNames.Add(name, _fieldNames[name]);
                }

                _fieldNames.Keys
                    .Where(key => !detectedFieldNames.ContainsKey(key))
                    .ToArray()
                    .ForEach(key => _fieldNames.Remove(key));
            }
            else
            {
                Int32 colCount = ColumnCount();
                for (Int32 i = 1; i <= colCount; i++)
                {
                    if (_fieldNames.Values.All(f => f.Index != i - 1))
                    {
                        var name = "Column" + i;

                        _fieldNames.Add(name, new XLTableField(this, name) { Index = i - 1 });
                    }
                }
            }
        }

        internal void AddFields(IEnumerable<String> fieldNames)
        {
            _fieldNames = new Dictionary<String, IXLTableField>();

            Int32 cellPos = 0;
            foreach (var name in fieldNames)
            {
                _fieldNames.Add(name, new XLTableField(this, name) { Index = cellPos++ });
            }
        }

        internal void RenameField(String oldName, String newName)
        {
            if (!_fieldNames.ContainsKey(oldName))
                throw new ArgumentException("The field does not exist in this table", "oldName");

            var field = _fieldNames[oldName];
            _fieldNames.Remove(oldName);
            _fieldNames.Add(newName, field);
        }

        internal String RelId { get; set; }

        public IXLTableRange DataRange
        {
            get
            {
                XLRange range;

                if (_showHeaderRow)
                {
                    range = _showTotalsRow
                                ? Range(2, 1, RowCount() - 1, ColumnCount())
                                : Range(2, 1, RowCount(), ColumnCount());
                }
                else
                {
                    range = _showTotalsRow
                                ? Range(1, 1, RowCount() - 1, ColumnCount())
                                : Range(1, 1, RowCount(), ColumnCount());
                }

                return new XLTableRange(range, this);
            }
        }

        private XLAutoFilter _autoFilter;

        public XLAutoFilter AutoFilter
        {
            get
            {
                if (_autoFilter == null)
                    _autoFilter = new XLAutoFilter();

                _autoFilter.Range = ShowTotalsRow ? Range(1, 1, RowCount() - 1, ColumnCount()) : AsRange();
                return _autoFilter;
            }
        }

        public new IXLBaseAutoFilter SetAutoFilter()
        {
            return AutoFilter;
        }

        protected override void OnRangeAddressChanged(XLRangeAddress oldAddress, XLRangeAddress newAddress)
        {
            //Do nothing for table
        }

        #region IXLTable Members

        public Boolean EmphasizeFirstColumn { get; set; }
        public Boolean EmphasizeLastColumn { get; set; }
        public Boolean ShowRowStripes { get; set; }
        public Boolean ShowColumnStripes { get; set; }

        private Boolean _showAutoFilter;

        public Boolean ShowAutoFilter
        {
            get { return _showHeaderRow && _showAutoFilter; }
            set { _showAutoFilter = value; }
        }

        public XLTableTheme Theme { get; set; }

        public String Name
        {
            get { return _name; }
            set
            {
                if (_name == value) return;

                // Validation rules for table names
                var oldname = _name ?? string.Empty;

                if (!XLHelper.ValidateName("table", value, oldname, Worksheet.Tables.Select(t => t.Name), out String message))
                    throw new ArgumentException(message, nameof(value));

                _name = value;

                // Some totals row formula depend on the table name. Update them.
                if (_fieldNames?.Any() ?? false)
                    this.Fields.ForEach(f => (f as XLTableField).UpdateTableFieldTotalsRowFormula());

                if (!String.IsNullOrWhiteSpace(oldname) && !String.Equals(oldname, _name, StringComparison.OrdinalIgnoreCase))
                {
                    Worksheet.Tables.Add(this);
                    Worksheet.Tables.Remove(oldname);
                }
            }
        }

        public Boolean ShowTotalsRow
        {
            get { return _showTotalsRow; }
            set
            {
                if (value && !_showTotalsRow)
                    InsertRowsBelow(1);
                else if (!value && _showTotalsRow)
                    TotalsRow().Delete();

                _showTotalsRow = value;

                if (_showTotalsRow)
                {
                    AutoFilter.Range = Worksheet.Range(
                        RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber,
                        RangeAddress.LastAddress.RowNumber - 1, RangeAddress.LastAddress.ColumnNumber);
                }
                else
                    AutoFilter.Range = Worksheet.Range(RangeAddress);
            }
        }

        public IXLRangeRow HeadersRow()
        {
            return HeadersRow(true);
        }

        internal IXLRangeRow HeadersRow(Boolean scanForNewFieldsNames)
        {
            if (!ShowHeaderRow) return null;

            if (scanForNewFieldsNames)
            {
                var tempResult = FieldNames;
            }

            return FirstRow();
        }

        public IXLRangeRow TotalsRow()
        {
            return ShowTotalsRow ? LastRow() : null;
        }

        public IXLTableField Field(String fieldName)
        {
            return Field(GetFieldIndex(fieldName));
        }

        public IXLTableField Field(Int32 fieldIndex)
        {
            return FieldNames.Values.First(f => f.Index == fieldIndex);
        }

        public IEnumerable<IXLTableField> Fields
        {
            get
            {
                Int32 columnCount = ColumnCount();
                for (int co = 0; co < columnCount; co++)
                    yield return Field(co);
            }
        }

        public IXLTable Resize(IXLRangeAddress rangeAddress)
        {
            return Resize(Worksheet.Range(RangeAddress));
        }

        public IXLTable Resize(string rangeAddress)
        {
            return Resize(Worksheet.Range(RangeAddress));
        }

        public IXLTable Resize(IXLCell firstCell, IXLCell lastCell)
        {
            return Resize(Worksheet.Range(firstCell, lastCell));
        }

        public IXLTable Resize(string firstCellAddress, string lastCellAddress)
        {
            return Resize(Worksheet.Range(firstCellAddress, lastCellAddress));
        }

        public IXLTable Resize(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            return Resize(Worksheet.Range(firstCellAddress, lastCellAddress));
        }

        public IXLTable Resize(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            return Resize(Worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn));
        }

        public IXLTable Resize(IXLRange range)
        {
            if (!this.ShowHeaderRow)
                throw new NotImplementedException("Resizing of tables with no headers not supported yet.");

            if (this.Worksheet != range.Worksheet)
                throw new InvalidOperationException("You cannot resize a table to a range on a different sheet.");

            var totalsRowChanged = this.ShowTotalsRow ? range.LastRow().RowNumber() - this.TotalsRow().RowNumber() : 0;
            var oldTotalsRowNumber = this.ShowTotalsRow ? this.TotalsRow().RowNumber() : -1;

            var existingHeaders = this.FieldNames.Keys;
            var newHeaders = new HashSet<string>();

            // Force evaluation of f.Column field
            var tempArray = this.Fields.Select(f => f.Column).ToArray();

            var firstRow = range.Row(1);
            if (!firstRow.FirstCell().Address.Equals(this.HeadersRow().FirstCell().Address)
                || !firstRow.LastCell().Address.Equals(this.HeadersRow().LastCell().Address))
            {
                _uniqueNames.Clear();
                var co = 1;
                foreach (var c in firstRow.Cells())
                {
                    if (String.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                        c.Value = GetUniqueName("Column", co, true);

                    var header = c.GetString();
                    _uniqueNames.Add(header);

                    if (!existingHeaders.Contains(header))
                        newHeaders.Add(header);

                    co++;
                }
            }

            if (totalsRowChanged < 0)
            {
                range.Rows(r => r.RowNumber().Equals(this.TotalsRow().RowNumber() + totalsRowChanged)).Single().InsertRowsAbove(1);
                range = Worksheet.Range(range.FirstCell(), range.LastCell().CellAbove());
                oldTotalsRowNumber++;
            }
            else if (totalsRowChanged > 0)
            {
                this.TotalsRow().RowBelow(totalsRowChanged + 1).InsertRowsAbove(1);
                this.TotalsRow().AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
            }

            this.RangeAddress = (XLRangeAddress)range.RangeAddress;
            RescanFieldNames();

            if (this.ShowTotalsRow)
            {
                foreach (var f in this._fieldNames.Values)
                {
                    var c = this.TotalsRow().Cell(f.Index + 1);
                    if (!c.IsEmpty() && newHeaders.Contains(f.Name))
                    {
                        f.TotalsRowLabel = c.GetFormattedString();
                        c.DataType = XLDataType.Text;
                    }
                }

                if (totalsRowChanged != 0)
                {
                    foreach (var f in this._fieldNames.Values.Cast<XLTableField>())
                    {
                        f.UpdateTableFieldTotalsRowFormula();
                        var c = this.TotalsRow().Cell(f.Index + 1);
                        if (!String.IsNullOrWhiteSpace(f.TotalsRowLabel))
                        {
                            c.DataType = XLDataType.Text;

                            //Remove previous row's label
                            var oldTotalsCell = this.Worksheet.Cell(oldTotalsRowNumber, f.Column.ColumnNumber());
                            if (oldTotalsCell.Value.ToString() == f.TotalsRowLabel)
                                oldTotalsCell.Value = null;
                        }

                        if (f.TotalsRowFunction != XLTotalsRowFunction.None)
                            c.DataType = XLDataType.Number;
                    }
                }
            }

            return this;
        }

        public IXLTable SetEmphasizeFirstColumn()
        {
            EmphasizeFirstColumn = true;
            return this;
        }

        public IXLTable SetEmphasizeFirstColumn(Boolean value)
        {
            EmphasizeFirstColumn = value;
            return this;
        }

        public IXLTable SetEmphasizeLastColumn()
        {
            EmphasizeLastColumn = true;
            return this;
        }

        public IXLTable SetEmphasizeLastColumn(Boolean value)
        {
            EmphasizeLastColumn = value;
            return this;
        }

        public IXLTable SetShowRowStripes()
        {
            ShowRowStripes = true;
            return this;
        }

        public IXLTable SetShowRowStripes(Boolean value)
        {
            ShowRowStripes = value;
            return this;
        }

        public IXLTable SetShowColumnStripes()
        {
            ShowColumnStripes = true;
            return this;
        }

        public IXLTable SetShowColumnStripes(Boolean value)
        {
            ShowColumnStripes = value;
            return this;
        }

        public IXLTable SetShowTotalsRow()
        {
            ShowTotalsRow = true;
            return this;
        }

        public IXLTable SetShowTotalsRow(Boolean value)
        {
            ShowTotalsRow = value;
            return this;
        }

        public IXLTable SetShowAutoFilter()
        {
            ShowAutoFilter = true;
            return this;
        }

        public IXLTable SetShowAutoFilter(Boolean value)
        {
            ShowAutoFilter = value;
            return this;
        }

        public new IXLRange Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending,
                                 Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            var toSortBy = new StringBuilder();
            foreach (string coPairTrimmed in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
            {
                String coString;
                String order;
                if (coPairTrimmed.Contains(' '))
                {
                    var pair = coPairTrimmed.Split(' ');
                    coString = pair[0];
                    order = pair[1];
                }
                else
                {
                    coString = coPairTrimmed;
                    order = sortOrder == XLSortOrder.Ascending ? "ASC" : "DESC";
                }

                if (!Int32.TryParse(coString, out Int32 co))
                    co = Field(coString).Index + 1;

                if (toSortBy.Length > 0)
                    toSortBy.Append(',');

                toSortBy.Append(co);
                toSortBy.Append(' ');
                toSortBy.Append(order);
            }
            return DataRange.Sort(toSortBy.ToString(), sortOrder, matchCase, ignoreBlanks);
        }

        public new IXLTable Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        IXLBaseAutoFilter IXLTable.AutoFilter
        {
            get { return AutoFilter; }
        }

        #endregion IXLTable Members

        private void InitializeValues(Boolean setAutofilter)
        {
            ShowRowStripes = true;
            _showHeaderRow = true;
            Theme = XLTableTheme.TableStyleMedium2;
            if (setAutofilter)
                InitializeAutoFilter();

            AsRange().Row(1).DataType = XLDataType.Text;

            if (RowCount() == 1)
                InsertRowsBelow(1);
        }

        public void InitializeAutoFilter()
        {
            ShowAutoFilter = true;
        }

        internal void OnAddedToTables()
        {
            _uniqueNames = new HashSet<string>();
            Int32 co = 1;
            foreach (IXLCell c in Row(1).Cells())
            {
                if (String.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                    c.Value = GetUniqueName("Column", co, true);
                _uniqueNames.Add(c.GetString());
                co++;
            }
        }

        private String GetUniqueName(String originalName, Int32 initialOffset, Boolean enforceOffset)
        {
            String name = String.Concat(originalName, enforceOffset ? initialOffset.ToInvariantString() : string.Empty);
            if (_uniqueNames?.Contains(name) ?? false)
            {
                Int32 i = initialOffset;
                name = originalName + i.ToInvariantString();
                while (_uniqueNames.Contains(name))
                {
                    i++;
                    name = originalName + i.ToInvariantString();
                }
            }

            return name;
        }

        public Int32 GetFieldIndex(String name)
        {
            // There is a discrepancy in the way headers with line breaks are stored.
            // The entry in the table definition will contain \r\n
            // but the shared string value of the actual cell will contain only \n
            name = name.Replace("\r\n", "\n");
            if (FieldNames.ContainsKey(name))
                return FieldNames[name].Index;

            throw new ArgumentOutOfRangeException("The header row doesn't contain field name '" + name + "'.");
        }

        internal Boolean _showHeaderRow;

        public Boolean ShowHeaderRow
        {
            get { return _showHeaderRow; }
            set
            {
                if (_showHeaderRow == value) return;

                if (_showHeaderRow)
                {
                    var headersRow = HeadersRow();
                    _uniqueNames = new HashSet<string>();
                    Int32 co = 1;
                    foreach (IXLCell c in headersRow.Cells())
                    {
                        if (String.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                            c.Value = GetUniqueName("Column", co, true);
                        _uniqueNames.Add(c.GetString());
                        co++;
                    }

                    headersRow.Clear();
                    RangeAddress = new XLRangeAddress(
                        new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + 1,
                                      RangeAddress.FirstAddress.ColumnNumber,
                                      RangeAddress.FirstAddress.FixedRow,
                                      RangeAddress.FirstAddress.FixedColumn),
                        RangeAddress.LastAddress);
                }
                else
                {
                    var asRange = Worksheet.Range(
                        RangeAddress.FirstAddress.RowNumber - 1,
                        RangeAddress.FirstAddress.ColumnNumber,
                        RangeAddress.LastAddress.RowNumber,
                        RangeAddress.LastAddress.ColumnNumber);
                    var firstRow = asRange.FirstRow();
                    IXLRangeRow rangeRow;
                    if (firstRow.IsEmpty(true))
                    {
                        rangeRow = firstRow;
                        RangeAddress = new XLRangeAddress(
                            new XLAddress(Worksheet,
                                RangeAddress.FirstAddress.RowNumber - 1,
                                RangeAddress.FirstAddress.ColumnNumber,
                                RangeAddress.FirstAddress.FixedRow,
                                RangeAddress.FirstAddress.FixedColumn),
                            RangeAddress.LastAddress);
                    }
                    else
                    {
                        var fAddress = RangeAddress.FirstAddress;
                        //var lAddress = RangeAddress.LastAddress;

                        rangeRow = firstRow.InsertRowsBelow(1, false).First();

                        RangeAddress = new XLRangeAddress(
                            fAddress,
                            RangeAddress.LastAddress);
                    }

                    Int32 co = 1;
                    foreach (var name in FieldNames.Values.Select(f => f.Name))
                    {
                        rangeRow.Cell(co).SetValue(name);
                        co++;
                    }
                }

                _showHeaderRow = value;

                if (_showHeaderRow)
                    HeadersRow().DataType = XLDataType.Text;
            }
        }

        public IXLTable SetShowHeaderRow()
        {
            return SetShowHeaderRow(true);
        }

        public IXLTable SetShowHeaderRow(Boolean value)
        {
            ShowHeaderRow = value;
            return this;
        }

        public void ExpandTableRows(Int32 rows)
        {
            RangeAddress = new XLRangeAddress(
                RangeAddress.FirstAddress,
                new XLAddress(Worksheet, RangeAddress.LastAddress.RowNumber + rows,
                                         RangeAddress.LastAddress.ColumnNumber,
                                         RangeAddress.LastAddress.FixedRow,
                                         RangeAddress.LastAddress.FixedColumn));
        }

        public override XLRangeColumn Column(int columnNumber)
        {
            var column = base.Column(columnNumber);
            column.Table = this;
            return column;
        }

        public override XLRangeColumn Column(string columnName)
        {
            var column = base.Column(columnName);
            column.Table = this;
            return column;
        }

        public override IXLRangeColumns Columns(int firstColumn, int lastColumn)
        {
            var columns = base.Columns(firstColumn, lastColumn);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public override IXLRangeColumns Columns(Func<IXLRangeColumn, bool> predicate = null)
        {
            var columns = base.Columns(predicate);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public override IXLRangeColumns Columns(string columns)
        {
            var cols = base.Columns(columns);
            cols.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return cols;
        }

        public override IXLRangeColumns Columns(string firstColumn, string lastColumn)
        {
            var columns = base.Columns(firstColumn, lastColumn);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public override XLRangeColumns ColumnsUsed(bool includeFormats, Func<IXLRangeColumn, bool> predicate = null)
        {
            var columns = base.ColumnsUsed(includeFormats, predicate);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public override XLRangeColumns ColumnsUsed(Func<IXLRangeColumn, bool> predicate = null)
        {
            var columns = base.ColumnsUsed(predicate);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public IEnumerable<dynamic> AsDynamicEnumerable()
        {
            foreach (var row in this.DataRange.Rows())
            {
                dynamic expando = new ExpandoObject();
                foreach (var f in this.Fields)
                {
                    var value = row.Cell(f.Index + 1).Value;
                    // ExpandoObject supports IDictionary so we can extend it like this
                    var expandoDict = expando as IDictionary<string, object>;
                    if (expandoDict.ContainsKey(f.Name))
                        expandoDict[f.Name] = value;
                    else
                        expandoDict.Add(f.Name, value);
                }

                yield return expando;
            }
        }

        public DataTable AsNativeDataTable()
        {
            var table = new DataTable(this.Name);

            foreach (var f in Fields.Cast<XLTableField>())
            {
                Type type = typeof(object);
                if (f.IsConsistentDataType())
                {
                    var c = f.Column.Cells().Skip(this.ShowHeaderRow ? 1 : 0).First();
                    switch (c.DataType)
                    {
                        case XLDataType.Text:
                            type = typeof(String);
                            break;

                        case XLDataType.Boolean:
                            type = typeof(Boolean);
                            break;

                        case XLDataType.DateTime:
                            type = typeof(DateTime);
                            break;

                        case XLDataType.TimeSpan:
                            type = typeof(TimeSpan);
                            break;

                        case XLDataType.Number:
                            type = typeof(Double);
                            break;
                    }
                }

                table.Columns.Add(f.Name, type);
            }

            foreach (var row in this.DataRange.Rows())
            {
                var dr = table.NewRow();

                foreach (var f in this.Fields)
                {
                    dr[f.Name] = row.Cell(f.Index + 1).Value;
                }

                table.Rows.Add(dr);
            }

            return table;
        }

        public IXLTable CopyTo(IXLWorksheet targetSheet)
        {
            return CopyTo((XLWorksheet)targetSheet);
        }

        internal IXLTable CopyTo(XLWorksheet targetSheet, bool copyData = true)
        {
            if (targetSheet == Worksheet)
                throw new InvalidOperationException("Cannot copy table to the worksheet it already belongs to.");

            var targetRange = targetSheet.Range(RangeAddress.WithoutWorksheet());
            if (copyData)
                RangeUsed().CopyTo(targetRange);
            else
                HeadersRow().CopyTo(targetRange.FirstRow());

            String tableName = Name;
            var newTable = (XLTable)targetSheet.Table(targetRange, tableName, true);

            newTable.RelId = RelId;
            newTable.EmphasizeFirstColumn = EmphasizeFirstColumn;
            newTable.EmphasizeLastColumn = EmphasizeLastColumn;
            newTable.ShowRowStripes = ShowRowStripes;
            newTable.ShowColumnStripes = ShowColumnStripes;
            newTable.ShowAutoFilter = ShowAutoFilter;
            newTable.Theme = Theme;
            newTable._showTotalsRow = ShowTotalsRow;

            Int32 fieldCount = ColumnCount();
            for (Int32 f = 0; f < fieldCount; f++)
            {
                var tableField = newTable.Field(f) as XLTableField;
                var tField = Field(f) as XLTableField;
                tableField.Index = tField.Index;
                tableField.Name = tField.Name;
                tableField.totalsRowLabel = tField.totalsRowLabel;
                tableField.totalsRowFunction = tField.totalsRowFunction;
            }
            return newTable;
        }

        #region Append and replace data

        public IXLRange AppendData(IEnumerable data)
        {
            return AppendData(data, false);
        }

        public IXLRange AppendData(IEnumerable data, bool transpose)
        {
            if (data == null || data is String)
                return null;

            var numberOfNewRows = data.Cast<object>().Count();
            var lastRowOfOldRange = this.DataRange.LastRow();
            lastRowOfOldRange.InsertRowsBelow(numberOfNewRows);

            return OverwriteTableData(lastRowOfOldRange.RowBelow(), data, transpose);
        }

        public IXLRange AppendData(DataTable dataTable)
        {
            return AppendData(dataTable.Rows.Cast<DataRow>());
        }

        public IXLRange AppendData<T>(IEnumerable<T> data)
        {
            if (!data.Any() || data is String)
                return null;

            var numberOfNewRows = data.Count();
            var lastRowOfOldRange = this.DataRange.LastRow();
            lastRowOfOldRange.InsertRowsBelow(numberOfNewRows);

            return OverwriteTableData(lastRowOfOldRange.RowBelow(), data);
        }

        public IXLRange ReplaceData(IEnumerable data)
        {
            return ReplaceData(data, false);
        }

        public IXLRange ReplaceData(IEnumerable data, bool transpose)
        {
            if (data == null || data is String)
                throw new InvalidOperationException("Cannot replace table data with empty enumerable.");

            // Resize table
            var sizeDifference = data.Cast<object>().Count() - this.DataRange.RowCount();
            if (sizeDifference > 0)
                this.DataRange.LastRow().InsertRowsBelow(sizeDifference);
            else if (sizeDifference < 0)
                this.DataRange.Rows(this.LastRow().RowNumber() + sizeDifference + 1, this.LastRow().RowNumber())
                    .OrderByDescending(r => r.RowNumber())
                    .ForEach(r => r.Delete());

            this.DataRange.Clear(XLClearOptions.Contents);

            return this.OverwriteTableData(this.DataRange.FirstRow(), data, transpose);
        }

        public IXLRange ReplaceData(DataTable dataTable)
        {
            return ReplaceData(dataTable.Rows.Cast<DataRow>());
        }

        public IXLRange ReplaceData<T>(IEnumerable<T> data)
        {
            if (!data.Any() || data is String)
                throw new InvalidOperationException("Cannot replace table data with empty enumerable.");

            // Resize table
            var sizeDifference = data.Count() - this.DataRange.RowCount();
            if (sizeDifference > 0)
                this.DataRange.LastRow().InsertRowsBelow(sizeDifference);
            else if (sizeDifference < 0)
                this.DataRange.Rows(this.LastRow().RowNumber() + sizeDifference + 1, this.LastRow().RowNumber())
                    .OrderByDescending(r => r.RowNumber())
                    .ForEach(r => r.Delete());

            this.DataRange.Clear(XLClearOptions.Contents);

            return this.OverwriteTableData(this.DataRange.FirstRow(), data);
        }

        internal IXLRange OverwriteTableData(IXLTableRow firstRow, IEnumerable data, bool transpose)
        {
            if (data == null || data is String)
                return null;

            var firstCell = firstRow.FirstCell();
            var rowNumber = firstCell.Address.RowNumber;
            var columnNumber = firstCell.Address.ColumnNumber;

            var maxColumnNumber = 0;
            var maxRowNumber = 0;
            var isDataTable = false;
            var isDataReader = false;

            const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
            var memberCache = new Dictionary<Type, IEnumerable<MemberInfo>>();
            var accessorCache = new Dictionary<Type, TypeAccessor>();
            IEnumerable<MemberInfo> members = null;
            TypeAccessor accessor = null;

            foreach (var m in data)
            {
                var itemType = m.GetType();

                if (transpose)
                    rowNumber = firstCell.Address.RowNumber;
                else
                    columnNumber = firstCell.Address.ColumnNumber;

                if (itemType.IsPrimitive || itemType == typeof(String) || itemType == typeof(DateTime) || itemType.IsNumber())
                {
                    Worksheet.SetValue(m, rowNumber, columnNumber);

                    if (transpose)
                        rowNumber++;
                    else
                        columnNumber++;
                }
                else if (itemType.IsArray)
                {
                    foreach (var item in (Array)m)
                    {
                        Worksheet.SetValue(item, rowNumber, columnNumber);

                        if (transpose)
                            rowNumber++;
                        else
                            columnNumber++;
                    }
                }
                else if (isDataTable || m is DataRow)
                {
                    if (!isDataTable)
                        isDataTable = true;

                    foreach (var item in (m as DataRow).ItemArray)
                    {
                        Worksheet.SetValue(item, rowNumber, columnNumber);

                        if (transpose)
                            rowNumber++;
                        else
                            columnNumber++;
                    }
                }
                else if (isDataReader || m is IDataRecord)
                {
                    if (!isDataReader)
                        isDataReader = true;

                    var record = m as IDataRecord;

                    var fieldCount = record.FieldCount;
                    for (var i = 0; i < fieldCount; i++)
                    {
                        Worksheet.SetValue(record[i], rowNumber, columnNumber);

                        if (transpose)
                            rowNumber++;
                        else
                            columnNumber++;
                    }
                }
                else
                {
                    if (!memberCache.ContainsKey(itemType))
                    {
                        var _accessor = TypeAccessor.Create(itemType);

                        var _members = itemType.GetFields(bindingFlags).Cast<MemberInfo>()
                             .Concat(itemType.GetProperties(bindingFlags))
                             .Where(mi => !XLColumnAttribute.IgnoreMember(mi))
                             .OrderBy(mi => XLColumnAttribute.GetOrder(mi));

                        memberCache.Add(itemType, _members);
                        accessorCache.Add(itemType, _accessor);
                    }

                    accessor = accessorCache[itemType];
                    members = memberCache[itemType];

                    foreach (var mi in members)
                    {
                        if (mi.MemberType == MemberTypes.Property && (mi as PropertyInfo).GetGetMethod().IsStatic)
                            Worksheet.SetValue((mi as PropertyInfo).GetValue(null, null), rowNumber, columnNumber);
                        else if (mi.MemberType == MemberTypes.Field && (mi as FieldInfo).IsStatic)
                            Worksheet.SetValue((mi as FieldInfo).GetValue(null), rowNumber, columnNumber);
                        else
                            Worksheet.SetValue(accessor[m, mi.Name], rowNumber, columnNumber);

                        if (transpose)
                            rowNumber++;
                        else
                            columnNumber++;
                    }
                }

                if (transpose)
                    columnNumber++;
                else
                    rowNumber++;

                if (columnNumber > maxColumnNumber)
                    maxColumnNumber = columnNumber;

                if (rowNumber > maxRowNumber)
                    maxRowNumber = rowNumber;
            }

            var lastCell = this.Worksheet.Cell(maxRowNumber - 1, maxColumnNumber - 1);

            return Worksheet.Range(firstCell, lastCell);
        }

        internal IXLRange OverwriteTableData<T>(IXLTableRow firstRow, IEnumerable<T> data)
        {
            if (!(data?.Any() ?? false) || data is String)
                return null;

            var firstCell = firstRow.FirstCell();

            var currentRowNumber = firstCell.Address.RowNumber;
            var maximumColumnNumber = 0;
            var maximumRowNumber = 0;

            var isDataTable = false;
            var isDataReader = false;
            var itemType = data.GetItemType();

            if (itemType.IsPrimitive || itemType == typeof(String) || itemType == typeof(DateTime) || itemType.IsNumber())
            {
                foreach (object o in data)
                {
                    var currentColumnNumber = firstCell.Address.ColumnNumber;
                    Worksheet.SetValue(o, currentRowNumber, currentColumnNumber);
                    maximumColumnNumber = Math.Max(currentColumnNumber, maximumColumnNumber);

                    currentRowNumber++;
                }
            }
            else
            {
                const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
                var memberCache = new Dictionary<Type, IEnumerable<MemberInfo>>();
                var accessorCache = new Dictionary<Type, TypeAccessor>();
                IEnumerable<MemberInfo> members = null;
                TypeAccessor accessor = null;
                bool isPlainObject = itemType == typeof(object);

                if (!isPlainObject)
                {
                    members = itemType.GetFields(bindingFlags).Cast<MemberInfo>()
                        .Concat(itemType.GetProperties(bindingFlags))
                        .Where(mi => !XLColumnAttribute.IgnoreMember(mi))
                        .OrderBy(mi => XLColumnAttribute.GetOrder(mi));

                    accessor = TypeAccessor.Create(itemType);
                }

                foreach (T m in data)
                {
                    var currentColumnNumber = firstCell.Address.ColumnNumber;

                    if (isPlainObject)
                    {
                        // In this case data is just IEnumerable<object>, which means we have to determine the runtime type of each element
                        // This is very inefficient and we prefer type of T to be a concrete class or struct
                        var type = m.GetType();
                        if (!memberCache.ContainsKey(type))
                        {
                            var _accessor = TypeAccessor.Create(type);

                            var _members = type.GetFields(bindingFlags).Cast<MemberInfo>()
                                 .Concat(type.GetProperties(bindingFlags))
                                 .Where(mi => !XLColumnAttribute.IgnoreMember(mi))
                                 .OrderBy(mi => XLColumnAttribute.GetOrder(mi));

                            memberCache.Add(type, _members);
                            accessorCache.Add(type, _accessor);
                        }

                        members = memberCache[type];
                        accessor = accessorCache[type];
                    }
                    else if (members == null)
                        return null;

                    if (itemType.IsArray)
                    {
                        foreach (var item in (m as Array))
                        {
                            Worksheet.SetValue(item, currentRowNumber, currentColumnNumber);
                            currentColumnNumber++;
                        }
                    }
                    else if (isDataTable || m is DataRow)
                    {
                        var row = m as DataRow;
                        if (!isDataTable)
                            isDataTable = true;

                        foreach (var item in row.ItemArray)
                        {
                            Worksheet.SetValue(item, currentRowNumber, currentColumnNumber);
                            currentColumnNumber++;
                        }
                    }
                    else if (isDataReader || m is IDataRecord)
                    {
                        if (!isDataReader)
                            isDataReader = true;

                        var record = m as IDataRecord;

                        var fieldCount = record.FieldCount;

                        for (var i = 0; i < fieldCount; i++)
                        {
                            Worksheet.SetValue(record[i], currentRowNumber, currentColumnNumber);
                            currentColumnNumber++;
                        }
                    }
                    else
                    {
                        foreach (var mi in members)
                        {
                            if (mi.MemberType == MemberTypes.Property && (mi as PropertyInfo).GetGetMethod().IsStatic)
                                Worksheet.SetValue((mi as PropertyInfo).GetValue(null, null), currentRowNumber, currentColumnNumber);
                            else if (mi.MemberType == MemberTypes.Field && (mi as FieldInfo).IsStatic)
                                Worksheet.SetValue((mi as FieldInfo).GetValue(null), currentRowNumber, currentColumnNumber);
                            else
                                Worksheet.SetValue(accessor[m, mi.Name], currentRowNumber, currentColumnNumber);

                            currentColumnNumber++;
                        }
                    }

                    maximumColumnNumber = Math.Max(maximumColumnNumber, currentColumnNumber - 1);
                    maximumRowNumber = Math.Max(maximumRowNumber, currentRowNumber);

                    currentRowNumber++;
                }
            }

            ClearMerged();

            var lastCell = this.Worksheet.Cell(maximumRowNumber, maximumColumnNumber);

            return Worksheet.Range(firstCell, lastCell);
        }

        #endregion Append and replace data

        private void ClearMerged()
        {
            List<IXLRange> mergeToDelete = Worksheet.Internals.MergedRanges.GetIntersectedRanges(this.AsRange().RangeAddress).ToList();

            mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
        }
    }
}
