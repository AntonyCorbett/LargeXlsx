/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2025 Salvatore ISAJA. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED THE COPYRIGHT HOLDER ``AS IS'' AND ANY EXPRESS
OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN
NO EVENT SHALL THE COPYRIGHT HOLDER BE LIABLE FOR ANY DIRECT,
INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace LargeXlsx
{
    internal class Worksheet : IDisposable
    {
        private readonly Stream _stream;
        private readonly Stylesheet _stylesheet;
        private readonly SharedStringTable _sharedStringTable;
        private readonly bool _requireCellReferences;
        private readonly bool _skipInvalidCharacters;
        private readonly List<string> _mergedCellRefs;
        private readonly Dictionary<XlsxDataValidation, List<string>> _cellRefsByDataValidation;
        private readonly HashSet<int> _pageBreakRowNumbers;
        private readonly HashSet<int> _pageBreakColumnNumbers;
        private readonly ContentBuffer _contentBuffer;
        private string _autoFilterRef;
        private string _autoFilterAbsoluteRef;
        private XlsxSheetProtection _sheetProtection;
        private XlsxHeaderFooter _headerFooter;
        private bool _needsRef;
        private string _stringedCurrentRowNumber;

        public int Id { get; }
        public string Name { get; }
        public XlsxWorksheetState State { get; }
        public int CurrentRowNumber { get; private set; }
        public int CurrentColumnNumber { get; private set; }
        internal string AutoFilterAbsoluteRef => _autoFilterAbsoluteRef;

        public Worksheet(
            ZipArchive zipArchive,
            CompressionLevel compressionLevel,
            int id,
            string name,
            int splitRow,
            int splitColumn,
            bool rightToLeft,
            Stylesheet stylesheet,
            SharedStringTable sharedStringTable,
            IEnumerable<XlsxColumn> columns,
            bool showGridLines,
            bool showHeaders,
            bool requireCellReferences,
            bool skipInvalidCharacters,
            XlsxWorksheetState state)
        {
            Id = id;
            Name = name;
            State = state;
            CurrentRowNumber = 0;
            CurrentColumnNumber = 0;
            _stylesheet = stylesheet;
            _sharedStringTable = sharedStringTable;
            _requireCellReferences = requireCellReferences;
            _skipInvalidCharacters = skipInvalidCharacters;
            _mergedCellRefs = new List<string>();
            _pageBreakRowNumbers = new HashSet<int>();
            _pageBreakColumnNumbers = new HashSet<int>();
            _cellRefsByDataValidation = new Dictionary<XlsxDataValidation, List<string>>();
            var entry = zipArchive.CreateEntry($"xl/worksheets/sheet{id}.xml", compressionLevel);
            _stream = entry.Open();
            _contentBuffer = new ContentBuffer(new InvariantCultureStreamWriter(_stream));

            _contentBuffer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                                + "<sheetViews>"
                                + $"<sheetView showGridLines=\"{Util.BoolToInt(showGridLines)}\" showRowColHeaders=\"{Util.BoolToInt(showHeaders)}\""
                                + $" rightToLeft=\"{Util.BoolToInt(rightToLeft)}\" workbookViewId=\"0\">\n");

            if (splitRow > 0 || splitColumn > 0)
                FreezePanes(splitRow, splitColumn);
            _contentBuffer.Write("</sheetView></sheetViews>\n");
            WriteColumns(columns);
            _contentBuffer.Write("<sheetData>\n");
        }

        public void Dispose()
        {
            CloseLastRow();
            _contentBuffer.Write("</sheetData>\n");
            WriteSheetProtection();
            WriteAutoFilter();
            WriteMergedCells();
            WriteDataValidations();
            WriteHeaderFooter();
            WritePageBreaks();
            _contentBuffer.Write("</worksheet>\n");
            _contentBuffer.Dispose();
            _stream.Dispose();
        }

        public void BeginRow(double? height, bool hidden, XlsxStyle style)
        {
            CloseLastRow();
            if (CurrentRowNumber == Limits.MaxRowCount)
                throw new InvalidOperationException($"A worksheet can contain at most {Limits.MaxRowCount} rows ({CurrentRowNumber + 1} attempted)");
            CurrentRowNumber++;
            _stringedCurrentRowNumber = null;
            CurrentColumnNumber = 1;
            _contentBuffer.Write("<row");
            if (_requireCellReferences || _needsRef)
            {
                _contentBuffer.Write(" r=\"");
                WriteCurrentRowNumber();
                _contentBuffer.Write("\"");
                _needsRef = false;
            }
            if (height.HasValue)
                _contentBuffer.Write(" ht=\"{0}\" customHeight=\"1\"", height);
            if (hidden)
                _contentBuffer.Write(" hidden=\"1\"");
            if (style != null)
                _contentBuffer.Write(" s=\"{0}\" customFormat=\"1\"", _stylesheet.ResolveStyleId(style));
            _contentBuffer.Write(">\n");
        }

        public void SkipRows(int rowCount)
        {
            CloseLastRow();
            _needsRef = true;
            if (CurrentRowNumber + rowCount > Limits.MaxRowCount)
                throw new InvalidOperationException($"A worksheet can contain at most {Limits.MaxRowCount} rows ({CurrentRowNumber + rowCount} attempted)");
            CurrentRowNumber += rowCount;
        }

        public void SkipColumns(int columnCount)
        {
            EnsureRow();
            _needsRef = true;
            CurrentColumnNumber += columnCount;
        }


        public void Write(XlsxStyle style, int repeatCount)
        {
            EnsureRow();
            var styleId = _stylesheet.ResolveStyleId(style);
            for (var i = 0; i < repeatCount; i++)
            {
                // <c r="{0}{1}" s="{2}"/>
                _contentBuffer.Write("<c");
                WriteCellRef();
                WriteStyle(styleId);
                _contentBuffer.Write("/>\n");
                CurrentColumnNumber++;
            }
        }

        public void Write(string value, XlsxStyle style)
        {
            if (value == null)
            {
                Write(style, 1);
                return;
            }
            EnsureRow();
            // <c r="{0}{1}" s="{2}" t="inlineStr"><is><t xml:space="preserve">{3}</t></is></c>
            _contentBuffer.Write("<c");
            WriteCellRef();
            WriteStyle(style);
            _contentBuffer
                .Append(" t=\"inlineStr\"><is><t")
                .AddSpacePreserveIfNeeded(value)
                .Append(">")
                .AppendEscapedXmlText(value, _skipInvalidCharacters)
                .Append("</t></is></c>\n");
            CurrentColumnNumber++;
        }

        public void Write(double value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}"><v>{3}</v></c>
            _contentBuffer.Write("<c");
            WriteCellRef();
            WriteStyle(style);
            _contentBuffer.Append("><v>").Append(value).Append("</v></c>\n");
            CurrentColumnNumber++;
        }

        public void Write(decimal value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}"><v>{3}</v></c>
            _contentBuffer.Write("<c");
            WriteCellRef();
            WriteStyle(style);
            _contentBuffer.Append("><v>").Append(value).Append("</v></c>\n");
            CurrentColumnNumber++;
        }

        public void Write(int value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}"><v>{3}</v></c>
            _contentBuffer.Write("<c");
            WriteCellRef();
            WriteStyle(style);
            _contentBuffer.Append("><v>").Append(value).Append("</v></c>\n");
            CurrentColumnNumber++;
        }

        public void Write(bool value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}" t="b"><v>{3}</v></c>
            _contentBuffer.Write("<c");
            WriteCellRef();
            WriteStyle(style);
            _contentBuffer
                .Append(" t=\"b\"><v>")
                .Append(Util.BoolToInt(value))
                .Append("</v></c>\n");
            CurrentColumnNumber++;
        }

        public void WriteFormula(string formula, XlsxStyle style, IConvertible result)
        {
            // <c r="{0}{1}" s="{2}" t="str"><f>{3}</f><v>{4}</v></c>
            EnsureRow();
            _contentBuffer.Write("<c");
            WriteCellRef();
            WriteStyle(style);
            _contentBuffer.Append(" t=\"str\"><f>")
                .AppendEscapedXmlText(formula, _skipInvalidCharacters)
                .Append("</f>");
            if (result != null)
                _contentBuffer
                    .Append("<v>")
                    .AppendEscapedXmlText(result.ToString(CultureInfo.InvariantCulture), _skipInvalidCharacters)
                    .Append("</v>");
            _contentBuffer.Write("</c>\n");
            CurrentColumnNumber++;
        }

        public void WriteSharedString(string value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}" t="s"><v>{3}</v></c>
            _contentBuffer.Write("<c");
            WriteCellRef();
            WriteStyle(style);
            _contentBuffer
                .Append(" t=\"s\"><v>")
                .Append(_sharedStringTable.ResolveStringId(value))
                .Append("</v></c>\n");
            CurrentColumnNumber++;
        }

        public void AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            if (rowCount < 1 || columnCount < 1)
                throw new ArgumentOutOfRangeException();
            var toRow = fromRow + rowCount - 1;
            var fromColumnName = Util.GetColumnName(fromColumn);
            var toColumnName = Util.GetColumnName(fromColumn + columnCount - 1);
            _mergedCellRefs.Add($"{fromColumnName}{fromRow}:{toColumnName}{toRow}");
        }

        public void AddRowPageBreakBefore(int rowNumber)
        {
            if (rowNumber <= 1 || rowNumber > Limits.MaxRowCount)
                throw new ArgumentOutOfRangeException(nameof(rowNumber));
            _pageBreakRowNumbers.Add(rowNumber - 1);
        }

        public void AddColumnPageBreakBefore(int columnNumber)
        {
            if (columnNumber <= 1 || columnNumber > Limits.MaxColumnCount)
                throw new ArgumentOutOfRangeException(nameof(columnNumber));
            _pageBreakColumnNumbers.Add(columnNumber - 1);
        }

        public void SetAutoFilter(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            if (rowCount < 1 || columnCount < 1)
                throw new ArgumentOutOfRangeException();
            var toRow = fromRow + rowCount - 1;
            var fromColumnName = Util.GetColumnName(fromColumn);
            var toColumnName = Util.GetColumnName(fromColumn + columnCount - 1);
            _autoFilterRef = $"{fromColumnName}{fromRow}:{toColumnName}{toRow}";
            _autoFilterAbsoluteRef = $"'{Name.Replace("'", "''")}'!${fromColumnName}${fromRow}:${toColumnName}${toRow}";
        }

        public void AddDataValidation(int fromRow, int fromColumn, int rowCount, int columnCount, XlsxDataValidation dataValidation)
        {
            if (rowCount < 1 || columnCount < 1)
                throw new ArgumentOutOfRangeException();
            var cellRef = rowCount > 1 || columnCount > 1
                ? $"{Util.GetColumnName(fromColumn)}{fromRow}:{Util.GetColumnName(fromColumn + columnCount - 1)}{fromRow + rowCount - 1}"
                : $"{Util.GetColumnName(fromColumn)}{fromRow}";
            if (!_cellRefsByDataValidation.TryGetValue(dataValidation, out var cellRefs))
            {
                cellRefs = new List<string>();
                _cellRefsByDataValidation.Add(dataValidation, cellRefs);
            }
            cellRefs.Add(cellRef);
        }

        public void SetSheetProtection(XlsxSheetProtection sheetProtection)
        {
            if (sheetProtection.Password.Length < Limits.MinSheetProtectionPasswordLength || sheetProtection.Password.Length > Limits.MaxSheetProtectionPasswordLength)
                throw new ArgumentException("Invalid password length");
            _sheetProtection = sheetProtection;
        }

        public void SetHeaderFooter(XlsxHeaderFooter headerFooter)
        {
            _headerFooter = headerFooter;
        }

        private void WriteCellRef()
        {
            if (_requireCellReferences || _needsRef)
            {
                _contentBuffer.Write(" r=\"");
                _contentBuffer.Write(Util.GetColumnName(CurrentColumnNumber));
                WriteCurrentRowNumber();
                _contentBuffer.Write("\"");
                _needsRef = false;
            }
        }

        private void WriteCurrentRowNumber()
        {
            if (_stringedCurrentRowNumber == null)
                _stringedCurrentRowNumber = CurrentRowNumber.ToString();
            _contentBuffer.Write(_stringedCurrentRowNumber);
        }

        private void WriteStyle(int styleId)
        {
            if (styleId != 0)
                _contentBuffer.Append(" s=\"").Append(styleId).Append("\"");
        }

        private void WriteStyle(XlsxStyle style)
        {
            WriteStyle(_stylesheet.ResolveStyleId(style));
        }

        private void EnsureRow()
        {
            if (CurrentColumnNumber == 0)
                throw new InvalidOperationException($"{nameof(BeginRow)} not called");
        }

        private void CloseLastRow()
        {
            if (CurrentColumnNumber > 0)
            {
                _contentBuffer.Write("</row>\n");
                CurrentColumnNumber = 0;
            }
        }

        private void FreezePanes(int fromRow, int fromColumn)
        {
            var topLeftCell = $"{Util.GetColumnName(fromColumn + 1)}{fromRow + 1}";
            if (fromRow > 0 && fromColumn > 0)
            {
                _contentBuffer.Write("<pane xSplit=\"{0}\" ySplit=\"{1}\" topLeftCell=\"{2}\" activePane=\"bottomRight\" state=\"frozen\"/>"
                                        + "<selection pane=\"bottomRight\" activeCell=\"{2}\" sqref=\"{2}\"/>\n",
                    fromColumn, fromRow, topLeftCell);
            }
            else if (fromRow > 0)
            {
                _contentBuffer.Write("<pane ySplit=\"{0}\" topLeftCell=\"{1}\" activePane=\"bottomLeft\" state=\"frozen\"/>"
                                    + "<selection pane=\"bottomLeft\" activeCell=\"{1}\" sqref=\"{1}\"/>\n",
                    fromRow, topLeftCell);
            }
            else if (fromColumn > 0)
            {
                _contentBuffer.Write("<pane xSplit=\"{0}\" topLeftCell=\"{1}\" activePane=\"topRight\" state=\"frozen\"/>"
                                    + "<selection pane=\"topRight\" activeCell=\"{1}\" sqref=\"{1}\"/>\n",
                    fromColumn, topLeftCell);
            }
        }

        private void WriteColumns(IEnumerable<XlsxColumn> columns)
        {
            var columnIndex = 1;
            var colsWritten = false;
            foreach (var column in columns)
            {
                if (column.Hidden || column.Style != null || column.Width.HasValue)
                {
                    if (!colsWritten)
                    {
                        _contentBuffer.Write("<cols>");
                        colsWritten = true;
                    }
                    _contentBuffer.Write("<col min=\"{0}\" max=\"{1}\"", columnIndex, columnIndex + column.Count - 1);
                    if (column.Width.HasValue) _contentBuffer.Write(" width=\"{0}\"", column.Width.Value);
                    if (column.Hidden) _contentBuffer.Write(" hidden=\"1\"");
                    if (column.Width.HasValue) _contentBuffer.Write(" customWidth=\"1\"");
                    if (column.Style != null) _contentBuffer.Write(" style=\"{0}\"", _stylesheet.ResolveStyleId(column.Style));
                    _contentBuffer.Write("/>\n");
                }
                columnIndex += column.Count;
            }
            if (colsWritten)
                _contentBuffer.Write("</cols>\n");
        }

        private void WriteAutoFilter()
        {
            if (_autoFilterRef != null)
                _contentBuffer.Write("<autoFilter ref=\"{0}\"/>\n", _autoFilterRef);
        }

        private void WriteMergedCells()
        {
            if (!_mergedCellRefs.Any())
                return;
            _contentBuffer.Write("<mergeCells count=\"{0}\">\n", _mergedCellRefs.Count);
            foreach (var mergedCell in _mergedCellRefs)
                _contentBuffer.Write("<mergeCell ref=\"{0}\"/>\n", mergedCell);
            _contentBuffer.Write("</mergeCells>\n");
        }

        private void WriteDataValidations()
        {
            if (!_cellRefsByDataValidation.Any())
                return;
            _contentBuffer.Write("<dataValidations count=\"{0}\">\n", _cellRefsByDataValidation.Count);
            foreach (var kvp in _cellRefsByDataValidation)
            {
                _contentBuffer.Write("<dataValidation sqref=\"{0}\" allowBlank=\"{1}\"",
                    string.Join(" ", kvp.Value.Distinct()), Util.BoolToInt(kvp.Key.AllowBlank));
                if (kvp.Key.Error != null)
                    _contentBuffer.Append(" error=\"").AppendEscapedXmlAttribute(kvp.Key.Error, _skipInvalidCharacters).Write('"');
                if (kvp.Key.ErrorStyleValue.HasValue)
                    _contentBuffer.Write(" errorStyle=\"{0}\"", Util.EnumToAttributeValue(kvp.Key.ErrorStyleValue));
                if (kvp.Key.ErrorTitle != null)
                    _contentBuffer.Append(" errorTitle=\"").AppendEscapedXmlAttribute(kvp.Key.ErrorTitle, _skipInvalidCharacters).Write('"');
                if (kvp.Key.OperatorValue.HasValue)
                    _contentBuffer.Write(" operator=\"{0}\"", Util.EnumToAttributeValue(kvp.Key.OperatorValue));
                if (kvp.Key.Prompt != null)
                    _contentBuffer.Append(" prompt=\"").AppendEscapedXmlAttribute(kvp.Key.Prompt, _skipInvalidCharacters).Write('"');
                if (kvp.Key.PromptTitle != null)
                    _contentBuffer.Append(" promptTitle=\"").AppendEscapedXmlAttribute(kvp.Key.PromptTitle, _skipInvalidCharacters).Write('"');
                if (kvp.Key.ShowDropDown)
                    _contentBuffer.Write(" showDropDown=\"1\"");
                if (kvp.Key.ShowErrorMessage)
                    _contentBuffer.Write(" showErrorMessage=\"1\"");
                if (kvp.Key.ShowInputMessage)
                    _contentBuffer.Write(" showInputMessage=\"1\"");
                if (kvp.Key.ValidationTypeValue.HasValue)
                    _contentBuffer.Write(" type=\"{0}\"", Util.EnumToAttributeValue(kvp.Key.ValidationTypeValue));
                _contentBuffer.Write(">");
                if (kvp.Key.Formula1 != null)
                    _contentBuffer.Append("<formula1>").AppendEscapedXmlText(kvp.Key.Formula1, _skipInvalidCharacters).Append("</formula1>");
                if (kvp.Key.Formula2 != null)
                    _contentBuffer.Append("<formula2>").AppendEscapedXmlText(kvp.Key.Formula2, _skipInvalidCharacters).Append("</formula2>");
                _contentBuffer.Write("</dataValidation>\n");
            }
            _contentBuffer.Write("</dataValidations>\n");
        }

        private void WriteSheetProtection()
        {
            if (_sheetProtection == null)
                return;
            const int spinCount = 100000;
            var saltValue = Guid.NewGuid().ToByteArray();
            var hash = Util.ComputePasswordHash(_sheetProtection.Password, saltValue, spinCount);
            _contentBuffer.Write("<sheetProtection algorithmName=\"SHA-512\" hashValue=\"{0}\" saltValue=\"{1}\" spinCount=\"{2}\"", Convert.ToBase64String(hash), Convert.ToBase64String(saltValue), spinCount);
            if (_sheetProtection.Sheet) _contentBuffer.Write(" sheet=\"1\"");
            if (_sheetProtection.Objects) _contentBuffer.Write(" objects=\"1\"");
            if (_sheetProtection.Scenarios) _contentBuffer.Write(" scenarios=\"1\"");
            if (!_sheetProtection.FormatCells) _contentBuffer.Write(" formatCells=\"0\"");
            if (!_sheetProtection.FormatColumns) _contentBuffer.Write(" formatColumns=\"0\"");
            if (!_sheetProtection.FormatRows) _contentBuffer.Write(" formatRows=\"0\"");
            if (!_sheetProtection.InsertColumns) _contentBuffer.Write(" insertColumns=\"0\"");
            if (!_sheetProtection.InsertRows) _contentBuffer.Write(" insertRows=\"0\"");
            if (!_sheetProtection.InsertHyperlinks) _contentBuffer.Write(" insertHyperlinks=\"0\"");
            if (!_sheetProtection.DeleteColumns) _contentBuffer.Write(" deleteColumns=\"0\"");
            if (!_sheetProtection.DeleteRows) _contentBuffer.Write(" deleteRows=\"0\"");
            if (_sheetProtection.SelectLockedCells) _contentBuffer.Write(" selectLockedCells=\"1\"");
            if (!_sheetProtection.Sort) _contentBuffer.Write(" sort=\"0\"");
            if (!_sheetProtection.AutoFilter) _contentBuffer.Write(" autoFilter=\"0\"");
            if (!_sheetProtection.PivotTables) _contentBuffer.Write(" pivotTables=\"0\"");
            if (_sheetProtection.SelectUnlockedCells) _contentBuffer.Write(" selectUnlockedCells=\"1\"");
            _contentBuffer.Write("/>\n");
        }

        private void WriteHeaderFooter()
        {
            if (_headerFooter == null)
                return;
            var differentFirst = _headerFooter.FirstHeader != null || _headerFooter.FirstFooter != null;
            var differentOddEven = _headerFooter.EvenHeader != null || _headerFooter.EvenFooter != null;
            _contentBuffer.Write(
                "<headerFooter alignWithMargins=\"{0}\" differentFirst=\"{1}\" differentOddEven=\"{2}\" scaleWithDoc=\"{3}\">\n",
                Util.BoolToInt(_headerFooter.AlignWithMargins),
                Util.BoolToInt(differentFirst),
                Util.BoolToInt(differentOddEven),
                Util.BoolToInt(_headerFooter.ScaleWithDoc));
            if (_headerFooter.OddHeader != null)
                _contentBuffer.Append("<oddHeader>").AppendEscapedXmlText(_headerFooter.OddHeader, _skipInvalidCharacters).Append("</oddHeader>\n");
            if (_headerFooter.OddFooter != null)
                _contentBuffer.Append("<oddFooter>").AppendEscapedXmlText(_headerFooter.OddFooter, _skipInvalidCharacters).Append("</oddFooter>\n");
            if (_headerFooter.EvenHeader != null)
                _contentBuffer.Append("<evenHeader>").AppendEscapedXmlText(_headerFooter.EvenHeader, _skipInvalidCharacters).Append("</evenHeader>\n");
            if (_headerFooter.EvenFooter != null)
                _contentBuffer.Append("<evenFooter>").AppendEscapedXmlText(_headerFooter.EvenFooter, _skipInvalidCharacters).Append("</evenFooter>\n");
            if (_headerFooter.FirstHeader != null)
                _contentBuffer.Append("<firstHeader>").AppendEscapedXmlText(_headerFooter.FirstHeader, _skipInvalidCharacters).Append("</firstHeader>\n");
            if (_headerFooter.FirstFooter != null)
                _contentBuffer.Append("<firstFooter>").AppendEscapedXmlText(_headerFooter.FirstFooter, _skipInvalidCharacters).Append("</firstFooter>\n");
            _contentBuffer.Write("</headerFooter>\n");
        }

        private void WritePageBreaks()
        {
            if (_pageBreakRowNumbers.Count > 0)
            {
                _contentBuffer.Write($"<rowBreaks count=\"{_pageBreakRowNumbers.Count}\" manualBreakCount=\"{_pageBreakRowNumbers.Count}\">\n");
                foreach (var i in _pageBreakRowNumbers.OrderBy(r => r))
                    _contentBuffer.Write($"<brk id=\"{i}\" max=\"{Limits.MaxColumnCount}\" man=\"1\"/>\n");
                _contentBuffer.Write("</rowBreaks>\n");
            }
            if (_pageBreakColumnNumbers.Count > 0)
            {
                _contentBuffer.Write($"<colBreaks count=\"{_pageBreakColumnNumbers.Count}\" manualBreakCount=\"{_pageBreakColumnNumbers.Count}\">\n");
                foreach (var i in _pageBreakColumnNumbers.OrderBy(c => c))
                    _contentBuffer.Write($"<brk id=\"{i}\" max=\"{Limits.MaxRowCount}\" man=\"1\"/>\n");
                _contentBuffer.Write("</colBreaks>\n");
            }
        }
    }
}