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
using System.Threading.Tasks;

// disable warning regarding early exists from async fns
#pragma warning disable U2U1009

// ReSharper disable MethodHasAsyncOverload

// I have chosen the "bool hack" to implement the dual sync/async API
// https://learn.microsoft.com/en-us/archive/msdn-magazine/2015/july/async-programming-brownfield-async-development#the-flag-argument-hack

namespace LargeXlsx
{
    internal class Worksheet : IDisposable
    {
        private readonly Stream _stream;
        private readonly TextWriter _streamWriter;
        private readonly Stylesheet _stylesheet;
        private readonly SharedStringTable _sharedStringTable;
        private readonly bool _requireCellReferences;
        private readonly bool _skipInvalidCharacters;
        private readonly List<string> _mergedCellRefs;
        private readonly Dictionary<XlsxDataValidation, List<string>> _cellRefsByDataValidation;
        private readonly HashSet<int> _pageBreakRowNumbers;
        private readonly HashSet<int> _pageBreakColumnNumbers;
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

        private Worksheet(
            ZipArchive zipArchive,
            CompressionLevel compressionLevel,
            int id,
            string name,
            Stylesheet stylesheet,
            SharedStringTable sharedStringTable,
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
            _streamWriter = new InvariantCultureStreamWriter(_stream);
        }

        private async Task InitializeCoreAsync(
            bool showGridLines,
            bool showHeaders,
            bool rightToLeft,
            int splitRow,
            int splitColumn,
            IEnumerable<XlsxColumn> columns,
            bool sync)
        {
            var s =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<sheetViews>"
                + $"<sheetView showGridLines=\"{Util.BoolToInt(showGridLines)}\" showRowColHeaders=\"{Util.BoolToInt(showHeaders)}\""
                + $" rightToLeft=\"{Util.BoolToInt(rightToLeft)}\" workbookViewId=\"0\">\n";

            if (sync)
            {
                _streamWriter.Write(s);
            }
            else
            {
                await _streamWriter.WriteAsync(s);
            }

            if (splitRow > 0 || splitColumn > 0)
            {
                if (sync)
                {
                    FreezePanesCoreAsync(splitRow, splitColumn, sync: true).GetAwaiter().GetResult();
                }
                else
                {
                    await FreezePanesCoreAsync(splitRow, splitColumn, sync: false);
                }
            }

            const string closingXml = "</sheetView></sheetViews>\n";
            const string closingXml2 = "<sheetData>\n";
            if (sync)
            {
                _streamWriter.Write(closingXml);
                WriteColumnsCoreAsync(columns, sync: true).GetAwaiter().GetResult();
                _streamWriter.Write(closingXml2);
            }
            else
            {
                await _streamWriter.WriteAsync(closingXml);
                await WriteColumnsCoreAsync(columns, sync: false);
                await _streamWriter.WriteAsync(closingXml2);
            }
        }

        public static Worksheet Create(
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
            var worksheet = new Worksheet(
                zipArchive,
                compressionLevel,
                id,
                name,
                stylesheet,
                sharedStringTable,
                requireCellReferences,
                skipInvalidCharacters,
                state);

            worksheet.InitializeCoreAsync(
                showGridLines, showHeaders, rightToLeft, splitRow, splitColumn, columns, sync: true)
                .GetAwaiter().GetResult();

            return worksheet;
        }

        public async Task<Worksheet> CreateAsync(
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
            var worksheet = new Worksheet(
                zipArchive,
                compressionLevel,
                id,
                name,
                stylesheet,
                sharedStringTable,
                requireCellReferences,
                skipInvalidCharacters,
                state);

            await worksheet.InitializeCoreAsync(
                showGridLines, showHeaders, rightToLeft, splitRow, splitColumn, columns, sync: false);

            return worksheet;
        }

        // consider async disposal if moving on from .NET Standard 2.0
        public void Dispose()
        {
            CloseLastRowCoreAsync(sync: true).GetAwaiter().GetResult();
            _streamWriter.Write("</sheetData>\n");
            
            WriteSheetProtectionCoreAsync(sync: true).GetAwaiter().GetResult();
            WriteAutoFilterCoreAsync(sync: true).GetAwaiter().GetResult();
            WriteMergedCellsCoreAsync(sync: true).GetAwaiter().GetResult();
            WriteDataValidationsCoreAsync(sync: true).GetAwaiter().GetResult();
            WriteHeaderFooterCoreAsync(sync: true).GetAwaiter().GetResult();
            WritePageBreaksCoreAsync(sync: true).GetAwaiter().GetResult();
            
            _streamWriter.Write("</worksheet>\n");
            _streamWriter.Dispose();
            _stream.Dispose();
        }

        public void BeginRow(double? height, bool hidden, XlsxStyle style)
        {
            BeginRowCoreAsync(height, hidden, style, sync: true).GetAwaiter().GetResult();
        }

        private async Task BeginRowCoreAsync(double? height, bool hidden, XlsxStyle style, bool sync)
        {
            if (sync)
            {
                CloseLastRowCoreAsync(sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await CloseLastRowCoreAsync(sync: false);
            }

            if (CurrentRowNumber == Limits.MaxRowCount)
                throw new InvalidOperationException($"A worksheet can contain at most {Limits.MaxRowCount} rows ({CurrentRowNumber + 1} attempted)");
            CurrentRowNumber++;
            _stringedCurrentRowNumber = null;
            CurrentColumnNumber = 1;

            const string rowStartString = "<row";
            if (sync)
            {
                _streamWriter.Write(rowStartString);
            }
            else
            {
                await _streamWriter.WriteAsync(rowStartString);
            }

            if (_requireCellReferences || _needsRef)
            {
                const string ref1 = " r=\"";
                const string ref2 = "\"";
                if (sync)
                {
                    _streamWriter.Write(ref1);
                    WriteCurrentRowNumberCoreAsync(sync: true).GetAwaiter().GetResult();
                    _streamWriter.Write(ref2);
                }
                else
                {
                    await _streamWriter.WriteAsync(ref1);
                    await WriteCurrentRowNumberCoreAsync(sync: false);
                    await _streamWriter.WriteAsync(ref2);
                }
                _needsRef = false;
            }

            if (height.HasValue)
            {
                var heightString = $" ht=\"{height}\" customHeight=\"1\"";

                if (sync)
                {
                    _streamWriter.Write(heightString);
                }
                else
                {
                    await _streamWriter.WriteAsync(heightString);
                }
            }

            if (hidden)
            {
                const string hiddenString = " hidden=\"1\"";
                if (sync)
                {
                    _streamWriter.Write(hiddenString);
                }
                else
                {
                    await _streamWriter.WriteAsync(hiddenString);
                }
            }

            if (style != null)
            {
                var styleString = $" s=\"{_stylesheet.ResolveStyleId(style)}\" customFormat=\"1\"";
                if (sync)
                {
                    _streamWriter.Write(styleString);
                }
                else
                {
                    await _streamWriter.WriteAsync(styleString);
                }
            }

            const string closingString = ">\n";
            if (sync)
            {
                _streamWriter.Write(closingString);
            }
            else
            {
                await _streamWriter.WriteAsync(closingString);
            }
        }

        private async Task SkipRowsCoreAsync(int rowCount, bool sync)
        {
            if (sync)
            {
                CloseLastRowCoreAsync(sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await CloseLastRowCoreAsync(sync: false);
            }
            _needsRef = true;
            if (CurrentRowNumber + rowCount > Limits.MaxRowCount)
                throw new InvalidOperationException($"A worksheet can contain at most {Limits.MaxRowCount} rows ({CurrentRowNumber + rowCount} attempted)");
            CurrentRowNumber += rowCount;
        }

        public void SkipRows(int rowCount)
        {
            SkipRowsCoreAsync(rowCount, sync: true).GetAwaiter().GetResult();
        }

        public Task SkipRowsAsync(int rowCount)
        {
            return SkipRowsCoreAsync(rowCount, sync: false);
        }

        // no async version needed.
        public void SkipColumns(int columnCount)
        {
            EnsureRow();
            _needsRef = true;
            CurrentColumnNumber += columnCount;
        }

        public void Write(XlsxStyle style, int repeatCount)
        {
            WriteCoreAsync(style, repeatCount, sync: true).GetAwaiter().GetResult();
        }

        public Task WriteAsync(XlsxStyle style, int repeatCount)
        {
            return WriteCoreAsync(style, repeatCount, sync: false);
        }

        private async Task WriteCoreAsync(XlsxStyle style, int repeatCount, bool sync)
        {
            EnsureRow();
            var styleId = _stylesheet.ResolveStyleId(style);
            for (var i = 0; i < repeatCount; i++)
            {
                // <c r="{0}{1}" s="{2}"/>

                const string openingString = "<c";
                const string closingString = "/>\n";
                if (sync)
                {
                    _streamWriter.Write(openingString);
                    WriteCellRefCoreAsync(sync: true).GetAwaiter().GetResult();
                    WriteStyleCoreAsync(styleId, sync: true).GetAwaiter().GetResult();
                    _streamWriter.Write(closingString);
                }
                else
                {
                    await _streamWriter.WriteAsync(openingString);
                    await WriteCellRefCoreAsync(sync: false);
                    await WriteStyleCoreAsync(styleId, sync: false);
                    await _streamWriter.WriteAsync(closingString);
                }
                CurrentColumnNumber++;
            }
        }

        public void Write(string value, XlsxStyle style)
        {
            WriteCoreAsync(value, style, sync: true).GetAwaiter().GetResult();
        }

        public Task WriteAsync(string value, XlsxStyle style)
        {
            return WriteCoreAsync(value, style, sync: false);
        }

        private async Task WriteCoreAsync(string value, XlsxStyle style, bool sync)
        {
            if (value == null)
            {
                if (sync)
                {
                    Write(style, 1);
                }
                else
                {
                    await WriteAsync(style, 1);
                }

                return;
            }

            EnsureRow();
            // <c r="{0}{1}" s="{2}" t="inlineStr"><is><t xml:space="preserve">{3}</t></is></c>
            const string openingString = "<c";
            if (sync)
            {
                _streamWriter.Write(openingString);
                WriteCellRefCoreAsync(sync: true).GetAwaiter().GetResult();
                WriteStyleCoreAsync(style, sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
                await WriteCellRefCoreAsync(sync: false);
                await WriteStyleCoreAsync(style, sync: false);
            }

            const string openingTypeString = " t=\"inlineStr\"><is><t";
            const string closingString = "</t></is></c>\n";
            if (sync)
            {
                _streamWriter.Append(openingTypeString);
                _streamWriter.AddSpacePreserveIfNeeded(value);
                _streamWriter.Append(">");
                _streamWriter.AppendEscapedXmlText(value, _skipInvalidCharacters);
                _streamWriter.Append(closingString);
            }
            else
            {
                await _streamWriter.AppendAsync(openingTypeString);
                await _streamWriter.AddSpacePreserveIfNeededAsync(value);
                await _streamWriter.AppendAsync(">");
                await _streamWriter.AppendEscapedXmlTextAsync(value, _skipInvalidCharacters);
                await _streamWriter.AppendAsync(closingString);
            }
            
            CurrentColumnNumber++;
        }

        public void Write(double value, XlsxStyle style)
        {
            WriteCoreAsync(value, style, sync: true).GetAwaiter().GetResult();
        }

        public Task WriteAsync(double value, XlsxStyle style)
        {
            return WriteCoreAsync(value, style, sync: false);
        }
        
        private async Task WriteCoreAsync(double value, XlsxStyle style, bool sync)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}"><v>{3}</v></c>

            const string openingString = "<c";
            if (sync)
            {
                _streamWriter.Write(openingString);
                WriteCellRefCoreAsync(sync: true).GetAwaiter().GetResult();
                WriteStyleCoreAsync(style, sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
                await WriteCellRefCoreAsync(sync: false);
                await WriteStyleCoreAsync(style, sync: false);
            }

            const string closingString1 = "><v>";
            const string closingString2 = "</v></c>\n";

            if (sync)
            {
                _streamWriter.Append(closingString1);
                _streamWriter.Append(value);
                _streamWriter.Append(closingString2);
            }
            else
            {
                await _streamWriter.AppendAsync(closingString1);
                await _streamWriter.AppendAsync(value);
                await _streamWriter.AppendAsync(closingString2);
            }

            CurrentColumnNumber++;
        }

        public void Write(decimal value, XlsxStyle style)
        {
            WriteCoreAsync(value, style, sync: true).GetAwaiter().GetResult();
        }

        public Task WriteAsync(decimal value, XlsxStyle style)
        {
            return WriteCoreAsync(value, style, sync: false);
        }
        
        private async Task WriteCoreAsync(decimal value, XlsxStyle style, bool sync)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}"><v>{3}</v></c>
            
            const string openingString = "<c";
            if (sync)
            {
                _streamWriter.Write(openingString);
                WriteCellRefCoreAsync(sync: true).GetAwaiter().GetResult();
                WriteStyleCoreAsync(style, sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
                await WriteCellRefCoreAsync(sync: false);
                await WriteStyleCoreAsync(style, sync: false);
            }

            const string closingString1 = "><v>";
            const string closingString2 = "</v></c>\n";

            if (sync)
            {
                _streamWriter.Append(closingString1);
                _streamWriter.Append(value);
                _streamWriter.Append(closingString2);
            }
            else
            {
                await _streamWriter.AppendAsync(closingString1);
                await _streamWriter.AppendAsync(value);
                await _streamWriter.AppendAsync(closingString2);
            }
            
            CurrentColumnNumber++;
        }

        public void Write(int value, XlsxStyle style)
        {
            WriteCoreAsync(value, style, sync: true).GetAwaiter().GetResult();
        }

        public Task WriteAsync(int value, XlsxStyle style)
        {
            return WriteCoreAsync(value, style, sync: false);
        }

        private async Task WriteCoreAsync(int value, XlsxStyle style, bool sync)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}"><v>{3}</v></c>
            const string openingString = "<c";
            if (sync)
            {
                _streamWriter.Write(openingString);
                WriteCellRefCoreAsync(sync: true).GetAwaiter().GetResult();
                WriteStyleCoreAsync(style, sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
                await WriteCellRefCoreAsync(sync: false);
                await WriteStyleCoreAsync(style, sync: false);
            }

            const string closingString1 = "><v>";
            const string closingString2 = "</v></c>\n";

            if (sync)
            {
                _streamWriter.Append(closingString1);
                _streamWriter.Append(value);
                _streamWriter.Append(closingString2);
            }
            else
            {
                await _streamWriter.AppendAsync(closingString1);
                await _streamWriter.AppendAsync(value);
                await _streamWriter.AppendAsync(closingString2);
            }

            CurrentColumnNumber++;
        }

        public void Write(bool value, XlsxStyle style)
        {
            WriteCoreAsync(value, style, sync: true).GetAwaiter().GetResult();
        }

        public Task WriteAsync(bool value, XlsxStyle style)
        {
            return WriteCoreAsync(value, style, sync: false);
        }

        private async Task WriteCoreAsync(bool value, XlsxStyle style, bool sync)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}" t="b"><v>{3}</v></c>
            
            const string openingString = "<c";
            if (sync)
            {
                _streamWriter.Write(openingString);
                WriteCellRefCoreAsync(sync: true).GetAwaiter().GetResult();
                WriteStyleCoreAsync(style, sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
                await WriteCellRefCoreAsync(sync: false);
                await WriteStyleCoreAsync(style, sync: false);
            }

            const string closingString1 = " t=\"b\"><v>";
            const string closingString2 = "</v></c>\n";

            if (sync)
            {
                _streamWriter.Append(closingString1);
                _streamWriter.Append(Util.BoolToInt(value));
                _streamWriter.Append(closingString2);
            }
            else
            {
                await _streamWriter.AppendAsync(closingString1);
                await _streamWriter.AppendAsync(Util.BoolToInt(value));
                await _streamWriter.AppendAsync(closingString2);
            }

            CurrentColumnNumber++;
        }

        public void WriteFormula(string formula, XlsxStyle style, IConvertible result)
        {
            WriteFormulaCoreAsync(formula, style, result, sync: true).GetAwaiter().GetResult();
        }

        public Task WriteFormulaAsync(string formula, XlsxStyle style, IConvertible result)
        {
            return WriteFormulaCoreAsync(formula, style, result, sync: false);
        }
        
        private async Task WriteFormulaCoreAsync(
            string formula, XlsxStyle style, IConvertible result, bool sync)
        {
            // <c r="{0}{1}" s="{2}" t="str"><f>{3}</f><v>{4}</v></c>
            EnsureRow();
            
            const string openingString = "<c";

            if (sync)
            {
                _streamWriter.Write(openingString);
                WriteCellRefCoreAsync(sync: true).GetAwaiter().GetResult();
                WriteStyleCoreAsync(style, sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
                await WriteCellRefCoreAsync(sync: false);
                await WriteStyleCoreAsync(style, sync: false);
            }

            const string formulaOpeningString = " t=\"str\"><f>";
            const string formulaClosingString = "</f>";

            if (sync)
            {
                _streamWriter.Append(formulaOpeningString);
                _streamWriter.AppendEscapedXmlText(formula, _skipInvalidCharacters);
                _streamWriter.Append(formulaClosingString);
            }
            else
            {
                await _streamWriter.AppendAsync(formulaOpeningString);
                await _streamWriter.AppendEscapedXmlTextAsync(formula, _skipInvalidCharacters);
                await _streamWriter.AppendAsync(formulaClosingString);
            }
        
            if (result != null)
            {
                const string resultOpeningString = "<v>";
                const string resultClosingString = "</v>";

                if (sync)
                {
                    _streamWriter.Append(resultOpeningString);
                    _streamWriter.AppendEscapedXmlText(result.ToString(CultureInfo.InvariantCulture), _skipInvalidCharacters);
                    _streamWriter.Append(resultClosingString);
                }
                else
                {
                    await _streamWriter.AppendAsync(resultOpeningString);
                    await _streamWriter.AppendEscapedXmlTextAsync(result.ToString(CultureInfo.InvariantCulture), _skipInvalidCharacters);
                    await _streamWriter.AppendAsync(resultClosingString);
                }
            }

            const string closingString = "</c>\n";
            if (sync)
            {
                _streamWriter.Write(closingString);
            }
            else
            {
                await _streamWriter.WriteAsync(closingString);
            }
            
            CurrentColumnNumber++;
        }

        public void WriteSharedString(string value, XlsxStyle style)
        {
            WriteSharedStringCoreAsync(value, style, sync: true).GetAwaiter().GetResult();
        }

        public Task WriteSharedStringTask(string value, XlsxStyle style)
        {
            return WriteSharedStringCoreAsync(value, style, sync: false);
        }
        
        private async Task WriteSharedStringCoreAsync(
            string value, XlsxStyle style, bool sync)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}" t="s"><v>{3}</v></c>
            
            const string openingString = "<c";

            if (sync)
            {
                _streamWriter.Write(openingString);
                WriteCellRefCoreAsync(sync: true).GetAwaiter().GetResult();
                WriteStyleCoreAsync(style, sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
                await WriteCellRefCoreAsync(sync: false);
                await WriteStyleCoreAsync(style, sync: false);
            }

            const string sharedStringOpeningString = " t=\"s\"><v>";
            const string sharedStringClosingString = "</v></c>\n";
            if (sync)
            {
                _streamWriter.Append(sharedStringOpeningString);
                _streamWriter.Append(_sharedStringTable.ResolveStringId(value));
                _streamWriter.Append(sharedStringClosingString);
            }
            else
            {
                await _streamWriter.AppendAsync(sharedStringOpeningString);
                await _streamWriter.AppendAsync(_sharedStringTable.ResolveStringId(value));
                await _streamWriter.AppendAsync(sharedStringClosingString);

            }

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

        private async Task WriteCellRefCoreAsync(bool sync)
        {
            if (_requireCellReferences || _needsRef)
            {
                const string openingString = " r=\"";
                const string closingString = "\"";
                var columnName = Util.GetColumnName(CurrentColumnNumber);
                if (sync)
                {
                    _streamWriter.Write(openingString);
                    _streamWriter.Write(columnName);
                    WriteCurrentRowNumberCoreAsync(sync: true).GetAwaiter().GetResult();
                    _streamWriter.Write(closingString);
                }
                else
                {
                    await _streamWriter.WriteAsync(openingString);
                    await _streamWriter.WriteAsync(columnName);
                    await WriteCurrentRowNumberCoreAsync(sync: false);
                    await _streamWriter.WriteAsync(closingString);
                }

                _needsRef = false;
            }
        }

        private async Task WriteCurrentRowNumberCoreAsync(bool sync)
        {
            if (_stringedCurrentRowNumber == null)
                _stringedCurrentRowNumber = CurrentRowNumber.ToString();

            if (sync)
            {
                _streamWriter.Write(_stringedCurrentRowNumber);
            }
            else
            {
                await _streamWriter.WriteAsync(_stringedCurrentRowNumber);
            }
        }

        private async Task WriteStyleCoreAsync(int styleId, bool sync)
        {
            if (styleId != 0)
            {
                const string openingString = " s=\"";
                const string closingString = "\"";

                if (sync)
                {
                    _streamWriter.Append(openingString).Append(styleId).Append(closingString);
                }
                else
                {
                    await _streamWriter.WriteAsync(openingString);
                    await _streamWriter.WriteAsync(styleId.ToString());
                    await _streamWriter.WriteAsync(closingString);
                }
            }
        }

        private async Task WriteStyleCoreAsync(XlsxStyle style, bool sync)
        {
            var id = _stylesheet.ResolveStyleId(style);

            if (sync)
            {
                WriteStyleCoreAsync(id, sync: true).GetAwaiter().GetResult();
            }
            else
            {
                await WriteStyleCoreAsync(id, sync: false);
            }
        }

        private void EnsureRow()
        {
            if (CurrentColumnNumber == 0)
                throw new InvalidOperationException($"{nameof(BeginRow)} not called");
        }

        private async Task CloseLastRowCoreAsync(bool sync)
        {
            if (CurrentColumnNumber > 0)
            {
                const string closingString = "</row>\n";
                if (sync)
                {
                    _streamWriter.Write(closingString);
                }
                else
                {
                    await _streamWriter.WriteAsync(closingString);
                }

                CurrentColumnNumber = 0;
            }
        }

        private async Task FreezePanesCoreAsync(int fromRow, int fromColumn, bool sync)
        {
            var topLeftCell = $"{Util.GetColumnName(fromColumn + 1)}{fromRow + 1}";
            if (fromRow > 0 && fromColumn > 0)
            {
                var s = string.Format(
                    "<pane xSplit=\"{0}\" ySplit=\"{1}\" topLeftCell=\"{2}\" activePane=\"bottomRight\" state=\"frozen\"/>" +
                    "<selection pane=\"bottomRight\" activeCell=\"{2}\" sqref=\"{2}\"/>\n",
                    fromColumn, fromRow, topLeftCell);

                if (sync)
                {
                    _streamWriter.Write(s);
                }
                else
                {
                    await _streamWriter.WriteAsync(s);
                }
            }

            if (fromRow > 0)
            {
                var s = string.Format(
                    "<pane ySplit=\"{0}\" topLeftCell=\"{1}\" activePane=\"bottomLeft\" state=\"frozen\"/>" +
                    "<selection pane=\"bottomLeft\" activeCell=\"{1}\" sqref=\"{1}\"/>\n",
                    fromRow, topLeftCell);

                if (sync)
                {
                    _streamWriter.Write(s);
                }
                else
                {
                    await _streamWriter.WriteAsync(s);
                }
            }
            if (fromColumn > 0)
            {
                var s = string.Format(
                    "<pane xSplit=\"{0}\" topLeftCell=\"{1}\" activePane=\"topRight\" state=\"frozen\"/>" +
                    "<selection pane=\"topRight\" activeCell=\"{1}\" sqref=\"{1}\"/>\n",
                    fromColumn, topLeftCell);

                if (sync)
                {
                    _streamWriter.Write(s);
                }
                else
                {
                    await _streamWriter.WriteAsync(s);
                }
            }
        }

        private async Task WriteColumnsCoreAsync(IEnumerable<XlsxColumn> columns, bool sync)
        {
            var columnIndex = 1;
            var colsWritten = false;
            foreach (var column in columns)
            {
                if (column.Hidden || column.Style != null || column.Width.HasValue)
                {
                    if (!colsWritten)
                    {
                        const string openingCols = "<cols>";
                        if (sync)
                        {
                            _streamWriter.Write(openingCols);
                        }
                        else
                        {
                            await _streamWriter.WriteAsync(openingCols);
                        }
                        colsWritten = true;
                    }

                    var openingCol = $"<col min=\"{columnIndex}\" max=\"{columnIndex + column.Count - 1}\"";
                    if (sync)
                    {
                        _streamWriter.Write(openingCol);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(openingCol);
                    }

                    if (column.Width.HasValue)
                    {
                        var widthString = $" width=\"{column.Width.Value}\"";
                        if (sync)
                        {
                            _streamWriter.Write(widthString);
                        }
                        else
                        {
                            await _streamWriter.WriteAsync(widthString);
                        }
                    }

                    if (column.Hidden)
                    {
                        const string hiddenString = " hidden=\"1\"";
                        if (sync)
                        {
                            _streamWriter.Write(hiddenString);
                        }
                        else
                        {
                            await _streamWriter.WriteAsync(hiddenString);
                        }
                    }

                    if (column.Width.HasValue)
                    {
                        const string customWidthString = " customWidth=\"1\"";
                        if (sync)
                        {
                            _streamWriter.Write(customWidthString);
                        }
                        else
                        {
                            await _streamWriter.WriteAsync(customWidthString);
                        }
                    }

                    if (column.Style != null)
                    {
                        var styleString = $" style=\"{_stylesheet.ResolveStyleId(column.Style)}\"";
                        if (sync)
                        {
                            _streamWriter.Write(styleString);
                        }
                        else
                        {
                            await _streamWriter.WriteAsync(styleString);
                        }
                    }

                    const string columnClosingString = "/>\n";
                    if (sync)
                    {
                        _streamWriter.Write(columnClosingString);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(columnClosingString);
                    }
                }
                columnIndex += column.Count;
            }

            if (colsWritten)
            {
                const string colsClosingString = "</cols>\n";
                if (sync)
                {
                    _streamWriter.Write(colsClosingString);
                }
                else
                {
                    await _streamWriter.WriteAsync(colsClosingString);
                }
            }
        }

        private async Task WriteAutoFilterCoreAsync(bool sync)
        {
            if (_autoFilterRef != null)
            {
                var s = $"<autoFilter ref=\"{_autoFilterRef}\"/>\n";

                if (sync)
                {
                    _streamWriter.Write(s);
                }
                else
                {
                    await _streamWriter.WriteAsync(s);
                }
            }
        }

        private async Task WriteMergedCellsCoreAsync(bool sync)
        {
            if (_mergedCellRefs.Count == 0)
                return;

            var s = $"<mergeCells count=\"{_mergedCellRefs.Count}\">\n";
            
            if (sync)
            {
                _streamWriter.Write(s);
            }
            else
            {
                await _streamWriter.WriteAsync(s);
            }

            foreach (var mergedCell in _mergedCellRefs)
            {
                var mergeCellString = $"<mergeCell ref=\"{mergedCell}\"/>\n";
                if (sync)
                {
                    _streamWriter.Write(mergeCellString);
                }
                else
                {
                    await _streamWriter.WriteAsync(mergeCellString);
                }
            }

            const string closingString = "</mergeCells>\n";
            if (sync)
            {
                _streamWriter.Write(closingString);
            }
            else
            {
                await _streamWriter.WriteAsync(closingString);
            }
        }

        private async Task WriteDataValidationsCoreAsync(bool sync)
        {
            if (_cellRefsByDataValidation.Count == 0)
                return;

            var openingString = $"<dataValidations count=\"{_cellRefsByDataValidation.Count}\">\n";
            if (sync)
            {
                _streamWriter.Write(openingString);
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
            }
            
            foreach (var kvp in _cellRefsByDataValidation)
            {
                var dataValidationString =
                    $"<dataValidation sqref=\"{string.Join(" ", kvp.Value.Distinct())}\" allowBlank=\"{Util.BoolToInt(kvp.Key.AllowBlank)}\"";
                if (sync)
                {
                    _streamWriter.Write(dataValidationString);
                }
                else
                {
                    await _streamWriter.WriteAsync(dataValidationString);
                }
                
                if (kvp.Key.Error != null)
                {
                    const string errorOpeningString = " error=\"";
                                                      
                    if (sync)
                    {
                        _streamWriter.Append(errorOpeningString);
                        _streamWriter.AppendEscapedXmlAttribute(kvp.Key.Error, _skipInvalidCharacters);
                        _streamWriter.Write('"');
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(errorOpeningString);
                        await _streamWriter.AppendEscapedXmlAttributeAsync(kvp.Key.Error, _skipInvalidCharacters);
                        await _streamWriter.WriteAsync('"');
                    }
                }

                if (kvp.Key.ErrorStyleValue.HasValue)
                {
                    var styleString = $" errorStyle=\"{Util.EnumToAttributeValue(kvp.Key.ErrorStyleValue)}\"";
                    if (sync)
                    {
                        _streamWriter.Write(styleString);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(styleString);
                    }
                }

                if (kvp.Key.ErrorTitle != null)
                {
                    const string errorTitleString = " errorTitle=\"";
                    if (sync)
                    {
                        _streamWriter.Append(errorTitleString);
                        _streamWriter.AppendEscapedXmlAttribute(kvp.Key.ErrorTitle, _skipInvalidCharacters);
                        _streamWriter.Write('"');
                    }
                    else
                    {
                        await _streamWriter.AppendAsync(errorTitleString);
                        await _streamWriter.AppendEscapedXmlAttributeAsync(kvp.Key.ErrorTitle, _skipInvalidCharacters);
                        await _streamWriter.WriteAsync('"');
                    }
                }

                if (kvp.Key.OperatorValue.HasValue)
                {
                    var operatorString = $" operator=\"{Util.EnumToAttributeValue(kvp.Key.OperatorValue)}\"";
                    if (sync)
                    {
                        _streamWriter.Write(operatorString);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(operatorString);
                    }
                }

                if (kvp.Key.Prompt != null)
                {
                    const string promptString = " prompt=\"";
                    if (sync)
                    {
                        _streamWriter.Append(promptString);
                        _streamWriter.AppendEscapedXmlAttribute(kvp.Key.Prompt, _skipInvalidCharacters);
                        _streamWriter.Write('"');
                    }
                    else
                    {
                        await _streamWriter.AppendAsync(promptString);
                        await _streamWriter.AppendEscapedXmlAttributeAsync(kvp.Key.Prompt, _skipInvalidCharacters);
                        await _streamWriter.WriteAsync('"');
                    }
                }

                if (kvp.Key.PromptTitle != null)
                {
                    const string promptTitleString = " promptTitle=\"";

                    if (sync)
                    {
                        _streamWriter.Append(promptTitleString);
                        _streamWriter.AppendEscapedXmlAttribute(kvp.Key.PromptTitle, _skipInvalidCharacters);
                        _streamWriter.Write('"');
                    }
                    else
                    {
                        await _streamWriter.AppendAsync(promptTitleString);
                        await _streamWriter.AppendEscapedXmlAttributeAsync(kvp.Key.PromptTitle, _skipInvalidCharacters);
                        await _streamWriter.WriteAsync('"');
                    }
                }

                if (kvp.Key.ShowDropDown)
                {
                    const string dropDownString = " showDropDown=\"1\"";
                    if (sync)
                    {
                        _streamWriter.Write(dropDownString);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(dropDownString);
                    }
                }

                if (kvp.Key.ShowErrorMessage)
                {
                    const string showErrorString = " showErrorMessage=\"1\"";
                    if (sync)
                    {
                        _streamWriter.Write(showErrorString);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(showErrorString);
                    }
                }

                if (kvp.Key.ShowInputMessage)
                {
                    const string showInputString = " showInputMessage=\"1\"";
                    if (sync)
                    {
                        _streamWriter.Write(showInputString);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(showInputString);
                    }
                }

                if (kvp.Key.ValidationTypeValue.HasValue)
                {
                    var validationTypeValueStrong =
                        $" type=\"{Util.EnumToAttributeValue(kvp.Key.ValidationTypeValue)}\"";
                    if (sync)
                    {
                        _streamWriter.Write(validationTypeValueStrong);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(validationTypeValueStrong);
                    }
                }

                if (sync)
                {
                    _streamWriter.Write(">");
                }
                else
                {
                    await _streamWriter.WriteAsync(">");
                }

                if (kvp.Key.Formula1 != null)
                {
                    const string formula1OpeningString = "<formula1>";
                    const string formula1ClosingString = "</formula1>";
                    if (sync)
                    {
                        _streamWriter.Append(formula1OpeningString);
                        _streamWriter.AppendEscapedXmlText(kvp.Key.Formula1, _skipInvalidCharacters);
                        _streamWriter.Append(formula1ClosingString);
                    }
                    else
                    {
                        await _streamWriter.AppendAsync(formula1OpeningString);
                        await _streamWriter.AppendEscapedXmlTextAsync(kvp.Key.Formula1, _skipInvalidCharacters);
                        await _streamWriter.AppendAsync(formula1ClosingString);
                    }
                }

                if (kvp.Key.Formula2 != null)
                {
                    const string formula2OpeningString = "<formula2>";
                    const string formula2ClosingString = "</formula2>";
                    if (sync)
                    {
                        _streamWriter.Append(formula2OpeningString);
                        _streamWriter.AppendEscapedXmlText(kvp.Key.Formula2, _skipInvalidCharacters);
                        _streamWriter.Append(formula2ClosingString);
                    }
                    else
                    {
                        await _streamWriter.AppendAsync(formula2OpeningString);
                        await _streamWriter.AppendEscapedXmlTextAsync(kvp.Key.Formula2, _skipInvalidCharacters);
                        await _streamWriter.AppendAsync(formula2ClosingString);
                    }
                }

                const string dataValidationClosingString = "</dataValidation>\n";
                if (sync)
                {
                    _streamWriter.Write(dataValidationClosingString);
                }
                else
                {
                    await _streamWriter.WriteAsync(dataValidationClosingString);
                }
            }

            const string dataValidationsClosingString = "</dataValidations>\n";
            if (sync)
            {
                _streamWriter.Write(dataValidationsClosingString);
            }
            else
            {
                await _streamWriter.WriteAsync(dataValidationsClosingString);
            }
        }
        
        private async Task WriteSheetProtectionCoreAsync(bool sync)
        {
            if (_sheetProtection == null)
                return;
            
            const int spinCount = 100000;
            var saltValue = Guid.NewGuid().ToByteArray();
            var hash = Util.ComputePasswordHash(_sheetProtection.Password, saltValue, spinCount);

            var openingString =
                $"<sheetProtection algorithmName=\"SHA-512\" hashValue=\"{Convert.ToBase64String(hash)}\" saltValue=\"{Convert.ToBase64String(saltValue)}\" spinCount=\"{spinCount}\"";
            
            if (sync)
            {
                _streamWriter.Write(openingString);
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
            }

            if (_sheetProtection.Sheet)
            {
                const string sheetProtectionString = " sheet=\"1\"";
                if (sync)
                {
                    _streamWriter.Write(sheetProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(sheetProtectionString);
                }
            }

            if (_sheetProtection.Objects)
            {
                const string objectsProtectionString = " objects=\"1\"";
                if (sync)
                {
                    _streamWriter.Write(objectsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(objectsProtectionString);
                }
            }

            if (_sheetProtection.Scenarios)
            {
                const string scenariosProtectionString = " scenarios=\"1\"";
                if (sync)
                {
                    _streamWriter.Write(scenariosProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(scenariosProtectionString);
                }
            }

            if (!_sheetProtection.FormatCells)
            {
                const string formatCellsProtectionString = " formatCells=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(formatCellsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(formatCellsProtectionString);
                }
            }

            if (!_sheetProtection.FormatColumns)
            {
                const string formatColumnsProtectionString = " formatColumns=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(formatColumnsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(formatColumnsProtectionString);
                }
            }

            if (!_sheetProtection.FormatRows)
            {
                const string formatRowsProtectionString = " formatRows=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(formatRowsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(formatRowsProtectionString);
                }
            }

            if (!_sheetProtection.InsertColumns)
            {
                const string insertColumnsProtectionString = " insertColumns=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(insertColumnsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(insertColumnsProtectionString);
                }
            }

            if (!_sheetProtection.InsertRows)
            {
                const string insertRowsProtectionString = " insertRows=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(insertRowsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(insertRowsProtectionString);
                }
            }

            if (!_sheetProtection.InsertHyperlinks)
            {
                const string insertHyperlinksProtectionString = " insertHyperlinks=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(insertHyperlinksProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(insertHyperlinksProtectionString);
                }
            }

            if (!_sheetProtection.DeleteColumns)
            {
                const string deleteColumnsProtectionString = " deleteColumns=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(deleteColumnsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(deleteColumnsProtectionString);
                }
            }

            if (!_sheetProtection.DeleteRows)
            {
                const string deleteRowsProtectionString = " deleteRows=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(deleteRowsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(deleteRowsProtectionString);
                }
            }

            if (_sheetProtection.SelectLockedCells)
            {
                const string selectLockedCellsProtectionString = " selectLockedCells=\"1\"";
                if (sync)
                {
                    _streamWriter.Write(selectLockedCellsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(selectLockedCellsProtectionString);
                }
            }

            if (!_sheetProtection.Sort)
            {
                const string sortProtectionString = " sort=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(sortProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(sortProtectionString);
                }
            }

            if (!_sheetProtection.AutoFilter)
            {
                const string autoFilterProtectionString = " autoFilter=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(autoFilterProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(autoFilterProtectionString);
                }
            }

            if (!_sheetProtection.PivotTables)
            {
                const string pivotTablesProtectionString = " pivotTables=\"0\"";
                if (sync)
                {
                    _streamWriter.Write(pivotTablesProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(pivotTablesProtectionString);
                }
            }

            if (_sheetProtection.SelectUnlockedCells)
            {
                const string selectUnlockedCellsProtectionString = " selectUnlockedCells=\"1\"";
                if (sync)
                {
                    _streamWriter.Write(selectUnlockedCellsProtectionString);
                }
                else
                {
                    await _streamWriter.WriteAsync(selectUnlockedCellsProtectionString);
                }
            }
            
            const string closingString = "/>\n";
            if (sync)
            {
                _streamWriter.Write(closingString);
            }
            else
            {
                await _streamWriter.WriteAsync(closingString);
            }
        }

        private async Task WriteHeaderFooterCoreAsync(bool sync)
        {
            if (_headerFooter == null)
                return;
            
            var differentFirst = _headerFooter.FirstHeader != null || _headerFooter.FirstFooter != null;
            var differentOddEven = _headerFooter.EvenHeader != null || _headerFooter.EvenFooter != null;
            
            var openingString = $"<headerFooter alignWithMargins=\"{Util.BoolToInt(_headerFooter.AlignWithMargins)}\" differentFirst=\"{Util.BoolToInt(differentFirst)}\" differentOddEven=\"{Util.BoolToInt(differentOddEven)}\" scaleWithDoc=\"{Util.BoolToInt(_headerFooter.ScaleWithDoc)}\">\n";

            if (sync)
            {
                _streamWriter.Write(openingString);
            }
            else
            {
                await _streamWriter.WriteAsync(openingString);
            }

            if (_headerFooter.OddHeader != null)
            {
                const string oddHeaderOpeningString = "<oddHeader>";
                const string oddHeaderClosingString = "</oddHeader>\n";
                
                if (sync)
                {
                    _streamWriter.Append(oddHeaderOpeningString);
                    _streamWriter.AppendEscapedXmlText(_headerFooter.OddHeader, _skipInvalidCharacters);
                    _streamWriter.Append(oddHeaderClosingString);
                }
                else
                {
                    await _streamWriter.AppendAsync(oddHeaderOpeningString);
                    await _streamWriter.AppendEscapedXmlTextAsync(_headerFooter.OddHeader, _skipInvalidCharacters);
                    await _streamWriter.AppendAsync(oddHeaderClosingString);
                }
            }

            if (_headerFooter.OddFooter != null)
            {
                const string oddFooterOpeningString = "<oddFooter>";
                const string oddFooterClosingString = "</oddFooter>\n";
                
                if(sync)
                {
                    _streamWriter.Append(oddFooterOpeningString);
                    _streamWriter.AppendEscapedXmlText(_headerFooter.OddFooter, _skipInvalidCharacters);
                    _streamWriter.Append(oddFooterClosingString);
                }
                else
                {
                    await _streamWriter.AppendAsync(oddFooterOpeningString);
                    await _streamWriter.AppendEscapedXmlTextAsync(_headerFooter.OddFooter, _skipInvalidCharacters);
                    await _streamWriter.AppendAsync(oddFooterClosingString);
                }
            }

            if (_headerFooter.EvenHeader != null)
            {
                const string evenHeaderOpeningString = "<evenHeader>";
                const string evenHeaderClosingString = "</evenHeader>\n";
                if (sync)
                {
                    _streamWriter.Append(evenHeaderOpeningString);
                    _streamWriter.AppendEscapedXmlText(_headerFooter.EvenHeader, _skipInvalidCharacters);
                    _streamWriter.Append(evenHeaderClosingString);
                }
                else
                {
                    await _streamWriter.AppendAsync(evenHeaderOpeningString);
                    await _streamWriter.AppendEscapedXmlTextAsync(_headerFooter.EvenHeader, _skipInvalidCharacters);
                    await _streamWriter.AppendAsync(evenHeaderClosingString);
                }
            }

            if (_headerFooter.EvenFooter != null)
            {
                const string evenFooterOpeningString = "<evenFooter>";
                const string evenFooterClosingString = "</evenFooter>\n";
                
                if(sync)
                {
                    _streamWriter.Append(evenFooterOpeningString);
                    _streamWriter.AppendEscapedXmlText(_headerFooter.EvenFooter, _skipInvalidCharacters);
                    _streamWriter.Append(evenFooterClosingString);
                }
                else
                {
                    await _streamWriter.AppendAsync(evenFooterOpeningString);
                    await _streamWriter.AppendEscapedXmlTextAsync(_headerFooter.EvenFooter, _skipInvalidCharacters);
                    await _streamWriter.AppendAsync(evenFooterClosingString);
                }
            }

            if (_headerFooter.FirstHeader != null)
            {
                const string firstHeaderOpeningString = "<firstHeader>";
                const string firstHeaderClosingString = "</firstHeader>\n";
                if (sync)
                {
                    _streamWriter.Append(firstHeaderOpeningString);
                    _streamWriter.AppendEscapedXmlText(_headerFooter.FirstHeader, _skipInvalidCharacters);
                    _streamWriter.Append(firstHeaderClosingString);
                }
                else
                {
                    await _streamWriter.AppendAsync(firstHeaderOpeningString);
                    await _streamWriter.AppendEscapedXmlTextAsync(_headerFooter.FirstHeader, _skipInvalidCharacters);
                    await _streamWriter.AppendAsync(firstHeaderClosingString);
                }
            }

            if (_headerFooter.FirstFooter != null)
            {
                const string firstFooterOpeningString = "<firstFooter>";
                const string firstFooterClosingString = "</firstFooter>\n";
                
                if(sync)
                {
                    _streamWriter.Append(firstFooterOpeningString);
                    _streamWriter.AppendEscapedXmlText(_headerFooter.FirstFooter, _skipInvalidCharacters);
                    _streamWriter.Append(firstFooterClosingString);
                }
                else
                {
                    await _streamWriter.AppendAsync(firstFooterOpeningString);
                    await _streamWriter.AppendEscapedXmlTextAsync(_headerFooter.FirstFooter, _skipInvalidCharacters);
                    await _streamWriter.AppendAsync(firstFooterClosingString);
                }
            }

            const string closingString = "</headerFooter>\n";
            if (sync)
            {
                _streamWriter.Write(closingString);
            }
            else
            {
                await _streamWriter.WriteAsync(closingString);
            }
        }

        private async Task WritePageBreaksCoreAsync(bool sync)
        {
            if (_pageBreakRowNumbers.Count > 0)
            {
                var rowBreaksOpeningString = $"<rowBreaks count=\"{_pageBreakRowNumbers.Count}\" manualBreakCount=\"{_pageBreakRowNumbers.Count}\">\n";
                if (sync)
                {
                    _streamWriter.Write(rowBreaksOpeningString);
                }
                else
                {
                    await _streamWriter.WriteAsync(rowBreaksOpeningString);
                }

                foreach (var i in _pageBreakRowNumbers.OrderBy(r => r))
                {
                    var s = $"<brk id=\"{i}\" max=\"{Limits.MaxColumnCount}\" man=\"1\"/>\n";
                    if (sync)
                    {
                        _streamWriter.Write(s);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(s);
                    }
                }

                const string closingBreakString = "</rowBreaks>\n";
                if (sync)
                {
                    _streamWriter.Write(closingBreakString);
                }
                else
                {
                    await _streamWriter.WriteAsync(closingBreakString);
                }
            }
            if (_pageBreakColumnNumbers.Count > 0)
            {
                var colBreaksOpeningString = $"<colBreaks count=\"{_pageBreakColumnNumbers.Count}\" manualBreakCount=\"{_pageBreakColumnNumbers.Count}\">\n";
                if (sync)
                {
                    _streamWriter.Write(colBreaksOpeningString);
                }
                else
                {
                    await _streamWriter.WriteAsync(colBreaksOpeningString);
                }

                foreach (var i in _pageBreakColumnNumbers.OrderBy(c => c))
                {
                    var s = $"<brk id=\"{i}\" max=\"{Limits.MaxRowCount}\" man=\"1\"/>\n";

                    if (sync)
                    {
                        _streamWriter.Write(s);
                    }
                    else
                    {
                        await _streamWriter.WriteAsync(s);
                    }
                }
                
                const string colBreaksClosingString = "</colBreaks>\n";
                if (sync)
                {
                    _streamWriter.Write(colBreaksClosingString);
                }
                else
                {
                    await _streamWriter.WriteAsync(colBreaksClosingString);
                }
            }
        }
    }
}