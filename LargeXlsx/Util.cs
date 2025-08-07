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
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
// ReSharper disable MethodHasAsyncOverload

namespace LargeXlsx
{
    internal static class Util
    {
        private static readonly DateTime ExcelEpoch = new DateTime(1900, 1, 1);
        private static readonly DateTime Date19000301 = new DateTime(1900, 3, 1);
        private static readonly string[] CachedColumnNames = new string[Limits.MaxColumnCount];

        public static TextWriter Append(this TextWriter textWriter, string value)
        {
            textWriter.Write(value);
            return textWriter;
        }

        public static async Task<TextWriter> AppendAsync(this TextWriter textWriter, string value)
        {
            await textWriter.WriteAsync(value);
            return textWriter;
        }
        
        public static TextWriter Append(this TextWriter textWriter, double value)
        {
            textWriter.Write(value);
            return textWriter;
        }

        public static async Task<TextWriter> AppendAsync(this TextWriter textWriter, double value)
        {
            await textWriter.WriteAsync(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            return textWriter;
        }

        public static TextWriter Append(this TextWriter textWriter, decimal value)
        {
            textWriter.Write(value);
            return textWriter;
        }

        public static async Task<TextWriter> AppendAsync(this TextWriter textWriter, decimal value)
        {
            await textWriter.WriteAsync(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            return textWriter;
        }

        public static TextWriter Append(this TextWriter textWriter, int value)
        {
            textWriter.Write(value);
            return textWriter;
        }

        public static async Task<TextWriter> AppendAsync(this TextWriter textWriter, int value)
        {
            await textWriter.WriteAsync(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            return textWriter;
        }

        public static TextWriter AppendEscapedXmlText(this TextWriter textWriter, string value, bool skipInvalidCharacters)
        {
            return AppendEscapedXmlTextCoreAsync(
                textWriter, value, skipInvalidCharacters, sync: true).GetAwaiter().GetResult();
        }

        public static Task<TextWriter> AppendEscapedXmlTextAsync(this TextWriter textWriter, string value, bool skipInvalidCharacters)
        {
            return AppendEscapedXmlTextCoreAsync(
                textWriter, value, skipInvalidCharacters, sync: false);
        }

#pragma warning disable U2U1009
        private static async Task<TextWriter> AppendEscapedXmlTextCoreAsync(
#pragma warning restore U2U1009
            this TextWriter textWriter, string value, bool skipInvalidCharacters, bool sync)
        {
            // A plain old for provides a measurable improvement on garbage collection
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (XmlConvert.IsXmlChar(c))
                {
                    if (c == '<')
                    {
                        const string lessThanString = "&lt;";
                        if (sync)
                        {
                            textWriter.Write(lessThanString);
                        }
                        else
                        {
                            await textWriter.WriteAsync(lessThanString);
                        }
                    }
                    else if (c == '>')
                    {
                        const string greaterThanString = "&gt;";
                        
                        if (sync)
                        {
                            textWriter.Write(greaterThanString);
                        }
                        else
                        {
                            await textWriter.WriteAsync(greaterThanString);
                        }
                    }
                    else if (c == '&')
                    {
                        const string ampersandString = "&amp;";
                        if (sync)
                        {
                            textWriter.Write(ampersandString);
                        }
                        else
                        {
                            await textWriter.WriteAsync(ampersandString);
                        }
                    }
                    else
                    {
                        if (sync)
                        {
                            textWriter.Write(c);
                        }
                        else
                        {
                            await textWriter.WriteAsync(c);
                        }
                    }
                }
                else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
                {
                    if (sync)
                    {
                        textWriter.Write(c);
                        textWriter.Write(value[i + 1]);
                    }
                    else
                    {
                        await textWriter.WriteAsync(c);
                        await textWriter.WriteAsync(value[i + 1]);
                    }
                    i++;
                }
                else if (!skipInvalidCharacters)
                    throw new XmlException($"Invalid XML character at position {i} in \"{value}\"");
            }
            return textWriter;
        }
        
        public static TextWriter AppendEscapedXmlAttribute(
            this TextWriter textWriter, string value, bool skipInvalidCharacters)
        {
            return AppendEscapedXmlAttributeCoreAsync(
                textWriter, value, skipInvalidCharacters, sync: true).GetAwaiter().GetResult();
        }

        public static Task<TextWriter> AppendEscapedXmlAttributeAsync(
            this TextWriter textWriter, string value, bool skipInvalidCharacters)
        {
            return AppendEscapedXmlAttributeCoreAsync(
                textWriter, value, skipInvalidCharacters, sync: false);
        }

#pragma warning disable U2U1009
        private static async Task<TextWriter> AppendEscapedXmlAttributeCoreAsync(
#pragma warning restore U2U1009
            this TextWriter textWriter, string value, bool skipInvalidCharacters, bool sync)
        {
            // A plain old for provides a measurable improvement on garbage collection
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (XmlConvert.IsXmlChar(c))
                {
                    if (c == '<')
                    {
                        const string lessThanString = "&lt;";
                        if (sync)
                        {
                            textWriter.Write(lessThanString);
                        }
                        else
                        {
                            await textWriter.WriteAsync(lessThanString);
                        }
                    }
                    else if (c == '>')
                    {
                        const string greaterThanString = "&gt;";
                        if (sync)
                        {
                            textWriter.Write(greaterThanString);
                        }
                        else
                        {
                            await textWriter.WriteAsync(greaterThanString);
                        }
                    }
                    else if (c == '&')
                    {
                        const string ampersandString = "&amp;";
                        if (sync)
                        {
                            textWriter.Write(ampersandString);
                        }
                        else
                        {
                            await textWriter.WriteAsync(ampersandString);
                        }
                    }
                    else if (c == '\'')
                    {
                        const string apostropheString = "&apos;";
                        if (sync)
                        {
                            textWriter.Write(apostropheString);
                        }
                        else
                        {
                            await textWriter.WriteAsync(apostropheString);
                        }
                    }
                    else if (c == '"')
                    {
                        const string quoteString = "&quot;";
                        if (sync)
                        {
                            textWriter.Write(quoteString);
                        }
                        else
                        {
                            await textWriter.WriteAsync(quoteString);
                        }
                    }
                    else
                    {
                        if (sync)
                        {
                            textWriter.Write(c);
                        }
                        else
                        {
                            await textWriter.WriteAsync(c);
                        }
                    }
                }
                else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
                {
                    if (sync)
                    {
                        textWriter.Write(c);
                        textWriter.Write(value[i + 1]);
                    }
                    else
                    {
                        await textWriter.WriteAsync(c);
                        await textWriter.WriteAsync(value[i + 1]);
                    }

                    i++;
                }
                else if (!skipInvalidCharacters)
                    throw new XmlException($"Invalid XML character at position {i} in \"{value}\"");
            }
            return textWriter;
        }
        
        public static string GetColumnName(int columnIndex)
        {
            if (columnIndex < 1 || columnIndex > Limits.MaxColumnCount)
                throw new InvalidOperationException($"A worksheet can contain at most {Limits.MaxColumnCount} columns ({columnIndex} attempted)");
            var columnName = CachedColumnNames[columnIndex - 1];
            if (columnName == null)
            {
                columnName = GetColumnNameInternal(columnIndex);
                CachedColumnNames[columnIndex - 1] = columnName;
            }
            return columnName;
        }

        private static string GetColumnNameInternal(int columnIndex)
        {
            var columnName = new StringBuilder(3); // This has been measured to be faster than string concatenation
            while (true)
            {
                if (columnIndex > 26)
                {
                    columnIndex = Math.DivRem(columnIndex - 1, 26, out var rem);
                    columnName.Insert(0, (char)('A' + rem));
                }
                else
                {
                    columnName.Insert(0, (char)('A' + columnIndex - 1));
                    return columnName.ToString();
                }
            }
        }

        public static double DateToDouble(DateTime date)
        {
            var days = date.Subtract(ExcelEpoch).TotalDays + 1;
            // Excel wrongly assumes that 1900 is a leap year:
            // https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year
            if (date >= Date19000301) days++;
            return days;
        }

        public static int BoolToInt(bool value) => value ? 1 : 0;

        public static string EnumToAttributeValue<T>(T enumValue)
        {
            var s = enumValue.ToString();
            return char.ToLowerInvariant(s[0]) + s.Substring(1);
        }

        // Hashing procedure courtesy of https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/1357ea58-646e-4483-92ef-95d718079d6f
        public static byte[] ComputePasswordHash(string password, byte[] saltValue, int spinCount)
        {
            var hasher = new SHA512Managed();
            var hash = hasher.ComputeHash(saltValue.Concat(Encoding.Unicode.GetBytes(password)).ToArray());
            for (var i = 0; i < spinCount; i++)
            {
                var iterator = BitConverter.GetBytes(i);
                if (!BitConverter.IsLittleEndian)
                    Array.Reverse(iterator);
                hash = hasher.ComputeHash(hash.Concat(iterator).ToArray());
            }
            return hash;
        }

        public static TextWriter AddSpacePreserveIfNeeded(
            this TextWriter textWriter, string value)
        {
            AddSpacePreserveIfNeededCoreAsync(textWriter, value, sync: true).GetAwaiter().GetResult();
            return textWriter;
        }

        public static Task<TextWriter> AddSpacePreserveIfNeededAsync(
            this TextWriter textWriter, string value)
        {
            return AddSpacePreserveIfNeededCoreAsync(textWriter, value, sync: false);
        }

        private static async Task<TextWriter> AddSpacePreserveIfNeededCoreAsync(
            this TextWriter textWriter, string value, bool sync)
        {
            if (value.Length > 0 && (XmlConvert.IsWhitespaceChar(value[0]) ||
                                     XmlConvert.IsWhitespaceChar(value[value.Length - 1])))
            {
                const string spacePreserve = " xml:space=\"preserve\"";
                if (sync)
                {
                    textWriter.Write(spacePreserve);
                }
                else
                {
                    await textWriter.WriteAsync(spacePreserve);
                }
            }
            return textWriter;
        }
    }
}