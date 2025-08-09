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

// disable warning regarding early exists from async fns
#pragma warning disable U2U1009
// prefer string interpolation over String.Format
#pragma warning disable U2U1104

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

        public static TextWriter Append(this TextWriter textWriter, double value)
        {
            textWriter.Write(value);
            return textWriter;
        }

        public static TextWriter Append(this TextWriter textWriter, decimal value)
        {
            textWriter.Write(value);
            return textWriter;
        }

        public static TextWriter Append(this TextWriter textWriter, int value)
        {
            textWriter.Write(value);
            return textWriter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textWriter"></param>
        /// <param name="value"></param>
        /// <param name="skipInvalidCharacters"></param>
        /// <returns></returns>
        /// <remarks>See also <see cref="AppendEscapedXmlTextAsync"/>.</remarks>
        /// <exception cref="XmlException"></exception>
        public static TextWriter AppendEscapedXmlText(this TextWriter textWriter, string value, bool skipInvalidCharacters)
        {
            // A plain old for provides a measurable improvement on garbage collection
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (XmlConvert.IsXmlChar(c))
                {
                    if (c == '<') textWriter.Write("&lt;");
                    else if (c == '>') textWriter.Write("&gt;");
                    else if (c == '&') textWriter.Write("&amp;");
                    else textWriter.Write(c);
                }
                else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
                {
                    textWriter.Write(c);
                    textWriter.Write(value[i + 1]);
                    i++;
                }
                else if (!skipInvalidCharacters)
                    throw new XmlException($"Invalid XML character at position {i} in \"{value}\"");
            }
            return textWriter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textWriter"></param>
        /// <param name="value"></param>
        /// <param name="skipInvalidCharacters"></param>
        /// <returns></returns>
        /// <remarks>See also <see cref="AppendEscapedXmlText"/>.</remarks>
        /// <exception cref="XmlException"></exception>
        public static async Task AppendEscapedXmlTextAsync(this TextWriter textWriter, string value, bool skipInvalidCharacters)
        {
            // A plain old for provides a measurable improvement on garbage collection
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (XmlConvert.IsXmlChar(c))
                {
                    if (c == '<') await textWriter.WriteAsync("&lt;").ConfigureAwait(false);
                    else if (c == '>') await textWriter.WriteAsync("&gt;").ConfigureAwait(false);
                    else if (c == '&') await textWriter.WriteAsync("&amp;").ConfigureAwait(false);
                    else await textWriter.WriteAsync(c).ConfigureAwait(false);
                }
                else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
                {
                    await textWriter.WriteAsync(c).ConfigureAwait(false);
                    await textWriter.WriteAsync(value[i + 1]).ConfigureAwait(false);
                    i++;
                }
                else if (!skipInvalidCharacters)
                    throw new XmlException($"Invalid XML character at position {i} in \"{value}\"");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textWriter"></param>
        /// <param name="value"></param>
        /// <param name="skipInvalidCharacters"></param>
        /// <returns></returns>
        /// <remarks>See also <see cref="AppendEscapedXmlTextAsync"/>.</remarks>
        /// <exception cref="XmlException"></exception>
        public static TextWriter AppendEscapedXmlAttribute(this TextWriter textWriter, string value, bool skipInvalidCharacters)
        {
            // A plain old for provides a measurable improvement on garbage collection
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (XmlConvert.IsXmlChar(c))
                {
                    if (c == '<') textWriter.Write("&lt;");
                    else if (c == '>') textWriter.Write("&gt;");
                    else if (c == '&') textWriter.Write("&amp;");
                    else if (c == '\'') textWriter.Write("&apos;");
                    else if (c == '"') textWriter.Write("&quot;");
                    else textWriter.Write(c);
                }
                else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
                {
                    textWriter.Write(c);
                    textWriter.Write(value[i + 1]);
                    i++;
                }
                else if (!skipInvalidCharacters)
                    throw new XmlException($"Invalid XML character at position {i} in \"{value}\"");
            }
            return textWriter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textWriter"></param>
        /// <param name="value"></param>
        /// <param name="skipInvalidCharacters"></param>
        /// <returns></returns>
        /// <remarks>See also <see cref="AppendEscapedXmlTextAsync"/>.</remarks>
        /// <exception cref="XmlException"></exception>
        public static async Task AppendEscapedXmlAttributeAsync(this TextWriter textWriter, string value, bool skipInvalidCharacters)
        {
            // A plain old for provides a measurable improvement on garbage collection
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (XmlConvert.IsXmlChar(c))
                {
                    if (c == '<') await textWriter.WriteAsync("&lt;").ConfigureAwait(false);
                    else if (c == '>') await textWriter.WriteAsync("&gt;").ConfigureAwait(false);
                    else if (c == '&') await textWriter.WriteAsync("&amp;").ConfigureAwait(false);
                    else if (c == '\'') await textWriter.WriteAsync("&apos;").ConfigureAwait(false);
                    else if (c == '"') await textWriter.WriteAsync("&quot;").ConfigureAwait(false);
                    else await textWriter.WriteAsync(c).ConfigureAwait(false);
                }
                else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
                {
                    await textWriter.WriteAsync(c).ConfigureAwait(false);
                    await textWriter.WriteAsync(value[i + 1]).ConfigureAwait(false);
                    i++;
                }
                else if (!skipInvalidCharacters)
                    throw new XmlException($"Invalid XML character at position {i} in \"{value}\"");
            }
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

        public static string BoolToIntString(bool value) => value ? "1" : "0";

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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textWriter"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <remarks>See also <see cref="AddSpacePreserveIfNeededAsync"/>.</remarks>
        public static TextWriter AddSpacePreserveIfNeeded(this TextWriter textWriter, string value)
        {
            if (value.Length > 0 && (XmlConvert.IsWhitespaceChar(value[0]) || XmlConvert.IsWhitespaceChar(value[value.Length - 1])))
                textWriter.Write(" xml:space=\"preserve\"");
            return textWriter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textWriter"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <remarks>See also <see cref="AddSpacePreserveIfNeeded"/>.</remarks>
        public static async Task AddSpacePreserveIfNeededAsync(this TextWriter textWriter, string value)
        {
            if (value.Length > 0 && (XmlConvert.IsWhitespaceChar(value[0]) || XmlConvert.IsWhitespaceChar(value[value.Length - 1])))
                await textWriter.WriteAsync(" xml:space=\"preserve\"").ConfigureAwait(false);
        }

        public static Task WriteAsync(this TextWriter textWriter, double value)
            => textWriter.WriteAsync(value.ToString(System.Globalization.CultureInfo.InvariantCulture));

        public static Task WriteAsync(this TextWriter textWriter, decimal value)
            => textWriter.WriteAsync(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        
        public static Task WriteAsync(this TextWriter textWriter, int value)
            => textWriter.WriteAsync(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }
}