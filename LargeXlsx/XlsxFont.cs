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
using System.Drawing;

namespace LargeXlsx
{
    public readonly struct XlsxFont : IEquatable<XlsxFont>
    {
        public enum Underline
        {
            None,
            Single,
            Double,
            SingleAccounting,
            DoubleAccounting
        }

        public static readonly XlsxFont Default = new XlsxFont("Calibri", 11, Color.Black);

        public string Name { get; }
        public double Size { get; }
        public Color Color { get; }
        public bool Bold { get; }
        public bool Italic { get; }
        public bool Strike { get; }
        public Underline UnderlineType { get; }

        public XlsxFont(string name, double size, Color color, bool bold = false, bool italic = false, bool strike = false, Underline underline = Underline.None)
        {
            Name = name;
            Size = size;
            Color = color;
            Bold = bold;
            Italic = italic;
            Strike = strike;
            UnderlineType = underline;
        }

        public XlsxFont With(Color color) => new XlsxFont(Name, Size, color, Bold, Italic, Strike, UnderlineType);
        public XlsxFont WithName(string name) => new XlsxFont(name, Size, Color, Bold, Italic, Strike, UnderlineType);
        public XlsxFont WithSize(double size) => new XlsxFont(Name, size, Color, Bold, Italic, Strike, UnderlineType);
        public XlsxFont WithBold(bool bold = true) => new XlsxFont(Name, Size, Color, bold, Italic, Strike, UnderlineType);
        public XlsxFont WithItalic(bool italic = true) => new XlsxFont(Name, Size, Color, Bold, italic, Strike, UnderlineType);
        public XlsxFont WithStrike(bool strike = true) => new XlsxFont(Name, Size, Color, Bold, Italic, strike, UnderlineType);
        public XlsxFont WithUnderline(Underline underline = Underline.Single) => new XlsxFont(Name, Size, Color, Bold, Italic, Strike, underline);

        #region Equality members
        public override bool Equals(object obj)
        {
            return obj is XlsxFont other && Equals(other);
        }

        public bool Equals(XlsxFont other)
        {
            return Name == other.Name
                   && Size.Equals(other.Size)
                   && Color.Equals(other.Color)
                   && Bold == other.Bold
                   && Italic == other.Italic
                   && Strike == other.Strike
                   && UnderlineType == other.UnderlineType;
        }

        public override int GetHashCode()
        {
            var hashCode = -1006025245;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Name);
            hashCode = hashCode * -1521134295 + Size.GetHashCode();
            hashCode = hashCode * -1521134295 + Color.GetHashCode();
            hashCode = hashCode * -1521134295 + Bold.GetHashCode();
            hashCode = hashCode * -1521134295 + Italic.GetHashCode();
            hashCode = hashCode * -1521134295 + Strike.GetHashCode();
            hashCode = hashCode * -1521134295 + UnderlineType.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(XlsxFont font1, XlsxFont font2)
        {
            return font1.Equals(font2);
        }

        public static bool operator !=(XlsxFont font1, XlsxFont font2)
        {
            return !font1.Equals(font2);
        }
        #endregion
    }
}
