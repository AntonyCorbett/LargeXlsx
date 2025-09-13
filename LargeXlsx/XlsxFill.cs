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
using System.Drawing;

namespace LargeXlsx
{
    public readonly struct XlsxFill : IEquatable<XlsxFill>
    {
        public enum Pattern
        {
            None,
            Solid,
            Gray125,
            Gray0625,
            DarkDown,
            DarkGray,
            DarkGrid,
            DarkHorizontal,
            DarkTrellis,
            DarkUp,
            DarkVertical,
            LightDown,
            LightGray,
            LightGrid,
            LightHorizontal,
            LightTrellis,
            LightUp,
            LightVertical,
            MediumGray
        }

        public static readonly XlsxFill None = new XlsxFill(Color.White, Pattern.None);
        public static readonly XlsxFill Gray125 = new XlsxFill(Color.White, Pattern.Gray125);

        public Color Color { get; }
        public Pattern PatternType { get; }

        public XlsxFill(Color color, Pattern patternType = Pattern.Solid)
        {
            Color = color;
            PatternType = patternType;
        }

        #region Equality members
        public override bool Equals(object obj)
        {
            return obj is XlsxFill other && Equals(other);
        }

        public bool Equals(XlsxFill other)
        {
            return Color.Equals(other.Color) && PatternType == other.PatternType;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Color.GetHashCode();
                hashCode = (hashCode * 397) ^ (int)PatternType;
                return hashCode;
            }
        }

        public static bool operator ==(XlsxFill fill1, XlsxFill fill2)
        {
            return fill1.Equals(fill2);
        }

        public static bool operator !=(XlsxFill fill1, XlsxFill fill2)
        {
            return !fill1.Equals(fill2);
        }
        #endregion
    }
}
