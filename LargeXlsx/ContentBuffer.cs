using System;
using System.IO;
using System.Text;

namespace LargeXlsx
{
    internal sealed class ContentBuffer : IDisposable
    {
        const int BufferSize = 128 * 1024;
        const int BufferThreshold = BufferSize - 1024;
        private const int LengthOfDouble = 10; // not critical, just a rough estimate for the length of a double value
        private const int LengthOfDecimal = 5; // not critical, just a rough estimate for the length of a decimal value
        private const int LengthOfInt = 5; // not critical, just a rough estimate for the length of an int value
        private readonly StringBuilder _buffer = new StringBuilder(BufferSize);
        private readonly TextWriter _streamWriter;

        public ContentBuffer(TextWriter streamWriter)
        {
            _streamWriter = streamWriter ?? throw new ArgumentNullException(nameof(streamWriter));
        }

        public ContentBuffer Append(string content)
        {
            if (_buffer.Length + content.Length > BufferThreshold)
            {
                Flush();
            }

            _buffer.Append(content);
            return this;
        }

        public ContentBuffer AppendLine(string content)
        {
            if (_buffer.Length + content.Length + Environment.NewLine.Length > BufferThreshold)
            {
                Flush();
            }

            _buffer.AppendLine(content);
            return this;
        }

        public ContentBuffer AppendLine()
        {
            if (_buffer.Length + Environment.NewLine.Length > BufferThreshold)
            {
                Flush();
            }

            _buffer.AppendLine();
            return this;    
        }

        public ContentBuffer Append(char c)
        {
            if (_buffer.Length + 1 > BufferThreshold)
            {
                Flush();
            }
        
            _buffer.Append(c);
            return this;    
        }

        public ContentBuffer Append(double value)
        {
            if (_buffer.Length + LengthOfDouble > BufferThreshold)
            {
                Flush();
            }

            _buffer.Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            return this;
        }

        public ContentBuffer Append(decimal value)
        {
            if (_buffer.Length + LengthOfDecimal > BufferThreshold)
            {
                Flush();
            }

            _buffer.Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            return this;
        }

        public ContentBuffer Append(int value)
        {
            if (_buffer.Length + LengthOfInt > BufferThreshold)
            {
                Flush();
            }

            _buffer.Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            return this;
        }

        public ContentBuffer Append(char[] chars, int startIndex, int count)
        {
            if (_buffer.Length + count > BufferThreshold)
            {
                Flush();
            }
         
            _buffer.Append(chars, startIndex, count);
            return this;    
        }

        public ContentBuffer Append(string format, params object[] args)
            => Append(string.Format(System.Globalization.CultureInfo.InvariantCulture, format, args));

        public ContentBuffer AppendLine(string format, params object[] args)
            => AppendLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, format, args));

        public void Write(string content) => Append(content);

        public void Write(char content) => Append(content);

        public void Write(string format, params object[] args)
            => Write(string.Format(System.Globalization.CultureInfo.InvariantCulture, format, args));

        public void Flush()
        {
            if (_buffer.Length > 0)
            {
                _streamWriter.Write(_buffer.ToString());
                _buffer.Clear();
            }
        }

        public void Dispose()
        {
            Flush();
            _streamWriter?.Dispose();
        }
    }
}
