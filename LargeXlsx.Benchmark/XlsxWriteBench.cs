using BenchmarkDotNet.Attributes;
using SharpCompress.Compressors.Deflate;

namespace LargeXlsx.Benchmark;

[MemoryDiagnoser]
public class XlsxWriteBench
{
    [Params(1000)]
    public int Rows { get; set; }

    private string[] _strings =
    [
        "Lorem ipsum dolor sit amet",
        "consectetur adipiscing elit",
        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua",
        "Lorem ipsum dolor sit amet",
        "consectetur adipiscing elit",
        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua",
        "Lorem ipsum dolor sit amet",
        "consectetur adipiscing elit",
        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua",
    ];
    private readonly int[] _integers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    private readonly double[] _numbers = [123, 456, 789, 1011, 1213, 1415, 1617, 1819, 2021, 2223];
    private readonly decimal[] _decimals = [1.1m, 2.2m, 3.3m, 4.4m, 5.5m, 6.6m, 7.7m, 8.8m, 9.9m, 10.10m];
    private readonly bool[] _bools = [true, false, true, false, true, false, true, false, true, false];
    private readonly DateTime[] _dates = [DateTime.Now, DateTime.Today, DateTime.UtcNow, DateTime.Now.AddDays(1), DateTime.Today.AddDays(1), DateTime.UtcNow.AddDays(1), DateTime.Now.AddDays(2), DateTime.Today.AddDays(2), DateTime.UtcNow.AddDays(2), DateTime.Now.AddDays(3)];

    [GlobalSetup]
    public void Setup()
    {
    
    }

    [Benchmark]
    public void WriteInlineStrings()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);

        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            foreach (var s in _strings)
            {
                xw.Write(s);
            }
        }
    }

    [Benchmark]
    public void WriteSharedStrings()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);

        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();

            foreach (var s in _strings)
            {
                xw.Write(s);
            }
        }
    }

    [Benchmark]
    public void WriteNumbers()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);

        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            foreach (var number in _numbers)
            {
                xw.Write(number);
            }
        }
    }

    [Benchmark]
    public void WriteInts()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");
       
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();

            foreach (var integer in _integers)
            {
                xw.Write(integer);
            }
        }
    }

    [Benchmark]
    public void WriteDecimals()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");

        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            foreach (var dec in _decimals)
            {
                xw.Write(dec);
            }
        }
    }

    [Benchmark]
    public void WriteBools()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");
        
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();

            foreach (var b in _bools)
            {
                xw.Write(b);
            }
        }
    }

    [Benchmark]
    public void WriteDateTimes()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");

        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            foreach (var dt in _dates)
            {
                xw.Write(dt);
            }
        }
    }

    [Benchmark]
    public int ColumnNameLookup()
    {
        var sum = 0;
        for (var i = 1; i <= Limits.MaxColumnCount; i++)
        {
            var name = Util.GetColumnName(i);
            sum += name.Length;
        }
        return sum;
    }

    [Benchmark]
    public void WriteInlineStringsBatch()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            xw.WriteRow(_strings);
        }
    }

    [Benchmark]
    public void WriteNumbersBatch()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            xw.WriteRow(_numbers);
        }
    }

    [Benchmark]
    public void WriteIntBatch()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            xw.WriteRow(_integers);
        }
    }

    [Benchmark]
    public void WriteDecimalBatch()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            xw.WriteRow(_decimals);
        }
    }

    [Benchmark]
    public void WriteBoolBatch()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            xw.WriteRow(_bools);
        }
    }

    [Benchmark]
    public void WriteDateTimeBatch()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            xw.WriteRow(_dates);
        }
    }

    [Benchmark]
    public void WriteStingBatch()
    {
        using var stream = new MemoryStream();
        using var xw = new XlsxWriter(stream, CompressionLevel.Level1, useZip64: true, requireCellReferences: false, skipInvalidCharacters: false);
        xw.BeginWorksheet("Sheet1");
        for (var i = 0; i < Rows; i++)
        {
            xw.BeginRow();
            xw.WriteRow(_strings);
        }
    }
}