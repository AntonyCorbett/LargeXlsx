using BenchmarkDotNet.Attributes;
using System.Drawing;
using System.Text;

namespace LargeXlsx.Benchmarks
{
    [MemoryDiagnoser]
    public class Benchmarks
    {
        // Provide a hardware-dependent baseline for raw memory write speed, which is
        // independent of the XlsxLarge library code. All developers can compare the
        // relative cost of the library’s operations against this baseline, regardless
        // of their machine.
        [Benchmark(Baseline = true)]
        public void BaselineMethod()
        {
            using var stream = new MemoryStream();
            using var streamWriter = new StreamWriter(stream, Encoding.UTF8);

            for (var n = 0; n < 100; n++)
            {
                streamWriter.Write("This is a baseline method");
                var buffer = new char[128];
                streamWriter.Write(buffer, 0, buffer.Length);
            }
        }

        [Benchmark]
        public void MethodToBenchmark()
        {
            using var stream = new MemoryStream();
            using var xlsxWriter = new XlsxWriter(stream);
            var leftBorderStyle = XlsxStyle.Default.With(new XlsxBorder(left: new XlsxBorder.Line(Color.DeepPink, XlsxBorder.Style.Thin)));
            var allBorderStyle = XlsxStyle.Default.With(XlsxBorder.Around(new XlsxBorder.Line(Color.CornflowerBlue, XlsxBorder.Style.Dashed)));
            var diagonalBorderStyle = XlsxStyle.Default.With(
                new XlsxBorder(diagonal: new XlsxBorder.Line(Color.Red, XlsxBorder.Style.Dotted), diagonalDown: true, diagonalUp: true));

            xlsxWriter
                .BeginWorksheet("Sheet1")
                .SkipRows(1)
                .BeginRow(height: 50).SkipColumns(1).Write("B1", leftBorderStyle).SkipColumns(1).Write("D1", allBorderStyle).Write("E1", diagonalBorderStyle)
                .BeginRow().SkipColumns(1).Write("B2", leftBorderStyle).SkipColumns(1).Write("D2", allBorderStyle).Write("E2", diagonalBorderStyle)
                .BeginRow().SkipColumns(1).Write(leftBorderStyle).SkipColumns(1).Write(allBorderStyle).Write(diagonalBorderStyle);
        }
    }
}
