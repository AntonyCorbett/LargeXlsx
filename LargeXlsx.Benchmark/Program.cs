using BenchmarkDotNet.Running;

namespace LargeXlsx.Benchmark
{
    internal class Program
    {
        static void Main(string[] args)
        {
            BenchmarkRunner.Run<XlsxWriteBench>();
        }
    }
}
