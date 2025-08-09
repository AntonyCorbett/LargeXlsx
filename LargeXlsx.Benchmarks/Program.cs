using BenchmarkDotNet.Running;

namespace LargeXlsx.Benchmarks
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            var summary = BenchmarkRunner.Run<Benchmarks>();

            var baseline = summary.Reports.FirstOrDefault(r => r.BenchmarkCase.Descriptor.WorkloadMethod.Name == "BaselineMethod");
            var target = summary.Reports.FirstOrDefault(r => r.BenchmarkCase.Descriptor.WorkloadMethod.Name == "MethodToBenchmark");

            if (baseline != null && target != null)
            {
                var baselineMean = baseline.ResultStatistics?.Mean ?? throw new InvalidOperationException("Baseline mean is not available");
                var targetMean = target.ResultStatistics?.Mean ?? throw new InvalidOperationException("Target mean is not available");
                var ratio = targetMean / baselineMean;

                const double thresholdRatio = 40.0;

                Console.WriteLine($"\nRelative performance ratio: {ratio:F2}");

                if (ratio > thresholdRatio)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"WARNING: MethodToBenchmark is {ratio:F2}x slower than baseline (threshold: {thresholdRatio})");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"MethodToBenchmark is all good{ratio:F2}x slower than baseline (threshold: {thresholdRatio})");
                }

                Console.ResetColor();
            }
            else
            {
                Console.WriteLine("Could not find both baseline and target benchmark results.");
            }
        }
    }
}
