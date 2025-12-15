#nullable enable
using System.Collections.Generic;
using BenchmarkDotNet.Attributes;
using ServizioAntiCopieMultiple;
using Microsoft.VSDiagnostics;

namespace ServizioAntiCopieMultiple.Benchmarks
{
    [CPUUsageDiagnoser]
    public class PrintMonitorBenchmarks
    {
        private string nameWithJobId = "PrinterName, 123";
        private Dictionary<string, object?> propsWithCopies = null!;
        private Dictionary<string, object?> propsWithTotalPages = null!;
        [GlobalSetup]
        public void Setup()
        {
            propsWithCopies = new Dictionary<string, object?>
            {
                ["Copies"] = 5
            };
            propsWithTotalPages = new Dictionary<string, object?>
            {
                ["TotalPages"] = 10
            };
        }

        [Benchmark]
        public string ParseJobId() => PrintJobParser.ParseJobId(nameWithJobId);
        [Benchmark]
        public int GetCopies_FromCopies() => PrintJobParser.GetCopiesFromDictionary(propsWithCopies);
        [Benchmark]
        public int GetCopies_FromTotalPages() => PrintJobParser.GetCopiesFromDictionary(propsWithTotalPages);
    }
}