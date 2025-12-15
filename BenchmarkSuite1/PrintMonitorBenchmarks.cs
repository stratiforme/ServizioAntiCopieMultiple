#nullable enable
using System.Collections.Generic;
using BenchmarkDotNet.Attributes;
using ServizioAntiCopieMultiple;
using Microsoft.VSDiagnostics;
using System.Runtime.Versioning;

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
        [SupportedOSPlatform("windows")]
        public string ParseJobId() => PrintJobParser.ParseJobId(nameWithJobId);
        [Benchmark]
        [SupportedOSPlatform("windows")]
        public int GetCopies_FromCopies() => PrintJobParser.GetCopiesFromDictionary(propsWithCopies);
        [Benchmark]
        [SupportedOSPlatform("windows")]
        public int GetCopies_FromTotalPages() => PrintJobParser.GetCopiesFromDictionary(propsWithTotalPages);
    }
}