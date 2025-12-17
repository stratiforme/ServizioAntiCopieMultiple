using System;
using System.Linq;
using System.Runtime.Versioning;
using System.Threading.Channels;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using Microsoft.VSDiagnostics;

namespace ServizioAntiCopieMultiple.Benchmarks
{
    [CPUUsageDiagnoser]
    public class PrintJobProcessingBenchmarks
    {
        [Params(100, 1000)]
        public int Count;
        private static ValueTask SimulatedWorkAsync()
        {
            // small CPU work to simulate processing
            int x = 0;
            for (int i = 0; i < 10; i++)
                x += i;
            return ValueTask.CompletedTask;
        }

        [Benchmark]
        [SupportedOSPlatform("windows")]
        public async Task TaskRun()
        {
            var tasks = new Task[Count];
            for (int i = 0; i < Count; i++)
            {
                tasks[i] = Task.Run(() => SimulatedWorkAsync().AsTask());
            }

            await Task.WhenAll(tasks).ConfigureAwait(false);
        }

        // Local channel-based processor used for benchmark
        private sealed class LocalProcessor
        {
            private readonly Channel<Func<Task>> _channel = Channel.CreateUnbounded<Func<Task>>();
            private readonly Task _consumer;
            public LocalProcessor()
            {
                _consumer = Task.Run(ConsumeAsync);
            }

            public ValueTask EnqueueAsync(Func<Task> work)
            {
                if (!_channel.Writer.TryWrite(work))
                    return new ValueTask(Task.CompletedTask);
                return ValueTask.CompletedTask;
            }

            private async Task ConsumeAsync()
            {
                await foreach (var work in _channel.Reader.ReadAllAsync().ConfigureAwait(false))
                {
                    try
                    {
                        await work().ConfigureAwait(false);
                    }
                    catch
                    {
                        // swallow for benchmark simplicity
                    }
                }
            }

            public async Task StopAsync()
            {
                _channel.Writer.Complete();
                await _consumer.ConfigureAwait(false);
            }
        }

        [Benchmark]
        [SupportedOSPlatform("windows")]
        public async Task ChannelProcessing()
        {
            var processor = new LocalProcessor();
            for (int i = 0; i < Count; i++)
            {
                await processor.EnqueueAsync(() => SimulatedWorkAsync().AsTask()).ConfigureAwait(false);
            }

            await processor.StopAsync().ConfigureAwait(false);
        }
    }
}