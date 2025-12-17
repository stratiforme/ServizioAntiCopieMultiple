using System;
using System.Threading.Channels;
using System.Threading.Tasks;

namespace ServizioAntiCopieMultiple
{
    internal sealed class PrintJobProcessor : IAsyncDisposable
    {
        private readonly Channel<Func<Task>> _channel = Channel.CreateUnbounded<Func<Task>>(new UnboundedChannelOptions { SingleReader = true, SingleWriter = false });
        private readonly Task _consumer;
        private bool _disposed;

        public PrintJobProcessor()
        {
            _consumer = Task.Run(ConsumeAsync);
        }

        public ValueTask EnqueueAsync(Func<Task> work)
        {
            if (_disposed) return new ValueTask(Task.CompletedTask);
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
                    // swallow exceptions to avoid halting the consumer loop
                }
            }
        }

        public async Task StopAsync()
        {
            if (_disposed) return;
            _channel.Writer.Complete();
            try { await _consumer.ConfigureAwait(false); } catch { }
            _disposed = true;
        }

        public async ValueTask DisposeAsync()
        {
            await StopAsync().ConfigureAwait(false);
        }
    }
}
