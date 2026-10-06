using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvAsyncSaveContractTests
{
    [Theory]
    [InlineData(CsvCompressionType.None)]
    [InlineData(CsvCompressionType.GZip)]
    [InlineData(CsvCompressionType.Deflate)]
    public async Task SaveAsync_Uses_Async_Output_And_Leaves_Caller_Stream_Open(CsvCompressionType compression)
    {
        var document = CsvDocument.Parse("Name,Notes\nAlpha,\"line1\nline2\"\nBeta,\"a,b\"\n");
        using var output = new AsyncOnlyOutput();
        await document.SaveAsync(output, new CsvSaveOptions { CompressionType = compression });
        Assert.True(output.CanWrite);
        Assert.True(output.AsyncWrites > 0);
        using var input = new MemoryStream(output.Bytes);
        Assert.Equal(document.ToString(), CsvDocument.Load(input, new CsvLoadOptions { CompressionType = compression }).ToString());
    }

    [Fact]
    public async Task SaveAsync_Cancels_Pending_Destination_IO()
    {
        using var cancellation = new CancellationTokenSource();
        using var output = new AsyncOnlyOutput(block: true);
        var document = CsvDocument.Parse("Name\nAlpha\n");
        Task saving = document.SaveAsync(output, cancellationToken: cancellation.Token);
        Assert.Same(output.WriteStarted.Task, await Task.WhenAny(output.WriteStarted.Task, Task.Delay(5000)));
        cancellation.Cancel();
        Assert.Same(saving, await Task.WhenAny(saving, Task.Delay(5000)));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => saving);
    }

    [Fact]
    public async Task SaveAsync_Observes_Cancellation_During_Output()
    {
            using var cancellation = new CancellationTokenSource();
            var document = CsvDocument.Parse("Name\n" + new string('x', 300000) + "\n");
            using var output = new AsyncOnlyOutput(cancellation);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.SaveAsync(output, cancellationToken: cancellation.Token));
            Assert.True(output.CanWrite);
            Assert.True(output.Length > 0);
    }

    private sealed class AsyncOnlyOutput : Stream
    {
        private readonly MemoryStream _inner = new();
        private readonly CancellationTokenSource? _cancel;
        private readonly bool _block;
        internal TaskCompletionSource<bool> WriteStarted { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal AsyncOnlyOutput(CancellationTokenSource? cancel = null, bool block = false) { _cancel = cancel; _block = block; }
        internal int AsyncWrites { get; private set; }
        internal byte[] Bytes => _inner.ToArray();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override Task FlushAsync(CancellationToken cancellationToken) { cancellationToken.ThrowIfCancellationRequested(); return Task.CompletedTask; }
        public override void Write(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Synchronous I/O was used.");
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (_block) return WaitForCancellationAsync(cancellationToken);
            AsyncWrites++;
            _inner.Write(buffer, offset, count);
            _cancel?.Cancel();
            return Task.CompletedTask;
        }
#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) => throw new InvalidOperationException("Synchronous I/O was used.");
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (_block) { await WaitForCancellationAsync(cancellationToken); return; }
            AsyncWrites++;
            _inner.Write(buffer.Span);
            _cancel?.Cancel();
        }
#endif
        private async Task WaitForCancellationAsync(CancellationToken token)
        {
            var pending = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            WriteStarted.TrySetResult(true);
            using (token.Register(() => pending.TrySetCanceled())) await pending.Task;
        }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
