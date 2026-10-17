using BenchmarkDotNet.Attributes;
using OfficeIMO.Archives;
using System.IO.Compression;
using System.Security.Cryptography;

namespace OfficeIMO.Archives.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 1, iterationCount: 3)]
/// <summary>Compares complete verified reads and private-spool copying with identical SHA-256 output.</summary>
public class VerifiedZipCopy {
    private byte[] _zip = null!;
    private string _expected = null!;
    [GlobalSetup]
    public void Setup() {
        using var output = new MemoryStream();
        using var expected = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, true)) {
            using var entry = archive.CreateEntry("large.bin", CompressionLevel.Optimal).Open();
            byte[] buffer = Enumerable.Range(0, 65536).Select(i => (byte)(i % 251)).ToArray();
            for (int i = 0; i < 256; i++) { entry.Write(buffer); expected.AppendData(buffer); }
        }
        _zip = output.ToArray(); _expected = Convert.ToHexString(expected.GetHashAndReset());
        if (WholeMember() != _expected || VerifiedSpool() != _expected) throw new InvalidDataException("Benchmark work differs");
    }
    [Benchmark(Baseline = true)]
    public string WholeMember() {
        using var input = new MemoryStream(_zip);
        using var inspection = ArchiveInspector.Open(input);
        using var member = inspection.OpenRead(0);
        return Convert.ToHexString(SHA256.HashData(member));
    }
    [Benchmark]
    public string VerifiedSpool() {
        using var input = new MemoryStream(_zip);
        using var inspection = ArchiveInspector.Open(input);
        using var sink = new HashSink(); inspection.CopyTo(0, sink);
        return sink.Digest();
    }
    private sealed class HashSink : Stream {
        private readonly IncrementalHash _hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        internal string Digest() => Convert.ToHexString(_hash.GetHashAndReset());
        public override bool CanWrite => true; public override bool CanRead => false; public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) => _hash.AppendData(buffer, offset, count);
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _hash.Dispose(); base.Dispose(disposing); }
    }
}
