using System.IO;
using System.Security.Cryptography;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Operators.Utilities;

namespace OfficeIMO.Security;

internal sealed class PlatformRsaSignatureFactory : ISignatureFactory {
    private const int MaxSignedAttributeBytes = 4 * 1024 * 1024;
    private readonly RSA _rsa;
    private readonly HashAlgorithmName _digestAlgorithm;

    internal PlatformRsaSignatureFactory(RSA rsa, HashAlgorithmName digestAlgorithm) {
        _rsa = rsa ?? throw new ArgumentNullException(nameof(rsa));
        _digestAlgorithm = digestAlgorithm;
        string algorithmName = GetAlgorithmName(digestAlgorithm);
        AlgorithmDetails = DefaultSignatureAlgorithmFinder.Instance.Find(algorithmName);
    }

    public object AlgorithmDetails { get; }

    public IStreamCalculator<IBlockResult> CreateCalculator() =>
        new Calculator(_rsa, _digestAlgorithm);

    private static string GetAlgorithmName(HashAlgorithmName algorithm) {
        if (algorithm == HashAlgorithmName.SHA256) return "SHA256WITHRSA";
        if (algorithm == HashAlgorithmName.SHA384) return "SHA384WITHRSA";
        if (algorithm == HashAlgorithmName.SHA512) return "SHA512WITHRSA";
        if (algorithm == HashAlgorithmName.SHA1) return "SHA1WITHRSA";
        throw new NotSupportedException("CMS RSA signing supports SHA-1, SHA-256, SHA-384, and SHA-512.");
    }

    private sealed class Calculator : IStreamCalculator<IBlockResult>, IDisposable {
        private readonly RSA _rsa;
        private readonly HashAlgorithmName _digestAlgorithm;
        private readonly BoundedMemoryStream _stream = new BoundedMemoryStream(MaxSignedAttributeBytes);
        private bool _completed;

        internal Calculator(RSA rsa, HashAlgorithmName digestAlgorithm) {
            _rsa = rsa;
            _digestAlgorithm = digestAlgorithm;
        }

        public Stream Stream => _stream;

        public IBlockResult GetResult() {
            if (_completed) throw new InvalidOperationException("The CMS signature calculator has already completed.");
            _completed = true;
            byte[] signature = _rsa.SignData(
                _stream.ToArray(),
                _digestAlgorithm,
                RSASignaturePadding.Pkcs1);
            _stream.Dispose();
            return new SimpleBlockResult(signature);
        }

        public void Dispose() {
            _stream.Dispose();
        }
    }

    private sealed class BoundedMemoryStream : MemoryStream {
        private readonly long _maximumLength;

        internal BoundedMemoryStream(long maximumLength) {
            _maximumLength = maximumLength;
        }

        public override void Write(byte[] buffer, int offset, int count) {
            EnsureCapacity(count);
            base.Write(buffer, offset, count);
        }

        public override void WriteByte(byte value) {
            EnsureCapacity(1);
            base.WriteByte(value);
        }

#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) {
            EnsureCapacity(buffer.Length);
            base.Write(buffer);
        }
#endif

        private void EnsureCapacity(int additionalBytes) {
            if (additionalBytes < 0 || Position > _maximumLength - additionalBytes) {
                throw new InvalidDataException("CMS signed attributes exceed the configured internal safety limit.");
            }
        }
    }
}
