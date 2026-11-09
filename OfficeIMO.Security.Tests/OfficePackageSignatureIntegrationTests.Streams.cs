using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Security.Tests;

public sealed partial class OfficePackageSignatureIntegrationTests {
    [Theory]
    [InlineData("xlsx", true)]
    [InlineData("xlsx", false)]
    [InlineData("pptx", true)]
    [InlineData("pptx", false)]
    public void StreamValidationVerifiesOriginalPackageAndPreservesCallerOwnership(string extension, bool seekable) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.{extension}");
        try {
            CreatePackage(path, extension);
            using X509Certificate2 certificate = CreateCertificate();
            OfficePackageSigningResult signed = Sign(path, extension, certificate);
            Assert.True(signed.Succeeded, string.Join(" ", signed.Details));
            byte[] original = File.ReadAllBytes(path);
            TamperFirstContentPart(path);
            byte[] tampered = File.ReadAllBytes(path);
            File.Delete(path);

            using var source = new SignatureValidationSource(original, seekable);
            OfficePackageSignatureValidationReport valid = ValidateStream(source, extension);
            Assert.True(valid.IsCryptographicallyValid, string.Join(" ", valid.Findings));
            Assert.True(valid.IsValidUnderPolicy, string.Join(" ", valid.Findings));
            Assert.True(source.CanRead);
            Assert.Equal(original.LongLength, source.BytesRead);
            if (seekable) Assert.Equal(SignatureValidationSource.InitialPosition, source.Position);

            using var changedSource = new SignatureValidationSource(tampered, seekable);
            OfficePackageSignatureValidationReport changed = ValidateStream(changedSource, extension);
            Assert.False(changed.IsCryptographicallyValid);
            Assert.Contains(changed.Signatures.SelectMany(item => item.SignaturePart.SignedReferences),
                reference => reference.DigestVerificationStatus == OfficePackageSignatureValidationState.Failed);
            Assert.True(changedSource.CanRead);
            if (seekable) Assert.Equal(SignatureValidationSource.InitialPosition, changedSource.Position);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void StreamValidationBoundsEncodedInputAndLeavesSourceOpen(bool seekable) {
        using var source = new SignatureValidationSource(new byte[128], seekable);
        var options = new OfficePackageSignatureValidationOptions();
        options.Inspection.MaxPackageBytes = 32;

        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.ValidatePackageSignatures(source, OfficeSecurityProvider.Default, options));

        Assert.True(source.CanRead);
        Assert.Equal(seekable ? 0 : 33, source.BytesRead);
        if (seekable) Assert.Equal(SignatureValidationSource.InitialPosition, source.Position);
    }

    [Fact]
    public void StreamValidationRestoresPositionWhenPackageIsMalformed() {
        using var source = new SignatureValidationSource(new byte[128], seekable: true);

        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.ValidatePackageSignatures(source, OfficeSecurityProvider.Default));

        Assert.Equal(128, source.BytesRead);
        Assert.True(source.CanRead);
        Assert.Equal(SignatureValidationSource.InitialPosition, source.Position);
    }

    private static OfficePackageSignatureValidationReport ValidateStream(Stream source, string extension) {
        var options = new OfficePackageSignatureValidationOptions { ValidateCertificateTrust = false };
        return extension switch {
            "xlsx" => ExcelDocument.ValidatePackageSignatures(source, OfficeSecurityProvider.Default, options),
            "pptx" => PowerPointDocument.ValidateSignatures(source, OfficeSecurityProvider.Default, options),
            _ => throw new ArgumentOutOfRangeException(nameof(extension))
        };
    }

    private sealed class SignatureValidationSource : Stream {
        internal const long InitialPosition = 7;
        private readonly MemoryStream _inner;
        private readonly bool _seekable;

        internal SignatureValidationSource(byte[] package, bool seekable) {
            _seekable = seekable;
            if (seekable) {
                _inner = new MemoryStream(package, writable: false);
            } else {
                var prefixed = new byte[InitialPosition + package.LongLength];
                package.CopyTo(prefixed, InitialPosition);
                _inner = new MemoryStream(prefixed, writable: false);
            }
            _inner.Position = InitialPosition;
        }

        internal long BytesRead { get; private set; }
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => _seekable && _inner.CanSeek;
        public override bool CanWrite => false;
        public override long Length => CanSeek ? _inner.Length : throw new NotSupportedException();
        public override long Position {
            get => CanSeek ? _inner.Position : throw new NotSupportedException();
            set {
                if (!CanSeek) throw new NotSupportedException();
                _inner.Position = value;
            }
        }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = _inner.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }

        public override long Seek(long offset, SeekOrigin origin) =>
            CanSeek ? _inner.Seek(offset, origin) : throw new NotSupportedException();
        public override void Flush() => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
