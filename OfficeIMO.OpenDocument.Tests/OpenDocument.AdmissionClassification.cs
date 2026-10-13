using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentAdmissionClassificationTests {
    [Theory]
    [InlineData("MaxPackageBytes")]
    [InlineData("MaxEntries")]
    [InlineData("MaxDepth")]
    [InlineData("MaxEntryUncompressedBytes")]
    [InlineData("MaxTotalUncompressedBytes")]
    [InlineData("MaxXmlDepth")]
    public void ConfiguredIntegerAdmissionRetainsItsNativeIdentity(string limitName) {
        OdtDocument created = OdtDocument.Create(); created.AddParagraph("Readable content");
        byte[] bytes = created.ToBytes();
        var options = new OdfLoadOptions { DeferBinaryPayloads = true };
        long maximum = 1;
        switch (limitName) {
            case "MaxPackageBytes": options.MaxPackageBytes = maximum = bytes.Length - 1; break;
            case "MaxEntries": options.MaxEntries = 1; break;
            case "MaxDepth": options.MaxDepth = 1; break;
            case "MaxEntryUncompressedBytes": options.MaxEntryUncompressedBytes = 1; break;
            case "MaxTotalUncompressedBytes": options.MaxTotalUncompressedBytes = 1; break;
            case "MaxXmlDepth": options.MaxXmlDepth = 1; break;
        }
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() => {
            OdfDocument loaded = OdfDocument.Load(new MemoryStream(bytes), options);
            _ = ((OdtDocument)loaded).ContentBlocks;
        });
        Assert.Equal("opendocument", error.Owner); Assert.Equal(limitName, error.LimitName); Assert.Equal(maximum, error.Maximum);
        Assert.True(options.DeferBinaryPayloads);
    }

    [Theory]
    [InlineData("stream")]
    [InlineData("async")]
    [InlineData("path")]
    [InlineData("flat")]
    [InlineData("forward")]
    public async Task InputAdmissionUsesTheMarkedStreamGuardAndKeepsCallerOwnership(string route) {
        OdtDocument created = OdtDocument.Create(); created.AddParagraph("Readable content");
        byte[] bytes = route == "flat" ? Encoding.UTF8.GetBytes(created.ToFlatXml().ToString()) : created.ToBytes();
        var options = new OdfLoadOptions { MaxPackageBytes = bytes.Length - 1 };
        using var source = new MemoryStream(bytes); source.Position = Math.Min(7, bytes.Length);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".odt");
        try {
            OfficeInspectionResourceLimitException error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(async () => {
                if (route == "async") await OdfDocument.LoadAsync(source, options);
                else if (route == "path") { File.WriteAllBytes(path, bytes); OdfDocument.Load(path, options); }
                else if (route == "flat") OdfDocument.LoadFlatXml(source, options);
                else if (route == "forward") { source.Position = 0; OdfDocument.Load(new ForwardReadStream(source), options); }
                else OdfDocument.Load(source, options);
            });
            Assert.Equal("MaxPackageBytes", error.LimitName); Assert.Equal(bytes.Length - 1, error.Maximum);
            Assert.True(source.CanRead); if (route != "forward") Assert.Equal(7, source.Position);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void FractionalRatioAdmissionRetainsExactCorePolicyEvidence() {
        OdtDocument created = OdtDocument.Create(); created.AddParagraph(new string('x', 10000));
        byte[] bytes = created.ToBytes(); const double maximum = 1.25;
        OfficePackageSecurityException error = Assert.Throws<OfficePackageSecurityException>(() =>
            OdfDocument.Load(new MemoryStream(bytes), new OdfLoadOptions { MaxCompressionRatio = maximum, DeferBinaryPayloads = true }));
        Assert.Equal(OfficePackageSecurityRule.CompressionRatio, error.Rule); Assert.Equal(maximum, error.Limit);
        Assert.True(error.ObservedValue > maximum); Assert.False(string.IsNullOrEmpty(error.PartName));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BoundedXmlByteEnvelopeAcceptsUtf8Utf16AndReferences(bool utf16) {
        OdtDocument created = OdtDocument.Create(); created.AddParagraph("Café 中 & text");
        byte[] bytes = OdfTestPackageRewriter.Rewrite(created.ToBytes(), (name, original) => {
            if (name != "content.xml") return original;
            string xml = Encoding.UTF8.GetString(original).Replace("&amp;", "&#x26;");
            return utf16 ? Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml.Replace("utf-8", "utf-16"))).ToArray() : Encoding.UTF8.GetBytes(xml);
        });
        OdfDocument loaded = OdfDocument.Load(new MemoryStream(bytes), new OdfLoadOptions {
            DeferBinaryPayloads = true, MaxEntryUncompressedBytes = 8192, MaxXmlCharacters = 8192
        });
        Assert.Equal("Café 中 & text", ((OdtDocument)loaded).ContentBlocks.Single().Paragraph!.Text);
    }

    [Fact]
    public void BoundedEncryptedXmlEnvelopeRetainsDecryptionIntegrityChecks() {
        OdtDocument created = OdtDocument.Create(); created.AddParagraph("Café 中 & text");
        byte[] bytes = created.ToBytes(new OdfSaveOptions { Encryption = new OdfEncryptionOptions { Password = "fixture-password", IterationCount = 10000 } });
        OdfDocument loaded = OdfDocument.Load(new MemoryStream(bytes), new OdfLoadOptions {
            Password = "fixture-password", DeferBinaryPayloads = true, MaxEntryUncompressedBytes = 8192, MaxXmlCharacters = 8192
        });
        Assert.True(loaded.Security.SourceIsEncrypted);
        Assert.Equal("Café 中 & text", ((OdtDocument)loaded).ContentBlocks.Single().Paragraph!.Text);
    }

    [Theory]
    [InlineData("xml")]
    [InlineData("dtd")]
    [InlineData("zip")]
    public void MalformedDataAndDtdRefusalKeepTheirMalformedContract(string kind) {
        byte[] bytes = OdtDocument.Create().ToBytes();
        if (kind == "zip") bytes = new byte[] { 1, 2, 3 };
        else bytes = OdfTestPackageRewriter.Rewrite(bytes, replacements: new[] { new OdfTestPackageEntry("content.xml",
            Encoding.UTF8.GetBytes(kind == "xml" ? "<broken>" : "<!DOCTYPE root [<!ENTITY a 'expanded'>]><root>&a;</root>")) });
        Assert.Throws<InvalidDataException>(() => {
            OdfDocument loaded = OdfDocument.Load(new MemoryStream(bytes), new OdfLoadOptions { MaxEntryUncompressedBytes = 8192, MaxXmlCharacters = 8192 });
            _ = ((OdtDocument)loaded).ContentBlocks;
        });
    }

    private sealed class ForwardReadStream : Stream {
        private readonly Stream _source;
        internal ForwardReadStream(Stream source) => _source = source;
        public override bool CanRead => true; public override bool CanSeek => false; public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _source.Read(buffer, offset, count);
        public override void Flush() { } public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException(); public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
