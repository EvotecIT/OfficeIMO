using System.Collections.Generic;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectNativeContainerTests {
    [Fact]
    public void MetadataEditRetainsUnchangedRawSectionsAndPropertyValues() {
        var summaryId = OfficeOlePropertySetWriter.SummaryInformationFormatId; var customId = Guid.NewGuid();
        byte[] custom = OfficeOlePropertySetWriter.CreateSection(new[] { OfficeOleProperty.Integer(1, (short)1200), OfficeOleProperty.String(2, "Opaque custom metadata") });
        byte[] source = OfficeOlePropertySetWriter.CreatePropertySet((summaryId, OfficeOlePropertySetWriter.CreateSection(new[] {
            OfficeOleProperty.Integer(1, (short)1252), OfficeOleProperty.String(2, "Original"), OfficeOleProperty.String(4, "Retained author") })), (customId, custom));
        byte[] result = OfficeOlePropertySetEditor.RewriteStrings(source, summaryId, new Dictionary<uint, string?> { [2] = "Expanded café / Łódź / 日本語" }, default);
        var sections = OfficeOlePropertySetReader.ReadSections(result);
        Assert.Equal("Expanded café / Łódź / 日本語", sections.Single(s => s.FormatId == summaryId).Properties[2].AsString());
        Assert.Equal("Retained author", sections.Single(s => s.FormatId == summaryId).Properties[4].AsString());
        int offset = BitConverter.ToInt32(result, 64);
        Assert.Equal(custom, result.Skip(offset).Take(custom.Length).ToArray());
        Assert.Throws<InvalidDataException>(() => OfficeOlePropertySetEditor.RewriteStrings(new byte[32], summaryId, new Dictionary<uint, string?>(), default));
        Assert.Throws<OperationCanceledException>(() => OfficeOlePropertySetEditor.RewriteStrings(source, summaryId, new Dictionary<uint, string?>(), new CancellationToken(true)));
    }

    [Fact]
    public void CompoundRewriteBoundsTheContainerAndCanCancelDuringPayloadCopy() {
        var streams = new[] { new OfficeCompoundStream("Payload", new byte[512 * 1024]), new OfficeCompoundStream("Inert/VBA", new byte[35]) };
        byte[] source = OfficeCompoundFileWriter.Write(streams);
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? file, out var error), error);
        Assert.Throws<InvalidDataException>(() => OfficeCompoundFileWriter.Rewrite(file!, new Dictionary<string, byte[]>(), maxOutputBytes: source.Length - 1));
        using var cancellation = new CancellationTokenSource(); using var target = new CancelingOutput(cancellation);
        Assert.Throws<OperationCanceledException>(() => OfficeCompoundFileWriter.Write(target, streams, cancellationToken: cancellation.Token));
        Assert.InRange(target.Length, 512, 512 + 81920);
    }

    [Theory]
    [InlineData("MSProject.MPP14")]
    [InlineData("MSProject.MPT14")]
    public void OleIdentificationIsInertAndBoundsClipboardStrings(string format) {
        byte[] bytes = OfficeOleCompoundObjectWriter.Write(Guid.NewGuid(), "Project", format, "MSProject.Project.9");
        Assert.Equal(format, OfficeOleCompoundObjectReader.ClipboardFormat(bytes));
        Buffer.BlockCopy(BitConverter.GetBytes(uint.MaxValue), 0, bytes, 28, 4);
        Assert.Null(OfficeOleCompoundObjectReader.ClipboardFormat(bytes));
        Assert.Null(OfficeOleCompoundObjectReader.ClipboardFormat(new byte[20]));
    }

    private sealed class CancelingOutput : MemoryStream {
        private readonly CancellationTokenSource _source;
        internal CancelingOutput(CancellationTokenSource source) => _source = source;
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            if (Length > 512) _source.Cancel();
        }
    }
}
