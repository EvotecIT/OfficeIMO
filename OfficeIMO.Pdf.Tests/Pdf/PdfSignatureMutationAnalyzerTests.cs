using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSignatureMutationAnalyzerTests {
    [Fact]
    public void Analyze_ProvesPermittedDocMdpFormRevisionPreservesExistingSignature() {
        byte[] source = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(permissionLevel: 2);
        var values = new Dictionary<string, string> { ["Name"] = "Grace" };
        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(
            source,
            values,
            new PdfIncrementalFormFieldUpdateOptions {
                GenerateAppearanceStreams = true,
                KeepNeedAppearances = false
            });

        PdfSignatureMutationReport report = PdfSignatureMutationAnalyzer.Analyze(
            source,
            updated,
            PdfMutationOperation.FillFormFields,
            values.Keys);

        Assert.True(report.RequestedChangeIsPermitted);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, report.MutationPlan.ExecutionMode);
        Assert.True(report.OriginalBytesArePrefix);
        Assert.True(report.RevisionChainExtended);
        Assert.True(report.AllExistingSignaturesArePreserved);
        Assert.True(report.IsPreservedAppendOnlyMutation);
        PdfSignatureMutationResult signature = Assert.Single(report.Signatures);
        Assert.Equal(PdfSignatureMutationPermissionStatus.Permitted, signature.PermissionStatus);
        Assert.True(signature.IsPresentAfter);
        Assert.True(signature.ByteRangePreserved);
        Assert.True(signature.HasLaterRevisionsAfter);
        Assert.Contains("Signatures.StructurallyPreserved", report.Diagnostics);
    }

    [Fact]
    public void Analyze_ReportsForbiddenMutationSeparatelyFromUnchangedSignatureBytes() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();

        PdfSignatureMutationReport report = PdfSignatureMutationAnalyzer.Analyze(
            source,
            source,
            PdfMutationOperation.UpdateMetadata);

        Assert.False(report.RequestedChangeIsPermitted);
        Assert.Equal(PdfMutationExecutionMode.Blocked, report.MutationPlan.ExecutionMode);
        Assert.True(report.OriginalBytesArePrefix);
        Assert.False(report.RevisionChainExtended);
        PdfSignatureMutationResult signature = Assert.Single(report.Signatures);
        Assert.Equal(PdfSignatureMutationPermissionStatus.Forbidden, signature.PermissionStatus);
        Assert.True(signature.ByteRangePreserved);
        Assert.Contains("Mutation.Forbidden", report.Diagnostics);
    }

    [Fact]
    public void Analyze_DetectsChangedOriginalBytesEvenWhenSignatureDictionaryRemainsReadable() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();
        byte[] changed = ReplaceFirst(source, "(Alice)", "(Carol)");

        PdfSignatureMutationReport report = PdfSignatureMutationAnalyzer.Analyze(
            source,
            changed,
            PdfMutationOperation.FillFormFields,
            new[] { "Name" });

        Assert.False(report.OriginalBytesArePrefix);
        Assert.False(report.AllExistingSignaturesArePreserved);
        Assert.False(report.IsPreservedAppendOnlyMutation);
        Assert.Contains("Bytes.InputPrefixChanged", report.Diagnostics);
        Assert.Contains("Signatures.StructuralPreservationFailed", report.Diagnostics);
    }

    [Fact]
    public void Analyze_MapsByteRangeToSignedRevisionAndIdentifiesLaterRevision() {
        byte[] original = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Signed revision coverage"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            original,
            new PdfExternalSignatureOptions { ReservedSignatureContentsBytes = 256 });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x03, 0x02, 0x01, 0x01 });
        byte[] appended = AppendMetadataWithoutPolicy(signed, "Later unsigned revision");

        PdfSignatureMutationReport report = PdfSignatureMutationAnalyzer.Analyze(
            signed,
            appended,
            PdfMutationOperation.UpdateMetadata);

        PdfSignatureMutationResult signature = Assert.Single(report.Signatures);
        Assert.Equal(report.Before.Security.RevisionCount, signature.SignedRevisionNumberBefore);
        Assert.Equal(signature.SignedRevisionNumberBefore, signature.SignedRevisionNumberAfter);
        Assert.Equal(report.Before.Security.RevisionCount, signature.CoveredRevisionsBefore.Count);
        Assert.False(signature.HasLaterRevisionsBefore);
        Assert.True(signature.HasLaterRevisionsAfter);
        Assert.True(signature.IsStructurallyPreserved);
        Assert.False(report.RequestedChangeIsPermitted);
        Assert.True(report.RevisionChainExtended);
    }

    [Fact]
    public void Analyze_RejectsSupersededActiveSignatureDictionary() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(source);
        PdfSignatureInfo signature = Assert.Single(security.Signatures);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(source);
        string byteRange = string.Join(" ", signature.ByteRangeValues.Select(value =>
            value.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        string contents = ToHex(signature.ContentsBytes ?? Array.Empty<byte>());
        byte[] replacement = PdfObjectBytes.WrapIndirectObject(
            signature.ObjectNumber,
            "<< /Type /Sig /Filter /Tampered.Handler /SubFilter /adbe.pkcs7.detached /ByteRange [" +
            byteRange + "] /Contents <" + contents + "> >>");
        byte[] superseded = PdfIncrementalObjectWriter.Append(
            source,
            objects,
            security,
            trailerRaw,
            rawObjects: new[] { (signature.ObjectNumber, replacement) });

        PdfSignatureMutationReport report = PdfSignatureMutationAnalyzer.Analyze(
            source,
            superseded,
            PdfMutationOperation.UpdateMetadata);

        PdfSignatureMutationResult result = Assert.Single(report.Signatures);
        Assert.True(report.OriginalBytesArePrefix);
        Assert.True(result.ByteRangePreserved);
        Assert.False(result.ActiveDefinitionPreserved);
        Assert.False(result.IsStructurallyPreserved);
        Assert.False(report.AllExistingSignaturesArePreserved);
    }

    private static byte[] ReplaceFirst(byte[] source, string oldValue, string newValue) {
        byte[] oldBytes = Encoding.ASCII.GetBytes(oldValue);
        byte[] newBytes = Encoding.ASCII.GetBytes(newValue);
        Assert.Equal(oldBytes.Length, newBytes.Length);
        byte[] result = (byte[])source.Clone();
        for (int i = 0; i <= result.Length - oldBytes.Length; i++) {
            if (!result.AsSpan(i, oldBytes.Length).SequenceEqual(oldBytes)) {
                continue;
            }

            newBytes.CopyTo(result, i);
            return result;
        }

        throw new InvalidOperationException("Expected signature marker was not found.");
    }

    private static string ToHex(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2);
        for (int index = 0; index < bytes.Length; index++) {
            builder.Append(bytes[index].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }
        return builder.ToString();
    }

    private static byte[] AppendMetadataWithoutPolicy(byte[] source, string title) {
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(source);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(source);
        int infoObjectNumber = objects.Keys.Max() + 1;
        byte[] infoObject = PdfObjectBytes.WrapIndirectObject(
            infoObjectNumber,
            PdfInfoDictionaryBuilder.Build(new PdfMetadata { Title = title }));
        return PdfIncrementalObjectWriter.Append(
            source,
            objects,
            security,
            trailerRaw,
            rawObjects: new[] { (infoObjectNumber, infoObject) },
            infoObjectNumberOverride: infoObjectNumber);
    }
}
