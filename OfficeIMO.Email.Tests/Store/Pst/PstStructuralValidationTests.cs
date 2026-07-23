namespace OfficeIMO.Email.Store.Tests;

public sealed class PstStructuralValidationTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, true)]
    public void ValidatesPageAndBlockTrailersAcrossSupportedVariants(
        bool ansi, bool fourK, bool compressed) {
        using var source = new MemoryStream(PstTestFileBuilder.Create(
            ost: fourK, ansi: ansi, fourK: fourK, compressBlocks: compressed));
        using EmailStoreSession session = EmailStoreSession.Open(
            source, fourK ? "mailbox.ost" : "mailbox.pst");

        EmailStoreValidationReport report = session.Validate(
            new EmailStoreValidationOptions(
                mode: EmailStoreValidationMode.Shallow,
                verifyStructuralIntegrity: true,
                maxStructuralPages: 100,
                maxStructuralBlocks: 100,
                maxStructuralBytes: 1024 * 1024));

        Assert.True(report.StructuralIntegrityRequested);
        Assert.True(report.StructuralIntegritySupported);
        Assert.Equal(2, report.StructuralPagesExamined);
        Assert.Equal(4, report.StructuralBlocksExamined);
        Assert.True(report.StructuralBytesExamined > 0);
        Assert.True(report.StructuralFailures == 0,
            string.Join(" | ", report.Diagnostics.Select(item => item.Code + ":" + item.Message)));
        Assert.False(report.StructuralValidationWasTruncated);
        Assert.True(report.IsComplete);
        Assert.True(report.IsValid);
    }

    [Theory]
    [InlineData(true, false, "EMAIL_STORE_PST_PAGE_CRC")]
    [InlineData(false, true, "EMAIL_STORE_PST_BLOCK_CRC")]
    public void ReportsCorruptPageAndBlockCrcs(
        bool corruptPageCrc, bool corruptBlockCrc, string expectedCode) {
        using var source = new MemoryStream(PstTestFileBuilder.Create(
            corruptPageCrc: corruptPageCrc,
            corruptBlockCrc: corruptBlockCrc));
        using EmailStoreSession session = EmailStoreSession.Open(source, "mailbox.pst");

        EmailStoreValidationReport report = session.Validate(
            new EmailStoreValidationOptions(
                mode: EmailStoreValidationMode.Shallow,
                verifyStructuralIntegrity: true,
                maxStructuralPages: 100,
                maxStructuralBlocks: 100,
                maxStructuralBytes: 1024 * 1024));

        Assert.True(report.StructuralFailures > 0);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == expectedCode);
        Assert.False(report.IsComplete);
        Assert.False(report.IsValid);
    }

    [Fact]
    public void ReportsWhenStructuralValidationStopsAtABound() {
        using var source = new MemoryStream(PstTestFileBuilder.Create());
        using EmailStoreSession session = EmailStoreSession.Open(source, "mailbox.pst");

        EmailStoreValidationReport report = session.Validate(
            new EmailStoreValidationOptions(
                mode: EmailStoreValidationMode.Shallow,
                verifyStructuralIntegrity: true,
                maxStructuralPages: 100,
                maxStructuralBlocks: 1,
                maxStructuralBytes: 1024 * 1024));

        Assert.Equal(1, report.StructuralBlocksExamined);
        Assert.True(report.StructuralValidationWasTruncated);
        Assert.False(report.IsComplete);
        Assert.True(report.IsValid);
    }

    [Fact]
    public void ReportsUnsignedChildOffsetsOutsideTheSupportedStreamRange() {
        const int nbtOffset = 1536;
        const int pageDataLength = 496;
        byte[] bytes = PstTestFileBuilder.Create();
        using var source = new MemoryStream(bytes);
        using EmailStoreSession session = EmailStoreSession.Open(source, "mailbox.pst");
        bytes[nbtOffset + 491] = 1;
        Buffer.BlockCopy(BitConverter.GetBytes(ulong.MaxValue), 0, bytes, nbtOffset + 16, 8);
        Buffer.BlockCopy(BitConverter.GetBytes(PstCrc32.Compute(bytes, nbtOffset, pageDataLength)),
            0, bytes, nbtOffset + 500, 4);

        EmailStoreValidationReport report = session.Validate(
            new EmailStoreValidationOptions(
                mode: EmailStoreValidationMode.Shallow,
                verifyStructuralIntegrity: true,
                maxStructuralPages: 100,
                maxStructuralBlocks: 100,
                maxStructuralBytes: 1024 * 1024));

        Assert.Contains(report.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_PST_PAGE_CHILD_OFFSET");
        Assert.False(report.IsValid);
    }
}
