using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstPasswordTests {
    [Fact]
    public void ComputesThePstCrc32VariantWithoutInitialOrFinalXor() {
        uint checksum = PstPassword.ComputeChecksum(Encoding.ASCII.GetBytes("123456789"));

        Assert.Equal(0x2DFD2D88U, checksum);
    }

    [Fact]
    public void RequiresAndValidatesTheConfiguredPassword() {
        const string password = "OfficeIMO";
        uint checksum = PstPassword.ComputeChecksum(Encoding.ASCII.GetBytes(password));
        var properties = new[] {
            new MapiProperty(0x67FF, MapiPropertyType.Integer32, unchecked((int)checksum))
        };

        EmailStorePasswordException missing = Assert.Throws<EmailStorePasswordException>(
            () => PstPassword.Validate(properties, new EmailStoreReaderOptions()));
        EmailStorePasswordException mismatch = Assert.Throws<EmailStorePasswordException>(
            () => PstPassword.Validate(properties, new EmailStoreReaderOptions(pstPassword: "wrong")));
        PstPassword.Validate(properties, new EmailStoreReaderOptions(pstPassword: password));

        Assert.False(missing.PasswordWasProvided);
        Assert.True(mismatch.PasswordWasProvided);
    }

    [Fact]
    public void DoesNotRequireAPasswordForZeroOrMissingChecksum() {
        PstPassword.Validate(Array.Empty<MapiProperty>(), new EmailStoreReaderOptions());
        PstPassword.Validate(
            new[] { new MapiProperty(0x67FF, MapiPropertyType.Integer32, 0) },
            new EmailStoreReaderOptions());
    }

    [Fact]
    public void ProtectionDetectionDistinguishesNonzeroChecksum() {
        Assert.False(PstPassword.IsProtected(Array.Empty<MapiProperty>()));
        Assert.False(PstPassword.IsProtected(
            new[] { new MapiProperty(0x67FF, MapiPropertyType.Integer32, 0) }));
        Assert.True(PstPassword.IsProtected(
            new[] { new MapiProperty(0x67FF, MapiPropertyType.Integer32, 1) }));
    }

    [Fact]
    public void OstOpenDoesNotApplyTheLegacyPstPasswordGate() {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(
            ost: true, storePasswordChecksum: 0x12345678));

        using EmailStoreSession session = EmailStoreSession.Open(stream, "mailbox.ost");

        Assert.Equal(EmailStoreFormat.Ost, session.Format);
    }
}
