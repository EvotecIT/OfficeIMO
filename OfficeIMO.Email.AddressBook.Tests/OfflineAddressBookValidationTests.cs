namespace OfficeIMO.Email.AddressBook.Tests;

public sealed class OfflineAddressBookValidationTests {
    [Fact]
    public void ValidatesChecksumFramingAndSchemaDecoding() {
        byte[] oab = new OabV4Fixture().Build();
        using (var stream = new MemoryStream(oab, writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            OfflineAddressBookValidationReport report = session.Validate();
            OfflineAddressBookValidationResult result = Assert.Single(report.Results);

            Assert.True(report.IsValid);
            Assert.True(result.IsChecksumValid);
            Assert.Equal(3, result.RecordsScanned);
            Assert.Equal(3, result.RecordsDecoded);
            Assert.True(result.FramingComplete);
            Assert.True(result.ConsumedDeclaredPayload);
        }
    }

    [Fact]
    public void ReportsChecksumMismatchWithoutExposingRecordContent() {
        byte[] oab = new OabV4Fixture().Build();
        oab[oab.Length - 1] ^= 0x40;
        using (var stream = new MemoryStream(oab, writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            OfflineAddressBookValidationReport report = session.Validate(
                new OfflineAddressBookValidationOptions(
                    mode: OfflineAddressBookValidationMode.ChecksumOnly));

            OfflineAddressBookValidationResult result = Assert.Single(report.Results);
            Assert.False(result.IsChecksumValid);
            Assert.False(result.IsValid);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OAB_CHECKSUM_MISMATCH");
        }
    }

    [Fact]
    public void FullDecodeRejectsEntriesMissingPrimaryKeyProperties() {
        byte[] oab = new OabV4Fixture()
            .RemoveEntryProperty(0, OabPropertyTags.SmtpAddress, MapiPropertyType.Unicode)
            .Build();
        using (var stream = new MemoryStream(oab, writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            OfflineAddressBookValidationResult result = Assert.Single(session.Validate(
                new OfflineAddressBookValidationOptions(
                    mode: OfflineAddressBookValidationMode.FullDecode)).Results);

            Assert.False(result.IsValid);
            Assert.Equal(1, result.RecordsSkipped);
            Assert.Contains(result.Diagnostics, diagnostic =>
                diagnostic.Code == "OAB_VALIDATION_ENTRY_SKIPPED" &&
                diagnostic.Message.Contains("Required OAB primary-key property", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void DetectsTrailingPayloadAndConfiguredValidationBounds() {
        byte[] original = new OabV4Fixture().Build();
        byte[] withTrailingData = new byte[original.Length + 4];
        Buffer.BlockCopy(original, 0, withTrailingData, 0, original.Length);
        withTrailingData[withTrailingData.Length - 1] = 1;
        WriteUInt32(withTrailingData, 4, OabV4Fixture.ComputeCrc(withTrailingData, 12));

        using (var stream = new MemoryStream(withTrailingData, writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            OfflineAddressBookValidationResult result = Assert.Single(session.Validate().Results);
            Assert.True(result.IsChecksumValid);
            Assert.False(result.ConsumedDeclaredPayload);
            Assert.False(result.IsValid);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OAB_VALIDATION_TRAILING_DATA");
        }

        using (var stream = new MemoryStream(original, writable: false))
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "synthetic.oab")) {
            OfflineAddressBookValidationResult result = Assert.Single(session.Validate(
                new OfflineAddressBookValidationOptions(
                    validateChecksum: false,
                    maxEntriesPerAddressList: 1)).Results);
            Assert.False(result.FramingComplete);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OAB_VALIDATION_ENTRY_LIMIT");

            Assert.Throws<OfflineAddressBookLimitExceededException>(() => session.Validate(
                new OfflineAddressBookValidationOptions(maxChecksumBytesPerFile: 1)));
        }
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)(value >> 16);
        data[offset + 3] = (byte)(value >> 24);
    }
}
