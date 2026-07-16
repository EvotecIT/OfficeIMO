namespace OfficeIMO.Email.AddressBook.Tests;

public sealed class OfflineAddressBookInspectorTests {
    [Fact]
    public void InventoriesReadableLegacyTemplateAndUnknownComponentsWithoutOpeningASession() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-oab-inspect-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            File.WriteAllBytes(Path.Combine(directory, "udetails.oab"), new OabV4Fixture().Build());
            File.WriteAllBytes(Path.Combine(directory, "uanrdex.oab"), new byte[] { 2, 0, 0, 0 });
            File.WriteAllBytes(Path.Combine(directory, "template.oab"), new byte[] { 7, 0, 0, 0 });
            File.WriteAllBytes(Path.Combine(directory, "future.oab"), new byte[] { 99, 0, 0, 0 });

            OfflineAddressBookDiscoveryReport report = OfflineAddressBookInspector.Inspect(directory);

            Assert.Equal(4, report.Files.Count);
            Assert.Equal(1, report.ReadableFullDetailsCount);
            Assert.Equal(3, report.NonEntryComponentCount);
            Assert.Contains(report.Files, file => file.Format == OfflineAddressBookFormat.LegacyAnrIndex);
            Assert.Contains(report.Files, file => file.Format == OfflineAddressBookFormat.DisplayTemplate);
            Assert.Contains(report.Files, file => file.Format == OfflineAddressBookFormat.Unknown);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void StreamInspectionRestoresCallerPosition() {
        byte[] oab = new OabV4Fixture().Build();
        using (var stream = new MemoryStream()) {
            stream.Write(new byte[9], 0, 9);
            stream.Write(oab, 0, oab.Length);
            stream.Position = 9;

            OfflineAddressBookFileInfo info = OfflineAddressBookInspector.Inspect(stream, "details.oab");

            Assert.Equal(OfflineAddressBookFormat.Version4FullDetails, info.Format);
            Assert.Equal(9, stream.Position);
        }
    }
}
