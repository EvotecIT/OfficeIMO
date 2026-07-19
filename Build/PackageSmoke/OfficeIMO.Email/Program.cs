using OfficeIMO.Email;
using OfficeIMO.Email.AddressBook;
using OfficeIMO.Email.Store;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Image;
using System.Text;

byte[] eml = Encoding.ASCII.GetBytes(
    "Subject: Package smoke\r\n" +
    "Content-Type: text/plain; charset=windows-1252\r\n" +
    "Content-Transfer-Encoding: quoted-printable\r\n\r\n" +
    "Caf=E9\r\n");

EmailReadResult result = new EmailDocumentReader().Read(eml);
if (!string.Equals(result.Document.Body.Text?.Trim(), "Café", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The packed OfficeIMO.Email dependency graph could not decode Windows-1252 text.");
}

byte[] storeMessage = Encoding.ASCII.GetBytes(
    "From: sender@example.test\r\n" +
    "Subject: Packed EMLX\r\n\r\n" +
    "Store body\r\n");
byte[] storePrefix = Encoding.ASCII.GetBytes(storeMessage.Length.ToString() + "\n");
var emlx = new byte[storePrefix.Length + storeMessage.Length];
Buffer.BlockCopy(storePrefix, 0, emlx, 0, storePrefix.Length);
Buffer.BlockCopy(storeMessage, 0, emlx, storePrefix.Length, storeMessage.Length);

using (var stream = new MemoryStream(emlx, writable: false)) {
    EmailStoreReadResult storeResult = new EmailStoreReader().Read(stream, "package-smoke.emlx");
    EmailDocument storeDocument = storeResult.Store.Folders.Single().Items.Single().Document;
    if (!string.Equals(storeDocument.Subject, "Packed EMLX", StringComparison.Ordinal) ||
        !string.Equals(storeDocument.Body.Text?.Trim(), "Store body", StringComparison.Ordinal)) {
        throw new InvalidOperationException("The packed OfficeIMO.Email dependency graph could not project EMLX through the Store API.");
    }
}

string storePath = Path.Combine(Path.GetTempPath(),
    "officeimo-reader-email-store-package-smoke-" + Guid.NewGuid().ToString("N") + ".emlx");
try {
    File.WriteAllBytes(storePath, emlx);
    OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
        .AddEmailStoreHandler()
        .Build();
    ReaderEmailStoreItemResult readerItem = reader.ReadEmailStoreItems(
        storePath,
        new ReaderOptions { ComputeHashes = false },
        new ReaderEmailStoreOptions { MaxItems = 1 }).Single();
    if (!readerItem.Succeeded ||
        !readerItem.Chunks.Any(chunk => chunk.Text.Contains("Store body", StringComparison.Ordinal))) {
        throw new InvalidOperationException(
            "The packed OfficeIMO.Reader.Email dependency graph could not project an EMLX item.");
    }
} finally {
    if (File.Exists(storePath)) File.Delete(storePath);
}

byte[] oab = CreateAddressBook();
using (var stream = new MemoryStream(oab, writable: false))
using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "package-smoke.oab")) {
    OfflineAddressBookEntry entry = session.EnumerateEntries().Single();
    if (!string.Equals(entry.SmtpAddress, "package@example.test", StringComparison.Ordinal) ||
        !session.Validate().IsValid) {
        throw new InvalidOperationException(
            "The packed OfficeIMO.Email dependency graph could not decode and validate OAB v4 through the AddressBook API.");
    }
}

using (var stream = new MemoryStream(oab, writable: false)) {
    OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
        .AddEmailAddressBookHandler(new ReaderEmailAddressBookOptions { MaxEntries = 1 })
        .Build();
    OfficeDocumentReadResult readerDocument = reader.ReadDocument(
        stream, "package-smoke.oab", new ReaderOptions { ComputeHashes = false });
    if (!readerDocument.Chunks.Single().Text.Contains("package@example.test", StringComparison.Ordinal) ||
        !readerDocument.CapabilitiesUsed.Contains(
            OfficeDocumentReaderBuilderEmailAddressBookExtensions.HandlerId, StringComparer.Ordinal)) {
        throw new InvalidOperationException(
            "The packed OfficeIMO.Reader.Email dependency graph could not project an OAB entry.");
    }
}

byte[] svg = Encoding.UTF8.GetBytes(
    "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"2\" height=\"1\"><rect width=\"2\" height=\"1\"/></svg>");
using (var stream = new MemoryStream(svg, writable: false)) {
    OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
        .AddImageHandler()
        .Build();
    OfficeDocumentReadResult image = reader.ReadDocument(
        stream, "package-smoke.svg", new ReaderOptions { ComputeHashes = false });
    if (image.Assets.Count != 1 ||
        !string.Equals(image.Assets[0].MediaType, "image/svg+xml", StringComparison.Ordinal)) {
        throw new InvalidOperationException(
            "The packed OfficeIMO.Reader.Image dependency graph could not execute OfficeIMO.Drawing image identification.");
    }
}

Console.WriteLine($"Unified OfficeIMO Email and selective Reader package smokes passed on {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}.");

static byte[] CreateAddressBook() {
    using var stream = new MemoryStream();
    WriteUInt32(stream, 0x20);
    WriteUInt32(stream, 0);
    WriteUInt32(stream, 1);
    using (var metadata = new MemoryStream()) {
        WriteUInt32(metadata, 4);
        WriteProperty(metadata, 0x6800001F);
        WriteProperty(metadata, 0x68010003);
        WriteProperty(metadata, 0x6802001F);
        WriteProperty(metadata, 0x6804001F);
        WriteUInt32(metadata, 2);
        WriteProperty(metadata, 0x3003001F);
        WriteProperty(metadata, 0x39FE001F);
        WriteUInt32(stream, checked((uint)metadata.Length + 4));
        metadata.Position = 0;
        metadata.CopyTo(stream);
    }
    WriteUInt32(stream, 5);
    stream.WriteByte(0);
    byte[] x500 = Encoding.UTF8.GetBytes("/o=Package/ou=Recipients/cn=smoke\0");
    byte[] smtp = Encoding.UTF8.GetBytes("package@example.test\0");
    WriteUInt32(stream, checked((uint)(5 + x500.Length + smtp.Length)));
    stream.WriteByte(0xC0);
    stream.Write(x500, 0, x500.Length);
    stream.Write(smtp, 0, smtp.Length);
    byte[] result = stream.ToArray();
    WriteUInt32At(result, 4, ComputeCrc(result));
    return result;
}

static void WriteProperty(Stream stream, uint tag) {
    WriteUInt32(stream, tag);
    WriteUInt32(stream, 0);
}

static void WriteUInt32(Stream stream, uint value) {
    stream.WriteByte((byte)value);
    stream.WriteByte((byte)(value >> 8));
    stream.WriteByte((byte)(value >> 16));
    stream.WriteByte((byte)(value >> 24));
}

static void WriteUInt32At(byte[] data, int offset, uint value) {
    data[offset] = (byte)value;
    data[offset + 1] = (byte)(value >> 8);
    data[offset + 2] = (byte)(value >> 16);
    data[offset + 3] = (byte)(value >> 24);
}

static uint ComputeCrc(byte[] data) {
    uint crc = 0xFFFFFFFFU;
    for (int index = 12; index < data.Length; index++) {
        crc ^= data[index];
        for (int bit = 0; bit < 8; bit++) {
            crc = (crc & 1U) != 0 ? (crc >> 1) ^ 0xEDB88320U : crc >> 1;
        }
    }
    return crc;
}
