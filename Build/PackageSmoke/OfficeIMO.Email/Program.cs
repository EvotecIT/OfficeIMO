using OfficeIMO.Email;
using OfficeIMO.Email.Store;
using OfficeIMO.Reader;
using OfficeIMO.Reader.EmailStore;
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
        throw new InvalidOperationException("The packed OfficeIMO.Email.Store dependency graph could not project EMLX through OfficeIMO.Email.");
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
            "The packed OfficeIMO.Reader.EmailStore dependency graph could not project an EMLX item.");
    }
} finally {
    if (File.Exists(storePath)) File.Delete(storePath);
}

Console.WriteLine($"OfficeIMO Email and Reader email-store package smoke passed on {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}.");
