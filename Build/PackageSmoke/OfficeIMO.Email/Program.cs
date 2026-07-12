using OfficeIMO.Email;
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

Console.WriteLine($"OfficeIMO.Email package smoke passed on {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}.");
