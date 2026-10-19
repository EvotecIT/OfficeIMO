using DocumentFormat.OpenXml.Packaging;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace OfficeIMO.Excel.Utilities;

/// <summary>Fingerprints part content and relationship identities without materializing another XML string.</summary>
internal static class ExcelPackagePartFingerprint {
    internal static string Compute(OpenXmlPart part, bool includeContent = true, Action<XmlWriter>? writeContent = null, Func<OpenXmlPart, bool>? includeChildPart = null) {
        using var hash = SHA256.Create();
        using (var stream = new CryptoStream(Stream.Null, hash, CryptoStreamMode.Write)) {
            using (var text = new StreamWriter(stream, new UTF8Encoding(false), 4096, leaveOpen: true)) {
                text.Write(part.ContentType); text.Write('\n');
                foreach (var pair in part.Parts.Where(item => item.OpenXmlPart is not SharedStringTablePart && (includeChildPart == null || includeChildPart(item.OpenXmlPart)))
                    .OrderBy(item => item.RelationshipId, StringComparer.Ordinal))
                    text.Write(pair.RelationshipId + "|" + pair.OpenXmlPart.RelationshipType + "|" + pair.OpenXmlPart.Uri.OriginalString + "\n");
                foreach (var relationship in part.ExternalRelationships.OrderBy(item => item.Id, StringComparer.Ordinal))
                    text.Write(relationship.Id + "|" + relationship.RelationshipType + "|" + relationship.Uri.OriginalString + "\n");
                foreach (var relationship in part.HyperlinkRelationships.OrderBy(item => item.Id, StringComparer.Ordinal))
                    text.Write(relationship.Id + "|hyperlink|" + relationship.Uri.OriginalString + "\n");
                text.Flush();
            }
            if (includeContent) {
                if (writeContent != null || part.RootElement != null) {
                    using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false, OmitXmlDeclaration = true });
                    if (writeContent != null) writeContent(writer); else part.RootElement!.WriteTo(writer);
                    writer.Flush();
                } else {
                    using var input = part.GetStream(FileMode.Open, FileAccess.Read); input.CopyTo(stream);
                }
            }
            stream.FlushFinalBlock();
        }
        return Convert.ToBase64String(hash.Hash!);
    }
}
