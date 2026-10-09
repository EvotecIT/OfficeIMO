using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal static partial class VisioLegacyXmlCodec {
    private const string ImageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    private const string ObjectRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    private static void ExtractForeignData(XElement content, PackagePart owner, VisioXmlConversionReport report, ref int resourceCount, ref long totalBytes, CancellationToken cancellationToken) {
        int index = 1;
        foreach (XElement foreign in content.Descendants(Modern + "ForeignData")) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++resourceCount > 1024) throw new InvalidDataException("Legacy foreign content exceeds 1024 resources.");
            if (foreign.Value.Length > 90_000_000) throw new InvalidDataException("Legacy foreign resource exceeds its encoded size limit.");
            byte[] bytes;
            try { bytes = Convert.FromBase64String(foreign.Value); }
            catch (FormatException exception) { throw new InvalidDataException("Legacy foreign data is not valid base64.", exception); }
            if (bytes.LongLength > 64L * 1024 * 1024 || (totalBytes += bytes.LongLength) > 128L * 1024 * 1024)
                throw new InvalidDataException("Legacy foreign content exceeds its decoded size limit.");
            bool claimedImage = (string?)foreign.Attribute("ForeignType") is "Bitmap" or "Metafile" or "EnhMetaFile";
            bool recognizedImage = OfficeImageReader.TryIdentifyByContent(VisioForeignImage.PrepareBitmap(bytes), null, out OfficeImageInfo info);
            bool image = claimedImage && recognizedImage;
            // Unknown bytes remain opaque embedded payloads even if the source labels them an image.
            string contentType = image ? info.MimeType : "application/vnd.openxmlformats-officedocument.oleObject";
            string directory = image ? "media" : "embeddings";
            Uri uri;
            do { uri = new Uri("/visio/" + directory + "/legacy" + index++.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".bin", UriKind.Relative); }
            while (owner.Package.PartExists(uri));
            PackagePart part = owner.Package.CreatePart(uri, contentType);
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write)) stream.Write(bytes, 0, bytes.Length);
            PackageRelationship rel = owner.CreateRelationship(PackUriHelper.GetRelativeUri(owner.Uri, uri), TargetMode.Internal, image ? ImageRelationship : ObjectRelationship);
            foreign.ReplaceNodes(new XElement(Modern + "Rel", new XAttribute(Relationships + "id", rel.Id)));
            report.Add("VDX_FOREIGN_PRESERVED", "Embedded foreign bytes and metadata are preserved; image export reports unsupported payloads and placement separately.", OfficeConversionLossKind.Approximation,
                (string?)foreign.Parent?.Attribute("ID"));
        }
    }

    private static XElement InlineForeignData(PackagePart owner, HashSet<Uri> handled, VisioXmlConversionReport report) {
        XElement content = Read(owner);
        foreach (XElement foreign in content.Descendants(Modern + "ForeignData").ToList()) {
            string? id = (string?)foreign.Element(Modern + "Rel")?.Attribute(Relationships + "id");
            if (id == null) { report.Add("VDX_FOREIGN_DATA", "Foreign content has no binary relationship."); foreign.Remove(); continue; }
            PackageRelationship rel = owner.GetRelationship(id);
            if (rel.TargetMode != TargetMode.Internal) { report.Add("VDX_FOREIGN_DATA", "External foreign content cannot be embedded in legacy XML."); foreign.Remove(); continue; }
            PackagePart part = owner.Package.GetPart(PackUriHelper.ResolvePartUri(owner.Uri, rel.TargetUri));
            using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
            using var output = new MemoryStream(); input.CopyTo(output);
            foreign.ReplaceNodes(Convert.ToBase64String(output.ToArray())); handled.Add(part.Uri);
        }
        return content;
    }
}
