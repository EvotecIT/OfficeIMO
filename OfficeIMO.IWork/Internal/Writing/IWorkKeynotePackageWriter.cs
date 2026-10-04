using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml;
using OfficeIMO.Provenance;

namespace OfficeIMO.IWork.Internal;

/// <summary>Encodes native archives and deterministic stored ZIP entries through the shared Core ZIP owner.</summary>
internal static class IWorkKeynotePackageWriter {
    internal static byte[] Write(IWorkKeynoteDocument document, IWorkKeynoteWriteOptions limits, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document.Slides.Count == 0) throw new InvalidOperationException("A native Keynote presentation requires at least one slide.");
        if (document.Slides.Count > limits.MaximumSlides) throw new InvalidDataException("The native Keynote slide limit was exceeded.");
        IWorkKeynoteSlideBuilder[] input = document.Slides.ToArray();
        var slides = new List<(IWorkKeynoteSlideBuilder Slide, IWorkKeynoteTextBox[] Boxes)>(input.Length);
        long totalBoxes = 0, totalCharacters = 0;
        foreach (IWorkKeynoteSlideBuilder slide in input) {
            cancellationToken.ThrowIfCancellationRequested();
            totalBoxes += slide.TextBoxes.Count;
            if (totalBoxes > limits.MaximumTextBoxes) throw new InvalidDataException("The native Keynote text-box limit was exceeded.");
            IWorkKeynoteTextBox[] boxes = slide.TextBoxes.ToArray();
            foreach (IWorkKeynoteTextBox box in boxes) {
                cancellationToken.ThrowIfCancellationRequested();
                totalCharacters += box.Text.Length;
                if (totalCharacters > limits.MaximumTextCharacters) throw new InvalidDataException("The native Keynote text-character limit was exceeded.");
            }
            slides.Add((slide, boxes));
        }
        byte[] hash = ModelHash(document.SlideSize, slides, limits, cancellationToken);
        var builder = new IWorkKeynoteArchiveBuilder(limits, cancellationToken, hash);
        var archives = builder.Build(document.SlideSize, slides);
        var entries = new SortedDictionary<string, byte[]>(StringComparer.Ordinal) {
            ["Index/Document.iwa"] = archives.Document,
            ["Index/Metadata.iwa"] = archives.Metadata,
            ["Metadata/DocumentIdentifier"] = Encoding.ASCII.GetBytes(builder.DocumentIdentifier),
            ["Metadata/Properties.plist"] = Properties(builder.DocumentIdentifier),
            ["Metadata/BuildVersionHistory.plist"] = Plist(writer => {
                writer.WriteStartElement("array"); writer.WriteElementString("string", "OfficeIMO.IWork"); writer.WriteEndElement();
            })
        };
        // Core preserves local DOS timestamps. Supply local midnight so native bytes do not depend on the host time zone.
        var zipTimestamp = new DateTimeOffset(new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Local));
        var zipEntries = entries.Select(entry => new OfficeProvenanceZipWriteEntry(entry.Key, entry.Value.Length,
            false, zipTimestamp, 0, 0,
            Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<byte>(), () => new MemoryStream(entry.Value, writable: false))).ToArray();
        byte[] output = OfficeProvenanceZipWriter.Write(zipEntries, limits.MaximumPackageBytes,
            maximumOutputBytes: limits.MaximumPackageBytes, cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return output;
    }

    private static byte[] ModelHash(IWorkCanvasSize size,
        IEnumerable<(IWorkKeynoteSlideBuilder Slide, IWorkKeynoteTextBox[] Boxes)> slides,
        IWorkKeynoteWriteOptions limits, CancellationToken cancellationToken) {
        IWorkProtoWriter P() => new(limits.MaximumPackageBytes, cancellationToken);
        var input = P().Float(1, (float)size.WidthPoints).Float(2, (float)size.HeightPoints);
        foreach (var slide in slides) {
            var row = P().String(1, slide.Slide.BackgroundColor.RgbHex);
            foreach (IWorkKeynoteTextBox box in slide.Boxes) {
                row.Message(2, P().String(1, box.Text).String(2, box.FontName).Float(3, box.FontSizePoints).String(4, box.Color.RgbHex)
                    .Float(5, (float)box.Geometry.LeftPoints).Float(6, (float)box.Geometry.TopPoints)
                    .Float(7, (float)box.Geometry.WidthPoints).Float(8, (float)box.Geometry.HeightPoints));
            }
            input.Message(3, row);
        }
        using var sha = SHA256.Create();
        return sha.ComputeHash(input.ToArray());
    }

    private static byte[] Properties(string identifier) => Plist(writer => {
        writer.WriteStartElement("dict");
        var values = new SortedDictionary<string, string>(StringComparer.Ordinal) {
            ["documentUUID"] = identifier, ["stableDocumentUUID"] = identifier, ["versionUUID"] = identifier,
            ["privateUUID"] = identifier, ["shareUUID"] = identifier, ["revision"] = "0::" + identifier,
            ["fileFormatVersion"] = "14.4.1"
        };
        foreach (var item in values) { writer.WriteElementString("key", item.Key); writer.WriteElementString("string", item.Value); }
        writer.WriteElementString("key", "isMultiPage"); writer.WriteStartElement("true"); writer.WriteEndElement();
        writer.WriteElementString("key", "hasExternalReferenceOrMissingOrUnmaterializedRemoteData"); writer.WriteStartElement("false"); writer.WriteEndElement();
        writer.WriteEndElement();
    });

    private static byte[] Plist(Action<XmlWriter> content) {
        using var output = new MemoryStream();
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false })) {
            writer.WriteStartDocument();
            writer.WriteDocType("plist", "-//Apple//DTD PLIST 1.0//EN", "http://www.apple.com/DTDs/PropertyList-1.0.dtd", null);
            writer.WriteStartElement("plist"); writer.WriteAttributeString("version", "1.0");
            content(writer);
            writer.WriteEndElement(); writer.WriteEndDocument();
        }
        return output.ToArray();
    }
}
