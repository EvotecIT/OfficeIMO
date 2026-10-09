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
        var slides = new List<(IWorkKeynoteSlideBuilder Slide, IWorkKeynoteCreationDrawable[] Drawables)>(input.Length);
        long totalBoxes = 0, totalCharacters = 0, totalRuns = 0;
        long totalImages = 0, encodedImageBytes = 0, decodedImageBytes = 0;
        long totalTables = 0, totalTableCells = 0;
        foreach (IWorkKeynoteSlideBuilder slide in input) {
            cancellationToken.ThrowIfCancellationRequested();
            totalBoxes += slide.TextBoxes.Count;
            if (totalBoxes > limits.MaximumTextBoxes) throw new InvalidDataException("The native Keynote text-box limit was exceeded.");
            IWorkKeynoteCreationDrawable[] drawables = slide.Drawables.ToArray();
            foreach (IWorkKeynoteTextBox box in drawables.OfType<IWorkKeynoteTextBox>()) {
                cancellationToken.ThrowIfCancellationRequested();
                totalRuns += box.Runs.Count;
                if (totalRuns > limits.MaximumTextRuns) throw new InvalidDataException("The native Keynote character-run limit was exceeded.");
                totalCharacters += box.Text.Length;
                if (totalCharacters > limits.MaximumTextCharacters) throw new InvalidDataException("The native Keynote text-character limit was exceeded.");
            }
            foreach (IWorkKeynoteImage image in drawables.OfType<IWorkKeynoteImage>()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (++totalImages > limits.MaximumImages) throw new InvalidDataException("The native Keynote image limit was exceeded.");
                encodedImageBytes = checked(encodedImageBytes + image.Length);
                decodedImageBytes = checked(decodedImageBytes + image.DecodedBytes);
                if (encodedImageBytes > limits.MaximumEncodedImageBytes)
                    throw new InvalidDataException("The native Keynote encoded-image limit was exceeded.");
                if (decodedImageBytes > limits.MaximumDecodedImageBytes)
                    throw new InvalidDataException("The native Keynote decoded-image limit was exceeded.");
                if (encodedImageBytes > limits.MaximumPackageBytes)
                    throw new InvalidDataException("The native Keynote package limit was exceeded by image payloads.");
            }
            foreach (IWorkKeynoteTable table in drawables.OfType<IWorkKeynoteTable>()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (++totalTables > limits.MaximumTables) throw new InvalidDataException("The native Keynote table limit was exceeded.");
                totalTableCells += (long)table.RowCount * table.ColumnCount;
                if (totalTableCells > limits.MaximumTableCells) throw new InvalidDataException("The native Keynote table-cell limit was exceeded.");
                totalCharacters += table.Name.Length;
                foreach (IWorkTableCell cell in table.Cells) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (cell.Kind == IWorkCellKind.Text) totalCharacters += ((string)cell.Value!).Length;
                    if (totalCharacters > limits.MaximumTextCharacters) throw new InvalidDataException("The native Keynote text-character limit was exceeded.");
                }
                if (totalCharacters > limits.MaximumTextCharacters) throw new InvalidDataException("The native Keynote text-character limit was exceeded.");
            }
            slides.Add((slide, drawables));
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
        foreach (var resource in builder.ImageResources) entries.Add("Data/" + resource.FileName, resource.Bytes);
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
        IEnumerable<(IWorkKeynoteSlideBuilder Slide, IWorkKeynoteCreationDrawable[] Drawables)> slides,
        IWorkKeynoteWriteOptions limits, CancellationToken cancellationToken) {
        IWorkProtoWriter P() => new(limits.MaximumPackageBytes, cancellationToken);
        var input = P().Float(1, (float)size.WidthPoints).Float(2, (float)size.HeightPoints);
        foreach (var slide in slides) {
            var row = P().String(1, slide.Slide.BackgroundColor.RgbHex);
            bool hasMixedDrawables = slide.Drawables.Any(drawable => drawable is not IWorkKeynoteTextBox);
            foreach (IWorkKeynoteCreationDrawable drawable in slide.Drawables) {
                cancellationToken.ThrowIfCancellationRequested();
                if (drawable is IWorkKeynoteImage image) {
                    var data = P().Bytes(1, image.ContentHash).Float(2, (float)image.Geometry.LeftPoints)
                        .Float(3, (float)image.Geometry.TopPoints).Float(4, (float)image.Geometry.WidthPoints)
                        .Float(5, (float)image.Geometry.HeightPoints);
                    if (image.AccessibilityDescription != null) data.String(6, image.AccessibilityDescription);
                    row.Message(3, P().Message(2, data));
                    continue;
                }
                if (drawable is IWorkKeynoteTable table) {
                    var data = P().UInt(1, (uint)table.RowCount).UInt(2, (uint)table.ColumnCount).String(3, table.Name)
                        .Float(4, (float)table.Geometry.LeftPoints).Float(5, (float)table.Geometry.TopPoints)
                        .Float(6, (float)table.Geometry.WidthPoints).Float(7, (float)table.Geometry.HeightPoints)
                        .String(8, table.FontName).Float(9, table.FontSizePoints).String(10, table.Color.RgbHex);
                    foreach (IWorkTableCell cell in table.Cells) {
                        var value = P().UInt(1, (uint)cell.Row).UInt(2, (uint)cell.Column).UInt(3, (uint)cell.Kind);
                        if (cell.Kind == IWorkCellKind.Text) value.String(4, (string)cell.Value!);
                        else if (cell.Kind == IWorkCellKind.Number) value.Double(5, (double)cell.Value!);
                        else if (cell.Kind == IWorkCellKind.Boolean) value.Bool(6, (bool)cell.Value!);
                        data.Message(11, value);
                    }
                    row.Message(3, P().Message(3, data));
                    continue;
                }
                var box = (IWorkKeynoteTextBox)drawable;
                var textBox = P().String(1, box.Text).String(2, box.FontName).Float(3, box.FontSizePoints).String(4, box.Color.RgbHex)
                    .Float(5, (float)box.Geometry.LeftPoints).Float(6, (float)box.Geometry.TopPoints)
                    .Float(7, (float)box.Geometry.WidthPoints).Float(8, (float)box.Geometry.HeightPoints);
                if (box.Runs.Count > 1 || box.Runs[0].HasDecoration) {
                    foreach (IWorkKeynoteTextRun run in box.Runs) {
                        textBox.Message(9, P().String(1, run.Text).String(2, run.FontName).Float(3, run.FontSizePoints)
                            .String(4, run.Color.RgbHex).Bool(5, run.Bold).Bool(6, run.Italic).Bool(7, run.Underline).Bool(8, run.Strikethrough));
                    }
                }
                // Keep the established text-only hash and native bytes unchanged.
                if (hasMixedDrawables) row.Message(3, P().Message(1, textBox));
                else row.Message(2, textBox);
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
