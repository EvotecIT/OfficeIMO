using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Snapshots and bounds the public Pages model before native encoding.</summary>
internal static class IWorkPagesPackageWriter {
    internal static byte[] Write(IWorkPagesDocument document, IWorkPagesWriteOptions limits, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document.BodyBlocks.Count == 0) throw new InvalidOperationException("A native Pages document requires at least one body block.");
        if (document.BodyBlocks.Count > limits.MaximumParagraphs) throw new InvalidDataException("The native Pages paragraph limit was exceeded.");
        IWorkPagesBodyBlock[] blocks = document.BodyBlocks.ToArray();
        long runs = 0, characters = 0, tabs = 0, tables = 0, cells = 0;
        long images = 0, encodedImageBytes = 0, decodedImageBytes = 0;
        var numberedLists = new Dictionary<IWorkPagesNumberedList, uint>();
        var model = new IWorkProtoWriter(limits.MaximumPackageBytes, cancellationToken)
            .Float(1, (float)document.PageSize.WidthPoints).Float(2, (float)document.PageSize.HeightPoints)
            .Float(3, document.LeftMarginPoints).Float(4, document.RightMarginPoints)
            .Float(5, document.TopMarginPoints).Float(6, document.BottomMarginPoints);
        foreach (IWorkPagesBodyBlock block in blocks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (block is IWorkPagesImage image) {
                if (++images > limits.MaximumImages) throw new InvalidDataException("The native Pages image limit was exceeded.");
                encodedImageBytes = checked(encodedImageBytes + image.Length);
                decodedImageBytes = checked(decodedImageBytes + image.Data.DecodedBytes);
                if (encodedImageBytes > limits.MaximumEncodedImageBytes) throw new InvalidDataException("The native Pages encoded-image limit was exceeded.");
                if (decodedImageBytes > limits.MaximumDecodedImageBytes) throw new InvalidDataException("The native Pages decoded-image limit was exceeded.");
                characters += 2L + (image.AccessibilityDescription?.Length ?? 0);
                if (++runs > limits.MaximumTextRuns) throw new InvalidDataException("The native Pages character-run limit was exceeded.");
                if (characters > limits.MaximumTextCharacters) throw new InvalidDataException("The native Pages text-character limit was exceeded.");
                var imageSignature = new IWorkProtoWriter(limits.MaximumPackageBytes, cancellationToken)
                    .Bytes(1, image.Data.ContentHash).Float(2, (float)image.Geometry.WidthPoints).Float(3, (float)image.Geometry.HeightPoints);
                if (image.AccessibilityDescription != null) imageSignature.String(4, image.AccessibilityDescription);
                model.Message(9, imageSignature); continue;
            }
            if (block is IWorkPagesTable table) {
                if (++tables > limits.MaximumTables) throw new InvalidDataException("The native Pages table limit was exceeded.");
                cells += (long)table.RowCount * table.ColumnCount;
                if (cells > limits.MaximumTableCells) throw new InvalidDataException("The native Pages table-cell limit was exceeded.");
                characters += table.Name.Length + 2; // Attachment marker and paragraph terminator.
                if (++runs > limits.MaximumTextRuns) throw new InvalidDataException("The native Pages character-run limit was exceeded.");
                foreach (IWorkTableCell cell in table.Cells) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (cell.Kind == IWorkCellKind.Text) characters += ((string)cell.Value!).Length;
                    if (characters > limits.MaximumTextCharacters) throw new InvalidDataException("The native Pages text-character limit was exceeded.");
                }
                if (characters > limits.MaximumTextCharacters) throw new InvalidDataException("The native Pages text-character limit was exceeded.");
                model.Message(8, table.Data.ModelFields(limits.MaximumPackageBytes, cancellationToken));
                continue;
            }
            var paragraph = (IWorkPagesParagraphData)block;
            runs += paragraph.Runs.Count;
            characters += (long)paragraph.Text.Length + 1;
            tabs += paragraph.Layout.TabStops.Count;
            if (runs > limits.MaximumTextRuns) throw new InvalidDataException("The native Pages character-run limit was exceeded.");
            if (characters > limits.MaximumTextCharacters) throw new InvalidDataException("The native Pages text-character limit was exceeded.");
            if (tabs > limits.MaximumTabStops) throw new InvalidDataException("The native Pages custom-tab limit was exceeded.");
            var data = new IWorkProtoWriter(limits.MaximumPackageBytes, cancellationToken);
            foreach (IWorkTextRunData run in paragraph.Runs) {
                cancellationToken.ThrowIfCancellationRequested();
                data.Message(1, new IWorkProtoWriter(limits.MaximumPackageBytes, cancellationToken)
                    .String(1, run.Text).String(2, run.FontName).Float(3, run.FontSizePoints).String(4, run.Color.RgbHex)
                    .Bool(5, run.Bold).Bool(6, run.Italic).Bool(7, run.Underline).Bool(8, run.Strikethrough));
            }
            if (paragraph.NumberedList is { } numberedList) {
                if (!numberedLists.TryGetValue(numberedList, out uint listIndex)) {
                    if (numberedLists.Count >= limits.MaximumNumberedLists)
                        throw new InvalidDataException("The native Pages numbered-list limit was exceeded.");
                    listIndex = checked((uint)numberedLists.Count + 1); numberedLists.Add(numberedList, listIndex);
                }
                var numbering = new IWorkProtoWriter(limits.MaximumPackageBytes, cancellationToken)
                    .UInt(1, listIndex).UInt(2, (uint)paragraph.ListLevel).UInt(3, (uint)paragraph.ListNumber!.Value);
                if (paragraph.ListRestartNumber.HasValue) numbering.UInt(4, (uint)paragraph.ListRestartNumber.Value);
                data.Message(3, numbering);
            }
            if (paragraph.Layout.HasCustomProperties) {
                var layout = new IWorkProtoWriter(limits.MaximumPackageBytes, cancellationToken)
                    .Float(1, paragraph.Layout.LineSpacingMultiplier).Float(2, paragraph.Layout.SpaceBeforePoints)
                    .Float(3, paragraph.Layout.SpaceAfterPoints);
                foreach (IWorkTabStop stop in paragraph.Layout.TabStops) {
                    cancellationToken.ThrowIfCancellationRequested();
                    layout.Message(4, new IWorkProtoWriter(limits.MaximumPackageBytes, cancellationToken)
                        .Float(1, (float)stop.PositionPoints).UInt(2, (uint)stop.Alignment));
                }
                data.Message(2, layout);
            }
            model.Message(7, data);
        }
        using var sha = SHA256.Create();
        var builder = new IWorkPagesArchiveBuilder(limits.MaximumPackageBytes, cancellationToken, sha.ComputeHash(model.ToArray()));
        var archives = builder.Build(document, blocks);
        return IWorkNativePackageWriter.Write(archives.Document, archives.Metadata, builder.DocumentIdentifier,
            limits.MaximumPackageBytes, cancellationToken, builder.ImageResources);
    }
}
