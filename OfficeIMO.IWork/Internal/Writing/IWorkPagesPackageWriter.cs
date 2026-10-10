using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Snapshots and bounds the public Pages model before native encoding.</summary>
internal static class IWorkPagesPackageWriter {
    internal static byte[] Write(IWorkPagesDocument document, IWorkPagesWriteOptions limits, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document.Paragraphs.Count == 0) throw new InvalidOperationException("A native Pages document requires at least one paragraph.");
        if (document.Paragraphs.Count > limits.MaximumParagraphs) throw new InvalidDataException("The native Pages paragraph limit was exceeded.");
        IWorkPagesParagraphData[] paragraphs = document.Paragraphs.ToArray();
        long runs = 0, characters = 0, tabs = 0;
        var model = new IWorkProtoWriter(limits.MaximumPackageBytes, cancellationToken)
            .Float(1, (float)document.PageSize.WidthPoints).Float(2, (float)document.PageSize.HeightPoints)
            .Float(3, document.LeftMarginPoints).Float(4, document.RightMarginPoints)
            .Float(5, document.TopMarginPoints).Float(6, document.BottomMarginPoints);
        foreach (IWorkPagesParagraphData paragraph in paragraphs) {
            cancellationToken.ThrowIfCancellationRequested();
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
        var archives = builder.Build(document, paragraphs);
        return IWorkNativePackageWriter.Write(archives.Document, archives.Metadata, builder.DocumentIdentifier,
            limits.MaximumPackageBytes, cancellationToken);
    }
}
