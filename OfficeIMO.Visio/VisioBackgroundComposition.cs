using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Resolves cached page associations in paint order without modifying malformed source references.</summary>
internal static class VisioBackgroundComposition {
    internal static IReadOnlyList<VisioPage> Resolve(VisioPage page, CancellationToken token,
        ICollection<OfficeImageExportDiagnostic>? diagnostics, string? source, int maximumPages = int.MaxValue) {
        var pages = new List<VisioPage>();
        var visited = new HashSet<VisioPage>();
        for (VisioPage? current = page; current != null; current = current.BackgroundPage) {
            token.ThrowIfCancellationRequested();
            if (!visited.Add(current)) {
                AddLoss("VISIO_BACKGROUND_CYCLE", "A cyclic background reference is omitted; each reachable page is painted once.", current);
                break;
            }
            if (pages.Count >= maximumPages) throw new InvalidDataException("Visio background composition exceeds MaximumPages.");
            pages.Add(current);
            if (current.BackgroundPage == null && current.BackgroundPageId.HasValue) {
                if (current.BackgroundPageId.Value == current.Id)
                    AddLoss("VISIO_BACKGROUND_CYCLE", "A self-referencing background association is omitted.", current);
                else AddLoss("VISIO_BACKGROUND_MISSING", "The referenced background page is unavailable and cannot be composed.", current);
            }
        }
        pages.Reverse();
        return pages;

        void AddLoss(string code, string message, VisioPage current) => diagnostics?.Add(
            new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning, code, message,
                (source ?? page.Name) + ":background-reference:" + current.Id, OfficeConversionLossKind.Omission));
    }
}
