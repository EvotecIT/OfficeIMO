using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkLegacyXmlReader {
    private static readonly XNamespace Sf = IWorkLegacyXmlSource.Sf;
    private static readonly XNamespace Sfa = IWorkLegacyXmlSource.Sfa;
    private static readonly IWorkTextStyle PlainTextStyle = new(null, null, null, null, null, null, null, null, null);
    private static readonly IWorkParagraphStyle PlainParagraphStyle =
        new(null, null, null, null, null, null, null, null, null, null, PlainTextStyle);

    internal static IWorkPagesProjection ReadPages(IWorkSourceDocument source, IWorkLegacyXmlSource legacy) {
        XElement? body = legacy.Root.Element(Sf + "text-storage")?.Element(Sf + "text-body");
        var budget = new IWorkProjectionBudget(source.Options);
        IWorkTextContent content = ReadText(body, budget, source.CancellationToken);
        var diagnostics = new[] {
            LegacyLoss(legacy, "IWORK_LEGACY_PAGES_CONTENT_LIMITED",
                "The iWork '09 Pages profile recovers flowing body paragraphs. Styles, page layout, headers, "
                + "footers, and inline or positioned objects are outside the editable projection.")
        };
        return new IWorkPagesProjection(source, content, Array.Empty<IWorkPagesSection>(),
            Array.Empty<IWorkTextBox>(), Array.Empty<IWorkImageAsset>(), Array.Empty<IWorkTable>(),
            Array.Empty<IWorkPagesDrawable>(), null, diagnostics, supportsEditableReconstruction: false);
    }

    private static IWorkTextContent ReadText(XElement? body, IWorkProjectionBudget budget,
        CancellationToken cancellationToken) {
        var paragraphs = new List<IWorkTextParagraph>();
        if (body != null) ReadParagraphs(body, paragraphs, budget, cancellationToken);
        return new IWorkTextContent(paragraphs, isComplete: false, isTextComplete: false);
    }

    private static void ReadParagraphs(XElement parent, List<IWorkTextParagraph> paragraphs,
        IWorkProjectionBudget budget, CancellationToken cancellationToken) {
        foreach (XElement element in parent.Elements()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (element.Name == Sf + "section" || element.Name == Sf + "layout") {
                ReadParagraphs(element, paragraphs, budget, cancellationToken);
            } else if (element.Name == Sf + "p") {
                budget.AddTextItem();
                var runs = new List<IWorkTextRun>();
                ReadTextRuns(element, runs, budget, cancellationToken);
                paragraphs.Add(new IWorkTextParagraph(runs, PlainParagraphStyle, null, -1, null,
                    IWorkParagraphBreakKind.Paragraph));
            }
        }
    }

    private static void ReadTextRuns(XElement element, List<IWorkTextRun> runs, IWorkProjectionBudget budget,
        CancellationToken cancellationToken) {
        foreach (XNode node in element.Nodes()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (node is XElement child && (child.Name == Sf + "span" || child.Name == Sf + "link")) {
                ReadTextRuns(child, runs, budget, cancellationToken);
                continue;
            }
            string? value = node switch {
                XText text => text.Value,
                XElement control when control.Name.Namespace == Sf => control.Name.LocalName switch {
                    "tab" => "\t",
                    "lnbr" or "crbr" or "intratopicbr" => "\n",
                    _ => null
                },
                _ => null
            };
            if (value == null) continue;
            budget.AddTextItem();
            budget.AddTextCharacters(value.Length);
            runs.Add(new IWorkTextRun(value, PlainTextStyle, null));
        }
    }

    private static IWorkDiagnostic LegacyLoss(IWorkLegacyXmlSource legacy, string code, string message) =>
        new(IWorkDiagnosticSeverity.Warning, code, message, legacy.EntryPath);
}
