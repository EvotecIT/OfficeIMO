using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Tests;

/// <summary>
/// Shared test-owned HTML rendering corpus used by HTML image and PDF artifact gates.
/// </summary>
internal static class HtmlRenderingRepresentativeCorpus {
    internal const string RelativeRoot = "OfficeIMO.TestAssets/Documents/Html/Qualification/H4/representative";

    internal static IReadOnlyList<HtmlRenderingCorpusCase> All { get; } = new[] {
        new HtmlRenderingCorpusCase(
            "invoice",
            HtmlRenderMode.Continuous,
            LoadHtml("invoice"),
            new[] { "Invoice", "INV-", "1042", "Ada Lovelace", "$702.00", "PDF fidelity review" },
            linkUri: "https://example.test/invoices/1042",
            minimumVisualCount: 20,
            minimumHeadingCount: 3,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "account-statement",
            HtmlRenderMode.Paged,
            LoadHtml("account-statement"),
            new[] { "Account Statement", "Closing balance 975.00", "Cloud subscription", "Account notes" },
            expectedPageCount: 2,
            minimumVisualCount: 24,
            minimumHeadingCount: 3,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "quarterly-report",
            HtmlRenderMode.Continuous,
            LoadHtml("quarterly-report"),
            new[] { "Quarterly Report", "Revenue increased", "All regions", "Apply filters", "72%", "76%", "Asia Pacific" },
            minimumVisualCount: 40,
            minimumHeadingCount: 5,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending", "GridLayoutPending" },
            requiredVisualSources: new[] { "select#report-region", "input#report-owner", "button#apply-filter", "progress#margin-progress:value", "meter#pipeline-meter:value" }),
        new HtmlRenderingCorpusCase(
            "business-letter",
            HtmlRenderMode.Paged,
            LoadHtml("business-letter"),
            new[] { "Project Confirmation", "Dear Grace", "Atlas delivery schedule", "PowerPoint", "Compiler House" },
            linkUri: "mailto:office@example.test",
            minimumVisualCount: 18,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "certificate",
            HtmlRenderMode.Paged,
            LoadHtml("certificate"),
            new[] { "Certificate of Completion", "Ada Lovelace", "OfficeIMO document fidelity program", "Program Director" },
            minimumVisualCount: 18,
            minimumHeadingCount: 2,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "product-catalog",
            HtmlRenderMode.Continuous,
            LoadHtml("product-catalog"),
            new[] { "Product Catalog", "Atlas", "Visual reporting", "PDF/A checks", "Visual gates" },
            linkUri: "https://example.test/products/atlas",
            minimumVisualCount: 20,
            minimumHeadingCount: 5,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending", "PositionedLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "legal-contract",
            HtmlRenderMode.Paged,
            LoadHtml("legal-contract"),
            new[] { "Services Agreement", "Deliver document tooling", "Twelve months", "Ten business days", "Provider signature" },
            linkUri: "https://example.test/terms",
            minimumVisualCount: 26,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "email-render",
            HtmlRenderMode.Continuous,
            LoadHtml("email-render"),
            new[] { "Action Required", "Review the attached status update", "2 checks need attention", "Blocked status image" },
            linkUri: "https://example.test/action",
            diagnosticCodes: new[] { "ImageResourceRejectedByPolicy", "HyperlinkRejectedByPolicy" },
            minimumVisualCount: 20),
        new HtmlRenderingCorpusCase(
            "dashboard-print",
            HtmlRenderMode.Continuous,
            LoadHtml("dashboard-print"),
            new[] { "Operations Dashboard", "99.95%", "Alerts", "Documents processed", "PowerPoint 17%", "Open incident" },
            minimumVisualCount: 38,
            minimumHeadingCount: 5,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending", "GridLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "multilingual-bidi",
            HtmlRenderMode.Continuous,
            LoadHtml("multilingual-bidi"),
            new[] { "Multilingual Summary", "שלום 123", "سلام 456", "Zażółć", "Příliš žluťoučký", "Έγγραφο" },
            minimumVisualCount: 28,
            minimumHeadingCount: 3,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "static-standards-showcase",
            HtmlRenderMode.Paged,
            LoadHtml("static-standards-showcase"),
            new[] { "Static standards packet", "Inherited row A", "Clipped vector badge", "Second-page evidence", "artifact tagging" },
            expectedPageCount: 2,
            minimumVisualCount: 2,
            minimumHeadingCount: 2,
            forbiddenDiagnosticCodes: new[] { "HtmlRenderGridLayoutPending", "HtmlRenderClipPathValueUnsupported", "HtmlRenderPositioningModeUnsupported" },
            requireNoLoss: true),
        new HtmlRenderingCorpusCase(
            "book-extract",
            HtmlRenderMode.Paged,
            LoadHtml("book-extract"),
            new[] { "The Managed Document Handbook", "Chapter 2 · Fragmentation", "Handbook note", "Appendix marker: three explicit rendering intents" },
            expectedPageCount: 3,
            minimumVisualCount: 4,
            minimumHeadingCount: 4,
            forbiddenDiagnosticCodes: new[] { "HtmlRenderColumnsLayoutPending", "HtmlRenderFootnoteLayoutPending", "HtmlRenderNamedPageUnsupported" }),
        new HtmlRenderingCorpusCase(
            "application-form",
            HtmlRenderMode.Continuous,
            LoadHtml("application-form"),
            new[] { "Equipment Access Request", "Ada Lovelace", "Document systems", "Submit request", "Read the access policy" },
            linkUri: "https://example.test/access-policy",
            minimumVisualCount: 35,
            minimumHeadingCount: 1,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending" },
            requiredVisualSources: new[] { "input#requester-name", "select#department", "textarea#reason", "button#submit-request" }),
        new HtmlRenderingCorpusCase(
            "chart-report",
            HtmlRenderMode.Continuous,
            LoadHtml("chart-report"),
            new[] { "Capacity and Quality Review", "18,420", "Weekly throughput", "Completed documents", "On target" },
            minimumVisualCount: 38,
            minimumHeadingCount: 2,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending", "SvgContentUnsupported", "SvgRasterFallback" }),
        new HtmlRenderingCorpusCase(
            "browser-local-workbench",
            HtmlRenderMode.Continuous,
            LoadHtml("browser-local-workbench"),
            new[] { "Document Workbench", "Local conversion", "Browser-local", "Screen-to-PDF", "policy-bundle.zip" },
            minimumVisualCount: 38,
            minimumHeadingCount: 2,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending", "FlexLayoutPending", "PositionedLayoutPending" },
            requiredVisualSources: new[] { "button", "a" })
    };

    internal static string ResolveSourcePath(string repository, string id) =>
        Path.Combine(repository, RelativeSourcePath(id).Replace('/', Path.DirectorySeparatorChar));

    private static string LoadHtml(string id) {
        string relative = RelativeSourcePath(id).Replace('/', Path.DirectorySeparatorChar);
        string repositoryCandidate = Path.Combine(Environment.CurrentDirectory, relative);
        if (File.Exists(repositoryCandidate)) return File.ReadAllText(repositoryCandidate);

        string outputCandidate = Path.Combine(
            AppContext.BaseDirectory,
            "Documents",
            "Html",
            "Qualification",
            "H4",
            "representative",
            id + ".html");
        if (File.Exists(outputCandidate)) return File.ReadAllText(outputCandidate);

        throw new FileNotFoundException("Cannot locate the versioned H4 HTML rendering corpus source.", relative);
    }

    private static string RelativeSourcePath(string id) {
        if (string.IsNullOrEmpty(id) || id.Any(character =>
            !(character >= 'a' && character <= 'z')
            && !(character >= 'A' && character <= 'Z')
            && !(character >= '0' && character <= '9')
            && character != '-')) {
            throw new ArgumentException("Corpus identifiers may contain only ASCII letters, digits, and hyphens.", nameof(id));
        }
        return RelativeRoot + "/" + id + ".html";
    }
}

/// <summary>
/// Defines one stable HTML source and its observable cross-backend contract.
/// </summary>
internal sealed class HtmlRenderingCorpusCase {
    internal HtmlRenderingCorpusCase(
        string id,
        HtmlRenderMode mode,
        string html,
        IReadOnlyList<string> textMarkers,
        int expectedPageCount = 1,
        string? linkUri = null,
        IReadOnlyList<string>? diagnosticCodes = null,
        int minimumVisualCount = 2,
        int minimumHeadingCount = 1,
        IReadOnlyList<string>? forbiddenDiagnosticCodes = null,
        IReadOnlyList<string>? requiredVisualSources = null,
        bool requireNoLoss = false) {
        Id = id;
        Mode = mode;
        Html = html;
        TextMarkers = textMarkers;
        ExpectedPageCount = expectedPageCount;
        LinkUri = linkUri;
        DiagnosticCodes = diagnosticCodes ?? Array.Empty<string>();
        MinimumVisualCount = minimumVisualCount;
        MinimumHeadingCount = minimumHeadingCount;
        ForbiddenDiagnosticCodes = forbiddenDiagnosticCodes ?? Array.Empty<string>();
        RequiredVisualSources = requiredVisualSources ?? Array.Empty<string>();
        RequireNoLoss = requireNoLoss;
    }

    internal string Id { get; }
    internal string SourceRelativePath => HtmlRenderingRepresentativeCorpus.RelativeRoot + "/" + Id + ".html";
    internal HtmlRenderMode Mode { get; }
    internal string Html { get; }
    internal IReadOnlyList<string> TextMarkers { get; }
    internal int ExpectedPageCount { get; }
    internal string? LinkUri { get; }
    internal IReadOnlyList<string> DiagnosticCodes { get; }
    internal int MinimumVisualCount { get; }
    internal int MinimumHeadingCount { get; }
    internal IReadOnlyList<string> ForbiddenDiagnosticCodes { get; }
    internal IReadOnlyList<string> RequiredVisualSources { get; }
    internal bool RequireNoLoss { get; }
    internal double ExpectedSurfaceWidth => Mode == HtmlRenderMode.Paged
        ? 8.27D * HtmlRenderOptions.CssPixelsPerInch
        : Id switch {
            "application-form" => 680D,
            "browser-local-workbench" => 680D,
            "chart-report" => 652D,
            _ => 640D
        };

    internal HtmlRenderOptions CreateOptions() => new HtmlRenderOptions {
        Mode = Mode,
        ViewportWidth = 640D,
        PageSize = new OfficePageSize(8.27D, 11.69D),
        Margins = HtmlRenderMargins.All(40D),
        Scale = 0.5D,
        BackgroundColor = OfficeColor.White,
        UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
        FidelityPolicy = RequireNoLoss ? HtmlRenderFidelityPolicy.RequireNoLoss : HtmlRenderFidelityPolicy.AllowDiagnosedLoss
    };
}
