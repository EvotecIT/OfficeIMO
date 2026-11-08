using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Word.Html;

internal partial class HtmlToWordConverter {
    private readonly List<(string Css, string? Media)> _pageStylesheets = new();

    /// <summary>Sets the new document's shared print geometry before content layout; insertion keeps its existing sections.</summary>
    private void ApplyPageDefaults(WordDocument document, IHtmlDocument source, HtmlWordExportOptions options) {
        if (options.Theme == null && _pageStylesheets.Count == 0) return;
        WordPageSetup setup = document.PageSetup;
        WordMargins margins = setup.Margins;
        OfficeThemePage? themePage = options.Theme?.Page;
        var renderOptions = new HtmlRenderOptions {
            PageSize = new OfficePageSize(setup.Size.Width ?? WordPageSizes.A4.Width, setup.Size.Height ?? WordPageSizes.A4.Height),
            Margins = HtmlRenderMargins.FromCssPageRule(
                Pixels(themePage?.MarginLeft ?? margins.Left ?? OfficeLength.Inches(1)),
                Pixels(themePage?.MarginTop ?? margins.Top ?? OfficeLength.Inches(1)),
                Pixels(themePage?.MarginRight ?? margins.Right ?? OfficeLength.Inches(1)),
                Pixels(themePage?.MarginBottom ?? margins.Bottom ?? OfficeLength.Inches(1))),
            DefaultFontSize = _rootFontSizePixels
        };
        // Reuse accepted source text in the importer's stylesheet order. No resource is fetched again.
        var pageDocument = (IHtmlDocument)source.Implementation.CreateHtmlDocument(string.Empty);
        foreach (var stylesheet in _pageStylesheets) {
            _cancellationToken.ThrowIfCancellationRequested();
            var style = pageDocument.CreateElement("style");
            style.TextContent = stylesheet.Css;
            if (!string.IsNullOrWhiteSpace(stylesheet.Media)) style.SetAttribute("media", stylesheet.Media);
            pageDocument.Head!.AppendChild(style);
        }

        HtmlCssPageRuleSet pageRules;
        try {
            pageRules = HtmlCssPageSettingsResolver.Apply(pageDocument, renderOptions, options.ConversionReport, options.Limits, _cancellationToken);
        } catch (HtmlDomLimitException exception) {
            ThrowLimitExceeded(options, exception.Code, exception.Message, exception.LimitSource, exception.Actual, exception.Limit);
            throw;
        }
        if (pageRules.RuleCount == 0 && themePage == null) return;
        if (pageRules.HasPageSpecificRules) {
            AddDiagnostic(options, "html-page-specific-rules-omitted", "Named and page-specific CSS page rules are not mapped to Word sections.", "@page");
        }
        if (pageRules.HasMarginBoxRules) {
            AddDiagnostic(options, "html-page-margin-boxes-omitted", "CSS page margin boxes are not mapped to Word headers and footers.", "@page");
        }
        HtmlCssPageGeometry geometry = pageRules.ResolveGenericGeometry(renderOptions);
        if (geometry.PrintProduction != null) {
            AddDiagnostic(options, "html-page-production-omitted", "CSS bleed and printer marks are omitted; Word uses the trim page geometry.", "@page");
        }
        if (geometry.BackgroundColor != null) {
            AddDiagnostic(options, "html-page-background-omitted", "CSS page background paint is not mapped to the Word page.", "@page");
        }
        double inset = geometry.PrintProduction?.TrimInset ?? 0D;
        var size = new OfficePageSize(Length(geometry.Width - inset * 2D), Length(geometry.Height - inset * 2D));
        if (size.Width != setup.Size.Width || size.Height != setup.Size.Height) {
            setup.Size.SetSize(size, size.Width > size.Height ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait);
        }
        SetMargin(geometry.Margins.Top - inset, margins.Top, value => margins.Top = value);
        SetMargin(geometry.Margins.Bottom - inset, margins.Bottom, value => margins.Bottom = value);
        SetMargin(geometry.Margins.Left - inset, margins.Left, value => margins.Left = value, horizontal: true);
        SetMargin(geometry.Margins.Right - inset, margins.Right, value => margins.Right = value, horizontal: true);

        void SetMargin(double pixels, OfficeLength? current, Action<OfficeLength> assign, bool horizontal = false) {
            if (horizontal && pixels < 0D) {
                AddDiagnostic(options, "html-negative-horizontal-page-margin-omitted", "Negative horizontal CSS page margins cannot be represented as Word section margins.", "@page");
                return;
            }
            OfficeLength value = Length(pixels);
            if (value != current) assign(value);
        }
    }

    private static double Pixels(OfficeLength value) => value.ToInches() * HtmlRenderOptions.CssPixelsPerInch;
    private static OfficeLength Length(double pixels) => OfficeLength.Inches(pixels / HtmlRenderOptions.CssPixelsPerInch);
}
