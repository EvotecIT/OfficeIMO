using System;

namespace OfficeIMO;

internal sealed record OfficeConversionSupportAssessment(
    OfficeConversionSupportLevel Level,
    string Evidence,
    string KnownLimitations);

internal static class OfficeConversionSupportAssessments {
    internal static OfficeConversionSupportAssessment Get(string routeId) {
        if (IsImageRoute(routeId, out string source)) return Image(source);
        return routeId switch {
        "docx-pdf" => Advanced(
            "Realistic DOCX fixtures cover paragraphs, lists, tables, drawings, pagination, tagged output, portable fonts, and deterministic conversion reports.",
            "Complex floating layout, advanced DrawingML, SmartArt, field behavior, and exact Microsoft Word pagination are not fully reproduced."),
        "xlsx-pdf" => Advanced(
            "Structured workbook fixtures cover tables, values, formulas, styles, print layout, pagination, tagged output, and deterministic diagnostics.",
            "Pivot tables, advanced conditional formatting, charts, external links, and exact Excel print-layout behavior remain incomplete."),
        "pptx-pdf" => Advanced(
            "Slide fixtures cover text, tables, basic shapes, images, charts, themes, page geometry, tagged output, and deterministic diagnostics.",
            "SmartArt, animation, media playback, advanced effects, uncommon DrawingML, and exact PowerPoint composition remain incomplete."),
        "pdf-png" => Advanced(
            "Pixel-pinned page-render tests cover text, drawings, tables, images, portable fonts, and deterministic multi-page packaging.",
            "Type 3 glyph programs, advanced color profiles, patterns, masks, transparency groups, optional content, and scanned-text OCR remain incomplete."),

        "html-pdf" => Established(
            "Representative HTML documents cover headings, lists, tables, links, images, pagination, tagged output, portable fonts, and conversion reports.",
            "The renderer supports bounded HTML and CSS, not a full browser layout engine; scripts, dynamic layout, and advanced CSS are intentionally unsupported."),
        "markdown-html" => Established(
            "Typed Markdown fixtures cover common block and inline constructs, safe profiles, deterministic HTML, and explicit diagnostics.",
            "Raw HTML, extension-specific syntax, and browser-dependent styling are governed by the selected safety and rendering profile."),
        "html-markdown" => Established(
            "Representative HTML fixtures cover portable text structure, lists, tables, links, resources, and deterministic loss reporting.",
            "CSS layout, scripting, arbitrary embedded content, and presentation-only markup are simplified or omitted."),
        "markdown-docx" => Established(
            "Representative Markdown fixtures reopen as editable DOCX and verify headings, paragraphs, lists, tables, links, and images with diagnostics.",
            "Raw HTML, extension-specific syntax, and constructs without a Word equivalent may be simplified."),
        "docx-html" or "docx-markdown" => Established(
            "Representative Word fixtures verify readable structure, tables, lists, links, images, and deterministic semantic output.",
            "Page layout, floating objects, advanced fields, tracked changes, and unsupported DrawingML are simplified in semantic output."),
        "docx-rtf" or "rtf-docx" or "rtf-markdown" or "rtf-pdf" or "markdown-rtf" or "rtf-html" => Established(
            "Bounded RTF fixtures verify parsing, editable or semantic output, reopen behavior, and deterministic loss diagnostics.",
            "Unsupported destinations, fields, embedded objects, complex tables, and producer-specific RTF controls are simplified or reported."),
        "xlsx-html" => Established(
            "Representative workbooks verify sheets, values, formulas, tables, and styles in deterministic review HTML.",
            "Interactive workbook behavior, charts, pivots, macros, and exact print layout are not represented in semantic HTML."),
        "pptx-html" => Established(
            "Representative presentations verify slide order, text, tables, images, and geometry hints in deterministic review HTML.",
            "The result is review HTML rather than a slide renderer; animations, media, SmartArt, and advanced effects are simplified."),
        "markdown-pdf" => Established(
            "Typed Markdown fixtures cover common structure, pagination, tagged PDF output, portable fonts, and conversion diagnostics.",
            "Raw HTML and extension-specific constructs are bounded by the Markdown profile and the managed PDF layout engine."),
        "html-docx" or "html-xlsx" or "html-pptx" or "html-rtf" => Established(
            "Representative bounded HTML fixtures reopen in the destination format and verify native text, tables, links, and deterministic diagnostics.",
            "CSS layout, scripts, complex web components, and constructs without a native destination equivalent are simplified or omitted."),
        "pdf-pptx" => Established(
            "Reopen tests verify explicit editable, tables-only, hybrid, and visual page modes with native objects, pixel-pinned visual fallback, and loss reports.",
            "Editable reconstruction is bounded to supported text, tables, basic shapes, and images; complex graphics, clipping, reading order, and effects remain incomplete."),

        "pdf-docx" => Targeted(
            "Reopen tests verify parser-supported text, headings, lists, and detected tables in an editable Word document with loss diagnostics.",
            "This is semantic reconstruction, not page-layout recovery; scans need OCR and complex graphics, columns, clipping, annotations, and exact pagination are not preserved."),
        "pdf-xlsx" => Targeted(
            "Reopen tests verify detected PDF tables as editable worksheet cells with source-page evidence and loss diagnostics.",
            "Only table-shaped content is in scope; arbitrary page layout, charts, free text, scans without OCR, and workbook formulas are not reconstructed."),
        "pdf-html" => Targeted(
            "Semantic and positioned-review profiles verify readable text, detected tables, page geometry hints, images, links, forms, and deterministic diagnostics.",
            "The result is review HTML, not a full PDF renderer; complex graphics, optional content, scans without OCR, and pixel-perfect composition are not preserved."),
        "pdf-rtf" or "pdf-odt" => Targeted(
            "Reopen tests verify parser-supported logical text and basic structure in an editable text-document destination with loss diagnostics.",
            "Page layout, complex graphics, columns, clipping, annotations, scans without OCR, and exact pagination are not reconstructed."),
        "pdf-ods" => Targeted(
            "Reopen tests verify detected tables as editable OpenDocument spreadsheet cells with source-page diagnostics.",
            "Only table-shaped content is in scope; arbitrary layout, charts, formulas, and scans without OCR are not reconstructed."),
        "pdf-odp" => Targeted(
            "Reopen tests verify bounded PDF page content in an editable OpenDocument presentation profile with explicit loss diagnostics.",
            "Complex graphics, clipping, reading order, effects, annotations, and exact visual reconstruction remain incomplete."),
        "docx-odt" or "odt-docx" => Targeted(
            "Reopen tests verify common paragraphs, runs, lists, tables, links, and basic images with explicit approximation diagnostics.",
            "Fields, notes, section-specific layout, advanced drawings, nested styles, and producer-specific extensions remain incomplete."),
        "xlsx-ods" or "ods-xlsx" => Targeted(
            "Reopen tests verify typed values, formulas, annotations, basic styles, tables, and validation subsets with diagnostics.",
            "Advanced validation, conditional formatting, charts, pivots, external links, macros, and complex styles remain incomplete."),
        "pptx-odp" or "odp-pptx" => Targeted(
            "Reopen tests verify common slide text, tables, basic shapes, images, and page geometry with explicit approximation diagnostics.",
            "Masters, advanced DrawingML, SmartArt, media, transitions, animations, and theme inheritance remain incomplete."),
        "asciidoc-markdown" or "markdown-asciidoc" or "asciidoc-pdf" => Targeted(
            "Typed fixtures verify the documented bounded AsciiDoc subset, deterministic output, and explicit diagnostics.",
            "Includes, extensions, macros, complex attributes, custom processors, and unsupported inline or block syntax are not a full AsciiDoc implementation."),
        "latex-markdown" or "markdown-latex" or "latex-pdf" => Targeted(
            "Typed fixtures verify the documented bounded LaTeX subset, deterministic output, and explicit diagnostics.",
            "Macro expansion, package execution, TeX layout, arbitrary mathematics, and unsupported environments are intentionally outside the managed subset."),
        "onenote-html" or "html-onenote" or "onenote-markdown" or "onenote-pdf" => Targeted(
            "Typed section-model fixtures verify common page text, hierarchy, tables, images, and deterministic projection diagnostics.",
            "The API targets the OfficeIMO section model; full .one binary fidelity, ink, embedded files, rich positioning, and OneNote application rendering remain incomplete."),
        "odt-pdf" or "ods-pdf" or "odp-pdf" => Targeted(
            "Representative OpenDocument fixtures verify the supported typed model, fixed-layout output, and source diagnostics.",
            "Inherited styles, advanced drawings, charts, media, nested inline syntax, and exact office-suite pagination remain incomplete."),
        "mhtml-pdf" => Targeted(
            "Bounded archive fixtures verify MIME resource resolution, safe HTML projection, PDF output, and deterministic diagnostics.",
            "Scripts, remote execution, browser-only layout, malformed multipart edge cases, and advanced CSS are not a full browser/MHTML engine."),
        "visio-pdf" => Targeted(
            "Typed drawing fixtures verify supported pages, shapes, connectors, text, geometry, and fixed-layout diagnostics.",
            "Advanced masters, data graphics, themes, layers, embedded objects, and exact Visio rendering remain incomplete."),
        "docx-google-docs" or "google-docs-docx" => Established(
            "Native request/import fixtures verify run typography, paragraphs, tables, tabs, remote operation policies, and Drive fallback behavior without requiring live credentials.",
            "Google Docs supports single underline and strike rather than Word decoration variants; all-caps is materialized and unsupported Word-only layout uses explicit diagnostics or Drive fallback."),
        "xlsx-google-sheets" or "google-sheets-xlsx" => Established(
            "Native request/import fixtures verify cell and rich-run typography, values, formulas, tables, and remote operation policies without requiring live credentials.",
            "Google Sheets has no subscript, superscript, casing, or underline-variant fields in its cell text-format model; unsupported Excel-only behavior is diagnosed or uses Drive fallback."),
        "pptx-google-slides" or "google-slides-pptx" => Established(
            "Native request/import fixtures verify per-run typography, geometry, shapes, tables, notes, remote operation policies, and Drive fallback behavior without requiring live credentials.",
            "Google Slides supports single underline and strike rather than DrawingML variants; all-caps is materialized and complex PowerPoint-only composition uses explicit diagnostics or raster/Drive fallback."),
        "adf-markdown" or "markdown-adf" or "adf-html" or "html-adf" => Established(
            "Typed ADF fixtures verify recursive structure, strong/emphasis/strike/underline/subscript/superscript marks, links, tables, extension preservation, HTML projection, and deterministic fidelity diagnostics.",
            "ADF text/background colors and arbitrary HTML styling are retained only when the intermediate Markdown profile can represent them; unsupported nodes and marks use explicit diagnostics or extension preservation."),
        "markdown-confluence" or "html-confluence" or "confluence-markdown" or "confluence-html" => Targeted(
            "Materialized Confluence page and body fixtures exercise ADF and storage representations through the shared ADF, Markdown, and HTML converters with combined diagnostics.",
            "This is content-body conversion, not live page synchronization; Confluence-only macros, layouts, extensions, and presentation outside the supported ADF/storage subset may be simplified."),
        "csv-xlsx" or "xlsx-csv" => Established(
            "Delimited-data fixtures cover streams, files, delimiter detection, typed cell import, worksheet-range export, reopen behavior, and cancellation.",
            "CSV and TSV have no font, rich-text, formula, multi-sheet, drawing, or layout model; those workbook semantics are intentionally absent from CSV output."),
        "officemarkup-docx" or "officemarkup-xlsx" or "officemarkup-pptx" => Established(
            "Typed profile fixtures verify parser diagnostics, editable destination artifacts, font family/size/color, bold, italic, underline variants, strike, scripts, case transforms, small caps where native, and target-specific layout.",
            "OfficeIMO Markup is a directed authoring format rather than a lossless Office round trip; unsupported blocks and destination-only effects are diagnosed, simplified, or omitted according to exporter options."),
            _ => throw new InvalidOperationException($"Conversion route '{routeId}' does not have a support assessment.")
        };
    }

    private static bool IsImageRoute(string routeId, out string source) {
        source = string.Empty;
        string[] suffixes = { "-png", "-svg", "-jpeg", "-tiff", "-webp" };
        foreach (string suffix in suffixes) {
            if (!routeId.EndsWith(suffix, StringComparison.Ordinal)) continue;
            source = routeId.Substring(0, routeId.Length - suffix.Length);
            return true;
        }
        return false;
    }

    private static OfficeConversionSupportAssessment Image(string source) => source switch {
        "pdf" => Advanced(
            "Pixel-pinned page-render tests cover text, drawings, tables, images, portable fonts, all five image encoders, and deterministic multi-page packaging.",
            "Image output is intentionally flattened; Type 3 glyph programs, advanced color profiles, patterns, masks, transparency groups, optional content, and scanned-text OCR remain incomplete."),
        "docx" or "xlsx" or "pptx" or "html" or "visio" => Established(
            "Representative source fixtures exercise shared SVG and raster rendering, text styles, page or surface selection, all five image encoders, and deterministic diagnostics.",
            "Image output is intentionally flattened. Advanced source-specific layout and effects outside the corresponding managed renderer remain approximated or diagnosed."),
        "onenote" or "email" or "epub" or "odt" or "ods" or "odp" => Targeted(
            "Typed fixtures exercise the source adapter, shared SVG and raster renderer, all five image encoders, and deterministic diagnostics.",
            "Image output is intentionally flattened. Unsupported source structure, remote or missing resources, and advanced source-specific layout remain approximated or diagnosed."),
        _ => throw new InvalidOperationException($"Image conversion source '{source}' does not have a support assessment.")
    };

    private static OfficeConversionSupportAssessment Targeted(string evidence, string limits) =>
        new(OfficeConversionSupportLevel.Targeted, evidence, limits);

    private static OfficeConversionSupportAssessment Established(string evidence, string limits) =>
        new(OfficeConversionSupportLevel.Established, evidence, limits);

    private static OfficeConversionSupportAssessment Advanced(string evidence, string limits) =>
        new(OfficeConversionSupportLevel.Advanced, evidence, limits);
}
