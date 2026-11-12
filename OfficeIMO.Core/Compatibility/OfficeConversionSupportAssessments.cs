namespace OfficeIMO;

internal sealed record OfficeConversionSupportAssessment(
    OfficeConversionSupportLevel Level,
    string Evidence,
    string KnownLimitations);

internal static class OfficeConversionSupportAssessments {
    internal static OfficeConversionSupportAssessment Get(string routeId) {
        if (routeId.StartsWith("native-", StringComparison.Ordinal)) return Targeted(
            "Native format conversion fixtures exercise package types, source carriers, compatibility modes, encryption, loss rejection and reopened artifacts through the owning format codecs.",
            "The route is not a promise that every source feature can be written to every destination generation. Unsupported native structures and formats are rejected or reported according to explicit compatibility and loss policy.");
        if (routeId == "word-xml-docx") return Targeted(
            "Word XML fixture controls cover standalone package admission, editable content and reopened DOCX.",
            "DOCX output accepts ordinary Word documents. Macro-enabled and template Flat OPC packages require a matching native package destination through the Word API. Word 2003 XML import reports unsupported features.");
        if (routeId == "spreadsheet-xml-xlsx") return Targeted(
            "Native Excel XML Spreadsheet fixtures cover typed cells, styles, geometry and cached formula results in the supported profile.",
            "Uses the bounded Excel XML Spreadsheet 2003 codec. Unsupported workbook features are reported and formula execution is outside import.");
        if (routeId is "xls-pdf" or "xlsb-pdf" or "ppt-pdf") return Targeted(
            "Native binary import and existing PDF adapter fixtures cover recovered editable content and separately reported rendering findings.",
            "Binary source projection is bounded and incomplete. Unsupported source records, preserved-only content and PDF layout approximations remain explicit findings; native application visual equivalence is not promised.");
        if (routeId == "email-pdf") return Targeted(
            "Converter and host fixtures reopen the PDF and find the subject, sender, To and Cc recipients, date, HTML body and attachment list; the TextOnly profile keeps the Reader projection and its email policy finding.",
            "Renders through the HTML-to-PDF engine with the shared email body sanitization, blocked remote resources and bounded inline MIME images. RTF-only bodies are approximated through HTML. Attachment payloads are listed, not embedded or rendered; signed and encrypted payloads are not verified or decrypted. Mail-client visual equivalence is not promised.");
        if (routeId is "epub-pdf" or "reader-visio-pdf") return Targeted(
            "Reader projection fixtures cover searchable email text, EPUB chapter order, Visio semantic fallback and policy findings in reopened PDF output.",
            "Source-neutral semantic projection only. Attachments, pagination, resources, forms and source layout follow explicit projection policies and diagnostics; this route does not promise native visual equivalence.");
        if (routeId is "pages-pdf" or "numbers-pdf" or "keynote-pdf") return Targeted(
            "Composed workflow contracts retain complete editable iWork reconstruction and PDF-stage evidence, reopen the staged PDF, and enforce source identity, cancellation, and output limits.",
            "Requires the complete editable subset. Partial reconstructions, embedded previews, selection/clipping, and PDF-stage content omissions are rejected. Pagination, fonts, character spacing, and other rendering approximations remain reported; Apple visual equivalence is fixture-specific.");
        if (routeId is "visio-binary-vsdx" or "visio-binary-vssx" or "visio-binary-vstx" or "visio-binary-svg" or "visio-binary-pdf")
            return Targeted(
                "Independent Microsoft Visio version 11 drawings, stencils and templates are compared with paired native XML for shape hierarchy, masters, transforms and cached text; modern packages reopen and SVG/PDF output is inspected.",
                "Cached reconstruction only. Earlier binary generations, native binary writing, formulas, active content, unsupported geometry, custom data and full native visual fidelity are outside this profile. Import and subsequent projection have separate loss reports." +
                (routeId == "visio-binary-pdf" ? " PDF output uses the same vector diagram pages as VSDX." : string.Empty) +
                " Native Microsoft Visio acceptance remains unqualified.");
        if (IsImageRoute(routeId, out string source)) return Image(source);
        if (routeId.StartsWith("email-", StringComparison.Ordinal)) {
            return Established(
                "Cross-format EML, MSG, OFT, and TNEF fixtures verify typed message content, attachments, source selection, diagnostics, and explicit loss policy.",
                "Signed or encrypted messages, opaque MAPI properties, recurrence and time-zone blobs, and Outlook-only item semantics can require preservation, approximation, or an explicitly accepted loss policy.");
        }
        if (routeId.StartsWith("bibliography-", StringComparison.Ordinal)) {
            return Established(
                "Codec fixtures verify deterministic BibTeX, BibLaTeX, CSL JSON, RIS, NBIB/MEDLINE, and EndNote XML writes, reopen behavior, native-field handling, and conversion diagnostics.",
                "Destination codecs cannot represent every source-native field or structure; permissive writes report approximated or omitted data and ThrowIfLoss rejects known loss.");
        }
        if (routeId is "pages-docx" or "numbers-xlsx" or "keynote-pptx") {
            return Targeted(
                "Pinned Pages, Numbers, and Keynote fixtures verify bounded package inspection, typed projection, editable Office reconstruction, visual fallbacks, preserved records, and conversion reports.",
                "Modern IWA support is a bounded read-only source adapter; unsupported records, advanced layout, media, formulas, and producer-specific effects may require visual fallbacks or remain diagnostic evidence.");
        }
        return routeId switch {
        "chm-pdf" or "chm-markdown" or "chm-epub" => Targeted(
            "A Microsoft-compiled book matches an independent decoder byte-for-byte across every entry. Contract fixtures cover contents, index, encodings, bounded reads, topic links, rich Reader results and reopened PDF/EPUB output.",
            "ITSF 2/3 with uncompressed or standard LZX storage. ActiveX, scripts, merged-help books, the compiled search engine and Windows viewer behavior are not reproduced. Reflow can change CSS scope; multi-topic PDF is untagged and does not rebuild cross-topic links or the keyword index. See OfficeIMO.Chm/SUPPORT.md."),
        "book-project-epub" => Targeted(
            "Book-project batch contracts verify review gating, EPUB reopen, output limits, conflict handling and checkpoint reuse through the shared workflow runner.",
            "Exports the current publication only; project revisions and review records remain in the source project. Import failures block export, and non-fatal import losses require saved author acknowledgment. Reader presentation and comprehensive accessibility qualification remain separate."),
        "html-epub" or "markdown-epub" or "docx-epub" => Targeted(
            "Manuscript contracts cover chapter splitting, navigation, semantic structures, resource collection, import reports and re-opened EPUB output; representative artifacts pass independent EPUBCheck validation.",
            "Reflowable EPUB 3 only. Source layout, live fields, revisions, comments and controls can require reported approximations or omissions. Missing assets and unnamed images block export; schema conformance, accessibility metadata and reader presentation require independent review. DOCX and Markdown composition uses OfficeIMO.Workflows on .NET 8 or newer."),
        "doc-pdf" => Targeted(
            "Legacy DOC import and first-party PDF rendering retain separate fidelity reports; known import loss blocks conversion by default.",
            "Only the supported Word 97-2003 import subset is projected. Preserved-only graphics and richer structures cannot be rendered; Word pagination may differ."),
        "txt-pdf" => Targeted(
            "Literal text conversion uses the canonical PDF flow engine, strict decoding, source whitespace and bounded pagination.",
            "Long lines wrap to page width. Unicode appearance depends on supplied fonts; no Markdown, HTML or ANSI interpretation occurs."),
        "djvu-pdf" => Targeted(
            "Archival text and geometry, independent scanned-page comparisons, bounded malformed-input tests, and searchable PDF/OCR artifacts qualify the managed profile.",
            "JB2, MMR, IW44 1.2, palettes and managed JPEG profiles only. Annotations are reported without active rendering; indirect components require an explicit native resolver. OCR is caller-supplied and absent/empty-text-only. See OfficeIMO.DjVu/SUPPORT.md."),
        "xps-pdf" => Targeted(
            "Native package, drawing and PDF contracts cover both dialects, positioned Unicode, navigation and authored logical structure; representative artifacts are checked by independent PDF readers and renderers.",
            "Strict export rejects unsupported paint and semantic reconstruction. Search requires literal UnicodeString; glyph IDs, figure descriptions and PDF/UA conformance are not inferred. Native print tickets, metadata and signatures are not PDF preservation contracts."),
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
        "markdown-pptx" => Targeted(
            "Editable slide fixtures verify heading and thematic-break boundaries, front matter, speaker notes, slide dimensions, image policy, and conversion diagnostics in reopened PowerPoint output.",
            "Markdown is projected through the supported OfficeIMO Markup presentation profile. Unsupported blocks and styling are simplified or diagnosed; automatic layout, pagination, and native PowerPoint visual equivalence are not promised."),
        "docx-html" or "docx-markdown" => Established(
            "Representative Word fixtures verify readable structure, tables, lists, links, images, and deterministic semantic output.",
            "Page layout, floating objects, advanced fields, tracked changes, and unsupported DrawingML are simplified in semantic output."),
        "docx-rtf" or "rtf-docx" or "rtf-markdown" or "rtf-pdf" or "markdown-rtf" or "rtf-html" => Established(
            "Bounded RTF fixtures verify parsing, editable or semantic output, reopen behavior, and deterministic loss diagnostics.",
            "Unsupported destinations, fields, embedded objects, complex tables, and producer-specific RTF controls are simplified or reported."),
        "xlsx-html" => Established(
            "Representative workbooks verify sheets, values, formulas, tables, styles, and semantic chart data in deterministic review HTML.",
            "Interactive workbook behavior, chart rendering and unsupported chart styling, pivots, macros, and exact print layout are not represented in semantic HTML."),
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
            "Reopen tests verify editable text, headings, lists, and tables, plus visual page images with selected page order and physical sizes.",
            "Editable mode reconstructs supported logical content; scans need OCR and exact layout is not preserved. Visual mode embeds page images within the managed renderer's support and resource limits; text and interactive objects are not editable."),
        "pdf-xlsx" => Targeted(
            "Reopen tests verify detected PDF tables as editable worksheet cells with source-page evidence and loss diagnostics.",
            "Only table-shaped content is in scope; arbitrary page layout, charts, free text, scans without OCR, and workbook formulas are not reconstructed."),
        "pdf-html" => Targeted(
            "Semantic and positioned-review profiles verify readable text, detected tables, page geometry hints, images, links, forms, and deterministic diagnostics.",
            "The result is review HTML, not a full PDF renderer; complex graphics, optional content, scans without OCR, and pixel-perfect composition are not preserved."),
        "pdf-markdown" => Targeted(
            "Native logical readback fixtures verify Markdown headings, paragraphs, tables, selected page order, stream output, and configurable page separators.",
            "Reading order and structure are reconstructed from supported PDF content. Fixed page layout, graphics, interactive behavior, and scans without OCR are not preserved; the canonical conversion reports layout approximation."),
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
            "Converter contracts reopen headings, paragraph and run formatting, lists, tables and merges, inline images, page layout and breaks, bookmarks, default, first-page and even-page headers and footers, footnotes and endnotes, and PAGE, NUMPAGES, DATE and TIME fields. LibreOffice Writer and Microsoft Word ODT exports convert with body order and tables intact (OpenDocumentIndependentProducerContracts).",
            "Headers and footers come from the first section and first master page only. Tracked changes, section-specific layout, page-number restarts, header/footer tables, charts, content controls, other field instructions, note numbering and placement, and nested inline markup are approximated or reported as loss. The independent producer fixtures are minimal, so complex LibreOffice and Word documents remain unqualified."),
        "xlsx-ods" or "ods-xlsx" => Established(
            "LibreOffice Calc and Microsoft Excel ODS exports, Excel-produced chart and pivot workbooks, and converter contracts verify values, formulas, annotations, styles, validations, conditional formatting, charts, pivot and DataPilot tables, ODS print ranges and strict loss reports in reopened workbooks (OpenDocumentIndependentProducerContracts, SpreadsheetChartConversionTests, ExcelOdsDataPilotConversionTests).",
            "Conditional formatting covers up to 16 numeric cell-value rules per style with direct fills or fonts. Charts cover column, bar, line, pie and doughnut types with copied values, title and approximate placement. Pivots need one standard aggregate and a same-sheet source and output. Borders, value-dependent alignment, blank cells styled only by defaults, advanced validation, external links, macros, and Excel print-area names written to ODS remain explicit loss. LibreOffice and Excel files also report their page masters and the text formatting on row padding (blank-cell-tail-styles), so ThrowOnSkippedOrUnsupported and FailOnLoss reject ordinary producer files."),
        "pptx-odp" or "odp-pptx" => Targeted(
            "Converter contracts reopen slide size and order, hidden slides, text boxes and runs, hyperlinks, images and crops, tables and merges, basic shapes, solid and master backgrounds, common transitions, plain speaker notes, master and layout associations and placeholder roles. LibreOffice Impress and Microsoft PowerPoint ODP exports keep slide geometry and placeholders (OpenDocumentIndependentProducerContracts).",
            "Master and layout drawing content, placeholder geometry, theme and effect appearance, table styles, charts, SmartArt, media, note drawings, advanced animations and unsupported transition families remain explicit loss. The independent producer fixtures contain only placeholder slides, so complex producer decks remain unqualified."),
        "asciidoc-markdown" or "markdown-asciidoc" or "asciidoc-pdf" => Targeted(
            "Typed fixtures verify the documented bounded AsciiDoc subset, deterministic output, and explicit diagnostics.",
            "Includes, extensions, macros, complex attributes, custom processors, and unsupported inline or block syntax are not a full AsciiDoc implementation."),
        "latex-markdown" or "markdown-latex" or "latex-pdf" => Targeted(
            "Typed fixtures verify the documented bounded LaTeX subset, deterministic output, and explicit diagnostics.",
            "Macro expansion, package execution, TeX layout, arbitrary mathematics, and unsupported environments are intentionally outside the managed subset."),
        "onenote-html" or "html-onenote" or "onenote-markdown" => Targeted(
            "Typed section-model fixtures verify common page text, hierarchy, tables, images, and deterministic projection diagnostics.",
            "The API targets the OfficeIMO section model; full .one binary fidelity, ink, embedded files, rich positioning, and OneNote application rendering remain incomplete."),
        "onenote-pdf" => Targeted(
            "Native .one fixtures reopen the visual PDF and verify page geometry, extracted page text, hyperlink field results with link annotations, and per-picture embedding with alternative text (OfficeIMO.OneNote.Tests).",
            "SaveAsPdf writes the semantic profile: reflowable PDF in reading order that flattens the free-form canvas. SaveAsVisualPdf, or the route's OneNoteVisualPdf setting, writes visual pages as vector content with searchable text, embedded subset fonts, ink, tables and math. Pictures without a native PDF image form are embedded as PNG for that picture only, with a diagnostic. Text is not tagged with paragraph structure, fonts come from the host, and encrypted revisions, embedded files and OneNote application rendering remain incomplete."),
        "odt-pdf" => Targeted(
            "Renders through the ODT-to-Word bridge and the Word PDF engine. Contracts verify running headers and footers, tables, images, page breaks and strict loss policy; LibreOffice Writer and Microsoft Word ODT exports produce searchable PDF that carries the bridge's loss findings (OpenDocumentPdfConversionContracts, OpenDocumentIndependentProducerContracts).",
            "Inherits the ODT-to-DOCX limits: headers and footers come from the first master page, and tracked changes, section-specific layout, other fields and nested inline markup are approximated or reported. Pagination follows the Word layout engine, not LibreOffice or Word."),
        "ods-pdf" => Established(
            "Renders through the ODS-to-Excel bridge and the Excel PDF engine. LibreOffice Calc and Microsoft Excel ODS exports print one page per used sheet with searchable cell text and carried loss findings; contracts verify sheet print ranges and strict loss policy (OpenDocumentIndependentProducerContracts, OpenDocumentPdfConversionContracts).",
            "Inherits the ODS-to-XLSX limits and the Excel PDF print layout. FailOnLoss therefore rejects ordinary LibreOffice and Excel files, which report page masters and row-padding styles. Charts, pivots, conditional formatting outside the mapped subset and exact LibreOffice or Excel page breaks are approximated or reported."),
        "odp-pdf" => Targeted(
            "Renders through the ODP-to-PowerPoint bridge and the PowerPoint PDF engine. Contracts verify slide-sized pages in slide order, text, tables, images and animation loss; LibreOffice Impress and Microsoft PowerPoint ODP exports produce slide-sized pages that carry the bridge's loss findings (OpenDocumentPdfConversionContracts, OpenDocumentIndependentProducerContracts).",
            "Inherits the ODP-to-PPTX limits: master and layout drawing content, placeholder geometry, theme effects, charts, media and animations are omitted or reported. LibreOffice and PowerPoint slide composition is not reproduced exactly."),
        "mhtml-pdf" => Targeted(
            "Bounded archive fixtures verify MIME resource resolution, safe HTML projection, PDF output, and deterministic diagnostics.",
            "Scripts, remote execution, browser-only layout, malformed multipart edge cases, and advanced CSS are not a full browser/MHTML engine."),
        "odg-pdf" => Targeted(
            "Draw scene fixtures verify page dimensions, order, layer selection, searchable text, and page-qualified source and PDF diagnostics.",
            "Uses the supported Draw-to-Drawing profile. Master artwork, page backgrounds, advanced graphic styles, embedded objects, metadata and exact native typography remain incomplete."),
        "visio-pdf" => Established(
            "Visio-authored template drawings and typed drawing fixtures reopen the PDF and verify one vector page per Visio page, exact arc, ellipse and Bezier segments, cached linear and radial gradients, distinct arrowhead kinds, hyperlink annotations, inherited NURBS outlines, searchable text, layer selection and source/PDF loss diagnostics (OfficeIMO.Visio.Tests, Visio.Drawing.PdfFidelity.cs).",
            "Diagram pages are the default; a semantic text and topology report is available on request. Splines and NURBS are flattened, rectangular and path gradients are drawn as radial, theme-driven fills, shadows and unmapped arrow kinds are approximated and reported, formulas are not recalculated, and exact native typography is unqualified."),
        "publisher-pdf" => Targeted(
            "Native Contents/Quill/OfficeArt fixtures verify physical pages, searchable styled text, linked frames, native tables, embedded images and carried source/PDF diagnostics.",
            "Publisher 2002-and-later recovery uses the bounded shared drawing profile. Earlier generations, unsupported graphics, exact native text breaks and native Publisher pixel comparisons remain unqualified."),
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
        "dbf-csv" or "dbf-xlsx" => Targeted(
            "Independent Python-produced dBASE III, FoxPro 2 and Visual FoxPro fixtures exercise typed rows, deleted-row filtering, DBT/FPT memos, binary Base64, CSV reopen, and streaming/buffered XLSX reopen and Open XML validation.",
            "Requires DBAClientX.Dbf's bounded field profiles. Native application acceptance and Visual FoxPro 31/32 generations remain unqualified. Output omits DBF indexes, deletion flags and native schema metadata; CSV nulls need a selected marker, and XLSX follows Excel numeric precision and cell text limits."),
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
        "rtf" => Targeted(
            "RtfImageConversionTests render through the RTF PDF layout and the PDF rasterizer, covering page selection, saved encoding, carried source loss, strict rejection before publication and cancellation.",
            "Uses the RTF semantic PDF layout rather than the Word renderer behind DOCX images. Only PNG output is exercised directly; the other encoders come from the PDF image owner. Native Office pagination and unsupported source structures remain subject to reported conversion loss."),
        "mhtml" => Targeted(
            "OfficeIMO.Html.Tests Mhtml/MhtmlRenderResourcePolicyTests render archives through the MHTML PDF layout and the PDF rasterizer with embedded-resource and remote-resource policies.",
            "Uses the MHTML semantic PDF layout. Browser layout, scripts and unsupported archive structures remain subject to reported conversion loss."),
        "xps" => Targeted(
            "XpsImageConversionTests and XpsUnboundedRadialSpreadTests cover page dimensions, selection, encoding, source evidence and persistence through the native drawing projection.",
            "Uses the existing XPS drawing projection. Unsupported native paint and layout remain reported or rejected by that owner; this route does not add a separate renderer."),
        "publisher" => Targeted(
            "PublisherImageConversionTests cover page dimensions, selection, encoding, source evidence and persistence through the reconstructed Publisher drawing scene.",
            "Uses the existing Publisher drawing projection. Unsupported native paint and layout remain reported or rejected by that owner; this route does not add a separate renderer."),
        "odg" => Targeted(
            "OfficeIMO.OpenDocument.Tests OpenDocument.Odg.ImageConversion covers page dimensions, selection, encoding, source evidence and persistence through the Draw scene projection.",
            "Uses the existing Draw-to-Drawing projection. Unsupported native paint and layout remain reported or rejected by that owner; this route does not add a separate renderer."),
        "djvu" => Targeted(
            "Independent native decoder rasters are compared to encoded PNG pixels; other encodings use the shared raster owner.",
            "Retains the DjVu renderer's qualified codec, annotation and short-edge limitations. SVG embeds the rendered raster and does not reconstruct vectors."),
        "project" => Targeted(
            "Project image contracts exercise portable report views, typography, image density, selection and bounded output.",
            "Requires an explicit ProjectView. Scheduling requires a project calendar; portable reports do not reproduce native Project saved-view styling or filters."),
        "ods" => Established(
            "OpenDocumentImageExportContracts render ODS sheets in order through the ODS-to-Excel bridge and the same Excel renderer as XLSX images, in all five encodings. LibreOffice Calc and Microsoft Excel spreadsheets render every sheet with the bridge's loss findings attached (OpenDocumentIndependentProducerContracts).",
            "Image output is intentionally flattened. Inherits the ODS-to-XLSX conversion limits; unsupported source structure remains reported."),
        "odt" or "odp" => Targeted(
            "OpenDocumentImageExportContracts render ODT through the Word renderer with bridge loss diagnostics and aggregate loss policy, and ODP through the PowerPoint renderer in slide order, both in all five encodings.",
            "Image output is intentionally flattened. Inherits the Targeted ODT-to-DOCX and ODP-to-PPTX conversion limits; unsupported source structure remains approximated or reported."),
        "onenote" => Targeted(
            "OneNoteRenderingTests and OneNote.TextFormatting render typed sections through the shared SVG and raster renderer.",
            "Image output is intentionally flattened. Ink, embedded files, rich positioning and OneNote application rendering remain approximated or diagnosed."),
        "email" => Targeted(
            "OfficeIMO.Html.Tests Html.EmailImageExport renders message chrome, plain, HTML and RTF bodies and inline content-ID images through the HTML renderer in all five encodings.",
            "Image output is intentionally flattened. Remote resources are blocked, attachments are listed rather than rendered, and mail-client layout is not reproduced."),
        "epub" => Targeted(
            "OfficeIMO.Html.Tests Html.EpubImageExport, Html.EpubFallbackImageExport and Html.EpubEmptyImageExport render chapters through the HTML renderer with fallback and empty-publication handling.",
            "Image output is intentionally flattened. Reader pagination, scripts and missing resources remain approximated or diagnosed."),
        _ => throw new InvalidOperationException($"Image conversion source '{source}' does not have a support assessment.")
    };

    private static OfficeConversionSupportAssessment Targeted(string evidence, string limits) =>
        new(OfficeConversionSupportLevel.Targeted, evidence, limits);

    private static OfficeConversionSupportAssessment Established(string evidence, string limits) =>
        new(OfficeConversionSupportLevel.Established, evidence, limits);

    private static OfficeConversionSupportAssessment Advanced(string evidence, string limits) =>
        new(OfficeConversionSupportLevel.Advanced, evidence, limits);
}
