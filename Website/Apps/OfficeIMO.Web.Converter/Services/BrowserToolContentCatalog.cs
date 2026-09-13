using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

internal static class BrowserToolContentCatalog {
    internal static BrowserToolContent Provenance { get; } = new(
        "Check and remove file origin data",
        "See whether a supported file contains Content Credentials, links to credentials, or AI source declarations. Choose what to remove from a new copy.",
        "One JPEG, PNG, WebP, PDF, DOCX, XLSX, or PPTX file up to 25 MB.",
        "A readable inspection result and, when requested, a separate cleaned copy plus a JSON report.",
        "This does not remove visible watermarks, personal metadata, or image pixels. It also cannot prove who or what created a file.",
        "/provenance/",
        "Read the file origin guide",
        "Check file provenance and Content Credentials | OfficeIMO",
        [
            new("Choose a file", "Your file opens only in this browser tab; the original is never changed."),
            new("Inspect its origin data", "Review the supported Content Credentials, credential links, and AI source declarations found in the file."),
            new("Create and check a copy", "Select records to remove, create a separate copy, and review the re-inspection before downloading it.")
        ]);

    internal static BrowserToolContent For(ConversionRoute route) {
        (string summary, string output, string expectation) = route.Id switch {
            "docx-pdf" => (
                "Turn a Word document into a PDF you can preview and download, with warnings for content that could not be reproduced exactly.",
                "A PDF download, an in-browser preview, and conversion diagnostics you can review before saving the result.",
                "Common text, tables, images, headers, footers, lists, and page settings are supported. Complex floating layout, SmartArt, fields, and exact Microsoft Word pagination can differ."),
            "xlsx-pdf" => (
                "Turn an Excel workbook into a PDF while keeping worksheet layout decisions and conversion warnings visible.",
                "A PDF download with a preview and diagnostics for sheets, pagination, fonts, and unsupported workbook content.",
                "Large or highly visual workbooks may need a browser-safe preview or a server-side .NET workflow for full layout and media handling."),
            "pptx-pdf" => (
                "Turn PowerPoint slides into a fixed-layout PDF and review any approximations before downloading it.",
                "A PDF download, slide preview, and diagnostics for shapes, text, images, charts, and other presentation content.",
                "Animations, transitions, media playback, and some advanced DrawingML effects are not reproduced in a static PDF."),
            "html-pdf" => (
                "Paste HTML and create a tagged PDF using OfficeIMO's browser-safe HTML and CSS renderer.",
                "A PDF download, preview, and diagnostics for unsupported or adjusted HTML, CSS, fonts, and external resources.",
                "The renderer supports a bounded HTML and CSS profile. It does not run page scripts or behave like a full browser print engine."),
            "markdown-html" => (
                "Turn Markdown into safe, reviewable HTML without uploading the text to a server.",
                "An HTML file and rendered preview created from the supported Markdown profile.",
                "Raw HTML and features outside the selected safe profile may be removed or simplified."),
            "html-markdown" => (
                "Turn HTML into portable Markdown for editing, publishing, or version control.",
                "A Markdown download and text preview, with diagnostics for content that cannot be represented cleanly.",
                "Complex page layout, styling, scripts, forms, and browser behavior do not have direct Markdown equivalents."),
            "markdown-docx" => (
                "Create an editable Word document from Markdown while keeping conversion choices visible.",
                "A DOCX download with editable headings, paragraphs, lists, tables, links, and other supported Markdown content.",
                "Markdown does not carry every Word formatting or page-layout concept, so the result uses a deliberate document style rather than recreating a source Word file."),
            "pdf-docx" => (
                "Create an editable Word document from supported PDF content, or use rendered page images when visual fidelity matters more than editability.",
                "A DOCX download plus diagnostics describing native content, visual fallbacks, and reconstruction limits.",
                "PDF stores final page positions, not the original Word structure. Columns, reading order, fonts, and complex graphics may be reconstructed or preserved as images."),
            "pdf-xlsx" => (
                "Find tables in a PDF and place the detected rows and columns into an editable Excel workbook.",
                "An XLSX download containing detected tables, together with diagnostics about pages and table reconstruction.",
                "This extracts detected tables rather than recreating the entire PDF page. Scanned pages need OCR, which is outside this browser workflow."),
            "pdf-pptx" => (
                "Create a PowerPoint presentation from PDF pages using native, visual, hybrid, or tables-only content modes.",
                "A PPTX download plus diagnostics explaining which slide content was reconstructed and which content used a visual fallback.",
                "A PDF does not contain the original slide model, animations, or transitions. Choose the conversion mode based on whether editability or appearance matters most."),
            "pdf-html" => (
                "Turn a PDF into reviewable HTML using semantic reading order or a positioned page view.",
                "An HTML download and preview, with diagnostics about text, images, positioning, and unsupported PDF content.",
                "Semantic HTML favors reading and reuse; positioned HTML favors page appearance. Neither recreates the original authoring file or runs OCR on image-only pages."),
            "pdf-png" => (
                "Render every PDF page as a PNG image for previews, sharing, or image-based workflows.",
                "A PNG for a single-page PDF or a ZIP archive containing one detailed PNG per page, with page-level diagnostics.",
                "The result is made of images, so text is no longer selectable or editable. Large and image-heavy PDFs remain subject to browser memory and page limits."),
            _ => (route.Description, $"A {route.Target} download and the diagnostics produced by the conversion.", route.KnownLimitations)
        };

        return new BrowserToolContent(
            route.Title,
            summary,
            InputFor(route),
            output,
            expectation,
            GuideFor(route),
            "Read the conversion guide",
            $"{route.Title} in your browser | OfficeIMO",
            [
                new(route.InputKind == ConversionInputKind.File ? "Choose your file" : $"Add your {route.Source} content", route.InputKind == ConversionInputKind.File ? "Drop a supported file into the workspace or choose one from your device." : "Paste content into the editor or start with the built-in sample."),
                new("Review the settings", "Choose the available output profile or conversion mode before you run the tool."),
                new("Convert and review", "Preview the result, read any diagnostics, then download the new file when it meets your needs.")
            ]);
    }

    internal static BrowserToolContent For(PdfToolDefinition tool) {
        (string output, string expectation) = tool.Id switch {
            "inspect" => ("A readable summary plus downloadable JSON reports; the PDF itself is not rewritten.", "Inspection reports what OfficeIMO can read or safely rewrite. It does not validate certificate trust, decrypt a file without credentials, or run OCR."),
            "compare" => ("A downloadable HTML gallery showing expected, actual, and highlighted difference images, plus comparison details.", "The browser compares up to 25 pages at an exact visual threshold. A difference report is evidence to review, not a statement about which PDF is correct."),
            "merge" => ("One PDF containing all selected documents in the order shown, plus an operation report.", "Bookmarks, forms, encryption, signatures, and other document-level features may need policy decisions when files are combined."),
            "split" => ("A ZIP archive containing consecutive PDF parts and a report describing the split.", "The browser creates at most 100 output files. Choose a practical page count per file before running the split."),
            "extract" => ("A new PDF containing only the selected pages, in the order requested, plus an operation report.", "Page expressions are one-based. Use values such as 1-3,5,last and review the resolved page count before using the output."),
            "delete" => ("A new PDF that keeps every page except the selected pages, plus an operation report.", "Deletion applies only to the new copy. Confirm the selection carefully; use Extract pages when you want to keep only the named pages."),
            "reorder" => ("A new PDF whose pages follow the complete order you provide, plus an operation report.", "List every source page exactly once. The tool rejects repeated or omitted pages instead of creating an incomplete document."),
            "rotate" => ("A new PDF with the selected pages rotated by 90, 180, or 270 degrees, plus an operation report.", "Rotation changes page orientation in the new copy; it does not deskew scanned content inside a page image."),
            "optimize" => ("A losslessly rewritten PDF plus a report of applied and skipped actions. Balanced and archival profiles retain the original bytes when rewriting would not make the file smaller.", "Maximum compression and Fast Web View can return a larger file when their requested structure requires it. The tool does not rasterize pages or apply lossy scan compression."),
            "protect" => ("A separate AES-256 password-protected PDF plus preservation evidence.", "Store the owner password safely. Password protection controls compatible PDF readers; it is not digital signing or rights-management infrastructure."),
            "unlock" => ("A separate PDF without Standard password security, plus an operation report.", "You need the valid owner password. The tool does not bypass unsupported encryption or remove certificate-based restrictions."),
            "redact" => ("A rewritten PDF with matching literal text removed and a verification report checking rewritten content and streams.", "This is case-insensitive literal-text redaction, not OCR or pattern matching. Image-only text and text split into unexpected PDF objects may need a different workflow."),
            _ => ("A new downloadable result plus an operation report.", "Review the result and report before using the output.")
        };

        return new BrowserToolContent(
            tool.Label,
            tool.Description,
            InputFor(tool),
            output,
            expectation,
            GuideFor(tool),
            "Open the full guide",
            SeoTitleFor(tool),
            StepsFor(tool));
    }

    private static string InputFor(ConversionRoute route) => route.InputKind == ConversionInputKind.File
        ? route.Source == "PDF"
            ? $"One PDF file up to 25 MB and 500 pages ({route.Accept.Replace(",", ", ")})."
            : $"One {route.Source} file up to 25 MB ({route.Accept.Replace(",", ", ")})."
        : $"{route.Source} text pasted into the browser editor, up to 500,000 characters.";

    private static string SeoTitleFor(PdfToolDefinition tool) => tool.Id switch {
        "extract" => "Extract PDF pages online | OfficeIMO",
        "delete" => "Delete PDF pages online | OfficeIMO",
        "reorder" => "Reorder PDF pages online | OfficeIMO",
        "rotate" => "Rotate PDF pages online | OfficeIMO",
        "redact" => "Redact PDF text online | OfficeIMO",
        _ => $"{tool.Label} online in your browser | OfficeIMO"
    };

    private static string InputFor(PdfToolDefinition tool) => tool.InputMode switch {
        PdfToolInputMode.Pair => "Exactly two PDF files, up to 25 MB and 500 pages each.",
        PdfToolInputMode.Multiple => "Two to ten PDF files, up to 25 MB and 500 pages each, and 75 MB combined.",
        _ => "One PDF file up to 25 MB and 500 pages."
    };

    private static IReadOnlyList<BrowserToolStep> StepsFor(PdfToolDefinition tool) => [
        new(tool.InputMode == PdfToolInputMode.Pair ? "Choose two PDFs" : tool.InputMode == PdfToolInputMode.Multiple ? "Choose your PDFs" : "Choose a PDF", tool.InputMode == PdfToolInputMode.Multiple ? "Add the files and put them in the order the operation should use." : "Drop the PDF into the workspace or load the product sample."),
        new(SettingsStepFor(tool), SettingsDescriptionFor(tool)),
        new("Run, review, and download", "Read the visible result and operation evidence before downloading the new artifact.")
    ];

    private static string SettingsStepFor(PdfToolDefinition tool) => tool.Id switch {
        "inspect" => "Run the inspection",
        "compare" => "Check the comparison order",
        "merge" => "Set the file order",
        "split" => "Choose pages per file",
        "extract" or "delete" or "reorder" => "Describe the pages",
        "rotate" => "Choose pages and rotation",
        "optimize" => "Choose an optimization profile",
        "protect" => "Set the passwords",
        "unlock" => "Enter the owner password",
        "redact" => "Enter the exact text",
        _ => "Review the options"
    };

    private static string SettingsDescriptionFor(PdfToolDefinition tool) => tool.Id switch {
        "inspect" => "No rewrite settings are needed; inspection leaves the selected PDF unchanged.",
        "compare" => "The first PDF is the expected document and the second is the document being checked.",
        "merge" => "Move files until the displayed order matches the output you want.",
        "split" => "Set the maximum number of consecutive pages for each output PDF.",
        "extract" => "Enter the pages to keep, for example 1-3,5,last.",
        "delete" => "Enter the pages to remove and confirm the new-copy operation.",
        "reorder" => "Enter every source page exactly once, in the output order you want.",
        "rotate" => "Enter the pages and choose a clockwise rotation.",
        "optimize" => "Choose lossless compression, deduplication, or Fast Web View.",
        "protect" => "Set a user password and a distinct owner password for the new PDF.",
        "unlock" => "Supply the owner password required to remove Standard security from the copy.",
        "redact" => "Enter a literal text value and confirm that it should be permanently removed from the copy.",
        _ => "Review the available settings before running the operation."
    };

    private static string GuideFor(ConversionRoute route) => route.Id switch {
        "docx-pdf" => "/convert/word-to-pdf/",
        "xlsx-pdf" => "/convert/excel-to-pdf/",
        "pdf-html" => "/convert/pdf-to-html/",
        "pdf-docx" => "/convert/pdf-to-word/",
        "pdf-pptx" => "/convert/pdf-to-powerpoint/",
        "pdf-xlsx" => "/convert/pdf-tables-to-excel/",
        _ => "/convert/guides/"
    };

    private static string GuideFor(PdfToolDefinition tool) => $"/pdf/{tool.Id switch { "extract" => "extract-pages", "delete" => "delete-pages", "reorder" => "reorder-pages", "rotate" => "rotate-pages", _ => tool.Id }}/";
}
