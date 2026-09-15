using OfficeIMO;

namespace OfficeIMO.Workflows;

public static partial class OfficeOperationCapabilityCatalog {
    private const string NativeLifecycleCatalogId = "OfficeIMO.NativeLifecycle";

    private static void AddNativeLifecycleRows(ICollection<OfficeOperationCapability> rows) {
        AddNativeLifecycle(rows, "word-openxml", "OfficeIMO.Word", "Word.OpenXml",
            "WordDocument.Create / Load / Save",
            "OfficeIMO.Word.Tests create, load, edit, save, and preservation contracts",
            new[] { ".docx", ".docm", ".dotx", ".dotm" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Preserve, OfficeOperationKind.Inspect, OfficeOperationKind.Validate });
        AddNativeLifecycle(rows, "excel-openxml", "OfficeIMO.Excel", "Excel.OpenXml",
            "ExcelDocument.Create / Load / Save",
            "OfficeIMO.Excel.Tests create, load, edit, save, and preservation contracts",
            new[] { ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Preserve, OfficeOperationKind.Inspect, OfficeOperationKind.Validate });
        AddNativeLifecycle(rows, "powerpoint-openxml", "OfficeIMO.PowerPoint", "PowerPoint.OpenXml",
            "PowerPointPresentation.Create / Load / Save",
            "OfficeIMO.PowerPoint.Tests create, load, edit, save, and preservation contracts",
            new[] { ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".ppam" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Preserve, OfficeOperationKind.Inspect, OfficeOperationKind.Validate });
        AddNativeLifecycle(rows, "visio-openxml", "OfficeIMO.Visio", "Visio.OpenXml",
            "VisioDocument.Create / Load / Save; VisioValidator",
            "OfficeIMO.Visio.Tests create, load, edit, save, inspect, and validation contracts",
            new[] { ".vsdx", ".vsdm", ".vstx", ".vstm", ".vssx", ".vssm" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Preserve, OfficeOperationKind.Inspect, OfficeOperationKind.Validate });
        AddNativeLifecycle(rows, "onenote-native", "OfficeIMO.OneNote", "OneNote.Native",
            "OneNoteSectionReader / OneNoteSectionWriter / OneNotePackageReader / OneNotePackageWriter",
            "OfficeIMO.OneNote.Tests section, notebook, table-of-contents, and package round-trip contracts",
            new[] { ".one", ".onetoc2", ".onepkg" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Preserve, OfficeOperationKind.Inspect });
        AddNativeLifecycle(rows, "opendocument-native", "OfficeIMO.OpenDocument", "OpenDocument.Native",
            "OdfDocument.Create / Load / Save / Validate",
            "OfficeIMO.OpenDocument.Tests producer-corpus create, load, edit, save, and validation contracts",
            new[] { ".odt", ".ods", ".odp" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Preserve, OfficeOperationKind.Inspect, OfficeOperationKind.Validate });
        AddNativeLifecycle(rows, "rtf-native", "OfficeIMO.Rtf", "Rtf.Native",
            "RtfDocument.Create / Load / Save",
            "OfficeIMO.Rtf.Tests producer-corpus parsing, editing, preservation, and deterministic-write contracts",
            new[] { ".rtf" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Preserve, OfficeOperationKind.Inspect });
        AddNativeLifecycle(rows, "markdown-native", "OfficeIMO.Markdown", "Markdown.Native",
            "MarkdownDoc.Create / Load / Save / ToMarkdown",
            "OfficeIMO.Markdown.Tests typed-model parse, edit, and render contracts",
            new[] { ".md", ".markdown" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "Semantic edits serialize normalized Markdown; byte-identical source preservation is not promised.");
        AddNativeLifecycle(rows, "html-native", "OfficeIMO.Html", "Html.Native",
            "HtmlConversionDocument.Parse / ToHtml",
            "OfficeIMO.Html.Tests source parsing, editing, serialization, and rendering contracts",
            new[] { ".html", ".htm" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Inspect },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "Unedited source can be retained; semantic edits serialize normalized HTML rather than promising byte identity.");
        AddNativeLifecycle(rows, "pdf-native", "OfficeIMO.Pdf", "Pdf.Native",
            "PdfDocument.Create / Load / Save / Inspect",
            "OfficeIMO.Pdf.Tests authoritative interoperability read, render, diagnostics, and mutation contracts",
            new[] { ".pdf" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Inspect, OfficeOperationKind.Validate },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "Mutation and preservation depend on the parsed feature set and declared incremental or rewrite policy.");
        AddNativeLifecycle(rows, "email-native", "OfficeIMO.Email", "Email.Native",
            "EmailDocument.Load / Save",
            "OfficeIMO.Email.Tests typed read, edit, write, and fixture round-trip contracts",
            new[] { ".eml", ".mime", ".msg", ".oft", ".tnef", ".dat" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Inspect },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "Format-specific unsupported properties can remain opaque or require a documented conversion boundary.");
        AddNativeLifecycle(rows, "epub-native", "OfficeIMO.Epub", "Epub.Native",
            "EpubReader / EpubPackage",
            "OfficeIMO.Epub.Tests package extraction, metadata, and navigation inspection contracts",
            new[] { ".epub" },
            new[] { OfficeOperationKind.Read, OfficeOperationKind.Inspect },
            unsupported: new[] { OfficeOperationKind.Create, OfficeOperationKind.Edit },
            limitation: "The current package is an extraction and inspection surface; EPUB authoring is not implemented.");
    }

    private static void AddNativeLifecycle(
        ICollection<OfficeOperationCapability> rows,
        string capabilityId,
        string packageId,
        string formatId,
        string publicApi,
        string evidence,
        IEnumerable<string> extensions,
        IEnumerable<OfficeOperationKind> supported,
        IEnumerable<OfficeOperationKind>? partial = null,
        IEnumerable<OfficeOperationKind>? unsupported = null,
        string? limitation = null) {
        AddNativeLifecycleRows(rows, capabilityId, packageId, formatId, publicApi, evidence, extensions,
            supported, OfficeOperationSupportState.Supported, limitation: null);
        AddNativeLifecycleRows(rows, capabilityId, packageId, formatId, publicApi, evidence, extensions,
            partial ?? Array.Empty<OfficeOperationKind>(), OfficeOperationSupportState.Partial, limitation);
        AddNativeLifecycleRows(rows, capabilityId, packageId, formatId, publicApi, evidence, extensions,
            unsupported ?? Array.Empty<OfficeOperationKind>(), OfficeOperationSupportState.Unsupported, limitation);
    }

    private static void AddNativeLifecycleRows(
        ICollection<OfficeOperationCapability> rows,
        string capabilityId,
        string packageId,
        string formatId,
        string publicApi,
        string evidence,
        IEnumerable<string> extensions,
        IEnumerable<OfficeOperationKind> operations,
        OfficeOperationSupportState state,
        string? limitation) {
        foreach (OfficeOperationKind operation in operations) {
            rows.Add(new OfficeOperationCapability(
                "native:" + capabilityId + ":" + operation.ToString().ToLowerInvariant(),
                packageId,
                formatId,
                capabilityId,
                operation,
                state,
                publicApi,
                evidence + "::" + operation,
                NativeLifecycleCatalogId,
                extensions,
                limitation: limitation));
        }
    }
}
