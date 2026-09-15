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
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Inspect },
            limitedSupported: new[] { OfficeOperationKind.Edit },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "Unedited source can be retained; semantic edits serialize normalized HTML rather than promising byte identity.");
        AddNativeLifecycle(rows, "pdf-native", "OfficeIMO.Pdf", "Pdf.Native",
            "PdfDocument.Create / Load / Save / Inspect",
            "OfficeIMO.Pdf.Tests authoritative interoperability read, render, diagnostics, and mutation contracts",
            new[] { ".pdf" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Inspect, OfficeOperationKind.Validate },
            limitedSupported: new[] { OfficeOperationKind.Edit },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "Mutation and preservation depend on the parsed feature set and declared incremental or rewrite policy.");
        AddNativeLifecycle(rows, "email-native", "OfficeIMO.Email", "Email.Native",
            "EmailDocument / Load / Save",
            "OfficeIMO.Email.Tests typed read, edit, write, and fixture round-trip contracts",
            new[] { ".eml", ".mime", ".msg", ".oft", ".tnef", ".dat" },
            new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Inspect },
            limitedSupported: new[] { OfficeOperationKind.Edit },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "Format-specific unsupported properties can remain opaque or require a documented conversion boundary.");
        AddEmailStoreLifecycleRows(rows);
        AddProjectLifecycleRows(rows);
        AddNativeLifecycle(rows, "epub-native", "OfficeIMO.Epub", "Epub.Native",
            "EpubReader / EpubPackage",
            "OfficeIMO.Epub.Tests package extraction, metadata, and navigation inspection contracts",
            new[] { ".epub" },
            new[] { OfficeOperationKind.Read, OfficeOperationKind.Inspect },
            unsupported: new[] { OfficeOperationKind.Create, OfficeOperationKind.Edit },
            limitation: "The current package is an extraction and inspection surface; EPUB authoring is not implemented.");
    }

    private static void AddEmailStoreLifecycleRows(ICollection<OfficeOperationCapability> rows) {
        AddNativeLifecycle(rows, "email-store-pst", "OfficeIMO.Email", "Email.Store.Pst",
            "EmailStoreSession / EmailStorePstWriter / EmailStoreConverter.ConvertToPst / EmailStoreConverter.MergeToPst / EmailStorePstMutationTransaction",
            "OfficeIMO.Email.Tests PST read, inspection, validation, export, conversion, merge, and verified rewrite contracts",
            new[] { ".pst" },
            Array.Empty<OfficeOperationKind>(),
            limitedSupported: new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Inspect, OfficeOperationKind.Validate },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "PST creation writes a distinct destination. Mutation is a verified atomic rewrite of an existing Unicode PST; ANSI mutation, append, in-place NDB editing or repair, and password or encryption authoring are not supported. Opaque or unsupported semantics require explicit loss policy.");
        AddNativeLifecycle(rows, "email-store-ost", "OfficeIMO.Email", "Email.Store.Ost",
            "EmailStoreSession",
            "OfficeIMO.Email.Tests OST traversal, inspection, validation, and export contracts",
            new[] { ".ost" },
            Array.Empty<OfficeOperationKind>(),
            limitedSupported: new[] { OfficeOperationKind.Read, OfficeOperationKind.Inspect, OfficeOperationKind.Validate },
            unsupported: new[] { OfficeOperationKind.Create, OfficeOperationKind.Edit, OfficeOperationKind.Preserve },
            limitation: "OST is a bounded read, inspect, validate, and export source. OfficeIMO does not create, rewrite, append to, or preserve an OST as a newly serialized OST artifact.");
        AddNativeLifecycle(rows, "email-store-olm", "OfficeIMO.Email", "Email.Store.Olm",
            "EmailStoreSession",
            "OfficeIMO.Email.Tests OLM traversal, inspection, validation, and export contracts",
            new[] { ".olm" },
            Array.Empty<OfficeOperationKind>(),
            limitedSupported: new[] { OfficeOperationKind.Read, OfficeOperationKind.Inspect, OfficeOperationKind.Validate },
            unsupported: new[] { OfficeOperationKind.Create, OfficeOperationKind.Edit, OfficeOperationKind.Preserve },
            limitation: "OLM is a bounded read, inspect, validate, and export source. OfficeIMO does not create or rewrite OLM archives, and unsupported OLM entries remain diagnostic evidence rather than a round-trip promise.");
        AddNativeLifecycle(rows, "email-store-mbox", "OfficeIMO.Email", "Email.Store.Mbox",
            "EmailMailboxReader / EmailMailboxWriter / EmailStoreSession",
            "OfficeIMO.Email.Tests mboxo and mboxrd streaming read, edit, write, inspection, validation, and interoperability contracts",
            new[] { ".mbox", ".mbx" },
            Array.Empty<OfficeOperationKind>(),
            limitedSupported: new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Inspect, OfficeOperationKind.Validate },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "Writing emits deterministic mboxo or mboxrd from the typed messages. Envelope and escaping semantics are retained, but edited output is normalized and byte-identical aggregate preservation is not promised.");
        AddNativeLifecycle(rows, "email-store-emlx", "OfficeIMO.Email", "Email.Store.Emlx",
            "EmailStoreSession / EmailStoreEmlxWriter / ExportToNativeDirectory",
            "OfficeIMO.Email.Tests EMLX bounded read, edit, write, native export, metadata, inspection, and validation contracts",
            new[] { ".emlx" },
            Array.Empty<OfficeOperationKind>(),
            limitedSupported: new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit, OfficeOperationKind.Inspect, OfficeOperationKind.Validate },
            partial: new[] { OfficeOperationKind.Preserve },
            limitation: "EMLX writes regenerate the RFC message and supported Apple property-list metadata. Unsupported binary or unrecognized metadata remains diagnostic evidence rather than a complete byte-preserving round trip.");
    }

    private static void AddProjectLifecycleRows(ICollection<OfficeOperationCapability> rows) {
        const string packageId = "OfficeIMO.Project";
        const string publicApi = "ProjectDocument.Create / Load / AssessSave / Save";
        const string evidence = "OfficeIMO.Project.Tests lifecycle, preservation, conversion, and independent-reader contracts";
        string[] binaryExtensions = { ".mpp", ".mpt" };
        string[] mpxExtensions = { ".mpx" };

        AddProjectLifecycleRow(rows, "project-binary", packageId, "Project.MppMpt", OfficeOperationKind.Create,
            OfficeOperationSupportState.Supported, publicApi, evidence, binaryExtensions,
            "Creates generation profiles 8, 9, 12, and 14; a start date and named base calendar are required, unsupported fields require explicit omission permission, and path-based Global.mpt creation is rejected.");
        AddProjectLifecycleRow(rows, "project-binary", packageId, "Project.MppMpt", OfficeOperationKind.Read,
            OfficeOperationSupportState.Supported, publicApi, evidence, binaryExtensions,
            "Read coverage is generation- and producer-qualified; unsupported external variable-storage layouts and protected inputs are rejected.");
        AddProjectLifecycleRow(rows, "project-binary", packageId, "Project.MppMpt", OfficeOperationKind.Edit,
            OfficeOperationSupportState.Supported, publicApi, evidence, binaryExtensions,
            "Mapped scalar and structural edits are assessed; opaque references, unqualified native fields, and signed content require rejection or explicit caller permission.");
        AddProjectLifecycleRow(rows, "project-binary", packageId, "Project.MppMpt", OfficeOperationKind.Preserve,
            OfficeOperationSupportState.Supported, publicApi, evidence, binaryExtensions,
            "Unchanged saves retain exact bytes; edited saves preserve opaque streams without claiming semantic support or safe reference remapping.");
        AddProjectLifecycleRow(rows, "project-binary", packageId, "Project.MppMpt", OfficeOperationKind.Inspect,
            OfficeOperationSupportState.Supported, publicApi, evidence, binaryExtensions,
            "Inspection is bounded to container, generation, stream, producer, protection, macro, signature, and embedded-content inventory; it does not activate or semantically decode opaque content.");

        AddProjectLifecycleRow(rows, "project-mpx", packageId, "Project.Mpx", OfficeOperationKind.Create,
            OfficeOperationSupportState.Supported, publicApi, evidence, mpxExtensions,
            "Creates MPX 4.0 with qualified field tables and encodings; unsupported model fields and unrepresentable text require rejection or explicit omission permission.");
        AddProjectLifecycleRow(rows, "project-mpx", packageId, "Project.Mpx", OfficeOperationKind.Read,
            OfficeOperationSupportState.Supported, publicApi, evidence, mpxExtensions,
            "Read coverage is limited to MPX 4.0/4.1, qualified encodings, numeric or English field tables, and supported value vocabularies.");
        AddProjectLifecycleRow(rows, "project-mpx", packageId, "Project.Mpx", OfficeOperationKind.Edit,
            OfficeOperationSupportState.Supported, publicApi, evidence, mpxExtensions,
            "Canonical rewrites cover the mapped MPX 4 profile; unknown records, unsupported fields, and ambiguous source syntax require explicit loss permission or rejection.");
        AddProjectLifecycleRow(rows, "project-mpx", packageId, "Project.Mpx", OfficeOperationKind.Preserve,
            OfficeOperationSupportState.Supported, publicApi, evidence, mpxExtensions,
            "Unchanged saves retain exact source bytes; rewriting retained unknown records or noncanonical source placement requires explicit loss permission.");
        AddProjectLifecycleRow(rows, "project-mpx", packageId, "Project.Mpx", OfficeOperationKind.Inspect,
            OfficeOperationSupportState.Supported, publicApi, evidence, mpxExtensions,
            "Inspection reports bounded MPX framing, encoding, field tables, and retained unknown records without executing or resolving external content.");

        AddProjectLifecycleRow(rows, "project-binary-to-xml", packageId, "Project.MppMpt", OfficeOperationKind.Convert,
            OfficeOperationSupportState.Partial, publicApi, evidence, binaryExtensions,
            "Conversion uses the typed model and pre-write assessment; opaque native content and target-profile loss require explicit caller permission.", "Project.Xml");
        AddProjectLifecycleRow(rows, "project-binary-to-mpx", packageId, "Project.MppMpt", OfficeOperationKind.Convert,
            OfficeOperationSupportState.Partial, publicApi, evidence, binaryExtensions,
            "MPX 4 output supports a bounded mapped subset; unsupported native fields, opaque records, and representation changes require explicit caller permission.", "Project.Mpx");
        AddProjectLifecycleRow(rows, "project-mpx-to-xml", packageId, "Project.Mpx", OfficeOperationKind.Convert,
            OfficeOperationSupportState.Partial, publicApi, evidence, mpxExtensions,
            "Conversion uses the typed model; retained unknown MPX records and unsupported source semantics require explicit caller permission.", "Project.Xml");
        AddProjectLifecycleRow(rows, "project-mpx-to-binary", packageId, "Project.Mpx", OfficeOperationKind.Convert,
            OfficeOperationSupportState.Partial, publicApi, evidence, mpxExtensions,
            "The selected MPP or MPT generation controls the writable field profile; unsupported MPX semantics and target-profile loss require explicit caller permission.", "Project.MppMpt");
    }

    private static void AddProjectLifecycleRow(
        ICollection<OfficeOperationCapability> rows,
        string id,
        string packageId,
        string formatId,
        OfficeOperationKind operation,
        OfficeOperationSupportState state,
        string publicApi,
        string evidence,
        IEnumerable<string> extensions,
        string limitation,
        string? targetFormatId = null) {
        rows.Add(new OfficeOperationCapability(
            "native:" + id + ":" + operation.ToString().ToLowerInvariant(),
            packageId,
            formatId,
            id,
            operation,
            state,
            publicApi,
            evidence + "::" + operation,
            NativeLifecycleCatalogId,
            extensions,
            targetFormatId,
            limitation));
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
        IEnumerable<OfficeOperationKind>? limitedSupported = null,
        IEnumerable<OfficeOperationKind>? partial = null,
        IEnumerable<OfficeOperationKind>? unsupported = null,
        string? limitation = null) {
        AddNativeLifecycleRows(rows, capabilityId, packageId, formatId, publicApi, evidence, extensions,
            supported, OfficeOperationSupportState.Supported, limitation: null);
        AddNativeLifecycleRows(rows, capabilityId, packageId, formatId, publicApi, evidence, extensions,
            limitedSupported ?? Array.Empty<OfficeOperationKind>(), OfficeOperationSupportState.Supported, limitation);
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
