using System.Globalization;
using OfficeIMO.ContentSafety;

namespace OfficeIMO.Reader.FormatInternals;

/// <summary>Thin shared projection of existing native inspection contracts; parsing remains in format owners.</summary>
internal static class NativeInspectionProjection {
    internal static OfficeContentSafetyOptions SafetyOptions(ReaderInspectionOptions options) => new OfficeContentSafetyOptions {
        MaxFindings = (int)Math.Min(int.MaxValue, options.MaxRecords)
    };

    internal static void InspectContentSafety(ReaderInspectionMetadataBuilder builder,
        Func<OfficeContentSafetyOptions, OfficeContentSafetyReport> inspect, string? path) {
        OfficeContentSafetyReport report;
        try { report = inspect(SafetyOptions(builder.Options)); }
        catch (OfficeInspectionResourceLimitException limit) {
            bool readerFindings = limit.Owner == "content-safety" && limit.LimitName == nameof(OfficeContentSafetyOptions.MaxFindings)
                && limit.Maximum == builder.Options.MaxRecords;
            throw new ReaderResourceLimitException(readerFindings ? "Inspection.MaxRecords" : limit.Owner + "." + limit.LimitName,
                limit.Maximum, limit);
        }
        AddContentSafety(builder, report, path);
    }

    internal static void AddContentSafety(ReaderInspectionMetadataBuilder builder, OfficeContentSafetyReport report, string? path) {
        foreach (OfficeContentSafetyFinding finding in report.Findings) {
            builder.Add(OfficeDocumentInspectionCategories.ContentSafety, finding.Kind.ToString(), finding.TextPreview,
                "text-preview", new ReaderLocation { Path = path, SourceBlockKind = "content-safety" }, finding.Id,
                ReaderInspectionMetadataBuilder.Attributes(
                    ("nativeLocation", finding.Location), ("mechanism", finding.Kind.ToString()),
                    ("risk", finding.Risk.ToString()), ("evidence", finding.Evidence),
                    ("contentSha256", finding.ContentHash), ("textLength", finding.TextLength.ToString(CultureInfo.InvariantCulture)),
                    ("isInstructionLike", finding.IsInstructionLike ? "true" : "false"),
                    ("instructionSignals", string.Join(",", finding.InstructionSignals)),
                    ("sourceTextOffset", finding.SourceTextOffset?.ToString(CultureInfo.InvariantCulture)),
                    ("sourceTextLength", finding.SourceTextLength?.ToString(CultureInfo.InvariantCulture))));
        }
        foreach (var finding in report.TextIntegrityFindings) {
            builder.Add(OfficeDocumentInspectionCategories.TextIntegrity, finding.Kind.ToString(), finding.CodePoint.ToString("X", CultureInfo.InvariantCulture),
                "unicode-codepoint", new ReaderLocation { Path = path, SourceBlockKind = "text-integrity" },
                attributes: ReaderInspectionMetadataBuilder.Attributes(
                    ("nativeLocation", finding.Location), ("risk", finding.Risk.ToString()),
                    ("sourceTextOffset", finding.TextOffset.ToString(CultureInfo.InvariantCulture)),
                    ("sourceTextLength", finding.TextLength.ToString(CultureInfo.InvariantCulture))));
        }
        foreach (string diagnostic in report.Diagnostics)
            builder.Diagnostic("INSPECTION_CONTENT_SAFETY_BOUNDARY", diagnostic, new ReaderLocation { Path = path });
        builder.Coverage("content-safety", "bounded-heuristics", "Native visibility, non-primary stories and Unicode; previews remain bounded and are not full source text.", new ReaderLocation { Path = path });
    }

    internal static void AddPayloads(ReaderInspectionMetadataBuilder builder, IReadOnlyList<OfficeEmbeddedPayloadInfo> payloads, string? path) {
        foreach (OfficeEmbeddedPayloadInfo payload in payloads) {
            builder.Add(OfficeDocumentInspectionCategories.EmbeddedPayload, payload.SuggestedFileName, payload.PartUri,
                "package-part", new ReaderLocation { Path = path, SourceBlockKind = "embedded-payload" }, payload.Id,
                ReaderInspectionMetadataBuilder.Attributes(
                    ("kind", payload.Kind.ToString()), ("ownerPartUri", payload.OwnerPartUri),
                    ("partUri", payload.PartUri), ("relationshipId", payload.RelationshipId),
                    ("contentType", payload.ContentType), ("lengthBytes", payload.Length.ToString(CultureInfo.InvariantCulture)),
                    ("sha256", payload.Sha256), ("executed", "false")));
        }
        builder.Coverage("embedded-payloads", "complete", "Descriptors in the loaded package model; payloads were not executed or recursively interpreted.", new ReaderLocation { Path = path });
    }

    internal static void AddVba(ReaderInspectionMetadataBuilder builder, OfficeVbaProjectInfo? vba, string? path) {
        if (vba != null) builder.Add(OfficeDocumentInspectionCategories.VbaProject, "VBA project", vba.PartUri, "package-part",
            new ReaderLocation { Path = path, SourceBlockKind = "vba-project" }, vba.PartUri,
            ReaderInspectionMetadataBuilder.Attributes(
                ("partUri", vba.PartUri), ("contentType", vba.ContentType),
                ("lengthBytes", vba.Length.ToString(CultureInfo.InvariantCulture)),
                ("sha256", vba.Sha256), ("moduleNames", string.Join(",", vba.ModuleNames)),
                ("hasSignatureRelationship", vba.HasSignature ? "true" : "false"), ("executed", "false")));
        builder.Coverage("vba-project", "metadata-only", "Module names are best effort; a signature relationship is not signature validation.", new ReaderLocation { Path = path });
    }
}
