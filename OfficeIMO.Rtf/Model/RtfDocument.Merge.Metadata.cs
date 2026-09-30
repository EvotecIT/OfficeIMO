namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    private void ImportMergeMetadata(RtfDocument source, RtfConversionReport report, bool invalidateStatistics) {
        int fileId = _fileReferences.Count == 0 ? 0 : checked(_fileReferences.Max(item => item.Id) + 1);
        foreach (RtfFileReference file in source.FileReferences) {
            file.Id = fileId++;
            _fileReferences.Add(file);
        }
        int namespaceId = _xmlNamespaces.Count == 0 ? 0 : checked(_xmlNamespaces.Max(item => item.Id) + 1);
        foreach (RtfXmlNamespace ns in source.XmlNamespaces) {
            ns.Id = namespaceId++;
            _xmlNamespaces.Add(ns);
        }
        foreach (RtfUserProperty property in source.UserProperties) {
            RtfUserProperty? existing = _userProperties.FirstOrDefault(item => string.Equals(item.Name, property.Name, StringComparison.OrdinalIgnoreCase));
            if (existing == null) _userProperties.Add(property);
            else if (existing.TypeCode != property.TypeCode || existing.StaticValue != property.StaticValue || existing.LinkedValue != property.LinkedValue)
                ReportMergeMetadataConflict(report, "UserProperties/" + property.Name);
        }
        foreach (RtfDocumentVariable variable in source.DocumentVariables) {
            RtfDocumentVariable? existing = _documentVariables.FirstOrDefault(item => string.Equals(item.Name, variable.Name, StringComparison.OrdinalIgnoreCase));
            if (existing == null) _documentVariables.Add(variable);
            else if (existing.Value != variable.Value) ReportMergeMetadataConflict(report, "DocumentVariables/" + variable.Name);
        }
        foreach (int rsid in source.RevisionSaveIds) if (!_revisionSaveIds.Contains(rsid)) _revisionSaveIds.Add(rsid);
        RevisionRootSaveId = MergeMetadataValue(RevisionRootSaveId, source.RevisionRootSaveId, "RevisionRootSaveId", report);
        RtfDocumentInfo info = source.Info;
        Info.Generator = MergeMetadataValue(Info.Generator, info.Generator, "Info/Generator", report);
        Info.Title = MergeMetadataValue(Info.Title, info.Title, "Info/Title", report);
        Info.Subject = MergeMetadataValue(Info.Subject, info.Subject, "Info/Subject", report);
        Info.Author = MergeMetadataValue(Info.Author, info.Author, "Info/Author", report);
        Info.Manager = MergeMetadataValue(Info.Manager, info.Manager, "Info/Manager", report);
        Info.Company = MergeMetadataValue(Info.Company, info.Company, "Info/Company", report);
        Info.Operator = MergeMetadataValue(Info.Operator, info.Operator, "Info/Operator", report);
        Info.Category = MergeMetadataValue(Info.Category, info.Category, "Info/Category", report);
        Info.Keywords = MergeMetadataValue(Info.Keywords, info.Keywords, "Info/Keywords", report);
        Info.Comments = MergeMetadataValue(Info.Comments, info.Comments, "Info/Comments", report);
        Info.HyperlinkBase = MergeMetadataValue(Info.HyperlinkBase, info.HyperlinkBase, "Info/HyperlinkBase", report);
        Info.Created = MergeMetadataValue(Info.Created, info.Created, "Info/Created", report);
        Info.Revised = MergeMetadataValue(Info.Revised, info.Revised, "Info/Revised", report);
        Info.Printed = MergeMetadataValue(Info.Printed, info.Printed, "Info/Printed", report);
        Info.BackedUp = MergeMetadataValue(Info.BackedUp, info.BackedUp, "Info/BackedUp", report);
        Info.EditingMinutes = MergeMetadataValue(Info.EditingMinutes, info.EditingMinutes, "Info/EditingMinutes", report);
        if (invalidateStatistics) {
            bool hadCounts = Info.NumberOfPages.HasValue || Info.NumberOfWords.HasValue || Info.NumberOfCharacters.HasValue || Info.NumberOfCharactersWithSpaces.HasValue ||
                info.NumberOfPages.HasValue || info.NumberOfWords.HasValue || info.NumberOfCharacters.HasValue || info.NumberOfCharactersWithSpaces.HasValue;
            Info.NumberOfPages = null;
            Info.NumberOfWords = null;
            Info.NumberOfCharacters = null;
            Info.NumberOfCharactersWithSpaces = null;
            if (hadCounts) report.Add(RtfConversionSeverity.Warning, "RtfMergeStatisticsInvalidated",
                "Source and destination page, word, and character counts are omitted because they do not describe the combined content. Recalculate them in the consuming application.",
                RtfConversionAction.Omitted, sourcePath: "Info", feature: "DocumentStatistics");
        } else {
            Info.NumberOfPages = MergeMetadataValue(Info.NumberOfPages, info.NumberOfPages, "Info/NumberOfPages", report);
            Info.NumberOfWords = MergeMetadataValue(Info.NumberOfWords, info.NumberOfWords, "Info/NumberOfWords", report);
            Info.NumberOfCharacters = MergeMetadataValue(Info.NumberOfCharacters, info.NumberOfCharacters, "Info/NumberOfCharacters", report);
            Info.NumberOfCharactersWithSpaces = MergeMetadataValue(Info.NumberOfCharactersWithSpaces, info.NumberOfCharactersWithSpaces, "Info/NumberOfCharactersWithSpaces", report);
        }
        Info.InternalVersion = MergeMetadataValue(Info.InternalVersion, info.InternalVersion, "Info/InternalVersion", report);
    }

    private static T MergeMetadataValue<T>(T destination, T source, string path, RtfConversionReport report) {
        if (source is null) return destination;
        if (destination is null) return source;
        if (!EqualityComparer<T>.Default.Equals(destination, source)) ReportMergeMetadataConflict(report, path);
        return destination;
    }

    private static void ReportMergeMetadataConflict(RtfConversionReport report, string path) =>
        report.Add(RtfConversionSeverity.Warning, "RtfMergeMetadataConflict",
            "The destination's document metadata was retained because the source declared a conflicting value.",
            RtfConversionAction.Omitted, sourcePath: path, feature: "DocumentMetadata");
}
