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
        RtfDocumentInfo info = source.Properties;
        Properties.Generator = MergeMetadataValue(Properties.Generator, info.Generator, "Info/Generator", report);
        Properties.Title = MergeMetadataValue(Properties.Title, info.Title, "Info/Title", report);
        Properties.Subject = MergeMetadataValue(Properties.Subject, info.Subject, "Info/Subject", report);
        Properties.Author = MergeMetadataValue(Properties.Author, info.Author, "Info/Author", report);
        Properties.Manager = MergeMetadataValue(Properties.Manager, info.Manager, "Info/Manager", report);
        Properties.Company = MergeMetadataValue(Properties.Company, info.Company, "Info/Company", report);
        Properties.Operator = MergeMetadataValue(Properties.Operator, info.Operator, "Info/Operator", report);
        Properties.Category = MergeMetadataValue(Properties.Category, info.Category, "Info/Category", report);
        Properties.Keywords = MergeMetadataValue(Properties.Keywords, info.Keywords, "Info/Keywords", report);
        Properties.Comments = MergeMetadataValue(Properties.Comments, info.Comments, "Info/Comments", report);
        Properties.HyperlinkBase = MergeMetadataValue(Properties.HyperlinkBase, info.HyperlinkBase, "Info/HyperlinkBase", report);
        Properties.Created = MergeMetadataValue(Properties.Created, info.Created, "Info/Created", report);
        Properties.Revised = MergeMetadataValue(Properties.Revised, info.Revised, "Info/Revised", report);
        Properties.Printed = MergeMetadataValue(Properties.Printed, info.Printed, "Info/Printed", report);
        Properties.BackedUp = MergeMetadataValue(Properties.BackedUp, info.BackedUp, "Info/BackedUp", report);
        Properties.EditingMinutes = MergeMetadataValue(Properties.EditingMinutes, info.EditingMinutes, "Info/EditingMinutes", report);
        if (invalidateStatistics) {
            bool hadCounts = Properties.NumberOfPages.HasValue || Properties.NumberOfWords.HasValue || Properties.NumberOfCharacters.HasValue || Properties.NumberOfCharactersWithSpaces.HasValue ||
                info.NumberOfPages.HasValue || info.NumberOfWords.HasValue || info.NumberOfCharacters.HasValue || info.NumberOfCharactersWithSpaces.HasValue;
            Properties.NumberOfPages = null;
            Properties.NumberOfWords = null;
            Properties.NumberOfCharacters = null;
            Properties.NumberOfCharactersWithSpaces = null;
            if (hadCounts) report.Add(RtfConversionSeverity.Warning, "RtfMergeStatisticsInvalidated",
                "Source and destination page, word, and character counts are omitted because they do not describe the combined content. Recalculate them in the consuming application.",
                RtfConversionAction.Omitted, sourcePath: "Info", feature: "DocumentStatistics");
        } else {
            Properties.NumberOfPages = MergeMetadataValue(Properties.NumberOfPages, info.NumberOfPages, "Info/NumberOfPages", report);
            Properties.NumberOfWords = MergeMetadataValue(Properties.NumberOfWords, info.NumberOfWords, "Info/NumberOfWords", report);
            Properties.NumberOfCharacters = MergeMetadataValue(Properties.NumberOfCharacters, info.NumberOfCharacters, "Info/NumberOfCharacters", report);
            Properties.NumberOfCharactersWithSpaces = MergeMetadataValue(Properties.NumberOfCharactersWithSpaces, info.NumberOfCharactersWithSpaces, "Info/NumberOfCharactersWithSpaces", report);
        }
        Properties.InternalVersion = MergeMetadataValue(Properties.InternalVersion, info.InternalVersion, "Info/InternalVersion", report);
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
