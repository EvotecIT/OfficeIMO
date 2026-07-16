using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects XLSB worksheet hyperlinks and validates their preservation contract.</summary>
    internal static class XlsbWorksheetHyperlinkProjector {
        internal static void Apply(ExcelSheet sheet, XlsbWorksheet sourceSheet) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (sourceSheet == null) throw new ArgumentNullException(nameof(sourceSheet));
            if (sourceSheet.Hyperlinks.Count == 0) return;

            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            var hyperlinks = new Hyperlinks();
            foreach (XlsbHyperlink source in sourceSheet.Hyperlinks) {
                var hyperlink = new Hyperlink {
                    Reference = source.Range.ToA1Reference()
                };
                if (source.IsExternal) {
                    Uri target = new Uri(source.ExternalTarget!, UriKind.RelativeOrAbsolute);
                    HyperlinkRelationship relationship = sheet.WorksheetPart.AddHyperlinkRelationship(
                        target,
                        isExternal: true,
                        source.RelationshipId);
                    hyperlink.Id = relationship.Id;
                }
                if (!string.IsNullOrEmpty(source.Location)) hyperlink.Location = source.Location;
                if (!string.IsNullOrEmpty(source.Tooltip)) hyperlink.Tooltip = source.Tooltip;
                if (!string.IsNullOrEmpty(source.Display)) hyperlink.Display = source.Display;
                hyperlinks.Append(hyperlink);
            }

            worksheet.Append(hyperlinks);
            sheet.EnsureWorksheetElementOrder();
            worksheet.Save();
        }

        internal static void ValidateUnchanged(ExcelSheet sheet, XlsbWorksheet sourceSheet) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            Hyperlinks[] containers = worksheet.Elements<Hyperlinks>().ToArray();
            if (sourceSheet.Hyperlinks.Count == 0) {
                if (containers.Length != 0 || sheet.WorksheetPart.HyperlinkRelationships.Any()) {
                    ThrowMutation(sheet);
                }
                return;
            }
            if (containers.Length != 1) ThrowMutation(sheet);

            Hyperlink[] actualLinks = containers[0].Elements<Hyperlink>().ToArray();
            HyperlinkRelationship[] relationships = sheet.WorksheetPart.HyperlinkRelationships.ToArray();
            int expectedExternalCount = sourceSheet.Hyperlinks.Count(link => link.IsExternal);
            if (actualLinks.Length != sourceSheet.Hyperlinks.Count || relationships.Length != expectedExternalCount) {
                ThrowMutation(sheet);
            }
            var relationshipsById = relationships.ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
            for (int index = 0; index < actualLinks.Length; index++) {
                Hyperlink actual = actualLinks[index];
                XlsbHyperlink expected = sourceSheet.Hyperlinks[index];
                if (actual.HasChildren
                    || actual.GetAttributes().Any(attribute => !IsSupportedAttribute(attribute.LocalName))
                    || !string.Equals(actual.Reference?.Value, expected.Range.ToA1Reference(), StringComparison.Ordinal)
                    || !MatchesOptional(actual.Location?.Value, expected.Location)
                    || !MatchesOptional(actual.Tooltip?.Value, expected.Tooltip)
                    || !MatchesOptional(actual.Display?.Value, expected.Display)) {
                    ThrowMutation(sheet);
                }

                string? relationshipId = actual.Id?.Value;
                if (expected.IsExternal) {
                    if (string.IsNullOrWhiteSpace(relationshipId)
                        || !relationshipsById.TryGetValue(relationshipId!, out HyperlinkRelationship? relationship)
                        || !string.Equals(relationship.Uri.OriginalString, expected.ExternalTarget, StringComparison.Ordinal)) {
                        ThrowMutation(sheet);
                    }
                } else if (!string.IsNullOrWhiteSpace(relationshipId)) {
                    ThrowMutation(sheet);
                }
            }
        }

        private static bool IsSupportedAttribute(string localName) {
            return localName == "ref"
                || localName == "id"
                || localName == "location"
                || localName == "tooltip"
                || localName == "display";
        }

        private static bool MatchesOptional(string? actual, string expected) {
            return string.Equals(actual ?? string.Empty, expected, StringComparison.Ordinal);
        }

        private static void ThrowMutation(ExcelSheet sheet) {
            throw new NotSupportedException($"Native XLSB rewriting preserves but cannot modify hyperlinks on worksheet '{sheet.Name}'. Save as .xlsx to retain that change.");
        }
    }
}
