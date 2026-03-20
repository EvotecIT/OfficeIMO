using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void CleanupProtectionArtifacts() {
            var worksheet = _worksheetPart.Worksheet;

            var protections = worksheet.Elements<SheetProtection>().ToList();
            SheetProtection? primaryProtection = protections.FirstOrDefault();
            if (primaryProtection != null) {
                foreach (var duplicateProtection in protections.Skip(1)) {
                    duplicateProtection.Remove();
                }

                primaryProtection.Sheet = true;
            }

            var protectedRanges = worksheet.Elements<ProtectedRanges>().FirstOrDefault();
            if (protectedRanges == null) {
                return;
            }

            if (primaryProtection == null) {
                worksheet.RemoveChild(protectedRanges);
                return;
            }

            var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var protectedRange in protectedRanges.Elements<ProtectedRange>().ToList()) {
                string? name = protectedRange.Name?.Value;
                string? range = protectedRange.SequenceOfReferences?.InnerText;
                if (string.IsNullOrWhiteSpace(name) ||
                    string.IsNullOrWhiteSpace(range) ||
                    !seenNames.Add(name!) ||
                    !TryValidateProtectedRangeReferences(range!)) {
                    protectedRange.Remove();
                }
            }

            if (!protectedRanges.Elements<ProtectedRange>().Any()) {
                worksheet.RemoveChild(protectedRanges);
            }
        }

        private static bool TryValidateProtectedRangeReferences(string referenceList) {
            foreach (var candidate in referenceList.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) {
                if (string.IsNullOrWhiteSpace(candidate)) {
                    return false;
                }

                if (!TryParseProtectedRangeReference(candidate)) {
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseProtectedRangeReference(string candidate) {
            string normalized = candidate.Trim();
            int bangIndex = normalized.LastIndexOf('!');
            if (bangIndex >= 0) {
                normalized = normalized[(bangIndex + 1)..];
            }

            if (A1.TryParseRange(normalized, out _, out _, out _, out _)) {
                return true;
            }

            var cell = A1.ParseCellRef(normalized);
            return cell.Row > 0 && cell.Col > 0;
        }
    }
}
