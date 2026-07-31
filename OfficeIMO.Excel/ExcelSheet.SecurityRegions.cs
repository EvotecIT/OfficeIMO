using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>Excel error indicators that may be ignored for one or more cell regions.</summary>
    [Flags]
    public enum ExcelIgnoredErrorKind {
        /// <summary>No ignored error indicators.</summary>
        None = 0,
        /// <summary>Evaluated formula errors.</summary>
        EvaluationError = 1,
        /// <summary>References to empty cells.</summary>
        EmptyCellReference = 2,
        /// <summary>Numbers stored as text.</summary>
        NumberStoredAsText = 4,
        /// <summary>Inconsistent formulas across a range.</summary>
        FormulaRange = 8,
        /// <summary>Formula-specific warnings.</summary>
        Formula = 16,
        /// <summary>Two-digit text years.</summary>
        TwoDigitTextYear = 32,
        /// <summary>Unlocked formula cells.</summary>
        UnlockedFormula = 64,
        /// <summary>Values inconsistent with list data validation.</summary>
        ListDataValidation = 128,
        /// <summary>Calculated-column warnings.</summary>
        CalculatedColumn = 256
    }

    /// <summary>Format-neutral description of an allowed-edit range on a protected worksheet.</summary>
    public sealed class ExcelAllowedEditRangeInfo {
        internal ExcelAllowedEditRangeInfo(string name, IReadOnlyList<string> ranges, bool passwordProtected, string? securityDescriptor) {
            Name = name;
            Ranges = ranges;
            IsPasswordProtected = passwordProtected;
            SecurityDescriptor = securityDescriptor;
        }

        /// <summary>Unique worksheet-scoped name.</summary>
        public string Name { get; }

        /// <summary>Normalized A1 cells and ranges.</summary>
        public IReadOnlyList<string> Ranges { get; }

        /// <summary>Whether package metadata includes a password hash.</summary>
        public bool IsPasswordProtected { get; }

        /// <summary>Optional security descriptor used by Excel permission workflows.</summary>
        public string? SecurityDescriptor { get; }
    }

    /// <summary>Format-neutral description of ignored-error metadata.</summary>
    public sealed class ExcelIgnoredErrorRegionInfo {
        internal ExcelIgnoredErrorRegionInfo(IReadOnlyList<string> ranges, ExcelIgnoredErrorKind errors) {
            Ranges = ranges;
            Errors = errors;
        }

        /// <summary>Normalized A1 cells and ranges.</summary>
        public IReadOnlyList<string> Ranges { get; }

        /// <summary>Error indicators ignored in the regions.</summary>
        public ExcelIgnoredErrorKind Errors { get; }
    }

    public partial class ExcelSheet {
        /// <summary>Lists allowed-edit regions without exposing Open XML types.</summary>
        public IReadOnlyList<ExcelAllowedEditRangeInfo> GetAllowedEditRanges() {
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var result = new List<ExcelAllowedEditRangeInfo>();
                ProtectedRanges? ranges = WorksheetRoot.GetFirstChild<ProtectedRanges>();
                if (ranges == null) return (IReadOnlyList<ExcelAllowedEditRangeInfo>)result;
                foreach (ProtectedRange range in ranges.Elements<ProtectedRange>()) {
                    string? name = range.Name?.Value;
                    IReadOnlyList<string> references = ParseRegionReferences(range.SequenceOfReferences?.InnerText);
                    if (string.IsNullOrWhiteSpace(name) || references.Count == 0) continue;
                    result.Add(new ExcelAllowedEditRangeInfo(
                        name!,
                        references,
                        !string.IsNullOrWhiteSpace(range.Password?.Value)
                            || !string.IsNullOrWhiteSpace(range.HashValue?.Value),
                        range.SecurityDescriptor?.Value));
                }
                return new ReadOnlyCollection<ExcelAllowedEditRangeInfo>(result);
            });
        }

        /// <summary>Adds or replaces a named allowed-edit region on a protected worksheet.</summary>
        public void SetAllowedEditRange(
            string name,
            IEnumerable<string> ranges,
            string? password = null,
            string? securityDescriptor = null) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            IReadOnlyList<string> normalized = NormalizeRegionReferences(ranges);
            if (!IsProtected) {
                throw new InvalidOperationException("Allowed-edit ranges require worksheet protection. Call Protect first.");
            }

            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
                ProtectedRanges container = worksheet.GetFirstChild<ProtectedRanges>() ?? worksheet.AppendChild(new ProtectedRanges());
                ProtectedRange? existing = container.Elements<ProtectedRange>()
                    .FirstOrDefault(item => string.Equals(item.Name?.Value, name.Trim(), StringComparison.OrdinalIgnoreCase));
                ProtectedRange target = existing ?? container.AppendChild(new ProtectedRange());
                target.Name = name.Trim();
                target.SequenceOfReferences = new ListValue<StringValue> { InnerText = string.Join(" ", normalized) };
                string? hash = ExcelProtectionHash.ResolveLegacyHash(password, null);
                target.Password = hash;
                if (hash == null) target.RemoveAttribute("password", string.Empty);
                target.SecurityDescriptor = string.IsNullOrWhiteSpace(securityDescriptor) ? null : securityDescriptor;
                if (string.IsNullOrWhiteSpace(securityDescriptor)) target.RemoveAttribute("securityDescriptor", string.Empty);
                EnsureWorksheetElementOrder();
                worksheet.Save();
            });
        }

        /// <summary>Removes a named allowed-edit region.</summary>
        public bool RemoveAllowedEditRange(string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            bool removed = false;
            WriteLock(() => {
                ProtectedRanges? container = WorksheetRoot.GetFirstChild<ProtectedRanges>();
                ProtectedRange? range = container?.Elements<ProtectedRange>()
                    .FirstOrDefault(item => string.Equals(item.Name?.Value, name.Trim(), StringComparison.OrdinalIgnoreCase));
                if (range == null) return;
                range.Remove();
                if (!container!.Elements<ProtectedRange>().Any()) container.Remove();
                WorksheetRoot.Save();
                removed = true;
            });
            return removed;
        }

        /// <summary>Lists standard ignored-error regions without exposing Open XML types.</summary>
        public IReadOnlyList<ExcelIgnoredErrorRegionInfo> GetIgnoredErrorRegions() {
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var result = new List<ExcelIgnoredErrorRegionInfo>();
                IgnoredErrors? container = WorksheetRoot.GetFirstChild<IgnoredErrors>();
                if (container == null) return (IReadOnlyList<ExcelIgnoredErrorRegionInfo>)result;
                foreach (IgnoredError error in container.Elements<IgnoredError>()) {
                    IReadOnlyList<string> references = ParseRegionReferences(error.SequenceOfReferences?.InnerText);
                    if (references.Count == 0) continue;
                    result.Add(new ExcelIgnoredErrorRegionInfo(references, ReadIgnoredErrorKind(error)));
                }
                return new ReadOnlyCollection<ExcelIgnoredErrorRegionInfo>(result);
            });
        }

        /// <summary>Adds ignored-error metadata for one or more regions.</summary>
        public void AddIgnoredErrorRegion(IEnumerable<string> ranges, ExcelIgnoredErrorKind errors) {
            IReadOnlyList<string> normalized = NormalizeRegionReferences(ranges);
            if (errors == ExcelIgnoredErrorKind.None) throw new ArgumentOutOfRangeException(nameof(errors));
            WriteLock(() => {
                IgnoredErrors container = WorksheetRoot.GetFirstChild<IgnoredErrors>() ?? WorksheetRoot.AppendChild(new IgnoredErrors());
                var ignored = new IgnoredError {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = string.Join(" ", normalized) }
                };
                ApplyIgnoredErrorKind(ignored, errors);
                container.Append(ignored);
                EnsureWorksheetElementOrder();
                WorksheetRoot.Save();
            });
        }

        /// <summary>Removes every standard ignored-error region overlapping the supplied A1 cell or range.</summary>
        public int RemoveIgnoredErrorRegions(string range) {
            ExcelReference target = ExcelReference.Parse(range);
            int removed = 0;
            WriteLock(() => {
                IgnoredErrors? container = WorksheetRoot.GetFirstChild<IgnoredErrors>();
                if (container == null) return;
                foreach (IgnoredError ignored in container.Elements<IgnoredError>().ToList()) {
                    bool overlaps = ParseRegionReferences(ignored.SequenceOfReferences?.InnerText)
                        .Select(reference => ExcelReference.Parse(reference))
                        .Any(reference => reference.Intersects(target));
                    if (!overlaps) continue;
                    ignored.Remove();
                    removed++;
                }
                if (!container.Elements<IgnoredError>().Any()) container.Remove();
                WorksheetRoot.Save();
            });
            return removed;
        }

        /// <summary>Clears all standard ignored-error regions.</summary>
        public void ClearIgnoredErrorRegions() {
            WriteLock(() => {
                WorksheetRoot.GetFirstChild<IgnoredErrors>()?.Remove();
                WorksheetRoot.Save();
            });
        }

        private static IReadOnlyList<string> NormalizeRegionReferences(IEnumerable<string> ranges) {
            if (ranges == null) throw new ArgumentNullException(nameof(ranges));
            var result = new List<string>();
            foreach (string raw in ranges) {
                ExcelReference reference = ExcelReference.Parse(raw);
                if (reference.Kind != ExcelReferenceKind.Cell && reference.Kind != ExcelReferenceKind.Range) {
                    throw new ArgumentException("Security regions require cell or rectangular range references.", nameof(ranges));
                }
                string normalized = reference.ToString(ExcelReferenceStyle.A1);
                int separator = normalized.LastIndexOf('!');
                if (separator >= 0) normalized = normalized.Substring(separator + 1);
                if (!result.Contains(normalized, StringComparer.OrdinalIgnoreCase)) result.Add(normalized);
            }
            if (result.Count == 0) throw new ArgumentException("At least one security region is required.", nameof(ranges));
            return new ReadOnlyCollection<string>(result);
        }

        private static IReadOnlyList<string> ParseRegionReferences(string? references) {
            if (string.IsNullOrWhiteSpace(references)) return System.Array.Empty<string>();
            var result = new List<string>();
            foreach (string part in references!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                if (ExcelReference.TryParse(part, out ExcelReference? parsed)
                    && (parsed!.Kind == ExcelReferenceKind.Cell || parsed.Kind == ExcelReferenceKind.Range)) {
                    result.Add(parsed.ToString(ExcelReferenceStyle.A1));
                }
            }
            return new ReadOnlyCollection<string>(result);
        }

        private static ExcelIgnoredErrorKind ReadIgnoredErrorKind(IgnoredError error) {
            ExcelIgnoredErrorKind result = ExcelIgnoredErrorKind.None;
            if (error.EvalError?.Value == true) result |= ExcelIgnoredErrorKind.EvaluationError;
            if (error.EmptyCellReference?.Value == true) result |= ExcelIgnoredErrorKind.EmptyCellReference;
            if (error.NumberStoredAsText?.Value == true) result |= ExcelIgnoredErrorKind.NumberStoredAsText;
            if (error.FormulaRange?.Value == true) result |= ExcelIgnoredErrorKind.FormulaRange;
            if (error.Formula?.Value == true) result |= ExcelIgnoredErrorKind.Formula;
            if (error.TwoDigitTextYear?.Value == true) result |= ExcelIgnoredErrorKind.TwoDigitTextYear;
            if (error.UnlockedFormula?.Value == true) result |= ExcelIgnoredErrorKind.UnlockedFormula;
            if (error.ListDataValidation?.Value == true) result |= ExcelIgnoredErrorKind.ListDataValidation;
            if (error.CalculatedColumn?.Value == true) result |= ExcelIgnoredErrorKind.CalculatedColumn;
            return result;
        }

        private static void ApplyIgnoredErrorKind(IgnoredError error, ExcelIgnoredErrorKind value) {
            error.EvalError = value.HasFlag(ExcelIgnoredErrorKind.EvaluationError);
            error.EmptyCellReference = value.HasFlag(ExcelIgnoredErrorKind.EmptyCellReference);
            error.NumberStoredAsText = value.HasFlag(ExcelIgnoredErrorKind.NumberStoredAsText);
            error.FormulaRange = value.HasFlag(ExcelIgnoredErrorKind.FormulaRange);
            error.Formula = value.HasFlag(ExcelIgnoredErrorKind.Formula);
            error.TwoDigitTextYear = value.HasFlag(ExcelIgnoredErrorKind.TwoDigitTextYear);
            error.UnlockedFormula = value.HasFlag(ExcelIgnoredErrorKind.UnlockedFormula);
            error.ListDataValidation = value.HasFlag(ExcelIgnoredErrorKind.ListDataValidation);
            error.CalculatedColumn = value.HasFlag(ExcelIgnoredErrorKind.CalculatedColumn);
        }
    }
}
