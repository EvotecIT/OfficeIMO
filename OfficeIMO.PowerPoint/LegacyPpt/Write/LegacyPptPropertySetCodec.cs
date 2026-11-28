using System.Globalization;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    /// <summary>
    /// Projects and writes the OLE property-set streams shared by classic Office
    /// documents while keeping PowerPoint's public property surface package-based.
    /// </summary>
    internal static partial class LegacyPptPropertySetCodec {
        private static readonly ISet<uint> SupportedSummaryPropertyIds =
            new HashSet<uint> { 1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 18 };
        private static readonly ISet<uint> SupportedDocumentSummaryPropertyIds =
            new HashSet<uint> { 1, 2, 3, 7, 8, 9, 14, 15 };

        /// <summary>
        /// Writes the binary property sets into the presentation's document properties: in its package while it is
        /// held there, through the SDK otherwise.
        /// </summary>
        internal static void Apply(PowerPointDocument presentation,
            LegacyPptPackage package) {
            ArgumentNullException.ThrowIfNull(presentation);
            ArgumentNullException.ThrowIfNull(package);
            IReadOnlyList<OfficeOlePropertySection>? summarySections =
                TryReadSections(package, OfficeOlePropertySetWriter
                    .SummaryInformationStreamName, out IReadOnlyList<
                    OfficeOlePropertySection> summary) ? summary : null;
            IReadOnlyList<OfficeOlePropertySection>? documentSummarySections =
                TryReadSections(package, OfficeOlePropertySetWriter
                    .DocumentSummaryInformationStreamName, out IReadOnlyList<
                    OfficeOlePropertySection> documentSummary) ? documentSummary : null;
            if (presentation.EditablePackage is { } editable) {
                var properties = new PackagePropertyWriter(editable);
                properties.ClearMapped();
                if (summarySections != null) ApplySummary(properties, summarySections);
                if (documentSummarySections != null) ApplyDocumentSummary(properties, documentSummarySections);
                presentation.MarkPackageEdited();
                return;
            }
            ApplyThroughSdk(presentation, summarySections, documentSummarySections);
        }

        internal static void ApplyReadOnlyDateOverrides(
            PowerPointDocument presentation, LegacyPptPackage package) {
            DateTime? created = null;
            DateTime? modified = null;
            DateTime? lastPrinted = null;
            if (TryReadSections(package, OfficeOlePropertySetWriter
                    .SummaryInformationStreamName, out IReadOnlyList<
                    OfficeOlePropertySection> sections)) {
                OfficeOlePropertySection? summary = sections.FirstOrDefault(
                    item => item.FormatId == OfficeOlePropertySetWriter
                        .SummaryInformationFormatId);
                if (summary != null) {
                    lastPrinted = ReadDate(summary, 11);
                    created = ReadDate(summary, 12);
                    modified = ReadDate(summary, 13);
                }
            }
            presentation.Properties
                .SetReadOnlyLegacyDateOverrides(created, modified,
                    lastPrinted);
        }

        internal static LegacyPptPropertySetProjection CreateProjection(
            LegacyPptPackageView view, LegacyPptPackage package) {
            ArgumentNullException.ThrowIfNull(view);
            ArgumentNullException.ThrowIfNull(package);
            bool summarySafe = CanRewriteSummary(package);
            bool documentSummarySafe = CanRewriteDocumentSummary(package);
            return new LegacyPptPropertySetProjection(
                CreateSummaryFingerprint(view),
                CreateDocumentSummaryFingerprint(view),
                CreateUnmappedCoreFingerprint(view), summarySafe,
                documentSummarySafe);
        }

        internal static bool TryCreateFreshStreams(
            PowerPointDocument presentation,
            out IReadOnlyList<OfficeCompoundStream> streams,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(presentation);
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(presentation);
            return TryCreateFreshStreams(view, out streams, out reason);
        }

        internal static bool TryCreateFreshStreams(
            LegacyPptPackageView view,
            out IReadOnlyList<OfficeCompoundStream> streams,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(view);
            var result = new List<OfficeCompoundStream>(2);
            streams = result;
            reason = null;
            var properties = PackageProperties.Read(view);
            byte[]? summary = CreateSummaryInformation(properties);
            if (summary != null) {
                result.Add(new OfficeCompoundStream(OfficeOlePropertySetWriter
                    .SummaryInformationStreamName, summary));
            }
            if (!TryCreateDocumentSummaryInformation(properties,
                    out byte[]? documentSummary, out reason)) {
                return false;
            }
            if (documentSummary != null) {
                result.Add(new OfficeCompoundStream(OfficeOlePropertySetWriter
                    .DocumentSummaryInformationStreamName,
                    documentSummary));
            }
            return true;
        }

        internal static bool TryBuildReplacementStreams(
            LegacyPptPackageView view,
            LegacyPptPropertySetProjection projection,
            out IReadOnlyDictionary<string, byte[]> streams) {
            ArgumentNullException.ThrowIfNull(view);
            var result = new Dictionary<string, byte[]>(
                StringComparer.OrdinalIgnoreCase);
            streams = result;
            if (!projection.UnmappedCoreMatches(view)) return false;
            bool summaryChanged = !projection.SummaryMatches(view);
            bool documentSummaryChanged = !projection.DocumentSummaryMatches(view);
            var properties = PackageProperties.Read(view);
            if (summaryChanged) {
                if (!projection.CanRewriteSummary) return false;
                result[OfficeOlePropertySetWriter.SummaryInformationStreamName]
                    = CreateSummaryInformation(properties)
                    ?? CreateEmptyPropertySet(
                        OfficeOlePropertySetWriter.SummaryInformationFormatId);
            }
            if (documentSummaryChanged) {
                if (!projection.CanRewriteDocumentSummary
                    || !TryCreateDocumentSummaryInformation(properties,
                        out byte[]? documentSummary, out _)) {
                    return false;
                }
                result[OfficeOlePropertySetWriter
                    .DocumentSummaryInformationStreamName]
                    = documentSummary ?? CreateEmptyPropertySet(
                        OfficeOlePropertySetWriter
                            .DocumentSummaryInformationFormatId);
            }
            return true;
        }

        internal static string CreateSummaryFingerprint(
            LegacyPptPackageView view) {
            var properties = PackageProperties.Read(view);
            return string.Join("\u001F", new[] {
                properties.Core("title"), properties.Core("subject"), properties.Core("creator"),
                properties.Core("keywords"), properties.Core("description"),
                properties.Core("lastModifiedBy"), properties.Core("revision"),
                DateText(properties.CoreDate("lastPrinted")), DateText(properties.CoreDate("created")),
                DateText(properties.CoreDate("modified")), properties.Application("Application"),
                properties.Application("TotalTime")
            }.Select(value => value ?? string.Empty));
        }

        internal static string CreateDocumentSummaryFingerprint(
            LegacyPptPackageView view) {
            var properties = PackageProperties.Read(view);
            string custom = LegacyPptMasterProjection.Markup(properties.Custom);
            return string.Join("\u001F", new[] {
                properties.Core("category"), properties.Application("PresentationFormat"),
                properties.Application("Slides"), properties.Application("Notes"),
                properties.Application("HiddenSlides"), properties.Application("Manager"),
                properties.Application("Company"), custom
            }.Select(value => value ?? string.Empty));
        }

        internal static string CreateUnmappedCoreFingerprint(
            LegacyPptPackageView view) {
            var properties = PackageProperties.Read(view);
            return string.Join("\u001F", new[] {
                properties.Core("contentStatus"), properties.Core("contentType"),
                properties.Core("identifier"), properties.Core("language"),
                properties.Core("version")
            }.Select(value => value ?? string.Empty));
        }

        private static void ApplySummary(IPropertyTarget properties,
            IReadOnlyList<OfficeOlePropertySection> sections) {
            OfficeOlePropertySection? section = sections.FirstOrDefault(item =>
                item.FormatId == OfficeOlePropertySetWriter
                    .SummaryInformationFormatId);
            if (section == null) return;
            properties.SetCore("title", ReadString(section, 2));
            properties.SetCore("subject", ReadString(section, 3));
            properties.SetCore("creator", ReadString(section, 4));
            properties.SetCore("keywords", ReadString(section, 5));
            properties.SetCore("description", ReadString(section, 6));
            properties.SetCore("lastModifiedBy", ReadString(section, 8));
            properties.SetCore("revision", ReadString(section, 9));
            properties.SetCoreDate("lastPrinted", ReadDate(section, 11));
            properties.SetCoreDate("created", ReadDate(section, 12));
            properties.SetCoreDate("modified", ReadDate(section, 13));
            properties.EnsureApplication();
            string? applicationName = ReadString(section, 18);
            if (applicationName != null) {
                properties.SetApplication("Application", applicationName);
            }
            string? totalTime = ReadTotalMinutes(section);
            if (totalTime != null) properties.SetApplication("TotalTime", totalTime);
        }

        private static string? ReadTotalMinutes(OfficeOlePropertySection section) {
            DateTime? duration = ReadDate(section, 10);
            if (!duration.HasValue) return null;
            long ticks = duration.Value.ToFileTimeUtc();
            return Math.Round(TimeSpan.FromTicks(ticks).TotalMinutes,
                MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);
        }

        private static void ApplyDocumentSummary(IPropertyTarget properties,
            IReadOnlyList<OfficeOlePropertySection> sections) {
            OfficeOlePropertySection? builtIn = sections.FirstOrDefault(item =>
                item.FormatId == OfficeOlePropertySetWriter
                    .DocumentSummaryInformationFormatId);
            if (builtIn != null) {
                properties.SetCore("category", ReadString(builtIn, 2));
                properties.EnsureApplication();
                SetApplicationText(properties, "PresentationFormat", ReadString(builtIn, 3));
                SetApplicationText(properties, "Slides", ReadIntegerText(builtIn, 7));
                SetApplicationText(properties, "Notes", ReadIntegerText(builtIn, 8));
                SetApplicationText(properties, "HiddenSlides", ReadIntegerText(builtIn, 9));
                SetApplicationText(properties, "Manager", ReadString(builtIn, 14));
                SetApplicationText(properties, "Company", ReadString(builtIn, 15));
            }
            OfficeOlePropertySection? custom = sections.FirstOrDefault(item =>
                item.FormatId == OfficeOlePropertySetWriter
                    .UserDefinedPropertiesFormatId);
            if (custom != null) {
                IReadOnlyList<CustomPropertyValue> values = ReadCustomProperties(custom);
                if (values.Count > 0) properties.SetCustom(values);
            }
        }

        private static void SetApplicationText(IPropertyTarget properties,
            string name, string? value) {
            if (value != null) properties.SetApplication(name, value);
        }

        /// <summary>The user-defined properties of a binary section, as custom document properties.</summary>
        private static IReadOnlyList<CustomPropertyValue> ReadCustomProperties(
            OfficeOlePropertySection section) {
            var values = new List<CustomPropertyValue>();
            int propertyId = 2;
            foreach (KeyValuePair<uint, string> name in section.Dictionary
                         .Where(pair => pair.Key > 1)
                         .OrderBy(pair => pair.Key)) {
                if (!section.Properties.TryGetValue(name.Key,
                        out OfficeOlePropertyValue? source)) {
                    continue;
                }
                // The id is taken for every stored value, written or not.
                int id = propertyId++;
                if (!TryCreateCustomProperty(source,
                        out string? variant, out string? text)) {
                    continue;
                }
                values.Add(new CustomPropertyValue(name.Value, id,
                    variant!, text!));
            }
            return values;
        }

        /// <summary>The <c>vt:</c> element name and text a binary user-defined value is written as.</summary>
        private static bool TryCreateCustomProperty(OfficeOlePropertyValue source,
            out string? variant, out string? text) {
            (variant, text) = source.Value switch {
                string value => ("lpwstr", value),
                bool value => ("bool", value ? "true" : "false"),
                DateTime value => ("filetime", value.ToUniversalTime()
                    .ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)),
                float value => ("r4", OfficeRoundTripText.Format(value,
                    CultureInfo.InvariantCulture)),
                double value => ("r8", OfficeRoundTripText.Format(value,
                    CultureInfo.InvariantCulture)),
                sbyte or short or int => ("i4", Convert.ToInt32(source.Value,
                    CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture)),
                byte value => ("ui1", value.ToString(CultureInfo.InvariantCulture)),
                ushort value => ("ui2", value.ToString(CultureInfo.InvariantCulture)),
                uint value => ("ui4", value.ToString(CultureInfo.InvariantCulture)),
                long value => ("i8", value.ToString(CultureInfo.InvariantCulture)),
                ulong value => ("ui8", value.ToString(CultureInfo.InvariantCulture)),
                byte[] value => ("blob", Convert.ToBase64String(value)),
                _ => ((string?)null, (string?)null)
            };
            return variant != null;
        }

        private static bool CanRewriteSummary(LegacyPptPackage package) {
            if (!TryReadSections(package, OfficeOlePropertySetWriter
                    .SummaryInformationStreamName, out IReadOnlyList<
                    OfficeOlePropertySection> sections)) {
                return !HasStream(package, OfficeOlePropertySetWriter
                    .SummaryInformationStreamName);
            }
            return sections.Count == 1
                && sections[0].FormatId == OfficeOlePropertySetWriter
                    .SummaryInformationFormatId
                && sections[0].Dictionary.Count == 0
                && sections[0].Properties.All(pair =>
                    SupportedSummaryPropertyIds.Contains(pair.Key)
                    && IsSupportedSummaryValue(pair.Key, pair.Value));
        }

        private static bool CanRewriteDocumentSummary(
            LegacyPptPackage package) {
            if (!TryReadSections(package, OfficeOlePropertySetWriter
                    .DocumentSummaryInformationStreamName, out IReadOnlyList<
                    OfficeOlePropertySection> sections)) {
                return !HasStream(package, OfficeOlePropertySetWriter
                    .DocumentSummaryInformationStreamName);
            }
            if (sections.Count > 2 || sections.Select(section => section.FormatId)
                    .Distinct().Count() != sections.Count) return false;
            foreach (OfficeOlePropertySection section in sections) {
                if (section.FormatId == OfficeOlePropertySetWriter
                        .DocumentSummaryInformationFormatId) {
                    if (section.Dictionary.Count != 0
                        || section.Properties.Any(pair =>
                            !SupportedDocumentSummaryPropertyIds
                                .Contains(pair.Key)
                            || !IsSupportedDocumentSummaryValue(pair.Key,
                                pair.Value))) return false;
                } else if (section.FormatId == OfficeOlePropertySetWriter
                               .UserDefinedPropertiesFormatId) {
                    if (!CanRewriteCustomSection(section)) return false;
                } else {
                    return false;
                }
            }
            return true;
        }

        private static bool CanRewriteCustomSection(
            OfficeOlePropertySection section) {
            ISet<uint> namedIds = new HashSet<uint>(section.Dictionary.Keys);
            return namedIds.All(id => id > 1)
                && section.Properties.Keys.All(id => id == 1
                    || namedIds.Contains(id))
                && namedIds.All(id => section.Properties.TryGetValue(id,
                    out OfficeOlePropertyValue? value)
                    && IsSupportedCustomValue(value));
        }

        private static bool IsSupportedSummaryValue(uint id,
            OfficeOlePropertyValue value) => id switch {
                1 => value.Value is short or int,
                10 or 11 or 12 or 13 => value.Value is DateTime,
                _ => value.Value is string
            };

        private static bool IsSupportedDocumentSummaryValue(uint id,
            OfficeOlePropertyValue value) => id switch {
                1 => value.Value is short or int,
                7 or 8 or 9 => IsInteger(value.Value),
                _ => value.Value is string
            };

        private static bool IsSupportedCustomValue(
            OfficeOlePropertyValue value) => value.Value is string or bool
                or DateTime or float or double or sbyte or byte or short
                or ushort or int or uint or long or ulong or byte[];

        private static bool IsInteger(object? value) => value is sbyte or byte
            or short or ushort or int or uint or long or ulong;

        private static bool TryReadSections(LegacyPptPackage package,
            string streamName,
            out IReadOnlyList<OfficeOlePropertySection> sections) {
            sections = Array.Empty<OfficeOlePropertySection>();
            if (!package.CompoundFile.Streams.TryGetValue(streamName,
                    out byte[]? bytes) || bytes.Length == 0) return false;
            try {
                sections = OfficeOlePropertySetReader.ReadSections(bytes);
                return true;
            } catch (Exception exception) when (exception is IOException
                       or ArgumentException or InvalidDataException
                       or OverflowException) {
                return false;
            }
        }

        private static bool HasStream(LegacyPptPackage package,
            string streamName) => package.CompoundFile.Streams.ContainsKey(
                streamName);

        private static string? ReadString(OfficeOlePropertySection section,
            uint id) => section.Properties.TryGetValue(id,
            out OfficeOlePropertyValue? value) ? value.Value as string : null;

        private static DateTime? ReadDate(OfficeOlePropertySection section,
            uint id) => section.Properties.TryGetValue(id,
            out OfficeOlePropertyValue? value) && value.Value is DateTime date
                ? date
                : null;

        private static string? ReadIntegerText(
            OfficeOlePropertySection section, uint id) => section.Properties
            .TryGetValue(id, out OfficeOlePropertyValue? value)
            && IsInteger(value.Value)
                ? Convert.ToString(value.Value, CultureInfo.InvariantCulture)
                : null;

        private static string DateText(DateTime? value) => value.HasValue
            ? value.Value.ToUniversalTime().Ticks.ToString(
                CultureInfo.InvariantCulture)
            : string.Empty;

    }

    internal sealed class LegacyPptPropertySetProjection {
        internal LegacyPptPropertySetProjection(string summaryFingerprint,
            string documentSummaryFingerprint, string unmappedCoreFingerprint,
            bool canRewriteSummary,
            bool canRewriteDocumentSummary) {
            SummaryFingerprint = summaryFingerprint;
            DocumentSummaryFingerprint = documentSummaryFingerprint;
            UnmappedCoreFingerprint = unmappedCoreFingerprint;
            CanRewriteSummary = canRewriteSummary;
            CanRewriteDocumentSummary = canRewriteDocumentSummary;
        }

        private string SummaryFingerprint { get; }
        private string DocumentSummaryFingerprint { get; }
        private string UnmappedCoreFingerprint { get; }
        internal bool CanRewriteSummary { get; }
        internal bool CanRewriteDocumentSummary { get; }

        internal bool SummaryMatches(LegacyPptPackageView view) =>
            string.Equals(SummaryFingerprint,
                LegacyPptPropertySetCodec.CreateSummaryFingerprint(view),
                StringComparison.Ordinal);

        internal bool DocumentSummaryMatches(LegacyPptPackageView view) =>
            string.Equals(DocumentSummaryFingerprint,
                LegacyPptPropertySetCodec.CreateDocumentSummaryFingerprint(
                    view), StringComparison.Ordinal);

        internal bool UnmappedCoreMatches(LegacyPptPackageView view) =>
            string.Equals(UnmappedCoreFingerprint,
                LegacyPptPropertySetCodec.CreateUnmappedCoreFingerprint(
                    view), StringComparison.Ordinal);
    }
}
