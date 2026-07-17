using System.Globalization;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

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

        internal static void Apply(PowerPointPresentation presentation,
            LegacyPptPackage package) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (package == null) throw new ArgumentNullException(nameof(package));
            ClearMappedProperties(presentation.OpenXmlDocument);
            if (TryReadSections(package, OfficeOlePropertySetWriter
                    .SummaryInformationStreamName, out IReadOnlyList<
                    OfficeOlePropertySection> summarySections)) {
                ApplySummary(presentation.OpenXmlDocument, summarySections);
            }
            if (TryReadSections(package, OfficeOlePropertySetWriter
                    .DocumentSummaryInformationStreamName, out IReadOnlyList<
                    OfficeOlePropertySection> documentSummarySections)) {
                ApplyDocumentSummary(presentation.OpenXmlDocument,
                    documentSummarySections);
            }
        }

        internal static void ApplyReadOnlyDateOverrides(
            PowerPointPresentation presentation, LegacyPptPackage package) {
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
            presentation.BuiltinDocumentProperties
                .SetReadOnlyLegacyDateOverrides(created, modified,
                    lastPrinted);
        }

        internal static LegacyPptPropertySetProjection CreateProjection(
            PowerPointPresentation presentation, LegacyPptPackage package) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (package == null) throw new ArgumentNullException(nameof(package));
            bool summarySafe = CanRewriteSummary(package);
            bool documentSummarySafe = CanRewriteDocumentSummary(package);
            return new LegacyPptPropertySetProjection(
                CreateSummaryFingerprint(presentation.OpenXmlDocument),
                CreateDocumentSummaryFingerprint(
                    presentation.OpenXmlDocument),
                CreateUnmappedCoreFingerprint(
                    presentation.OpenXmlDocument), summarySafe,
                documentSummarySafe);
        }

        internal static bool TryCreateFreshStreams(
            PowerPointPresentation presentation,
            out IReadOnlyList<OfficeCompoundStream> streams,
            out string? reason) {
            var result = new List<OfficeCompoundStream>(2);
            streams = result;
            reason = null;
            byte[]? summary = CreateSummaryInformation(presentation);
            if (summary != null) {
                result.Add(new OfficeCompoundStream(OfficeOlePropertySetWriter
                    .SummaryInformationStreamName, summary));
            }
            if (!TryCreateDocumentSummaryInformation(presentation,
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
            PowerPointPresentation presentation,
            LegacyPptPropertySetProjection projection,
            out IReadOnlyDictionary<string, byte[]> streams) {
            var result = new Dictionary<string, byte[]>(
                StringComparer.OrdinalIgnoreCase);
            streams = result;
            if (!projection.UnmappedCoreMatches(
                    presentation.OpenXmlDocument)) return false;
            bool summaryChanged = !projection.SummaryMatches(
                presentation.OpenXmlDocument);
            bool documentSummaryChanged = !projection.DocumentSummaryMatches(
                presentation.OpenXmlDocument);
            if (summaryChanged) {
                if (!projection.CanRewriteSummary) return false;
                result[OfficeOlePropertySetWriter.SummaryInformationStreamName]
                    = CreateSummaryInformation(presentation)
                    ?? CreateEmptyPropertySet(
                        OfficeOlePropertySetWriter.SummaryInformationFormatId);
            }
            if (documentSummaryChanged) {
                if (!projection.CanRewriteDocumentSummary
                    || !TryCreateDocumentSummaryInformation(presentation,
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
            PresentationDocument document) {
            var properties = document.PackageProperties;
            Ap.Properties? application = document.ExtendedFilePropertiesPart?
                .Properties;
            return string.Join("\u001F", new[] {
                properties.Title, properties.Subject, properties.Creator,
                properties.Keywords, properties.Description,
                properties.LastModifiedBy, properties.Revision,
                DateText(properties.LastPrinted), DateText(properties.Created),
                DateText(properties.Modified), application?.Application?.Text,
                application?.TotalTime?.Text
            }.Select(value => value ?? string.Empty));
        }

        internal static string CreateDocumentSummaryFingerprint(
            PresentationDocument document) {
            var properties = document.PackageProperties;
            Ap.Properties? application = document.ExtendedFilePropertiesPart?
                .Properties;
            string custom = document.CustomFilePropertiesPart?.Properties?
                .OuterXml ?? string.Empty;
            return string.Join("\u001F", new[] {
                properties.Category, application?.PresentationFormat?.Text,
                application?.Slides?.Text, application?.Notes?.Text,
                application?.HiddenSlides?.Text, application?.Manager?.Text,
                application?.Company?.Text, custom
            }.Select(value => value ?? string.Empty));
        }

        internal static string CreateUnmappedCoreFingerprint(
            PresentationDocument document) {
            var properties = document.PackageProperties;
            return string.Join("\u001F", new[] {
                properties.ContentStatus, properties.ContentType,
                properties.Identifier, properties.Language,
                properties.Version
            }.Select(value => value ?? string.Empty));
        }

        private static void ClearMappedProperties(PresentationDocument document) {
            var properties = document.PackageProperties;
            properties.Title = null;
            properties.Subject = null;
            properties.Creator = null;
            properties.Keywords = null;
            properties.Description = null;
            properties.LastModifiedBy = null;
            properties.Revision = null;
            properties.LastPrinted = null;
            properties.Created = null;
            properties.Modified = null;
            properties.Category = null;
            properties.ContentStatus = null;
            properties.ContentType = null;
            properties.Identifier = null;
            properties.Language = null;
            properties.Version = null;
            Ap.Properties? application = document.ExtendedFilePropertiesPart?
                .Properties;
            if (application != null) {
                application.Application = null;
                application.TotalTime = null;
                application.PresentationFormat = null;
                application.Slides = null;
                application.Notes = null;
                application.HiddenSlides = null;
                application.Manager = null;
                application.Company = null;
            }
            if (document.CustomFilePropertiesPart != null) {
                document.DeletePart(document.CustomFilePropertiesPart);
            }
        }

        private static void ApplySummary(PresentationDocument document,
            IReadOnlyList<OfficeOlePropertySection> sections) {
            OfficeOlePropertySection? section = sections.FirstOrDefault(item =>
                item.FormatId == OfficeOlePropertySetWriter
                    .SummaryInformationFormatId);
            if (section == null) return;
            var properties = document.PackageProperties;
            properties.Title = ReadString(section, 2);
            properties.Subject = ReadString(section, 3);
            properties.Creator = ReadString(section, 4);
            properties.Keywords = ReadString(section, 5);
            properties.Description = ReadString(section, 6);
            properties.LastModifiedBy = ReadString(section, 8);
            properties.Revision = ReadString(section, 9);
            properties.LastPrinted = ReadDate(section, 11);
            properties.Created = ReadDate(section, 12);
            properties.Modified = ReadDate(section, 13);
            Ap.Properties application = EnsureApplicationProperties(document);
            string? applicationName = ReadString(section, 18);
            if (applicationName != null) {
                application.Application = new Ap.Application(applicationName);
            }
            DateTime? duration = ReadDate(section, 10);
            if (duration.HasValue) {
                long ticks = duration.Value.ToFileTimeUtc();
                application.TotalTime = new Ap.TotalTime(Math.Round(
                    TimeSpan.FromTicks(ticks).TotalMinutes,
                    MidpointRounding.AwayFromZero).ToString(
                        CultureInfo.InvariantCulture));
            }
        }

        private static void ApplyDocumentSummary(PresentationDocument document,
            IReadOnlyList<OfficeOlePropertySection> sections) {
            OfficeOlePropertySection? builtIn = sections.FirstOrDefault(item =>
                item.FormatId == OfficeOlePropertySetWriter
                    .DocumentSummaryInformationFormatId);
            if (builtIn != null) {
                document.PackageProperties.Category = ReadString(builtIn, 2);
                Ap.Properties application = EnsureApplicationProperties(document);
                SetApplicationText(ReadString(builtIn, 3), value =>
                    application.PresentationFormat =
                        new Ap.PresentationFormat(value));
                SetApplicationText(ReadIntegerText(builtIn, 7), value =>
                    application.Slides = new Ap.Slides(value));
                SetApplicationText(ReadIntegerText(builtIn, 8), value =>
                    application.Notes = new Ap.Notes(value));
                SetApplicationText(ReadIntegerText(builtIn, 9), value =>
                    application.HiddenSlides = new Ap.HiddenSlides(value));
                SetApplicationText(ReadString(builtIn, 14), value =>
                    application.Manager = new Ap.Manager(value));
                SetApplicationText(ReadString(builtIn, 15), value =>
                    application.Company = new Ap.Company(value));
            }
            OfficeOlePropertySection? custom = sections.FirstOrDefault(item =>
                item.FormatId == OfficeOlePropertySetWriter
                    .UserDefinedPropertiesFormatId);
            if (custom != null) ApplyCustomProperties(document, custom);
        }

        private static void ApplyCustomProperties(PresentationDocument document,
            OfficeOlePropertySection section) {
            var values = new List<CustomDocumentProperty>();
            int propertyId = 2;
            foreach (KeyValuePair<uint, string> name in section.Dictionary
                         .Where(pair => pair.Key > 1)
                         .OrderBy(pair => pair.Key)) {
                if (!section.Properties.TryGetValue(name.Key,
                        out OfficeOlePropertyValue? source)
                    || !TryCreateCustomProperty(name.Value, propertyId++,
                        source, out CustomDocumentProperty? property)) {
                    continue;
                }
                values.Add(property!);
            }
            if (values.Count == 0) return;
            CustomFilePropertiesPart part = document.CustomFilePropertiesPart
                ?? document.AddCustomFilePropertiesPart();
            part.Properties = new Properties(values);
            part.Properties.Save();
        }

        private static bool TryCreateCustomProperty(string name, int propertyId,
            OfficeOlePropertyValue source,
            out CustomDocumentProperty? property) {
            property = new CustomDocumentProperty {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = propertyId,
                Name = name
            };
            switch (source.Value) {
                case string value:
                    property.VTLPWSTR = new VTLPWSTR(value);
                    return true;
                case bool value:
                    property.VTBool = new VTBool(value ? "true" : "false");
                    return true;
                case DateTime value:
                    property.VTFileTime = new VTFileTime(value.ToUniversalTime()
                        .ToString("yyyy-MM-ddTHH:mm:ssZ",
                            CultureInfo.InvariantCulture));
                    return true;
                case float value:
                    property.VTFloat = new VTFloat(value.ToString("R",
                        CultureInfo.InvariantCulture));
                    return true;
                case double value:
                    property.VTDouble = new VTDouble(value.ToString("R",
                        CultureInfo.InvariantCulture));
                    return true;
                case sbyte or short or int:
                    property.VTInt32 = new VTInt32(Convert.ToInt32(source.Value,
                        CultureInfo.InvariantCulture).ToString(
                            CultureInfo.InvariantCulture));
                    return true;
                case byte value:
                    property.VTUnsignedByte = new VTUnsignedByte(value.ToString(
                        CultureInfo.InvariantCulture));
                    return true;
                case ushort value:
                    property.VTUnsignedShort = new VTUnsignedShort(value.ToString(
                        CultureInfo.InvariantCulture));
                    return true;
                case uint value:
                    property.VTUnsignedInt32 = new VTUnsignedInt32(value.ToString(
                        CultureInfo.InvariantCulture));
                    return true;
                case long value:
                    property.VTInt64 = new VTInt64(value.ToString(
                        CultureInfo.InvariantCulture));
                    return true;
                case ulong value:
                    property.VTUnsignedInt64 = new VTUnsignedInt64(value.ToString(
                        CultureInfo.InvariantCulture));
                    return true;
                case byte[] value:
                    property.VTBlob = new VTBlob(Convert.ToBase64String(value));
                    return true;
                default:
                    property = null;
                    return false;
            }
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

        private static void SetApplicationText(string? value,
            Action<string> setter) {
            if (value != null) setter(value);
        }

        private static Ap.Properties EnsureApplicationProperties(
            PresentationDocument document) {
            ExtendedFilePropertiesPart part = document
                .ExtendedFilePropertiesPart
                ?? document.AddExtendedFilePropertiesPart();
            return part.Properties ??= new Ap.Properties();
        }

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

        internal bool SummaryMatches(PresentationDocument document) =>
            string.Equals(SummaryFingerprint,
                LegacyPptPropertySetCodec.CreateSummaryFingerprint(document),
                StringComparison.Ordinal);

        internal bool DocumentSummaryMatches(PresentationDocument document) =>
            string.Equals(DocumentSummaryFingerprint,
                LegacyPptPropertySetCodec.CreateDocumentSummaryFingerprint(
                    document), StringComparison.Ordinal);

        internal bool UnmappedCoreMatches(PresentationDocument document) =>
            string.Equals(UnmappedCoreFingerprint,
                LegacyPptPropertySetCodec.CreateUnmappedCoreFingerprint(
                    document), StringComparison.Ordinal);
    }
}
