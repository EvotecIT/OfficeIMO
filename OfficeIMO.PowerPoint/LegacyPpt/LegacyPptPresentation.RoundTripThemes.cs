using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordRoundTripTheme12Atom = 0x040E;
        private const ushort RecordRoundTripColorMapping12Atom = 0x040F;
        private const int MaximumRoundTripThemeXmlBytes = 8 * 1024 * 1024;
        private const int MaximumRoundTripThemeEntryCount = 64;
        private const double MaximumRoundTripThemeCompressionRatio = 100D;
        private const string DrawingNamespace =
            "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly UTF8Encoding StrictUtf8 = new(false, true);

        private LegacyPptRoundTripTheme? ReadRoundTripTheme(
            LegacyPptRecord owner, string ownerName,
            LegacyPptImportOptions options) {
            LegacyPptRecord? themeRecord = owner.Children.FirstOrDefault(
                child => child.Type == RecordRoundTripTheme12Atom);
            LegacyPptRecord? mappingRecord = owner.Children.FirstOrDefault(
                child => child.Type == RecordRoundTripColorMapping12Atom);
            if (themeRecord == null && mappingRecord == null) return null;

            string? themeXml = null;
            string? mappingXml = null;
            XElement? themeRoot = null;
            if (themeRecord != null) {
                if (themeRecord.PayloadLength > options.MaxInputBytes) {
                    AddDiagnostic("PPT-ROUNDTRIP-THEME-INVALID",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"The {ownerName} DrawingML theme remains preserve-only: the payload exceeds the configured input limit.",
                        themeRecord.Offset);
                } else if (!TryReadRoundTripThemeXml(themeRecord,
                        _decodedStorageBudget, out themeXml,
                        out themeRoot, out string? themeReason)) {
                    AddDiagnostic("PPT-ROUNDTRIP-THEME-INVALID",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"The {ownerName} DrawingML theme remains preserve-only: {themeReason}",
                        themeRecord.Offset);
                }
            }
            if (mappingRecord != null) {
                if (mappingRecord.PayloadLength > options.MaxInputBytes) {
                    AddDiagnostic("PPT-ROUNDTRIP-COLOR-MAP-INVALID",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"The {ownerName} DrawingML color mapping remains preserve-only: the payload exceeds the configured input limit.",
                        mappingRecord.Offset);
                } else if (!TryReadRoundTripColorMapping(mappingRecord,
                        _decodedStorageBudget, out mappingXml,
                        out string? mappingReason)) {
                    AddDiagnostic("PPT-ROUNDTRIP-COLOR-MAP-INVALID",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"The {ownerName} DrawingML color mapping remains preserve-only: {mappingReason}",
                        mappingRecord.Offset);
                }
            }
            if (themeXml == null && mappingXml == null) return null;

            IReadOnlyDictionary<PowerPointThemeColor, string> colors =
                themeRoot == null
                    ? new Dictionary<PowerPointThemeColor, string>()
                    : ReadThemeColors(themeRoot);
            XElement? colorScheme = themeRoot?.Descendants(
                XName.Get("clrScheme", DrawingNamespace)).FirstOrDefault();
            XElement? fontScheme = themeRoot?.Descendants(
                XName.Get("fontScheme", DrawingNamespace)).FirstOrDefault();
            return new LegacyPptRoundTripTheme(themeXml, mappingXml,
                themeRoot?.Name.LocalName == "themeOverride",
                themeRoot?.Attribute("name")?.Value,
                colorScheme?.Attribute("name")?.Value,
                colors,
                ReadLatinTypeface(fontScheme, "majorFont"),
                ReadLatinTypeface(fontScheme, "minorFont"));
        }

        private static bool TryReadRoundTripThemeXml(
            LegacyPptRecord record,
            LegacyPptDecodedStorageBudget decodedStorageBudget,
            out string? xml, out XElement? root, out string? reason) {
            xml = null;
            root = null;
            reason = null;
            try {
                byte[] recordBytes = record.CopyRecordBytes();
                OfficeArchiveSafety.ZipCentralDirectoryScanResult directory =
                    OfficeArchiveSafety.ScanZipCentralDirectory(recordBytes,
                        8, record.PayloadLength,
                        MaximumRoundTripThemeEntryCount);
                if (!directory.IsValid) {
                    reason = directory.Error
                        ?? "the embedded package has a malformed central directory.";
                    return false;
                }
                if (directory.LimitExceeded) {
                    reason = "the embedded package contains too many entries.";
                    return false;
                }
                using var stream = new MemoryStream(recordBytes, 8,
                    record.PayloadLength, writable: false);
                using var archive = new ZipArchive(stream, ZipArchiveMode.Read,
                    leaveOpen: false);
                long totalUncompressedLength = 0;
                foreach (ZipArchiveEntry entry in archive.Entries) {
                    string name = OfficeArchiveSafety.NormalizeEntryName(
                        entry.FullName);
                    if (!OfficeArchiveSafety.TryGetLength(entry,
                            out long length)) {
                        reason = "the embedded package has an invalid entry length.";
                        return false;
                    }
                    if (OfficeArchiveSafety.IsUnsafePath(name)
                        || length < 0
                        || length > MaximumRoundTripThemeXmlBytes
                        || OfficeArchiveSafety.IsCompressionRatioExceeded(
                            entry, length,
                            MaximumRoundTripThemeCompressionRatio)) {
                        reason = "the embedded package violates archive safety limits.";
                        return false;
                    }
                    totalUncompressedLength = checked(
                        totalUncompressedLength + length);
                    if (totalUncompressedLength
                        > MaximumRoundTripThemeXmlBytes * 2L) {
                        reason = "the embedded package exceeds the total expansion limit.";
                        return false;
                    }
                    if (!name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                        || length == 0) continue;
                    if (length > decodedStorageBudget
                            .RemainingAllocationBytes) {
                        reason = "the decoded theme exceeds the remaining aggregate storage limit.";
                        return false;
                    }
                    using Stream entryStream = entry.Open();
                    byte[] entryBytes = OfficeArchiveSafety.ReadEntryBytes(
                        entryStream, length,
                        MaximumRoundTripThemeXmlBytes);
                    if (!TryParseDrawingXml(entryBytes,
                            out string? candidate,
                            out XElement? candidateRoot)
                        || candidateRoot == null) {
                        continue;
                    }
                    if (candidateRoot.Name.LocalName is not "theme"
                        and not "themeOverride") continue;
                    int retainedBytes = StrictUtf8.GetByteCount(candidate!);
                    if (retainedBytes > decodedStorageBudget
                            .RemainingAllocationBytes) {
                        reason = "the decoded theme exceeds the remaining aggregate storage limit.";
                        return false;
                    }
                    decodedStorageBudget.Consume(retainedBytes);
                    xml = candidate;
                    root = candidateRoot;
                    return true;
                }
                reason = "the embedded package has no DrawingML theme part.";
                return false;
            } catch (Exception exception) when (exception is InvalidDataException
                                                or IOException
                                                or XmlException
                                                or DecoderFallbackException
                                                or NotSupportedException) {
                reason = exception.Message;
                return false;
            }
        }

        private static bool TryReadRoundTripColorMapping(
            LegacyPptRecord record,
            LegacyPptDecodedStorageBudget decodedStorageBudget,
            out string? xml, out string? reason) {
            xml = null;
            reason = null;
            if (record.PayloadLength <= 0
                || record.PayloadLength > MaximumRoundTripThemeXmlBytes) {
                reason = "the mapping payload length is invalid.";
                return false;
            }
            if (record.PayloadLength > decodedStorageBudget
                    .RemainingAllocationBytes) {
                reason = "the decoded color mapping exceeds the remaining aggregate storage limit.";
                return false;
            }
            try {
                byte[] bytes = record.CopyRecordBytes();
                string candidate = StrictUtf8.GetString(bytes, 8,
                    record.PayloadLength);
                if (!TryParseDrawingXml(candidate, out XElement? root)
                    || root == null
                    || root.Name.LocalName is not "clrMap"
                        and not "clrMapOvr"
                        and not "clrMapOverride") {
                    reason = "the payload is not DrawingML color-mapping XML.";
                    return false;
                }
                int retainedBytes = StrictUtf8.GetByteCount(candidate);
                if (retainedBytes > decodedStorageBudget
                        .RemainingAllocationBytes) {
                    reason = "the decoded color mapping exceeds the remaining aggregate storage limit.";
                    return false;
                }
                decodedStorageBudget.Consume(retainedBytes);
                xml = candidate;
                return true;
            } catch (Exception exception) when (exception is XmlException
                                                or DecoderFallbackException) {
                reason = exception.Message;
                return false;
            }
        }

        private static bool TryParseDrawingXml(string xml,
            out XElement? root) {
            root = null;
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaximumRoundTripThemeXmlBytes
            };
            using var stringReader = new StringReader(xml);
            using XmlReader reader = XmlReader.Create(stringReader, settings);
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            if (document.Root?.Name.NamespaceName != DrawingNamespace) {
                return false;
            }
            root = document.Root;
            return true;
        }

        private static bool TryParseDrawingXml(byte[] bytes,
            out string? xml, out XElement? root) {
            xml = null;
            root = null;
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaximumRoundTripThemeXmlBytes
            };
            using var stream = new MemoryStream(bytes, writable: false);
            using XmlReader reader = XmlReader.Create(stream, settings);
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            if (document.Root?.Name.NamespaceName != DrawingNamespace) {
                return false;
            }
            root = document.Root;
            xml = document.ToString(SaveOptions.DisableFormatting);
            return true;
        }

        private static IReadOnlyDictionary<PowerPointThemeColor, string>
            ReadThemeColors(XElement root) {
            XElement? scheme = root.Descendants(
                XName.Get("clrScheme", DrawingNamespace)).FirstOrDefault();
            if (scheme == null) {
                return new Dictionary<PowerPointThemeColor, string>();
            }
            var result = new Dictionary<PowerPointThemeColor, string>();
            foreach ((PowerPointThemeColor Color, string Name) pair in
                     ThemeColorElements) {
                XElement? owner = scheme.Element(
                    XName.Get(pair.Name, DrawingNamespace));
                XElement? value = owner?.Elements().FirstOrDefault();
                string? color = value?.Name.LocalName == "sysClr"
                    ? value.Attribute("lastClr")?.Value
                    : value?.Attribute("val")?.Value;
                if (IsHexColor(color)) result[pair.Color] = color!.ToUpperInvariant();
            }
            return result;
        }

        private static readonly (PowerPointThemeColor Color, string Name)[]
            ThemeColorElements = {
                (PowerPointThemeColor.Dark1, "dk1"),
                (PowerPointThemeColor.Light1, "lt1"),
                (PowerPointThemeColor.Dark2, "dk2"),
                (PowerPointThemeColor.Light2, "lt2"),
                (PowerPointThemeColor.Accent1, "accent1"),
                (PowerPointThemeColor.Accent2, "accent2"),
                (PowerPointThemeColor.Accent3, "accent3"),
                (PowerPointThemeColor.Accent4, "accent4"),
                (PowerPointThemeColor.Accent5, "accent5"),
                (PowerPointThemeColor.Accent6, "accent6"),
                (PowerPointThemeColor.Hyperlink, "hlink"),
                (PowerPointThemeColor.FollowedHyperlink, "folHlink")
            };

        private static string? ReadLatinTypeface(XElement? fontScheme,
            string ownerName) => fontScheme?
            .Element(XName.Get(ownerName, DrawingNamespace))?
            .Element(XName.Get("latin", DrawingNamespace))?
            .Attribute("typeface")?.Value;

        private static bool IsHexColor(string? value) => value?.Length == 6
            && value.All(Uri.IsHexDigit);
    }
}
