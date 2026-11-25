using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    internal sealed partial class XlsxTabularWorkbook {
        private (Dictionary<string, string> Overrides, Dictionary<string, string> Defaults)
            ReadContentTypes() {
            return ReadXmlPart("[Content_Types].xml", _options.MaxMetadataPartBytes, reader => {
                if (!reader.IsElement(OfficeXmlNamespaces.ContentTypes, "Types")) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The package content-type manifest namespace is not supported by the native path.");
                }

                var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                while (reader.Read() && !IsRootEnd(reader)) {
                    _options.CancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2) continue;

                    if (reader.IsElement(OfficeXmlNamespaces.ContentTypes, "Override")) {
                        string? rawPartName = GetAttribute(reader, "PartName");
                        string? contentType = GetKnownAttribute(reader, "ContentType");
                        string normalizedPartName = NormalizeContentTypePartName(rawPartName);
                        if (normalizedPartName.Length == 0
                            || rawPartName![0] != '/'
                            || !IsValidContentType(contentType)
                            || overrides.ContainsKey(normalizedPartName)) {
                            throw new XlsxTabularFastPathNotSupportedException(
                                "The package content-type overrides require the Open XML SDK fallback path.");
                        }
                        overrides.Add(normalizedPartName, contentType!);
                    } else if (reader.IsElement(OfficeXmlNamespaces.ContentTypes, "Default")) {
                        string? extension = GetAttribute(reader, "Extension");
                        string? contentType = GetKnownAttribute(reader, "ContentType");
                        if (!IsValidContentTypeExtension(extension)
                            || !IsValidContentType(contentType)
                            || defaults.ContainsKey(extension!)) {
                            throw new XlsxTabularFastPathNotSupportedException(
                                "The package content-type defaults require the Open XML SDK fallback path.");
                        }
                        defaults.Add(extension!, contentType!);
                    } else {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "The package content-type manifest requires the Open XML SDK fallback path.");
                    }
                    reader.SkipElement();
                }
                return (overrides, defaults);
            });
        }

        private Dictionary<string, PackageRelationship> ReadRelationships(
            string sourcePartName) {
            string relationshipPartName = GetRelationshipPartName(sourcePartName);
            return ReadXmlPart(relationshipPartName, _options.MaxMetadataPartBytes, reader => {
                if (!reader.IsElement(OfficeXmlNamespaces.PackageRelationships, "Relationships")) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The package relationship namespace is not supported by the native path.");
                }

                var result = new Dictionary<string, PackageRelationship>(StringComparer.Ordinal);
                while (reader.Read() && !IsRootEnd(reader)) {
                    _options.CancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2) continue;
                    if (!reader.IsElement(OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "The package relationships require the Open XML SDK fallback path.");
                    }

                    string? id = GetAttribute(reader, "Id");
                    string? type = GetKnownAttribute(reader, "Type");
                    string? target = GetAttribute(reader, "Target");
                    if (string.IsNullOrWhiteSpace(id)
                        || string.IsNullOrWhiteSpace(type)
                        || string.IsNullOrWhiteSpace(target)
                        || !IsValidRelationshipId(id!)
                        || !Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out _)
                        || result.ContainsKey(id!)) {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "The package relationships require the Open XML SDK fallback path.");
                    }

                    result.Add(id!, new PackageRelationship(
                        type!, target!, ReadRelationshipTargetMode(GetAttribute(reader, "TargetMode"))));
                    reader.SkipElement();
                }
                return result;
            });
        }

        private (XlsxTabularSheet[] Sheets, ExcelDateSystem DateSystem) ReadWorkbook(
            string workbookPartName,
            Dictionary<string, PackageRelationship> relationships,
            bool metadataOnly) {
            var (date1904, hasSheets, sheetCount, workbookSheets) = ReadXmlPart(
                workbookPartName,
                _options.MaxMetadataPartBytes,
                reader => {
                    if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "workbook")) {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "The workbook XML namespace is not supported by the native path.");
                    }

                    bool foundProperties = false;
                    bool foundSheets = false;
                    string? dateFlag = null;
                    int count = 0;
                    var entries = new List<(string? Name, string? RelationshipId)>();
                    while (reader.Read() && !IsRootEnd(reader)) {
                        _options.CancellationToken.ThrowIfCancellationRequested();
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2) continue;

                        if (!foundProperties && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "workbookPr")) {
                            foundProperties = true;
                            dateFlag = GetAttribute(reader, "date1904");
                        } else if (!foundSheets && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "sheets")) {
                            foundSheets = true;
                            while (reader.Read() && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 2)) {
                                _options.CancellationToken.ThrowIfCancellationRequested();
                                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                                if (reader.Depth == 3 && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "sheet")) {
                                    count++;
                                    if (count <= _options.MaxWorksheets) {
                                        entries.Add((
                                            GetAttribute(reader, "name"),
                                            reader.TryGetAttribute(OfficeXmlNamespaces.Relationships, "id", out string relationshipId)
                                                ? relationshipId : null));
                                    }
                                }
                                reader.SkipElement();
                            }
                            continue;
                        }
                        reader.SkipElement();
                    }
                    return (dateFlag, foundSheets, count, entries);
                });

            ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred;
            if (!string.IsNullOrEmpty(date1904)) {
                if (date1904 == "1" || string.Equals(date1904, "true", StringComparison.OrdinalIgnoreCase)) {
                    dateSystem = ExcelDateSystem.NineteenFour;
                } else if (date1904 != "0" && !string.Equals(date1904, "false", StringComparison.OrdinalIgnoreCase)) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The workbook date-system flag requires the Open XML SDK fallback path.");
                }
            }

            if (!hasSheets) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The workbook has no sheets collection.");
            }
            if (sheetCount > _options.MaxWorksheets) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxWorksheets), _options.MaxWorksheets,
                    $"The workbook contains more than the configured {_options.MaxWorksheets} worksheet definitions.");
            }
            var sheets = new List<XlsxTabularSheet>();
            var worksheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach ((string? name, string? relationshipId) in workbookSheets) {
                _options.CancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrEmpty(name)
                    || string.IsNullOrEmpty(relationshipId)
                    || !relationships.TryGetValue(relationshipId!, out PackageRelationship? relationship)) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "A workbook sheet requires the Open XML SDK fallback path.");
                }
                if (relationship.IsExternal) {
                    throw new InvalidDataException(
                        $"The OpenXML worksheet '{name}' references external relationship '{relationshipId}'.");
                }
                if (!IsOfficeRelationship(relationship.Type, WorksheetRelationshipSuffix)) {
                    if (IsSupportedNonWorksheetRelationship(relationship.Type)) {
                        if (metadataOnly) {
                            continue;
                        }
                        throw new XlsxTabularFastPathNotSupportedException(
                            "Non-worksheet sheet relationships require the Open XML SDK fallback path.");
                    }

                    throw new XlsxTabularFastPathNotSupportedException(
                        "A workbook sheet relationship requires the Open XML SDK fallback path.");
                }
                if (!worksheetNames.Add(name!)) {
                    throw new InvalidDataException(
                        $"The workbook contains duplicate worksheet name '{name}' under case-insensitive matching.");
                }

                string partName = ResolveTarget(workbookPartName, relationship.Target);
                if (!_parts.ContainsPart(partName)) {
                    throw new InvalidDataException(
                        $"The OpenXML worksheet '{name}' references missing relationship '{relationshipId}'.");
                }
                ValidatePartContentType(partName, WorksheetContentType, "worksheet");
                sheets.Add(new XlsxTabularSheet(name!, partName));
            }

            _options.CancellationToken.ThrowIfCancellationRequested();
            return (sheets.ToArray(), dateSystem);
        }

        private static bool IsRootEnd(OfficeXmlPullReader reader) =>
            reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 1;

        // Relationship and content types repeat in every package; a value written exactly as one of these
        // reuses the constant instead of decoding a new string.
        private static readonly string[] KnownMetadataValues = {
            TransitionalOfficeRelationshipsNamespace + "/officeDocument",
            TransitionalOfficeRelationshipsNamespace + WorksheetRelationshipSuffix,
            TransitionalOfficeRelationshipsNamespace + StylesRelationshipSuffix,
            TransitionalOfficeRelationshipsNamespace + SharedStringsRelationshipSuffix,
            TransitionalOfficeRelationshipsNamespace + "/theme",
            TransitionalOfficeRelationshipsNamespace + "/extended-properties",
            TransitionalOfficeRelationshipsNamespace + "/calcChain",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "application/vnd.openxmlformats-package.relationships+xml",
            "application/xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            WorksheetContentType,
            StylesContentType,
            SharedStringsContentType,
            "application/vnd.openxmlformats-officedocument.theme+xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
        };

        private static readonly byte[][] KnownMetadataValueBytes =
            Array.ConvertAll(KnownMetadataValues, static value => System.Text.Encoding.ASCII.GetBytes(value));

        /// <summary>Reads an unprefixed attribute, returning a shared constant for a well-known value.</summary>
        private static string? GetKnownAttribute(OfficeXmlPullReader reader, string localName) {
            for (int index = 0; index < reader.AttributeCount; index++) {
                if (reader.GetAttributeNamespace(index, out _) != OfficeXmlNamespaces.None
                    || !string.Equals(reader.GetAttributeLocalName(index), localName, StringComparison.Ordinal)) continue;
                reader.GetAttributeRawValue(index, out int start, out int length);
                ReadOnlySpan<byte> raw = reader.Data.AsSpan(start, length);
                for (int known = 0; known < KnownMetadataValueBytes.Length; known++) {
                    // Known values hold no references or whitespace, so equal bytes decode to the same text.
                    if (raw.SequenceEqual(KnownMetadataValueBytes[known])) return KnownMetadataValues[known];
                }
                return reader.GetAttributeValue(index);
            }
            return null;
        }

        // An unprefixed attribute, as XmlReader.GetAttribute(name) reads it; null when absent.
        private static string? GetAttribute(OfficeXmlPullReader reader, string localName) =>
            reader.TryGetAttribute(OfficeXmlNamespaces.None, localName, out string value) ? value : null;

    }
}
