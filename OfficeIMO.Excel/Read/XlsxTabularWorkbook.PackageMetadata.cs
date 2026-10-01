#nullable enable

using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class XlsxTabularWorkbook {
        private (Dictionary<string, string> Overrides, Dictionary<string, string> Defaults)
            ReadContentTypes() {
            return ReadXmlPart("[Content_Types].xml", _options.MaxMetadataPartBytes, reader => {
                if (reader.MoveToContent() != XmlNodeType.Element
                    || reader.LocalName != "Types"
                    || reader.NamespaceURI != PackageContentTypesNamespace) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The package content-type manifest namespace is not supported by the native path.");
                }

                var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                bool empty = reader.IsEmptyElement;
                reader.Read();
                while (!empty && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == 0)) {
                    _options.CancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1) {
                        if (!reader.Read()) break;
                        continue;
                    }

                    if (reader.LocalName == "Override"
                        && reader.NamespaceURI == PackageContentTypesNamespace) {
                        string? rawPartName = reader.GetAttribute("PartName");
                        string? contentType = reader.GetAttribute("ContentType");
                        string normalizedPartName = NormalizeContentTypePartName(rawPartName);
                        if (normalizedPartName.Length == 0
                            || rawPartName![0] != '/'
                            || !IsValidContentType(contentType)
                            || overrides.ContainsKey(normalizedPartName)) {
                            throw new XlsxTabularFastPathNotSupportedException(
                                "The package content-type overrides require the Open XML SDK fallback path.");
                        }
                        overrides.Add(normalizedPartName, contentType!);
                    } else if (reader.LocalName == "Default"
                        && reader.NamespaceURI == PackageContentTypesNamespace) {
                        string? extension = reader.GetAttribute("Extension");
                        string? contentType = reader.GetAttribute("ContentType");
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
                    reader.Skip();
                }
                return (overrides, defaults);
            });
        }

        private IReadOnlyDictionary<string, PackageRelationship> ReadRelationships(
            string sourcePartName) {
            string relationshipPartName = GetRelationshipPartName(sourcePartName);
            return ReadXmlPart(relationshipPartName, _options.MaxMetadataPartBytes, reader => {
                if (reader.MoveToContent() != XmlNodeType.Element
                    || reader.LocalName != "Relationships"
                    || reader.NamespaceURI != PackageRelationshipsNamespace) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The package relationship namespace is not supported by the native path.");
                }

                var result = new Dictionary<string, PackageRelationship>(StringComparer.Ordinal);
                bool empty = reader.IsEmptyElement;
                reader.Read();
                while (!empty && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == 0)) {
                    _options.CancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1) {
                        if (!reader.Read()) break;
                        continue;
                    }
                    if (reader.LocalName != "Relationship"
                        || reader.NamespaceURI != PackageRelationshipsNamespace) {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "The package relationships require the Open XML SDK fallback path.");
                    }

                    string? id = reader.GetAttribute("Id");
                    string? type = reader.GetAttribute("Type");
                    string? target = reader.GetAttribute("Target");
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
                        type!, target!, ReadRelationshipTargetMode(reader.GetAttribute("TargetMode"))));
                    reader.Skip();
                }
                return result;
            });
        }

        private (XlsxTabularSheet[] Sheets, ExcelDateSystem DateSystem) ReadWorkbook(
            string workbookPartName,
            IReadOnlyDictionary<string, PackageRelationship> relationships,
            bool metadataOnly) {
            var (date1904, hasSheets, sheetCount, workbookSheets) = ReadXmlPart(
                workbookPartName,
                _options.MaxMetadataPartBytes,
                reader => {
                    if (reader.MoveToContent() != XmlNodeType.Element
                        || reader.LocalName != "workbook"
                        || (reader.NamespaceURI != TransitionalSpreadsheetNamespace
                            && reader.NamespaceURI != StrictSpreadsheetNamespace)) {
                        throw new XlsxTabularFastPathNotSupportedException(
                            "The workbook XML namespace is not supported by the native path.");
                    }

                    string spreadsheetNamespace = reader.NamespaceURI;
                    bool emptyWorkbook = reader.IsEmptyElement;
                    bool foundProperties = false;
                    bool foundSheets = false;
                    string? dateFlag = null;
                    int count = 0;
                    var entries = new List<(string? Name, string? RelationshipId)>();
                    reader.Read();
                    while (!emptyWorkbook && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == 0)) {
                        _options.CancellationToken.ThrowIfCancellationRequested();
                        if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1) {
                            if (!reader.Read()) break;
                            continue;
                        }

                        if (reader.NamespaceURI == spreadsheetNamespace
                            && reader.LocalName == "workbookPr"
                            && !foundProperties) {
                            foundProperties = true;
                            dateFlag = reader.GetAttribute("date1904");
                            reader.Skip();
                            continue;
                        }

                        if (reader.NamespaceURI == spreadsheetNamespace
                            && reader.LocalName == "sheets"
                            && !foundSheets) {
                            foundSheets = true;
                            bool emptySheets = reader.IsEmptyElement;
                            int sheetsDepth = reader.Depth;
                            reader.Read();
                            while (!emptySheets
                                && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == sheetsDepth)) {
                                _options.CancellationToken.ThrowIfCancellationRequested();
                                if (reader.NodeType == XmlNodeType.Element) {
                                    if (reader.Depth == sheetsDepth + 1
                                        && reader.NamespaceURI == spreadsheetNamespace
                                        && reader.LocalName == "sheet") {
                                        count++;
                                        if (count <= _options.MaxWorksheets) {
                                            entries.Add((
                                                reader.GetAttribute("name"),
                                                reader.GetAttribute("id", TransitionalOfficeRelationshipsNamespace)
                                                    ?? reader.GetAttribute("id", StrictOfficeRelationshipsNamespace)));
                                        }
                                    }
                                    reader.Skip();
                                } else if (!reader.Read()) {
                                    break;
                                }
                            }
                            if (!emptySheets) reader.Read();
                            continue;
                        }

                        reader.Skip();
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
                throw new InvalidDataException(
                    $"The workbook contains more than the configured {_options.MaxWorksheets} worksheet definitions.");
            }
            if (string.IsNullOrWhiteSpace(_options.SheetName)
                && !_options.SheetIndex.HasValue
                && workbookSheets.Count > 1
                && !metadataOnly) {
                // The public multi-result reader still uses the SDK path. Stop before
                // resolving every sheet and optional global part so that fallback does
                // not pay the complete native metadata probe first.
                throw new XlsxTabularFastPathNotSupportedException(
                    "Multi-result XLSX reads retain the complete Open XML SDK path.");
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

    }
}
