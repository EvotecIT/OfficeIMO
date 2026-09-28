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

    }
}
