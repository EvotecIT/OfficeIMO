using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using System.Globalization;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static class LegacyDocPropertySetWriter {
        private const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private const string ExtendedPropertiesRelationshipType = LegacyDocSourcePackage.RelationshipTypePrefix + "extended-properties";
        private const string CustomPropertiesRelationshipType = LegacyDocSourcePackage.RelationshipTypePrefix + "custom-properties";

        // The W3CDTF forms System.IO.Packaging accepts for core-property dates, as it lists them.
        private static readonly string[] CoreDateFormats = CreateCoreDateFormats();

        internal static IReadOnlyList<OfficeCompoundStream> CreateDocumentPropertyStreams(WordDocument document, LegacyDocSourcePackage source) {
            ArgumentNullException.ThrowIfNull(document);

            var streams = new List<OfficeCompoundStream>(2);
            var core = new CoreProperties(source);
            byte[]? summaryInformation = CreateSummaryInformation(core);
            if (summaryInformation != null) {
                streams.Add(new OfficeCompoundStream(OfficeOlePropertySetWriter.SummaryInformationStreamName, summaryInformation));
            }

            byte[]? documentSummaryInformation = CreateDocumentSummaryInformation(document, source, core);
            if (documentSummaryInformation != null) {
                streams.Add(new OfficeCompoundStream(OfficeOlePropertySetWriter.DocumentSummaryInformationStreamName, documentSummaryInformation));
            }

            return streams;
        }

        private static byte[]? CreateSummaryInformation(CoreProperties properties) {
            var oleProperties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            OfficeOlePropertySetWriter.AddString(oleProperties, 2, properties.Title);
            OfficeOlePropertySetWriter.AddString(oleProperties, 3, properties.Subject);
            OfficeOlePropertySetWriter.AddString(oleProperties, 4, properties.Creator);
            OfficeOlePropertySetWriter.AddString(oleProperties, 5, properties.Keywords);
            OfficeOlePropertySetWriter.AddString(oleProperties, 6, properties.Description);
            OfficeOlePropertySetWriter.AddString(oleProperties, 8, properties.LastModifiedBy);
            OfficeOlePropertySetWriter.AddString(oleProperties, 9, properties.Revision);
            OfficeOlePropertySetWriter.AddFileTime(oleProperties, 11, properties.LastPrinted);
            OfficeOlePropertySetWriter.AddFileTime(oleProperties, 12, properties.Created);
            OfficeOlePropertySetWriter.AddFileTime(oleProperties, 13, properties.Modified);

            return oleProperties.Count == 1
                ? null
                : OfficeOlePropertySetWriter.CreatePropertySet((OfficeOlePropertySetWriter.SummaryInformationFormatId, OfficeOlePropertySetWriter.CreateSection(oleProperties)));
        }

        private static byte[]? CreateDocumentSummaryInformation(WordDocument document, LegacyDocSourcePackage source, CoreProperties core) {
            var sections = new List<(Guid FormatId, byte[] Section)>();
            var documentSummaryProperties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            LegacyDocElement application = source.Root(source.RelatedPart(null, ExtendedPropertiesRelationshipType));
            OfficeOlePropertySetWriter.AddString(documentSummaryProperties, 2, core.Category);
            OfficeOlePropertySetWriter.AddString(documentSummaryProperties, 14, ReadApplicationText(application, "Manager"));
            OfficeOlePropertySetWriter.AddString(documentSummaryProperties, 15, ReadApplicationText(application, "Company"));
            if (documentSummaryProperties.Count > 1) {
                sections.Add((OfficeOlePropertySetWriter.DocumentSummaryInformationFormatId, OfficeOlePropertySetWriter.CreateSection(documentSummaryProperties)));
            }

            // Custom properties are written to the package from the document's collection before saving, so a package
            // without them has none; asking the collection then would open the SDK package for nothing.
            if (source.RelatedPart(null, CustomPropertiesRelationshipType) != null && document.CustomProperties.Count > 0) {
                var customProperties = new List<OfficeOleProperty> {
                    OfficeOleProperty.Int16(OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
                };
                var dictionary = new Dictionary<uint, string>();
                uint propertyId = 2;
                foreach (KeyValuePair<string, WordCustomProperty> pair in document.CustomProperties.OrderBy(property => property.Key, StringComparer.OrdinalIgnoreCase)) {
                    if (TryCreateCustomProperty(propertyId, pair.Value, out OfficeOleProperty property)) {
                        dictionary[propertyId] = pair.Key;
                        customProperties.Add(property);
                        propertyId++;
                    }
                }

                if (dictionary.Count > 0) {
                    customProperties.Insert(1, OfficeOleProperty.Dictionary(0, dictionary));
                    sections.Add((OfficeOlePropertySetWriter.UserDefinedPropertiesFormatId, OfficeOlePropertySetWriter.CreateSection(customProperties)));
                }
            }

            return sections.Count == 0 ? null : OfficeOlePropertySetWriter.CreatePropertySet(sections.ToArray());
        }

        private static string? ReadApplicationText(LegacyDocElement application, string localName) {
            LegacyDocElement element = application.Child(OfficeXmlNamespaces.ExtendedProperties, localName);
            return element.Exists ? element.InnerText : null;
        }

        private static string[] CreateCoreDateFormats() => new[] {
            @"yyyy-MM-ddTHH:mm:ss", @"yyyy-MM-ddTHH:mm:ssZ", @"yyyy-MM-ddTHH:mm:sszzz", @"\-yyyy-MM-ddTHH:mm:ss", @"\-yyyy-MM-ddTHH:mm:ssZ", @"\-yyyy-MM-ddTHH:mm:sszzz",
            @"yyyy-MM-ddTHH:mm:ss.ff", @"yyyy-MM-ddTHH:mm:ss.fZ", @"yyyy-MM-ddTHH:mm:ss.fzzz", @"\-yyyy-MM-ddTHH:mm:ss.f", @"\-yyyy-MM-ddTHH:mm:ss.fZ", @"\-yyyy-MM-ddTHH:mm:ss.fzzz",
            @"yyyy-MM-ddTHH:mm:ss.ff", @"yyyy-MM-ddTHH:mm:ss.ffZ", @"yyyy-MM-ddTHH:mm:ss.ffzzz", @"\-yyyy-MM-ddTHH:mm:ss.ff", @"\-yyyy-MM-ddTHH:mm:ss.ffZ", @"\-yyyy-MM-ddTHH:mm:ss.ffzzz",
            @"yyyy-MM-ddTHH:mm:ss.fff", @"yyyy-MM-ddTHH:mm:ss.fffZ", @"yyyy-MM-ddTHH:mm:ss.fffzzz", @"\-yyyy-MM-ddTHH:mm:ss.fff", @"\-yyyy-MM-ddTHH:mm:ss.fffZ", @"\-yyyy-MM-ddTHH:mm:ss.fffzzz",
            @"yyyy-MM-ddTHH:mm:ss.ffff", @"yyyy-MM-ddTHH:mm:ss.ffffZ", @"yyyy-MM-ddTHH:mm:ss.ffffzzz", @"\-yyyy-MM-ddTHH:mm:ss.ffff", @"\-yyyy-MM-ddTHH:mm:ss.ffffZ", @"\-yyyy-MM-ddTHH:mm:ss.ffffzzz",
            @"yyyy-MM-ddTHH:mm:ss.fffff", @"yyyy-MM-ddTHH:mm:ss.fffffZ", @"yyyy-MM-ddTHH:mm:ss.fffffzzz", @"\-yyyy-MM-ddTHH:mm:ss.fffff", @"\-yyyy-MM-ddTHH:mm:ss.fffffZ", @"\-yyyy-MM-ddTHH:mm:ss.fffffzzz",
            @"yyyy-MM-ddTHH:mm:ss.ffffff", @"yyyy-MM-ddTHH:mm:ss.ffffffZ", @"yyyy-MM-ddTHH:mm:ss.ffffffzzz", @"\-yyyy-MM-ddTHH:mm:ss.ffffff", @"\-yyyy-MM-ddTHH:mm:ss.ffffffZ", @"\-yyyy-MM-ddTHH:mm:ss.ffffffzzz",
            @"yyyy-MM-ddTHH:mm:ss.fffffff", @"yyyy-MM-ddTHH:mm:ss.fffffffZ", @"yyyy-MM-ddTHH:mm:ss.fffffffzzz", @"\-yyyy-MM-ddTHH:mm:ss.fffffff", @"\-yyyy-MM-ddTHH:mm:ss.fffffffZ", @"\-yyyy-MM-ddTHH:mm:ss.fffffffzzz",
        };

        /// <summary>The core properties of the package, read as System.IO.Packaging reads them for the SDK.</summary>
        private sealed class CoreProperties {
            internal CoreProperties(LegacyDocSourcePackage source) {
                LegacyDocElement root = source.Root(source.RelatedPart(null, CorePropertiesRelationshipType));
                foreach (LegacyDocElement element in root.Elements()) {
                    switch (element.Namespace) {
                        case OfficeXmlNamespaces.DublinCore:
                            switch (element.LocalName) {
                                case "title": Title = element.Text; break;
                                case "subject": Subject = element.Text; break;
                                case "creator": Creator = element.Text; break;
                                case "description": Description = element.Text; break;
                            }
                            break;
                        case OfficeXmlNamespaces.CoreProperties:
                            switch (element.LocalName) {
                                case "keywords": Keywords = element.Text; break;
                                case "lastModifiedBy": LastModifiedBy = element.Text; break;
                                case "revision": Revision = element.Text; break;
                                case "category": Category = element.Text; break;
                                case "lastPrinted": LastPrinted = ParseDate(element.Text); break;
                            }
                            break;
                        case OfficeXmlNamespaces.DublinCoreTerms:
                            switch (element.LocalName) {
                                case "created": Created = ParseDate(element.Text); break;
                                case "modified": Modified = ParseDate(element.Text); break;
                            }
                            break;
                    }
                }
            }

            internal string? Title { get; }
            internal string? Subject { get; }
            internal string? Creator { get; }
            internal string? Keywords { get; }
            internal string? Description { get; }
            internal string? LastModifiedBy { get; }
            internal string? Revision { get; }
            internal string? Category { get; }
            internal DateTime? LastPrinted { get; }
            internal DateTime? Created { get; }
            internal DateTime? Modified { get; }

            private static DateTime ParseDate(string value) =>
                DateTime.ParseExact(value, CoreDateFormats, DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
        }

        private static bool TryCreateCustomProperty(uint propertyId, WordCustomProperty customProperty, out OfficeOleProperty property) {
            object? value = customProperty.Value;
            switch (customProperty.PropertyType) {
                case WordCustomPropertyType.DateTime:
                    property = OfficeOleProperty.FileTime(propertyId, Convert.ToDateTime(value, CultureInfo.InvariantCulture));
                    return true;
                case WordCustomPropertyType.NumberInteger:
                    property = OfficeOleProperty.Integer(propertyId, value);
                    return true;
                case WordCustomPropertyType.NumberDouble:
                    property = OfficeOleProperty.Double(propertyId, OfficeIMO.Core.Internal.OfficeRoundTripText.ToDouble(value, CultureInfo.InvariantCulture));
                    return true;
                case WordCustomPropertyType.YesNo:
                    property = OfficeOleProperty.Boolean(propertyId, Convert.ToBoolean(value, CultureInfo.InvariantCulture));
                    return true;
                case WordCustomPropertyType.Binary:
                    property = OfficeOleProperty.Blob(propertyId, GetBinaryValue(value));
                    return true;
                default:
                    property = OfficeOleProperty.String(propertyId, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                    return true;
            }
        }

        private static byte[] GetBinaryValue(object? value) {
            return value is byte[] bytes
                ? bytes
                : Convert.FromBase64String(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
        }
    }
}
