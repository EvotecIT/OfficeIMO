using DocumentFormat.OpenXml.CustomProperties;
using OfficeIMO.Shared;
using System.Globalization;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static class LegacyDocPropertySetWriter {
        internal static IReadOnlyList<OfficeCompoundStream> CreateDocumentPropertyStreams(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var streams = new List<OfficeCompoundStream>(2);
            byte[]? summaryInformation = CreateSummaryInformation(document);
            if (summaryInformation != null) {
                streams.Add(new OfficeCompoundStream(OfficeOlePropertySetWriter.SummaryInformationStreamName, summaryInformation));
            }

            byte[]? documentSummaryInformation = CreateDocumentSummaryInformation(document);
            if (documentSummaryInformation != null) {
                streams.Add(new OfficeCompoundStream(OfficeOlePropertySetWriter.DocumentSummaryInformationStreamName, documentSummaryInformation));
            }

            return streams;
        }

        private static byte[]? CreateSummaryInformation(WordDocument document) {
            BuiltinDocumentProperties properties = document.BuiltinDocumentProperties;
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

        private static byte[]? CreateDocumentSummaryInformation(WordDocument document) {
            var sections = new List<(Guid FormatId, byte[] Section)>();
            var documentSummaryProperties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            OfficeOlePropertySetWriter.AddString(documentSummaryProperties, 2, document.BuiltinDocumentProperties.Category);
            OfficeOlePropertySetWriter.AddString(documentSummaryProperties, 14, document.ApplicationProperties.Manager?.Text);
            OfficeOlePropertySetWriter.AddString(documentSummaryProperties, 15, document.ApplicationProperties.Company);
            if (documentSummaryProperties.Count > 1) {
                sections.Add((OfficeOlePropertySetWriter.DocumentSummaryInformationFormatId, OfficeOlePropertySetWriter.CreateSection(documentSummaryProperties)));
            }

            if (document.CustomDocumentProperties.Count > 0) {
                var customProperties = new List<OfficeOleProperty> {
                    OfficeOleProperty.Int16(OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
                };
                var dictionary = new Dictionary<uint, string>();
                uint propertyId = 2;
                foreach (KeyValuePair<string, WordCustomProperty> pair in document.CustomDocumentProperties.OrderBy(property => property.Key, StringComparer.OrdinalIgnoreCase)) {
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

        private static bool TryCreateCustomProperty(uint propertyId, WordCustomProperty customProperty, out OfficeOleProperty property) {
            object? value = customProperty.Value;
            switch (customProperty.PropertyType) {
                case PropertyTypes.DateTime:
                    property = OfficeOleProperty.FileTime(propertyId, Convert.ToDateTime(value, CultureInfo.InvariantCulture));
                    return true;
                case PropertyTypes.NumberInteger:
                    property = OfficeOleProperty.Integer(propertyId, value);
                    return true;
                case PropertyTypes.NumberDouble:
                    property = OfficeOleProperty.Double(propertyId, Convert.ToDouble(value, CultureInfo.InvariantCulture));
                    return true;
                case PropertyTypes.YesNo:
                    property = OfficeOleProperty.Boolean(propertyId, Convert.ToBoolean(value, CultureInfo.InvariantCulture));
                    return true;
                default:
                    property = OfficeOleProperty.String(propertyId, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                    return true;
            }
        }
    }
}
