using System.Globalization;
using OfficeIMO.Shared;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyOlePropertySetWriter {
        internal static IReadOnlyList<OfficeCompoundStream> CreateDocumentPropertyStreams(ExcelDocument document) {
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

        private static byte[]? CreateSummaryInformation(ExcelDocument document) {
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

        private static byte[]? CreateDocumentSummaryInformation(ExcelDocument document) {
            var sections = new List<(Guid FormatId, byte[] Section)>();
            var documentSummaryProperties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            OfficeOlePropertySetWriter.AddString(documentSummaryProperties, 2, document.BuiltinDocumentProperties.Category);
            OfficeOlePropertySetWriter.AddString(documentSummaryProperties, 14, document.ApplicationProperties.Manager);
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
                foreach (KeyValuePair<string, ExcelCustomProperty> pair in document.CustomDocumentProperties.OrderBy(property => property.Key, StringComparer.OrdinalIgnoreCase)) {
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

        private static bool TryCreateCustomProperty(uint propertyId, ExcelCustomProperty customProperty, out OfficeOleProperty property) {
            object? value = customProperty.Value;
            switch (customProperty.PropertyType) {
                case ExcelCustomPropertyType.DateTime:
                    property = OfficeOleProperty.FileTime(propertyId, Convert.ToDateTime(value, CultureInfo.InvariantCulture));
                    return true;
                case ExcelCustomPropertyType.NumberInteger:
                    property = OfficeOleProperty.Integer(propertyId, value);
                    return true;
                case ExcelCustomPropertyType.NumberDouble:
                    property = OfficeOleProperty.Double(propertyId, Convert.ToDouble(value, CultureInfo.InvariantCulture));
                    return true;
                case ExcelCustomPropertyType.YesNo:
                    property = OfficeOleProperty.Boolean(propertyId, Convert.ToBoolean(value, CultureInfo.InvariantCulture));
                    return true;
                case ExcelCustomPropertyType.Binary:
                    property = OfficeOleProperty.Blob(propertyId, GetBinaryValue(value));
                    return true;
                default:
                    property = OfficeOleProperty.String(propertyId, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                    return true;
            }
        }

        private static byte[] GetBinaryValue(object? value) {
            return value switch {
                null => Array.Empty<byte>(),
                byte[] bytes => (byte[])bytes.Clone(),
                _ => Convert.FromBase64String(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
            };
        }
    }
}
