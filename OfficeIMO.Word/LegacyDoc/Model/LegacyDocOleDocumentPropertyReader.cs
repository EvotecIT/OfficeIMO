using OfficeIMO.Drawing.Internal;
using OfficeIMO.Word.LegacyDoc.Diagnostics;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocOleDocumentPropertyReader {
        private const string SummaryInformationStreamName = "\u0005SummaryInformation";
        private const string DocumentSummaryInformationStreamName = "\u0005DocumentSummaryInformation";
        private const uint PropertyDictionaryId = 0;
        private const uint CodePagePropertyId = 1;

        internal static void AddDocumentProperties(OfficeCompoundFile compoundFile, LegacyDocDocument document, LegacyDocImportOptions options) {
            if (compoundFile == null) throw new ArgumentNullException(nameof(compoundFile));
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (options == null) throw new ArgumentNullException(nameof(options));

            if (TryGetStream(compoundFile.Streams, SummaryInformationStreamName, out byte[]? summaryStream)) {
                TryReadSummaryInformation(summaryStream!, document);
            }

            if (TryGetStream(compoundFile.Streams, DocumentSummaryInformationStreamName, out byte[]? documentSummaryStream)) {
                TryReadDocumentSummaryInformation(documentSummaryStream!, document, options);
            }
        }

        private static bool TryGetStream(IReadOnlyDictionary<string, byte[]> streams, string name, out byte[]? bytes) {
            return streams.TryGetValue(name, out bytes) && bytes != null && bytes.Length > 0;
        }

        private static void TryReadSummaryInformation(byte[] bytes, LegacyDocDocument document) {
            try {
                IReadOnlyList<OfficeOlePropertySection> sections = OfficeOlePropertySetReader.ReadSections(bytes);
                foreach (OfficeOlePropertySection section in sections) {
                    foreach (KeyValuePair<uint, OfficeOlePropertyValue> property in section.Properties) {
                        switch (property.Key) {
                            case 2:
                                document.DocumentProperties.Title = property.Value.AsString();
                                break;
                            case 3:
                                document.DocumentProperties.Subject = property.Value.AsString();
                                break;
                            case 4:
                                document.DocumentProperties.Creator = property.Value.AsString();
                                break;
                            case 5:
                                document.DocumentProperties.Keywords = property.Value.AsString();
                                break;
                            case 6:
                                document.DocumentProperties.Description = property.Value.AsString();
                                break;
                            case 8:
                                document.DocumentProperties.LastModifiedBy = property.Value.AsString();
                                break;
                            case 9:
                                document.DocumentProperties.Revision = property.Value.AsString();
                                break;
                            case 11:
                                document.DocumentProperties.LastPrinted = property.Value.AsDateTime();
                                break;
                            case 12:
                                document.DocumentProperties.Created = property.Value.AsDateTime();
                                break;
                            case 13:
                                document.DocumentProperties.Modified = property.Value.AsDateTime();
                                break;
                        }
                    }
                }
            } catch (Exception ex) when (ex is IOException || ex is ArgumentException || ex is InvalidDataException || ex is OverflowException) {
                document.AddWarning("DOC-OLE-PROPERTIES-UNREADABLE", $"The OLE document property stream '{SummaryInformationStreamName}' could not be read. {ex.Message}");
            }
        }

        private static void TryReadDocumentSummaryInformation(byte[] bytes, LegacyDocDocument document, LegacyDocImportOptions options) {
            try {
                IReadOnlyList<OfficeOlePropertySection> sections = OfficeOlePropertySetReader.ReadSections(bytes);
                foreach (OfficeOlePropertySection section in sections) {
                    if (section.Dictionary.Count == 0) {
                        if (section.Properties.TryGetValue(2, out OfficeOlePropertyValue? category)) {
                            document.DocumentProperties.Category = category.AsString();
                        }

                        if (section.Properties.TryGetValue(14, out OfficeOlePropertyValue? manager)) {
                            document.DocumentProperties.Manager = manager.AsString();
                        }

                        if (section.Properties.TryGetValue(15, out OfficeOlePropertyValue? company)) {
                            document.DocumentProperties.Company = company.AsString();
                        }
                    }

                    foreach (KeyValuePair<uint, string> name in section.Dictionary) {
                        if (name.Key == PropertyDictionaryId || name.Key == CodePagePropertyId) {
                            continue;
                        }

                        if (!section.Properties.TryGetValue(name.Key, out OfficeOlePropertyValue? value)) {
                            continue;
                        }

                        if (TryCreateCustomPropertyValue(value, out LegacyDocDocumentPropertyValue? customValue)) {
                            document.DocumentProperties.SetCustomProperty(name.Value, customValue!);
                        } else {
                            document.AddUnsupportedFeature(
                                new LegacyDocUnsupportedFeature(
                                    LegacyDocUnsupportedFeatureKind.DocumentProperty,
                                    "DOC-OLE-CUSTOM-DOCUMENT-PROPERTY-UNSUPPORTED",
                                    $"The OLE custom document property '{name.Value}' uses unsupported VARTYPE 0x{value.Type:X4}; the property is preserved in the source file but is not projected into the OfficeIMO document.",
                                    entryPath: DocumentSummaryInformationStreamName,
                                    detailCode: $"OleProperty:VT=0x{value.Type:X4}"),
                                options.ReportUnsupportedContent);
                        }
                    }
                }
            } catch (Exception ex) when (ex is IOException || ex is ArgumentException || ex is InvalidDataException || ex is OverflowException) {
                document.AddWarning("DOC-OLE-PROPERTIES-UNREADABLE", $"The OLE document property stream '{DocumentSummaryInformationStreamName}' could not be read. {ex.Message}");
            }
        }

        private static bool TryCreateCustomPropertyValue(OfficeOlePropertyValue value, out LegacyDocDocumentPropertyValue? customValue) {
            customValue = null;
            object? rawValue = value.Value;
            if (rawValue == null) {
                return false;
            }

            switch (rawValue) {
                case string text:
                    customValue = new LegacyDocDocumentPropertyValue(text, LegacyDocDocumentPropertyValueKind.Text);
                    return true;
                case bool boolean:
                    customValue = new LegacyDocDocumentPropertyValue(boolean, LegacyDocDocumentPropertyValueKind.Boolean);
                    return true;
                case DateTime dateTime:
                    customValue = new LegacyDocDocumentPropertyValue(dateTime, LegacyDocDocumentPropertyValueKind.DateTime);
                    return true;
                case sbyte signedByte:
                    customValue = new LegacyDocDocumentPropertyValue((int)signedByte, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case byte unsignedByte:
                    customValue = new LegacyDocDocumentPropertyValue((int)unsignedByte, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case short signedShort:
                    customValue = new LegacyDocDocumentPropertyValue((int)signedShort, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case ushort unsignedShort:
                    customValue = new LegacyDocDocumentPropertyValue((int)unsignedShort, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case int integer:
                    customValue = new LegacyDocDocumentPropertyValue(integer, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case long integer64 when integer64 >= int.MinValue && integer64 <= int.MaxValue:
                    customValue = new LegacyDocDocumentPropertyValue((int)integer64, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case long integer64:
                    customValue = new LegacyDocDocumentPropertyValue(integer64, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case uint unsignedInteger when unsignedInteger <= int.MaxValue:
                    customValue = new LegacyDocDocumentPropertyValue((int)unsignedInteger, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case uint unsignedInteger:
                    customValue = new LegacyDocDocumentPropertyValue((long)unsignedInteger, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case ulong unsignedInteger64 when unsignedInteger64 <= int.MaxValue:
                    customValue = new LegacyDocDocumentPropertyValue((int)unsignedInteger64, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case ulong unsignedInteger64 when unsignedInteger64 <= long.MaxValue:
                    customValue = new LegacyDocDocumentPropertyValue((long)unsignedInteger64, LegacyDocDocumentPropertyValueKind.Integer);
                    return true;
                case double number:
                    customValue = new LegacyDocDocumentPropertyValue(number, LegacyDocDocumentPropertyValueKind.Number);
                    return true;
                case float number:
                    customValue = new LegacyDocDocumentPropertyValue((double)number, LegacyDocDocumentPropertyValueKind.Number);
                    return true;
                case byte[] bytes:
                    customValue = new LegacyDocDocumentPropertyValue((byte[])bytes.Clone(), LegacyDocDocumentPropertyValueKind.Binary);
                    return true;
            }

            return false;
        }
    }
}
