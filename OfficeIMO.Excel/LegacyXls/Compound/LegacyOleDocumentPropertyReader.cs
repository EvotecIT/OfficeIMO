using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Shared;
using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Compound {
    internal static class LegacyOleDocumentPropertyReader {
        private const string SummaryInformationStreamName = "\u0005SummaryInformation";
        private const string DocumentSummaryInformationStreamName = "\u0005DocumentSummaryInformation";
        private const uint PropertyDictionaryId = 0;
        private const uint CodePagePropertyId = 1;
        private const string UnsupportedCustomPropertyCode = "XLS-OLE-CUSTOM-DOCUMENT-PROPERTY-UNSUPPORTED";

        internal static void AddDocumentProperties(OfficeCompoundFile compoundFile, LegacyXlsWorkbook workbook, LegacyXlsImportOptions options) {
            if (compoundFile == null) throw new ArgumentNullException(nameof(compoundFile));
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var properties = new LegacyXlsDocumentProperties();
            bool parsedAny = false;

            if (TryGetStream(compoundFile.Streams, SummaryInformationStreamName, out byte[]? summaryStream)) {
                parsedAny |= TryReadSummaryInformation(summaryStream!, properties, workbook);
            }

            if (TryGetStream(compoundFile.Streams, DocumentSummaryInformationStreamName, out byte[]? documentSummaryStream)) {
                parsedAny |= TryReadDocumentSummaryInformation(documentSummaryStream!, properties, workbook, options);
            }

            if (parsedAny && properties.HasAnyProperties) {
                workbook.SetDocumentProperties(properties);
            }
        }

        private static bool TryGetStream(IReadOnlyDictionary<string, byte[]> streams, string name, out byte[]? bytes) {
            return streams.TryGetValue(name, out bytes) && bytes != null && bytes.Length > 0;
        }

        private static bool TryReadSummaryInformation(byte[] bytes, LegacyXlsDocumentProperties target, LegacyXlsWorkbook workbook) {
            try {
                IReadOnlyList<OfficeOlePropertySection> sections = OfficeOlePropertySetReader.ReadSections(bytes);
                foreach (OfficeOlePropertySection section in sections) {
                    foreach (KeyValuePair<uint, OfficeOlePropertyValue> property in section.Properties) {
                        switch (property.Key) {
                            case 2:
                                target.Title = property.Value.AsString();
                                break;
                            case 3:
                                target.Subject = property.Value.AsString();
                                break;
                            case 4:
                                target.Creator = property.Value.AsString();
                                break;
                            case 5:
                                target.Keywords = property.Value.AsString();
                                break;
                            case 6:
                                target.Description = property.Value.AsString();
                                break;
                            case 8:
                                target.LastModifiedBy = property.Value.AsString();
                                break;
                            case 9:
                                target.Revision = property.Value.AsString();
                                break;
                            case 11:
                                target.LastPrinted = property.Value.AsDateTime();
                                break;
                            case 12:
                                target.Created = property.Value.AsDateTime();
                                break;
                            case 13:
                                target.Modified = property.Value.AsDateTime();
                                break;
                        }
                    }
                }

                return true;
            } catch (Exception ex) when (ex is IOException || ex is ArgumentException || ex is InvalidDataException || ex is OverflowException) {
                AddPropertyWarning(workbook, SummaryInformationStreamName, ex);
                return false;
            }
        }

        private static bool TryReadDocumentSummaryInformation(byte[] bytes, LegacyXlsDocumentProperties target, LegacyXlsWorkbook workbook, LegacyXlsImportOptions options) {
            try {
                IReadOnlyList<OfficeOlePropertySection> sections = OfficeOlePropertySetReader.ReadSections(bytes);
                bool parsed = false;
                foreach (OfficeOlePropertySection section in sections) {
                    if (section.Dictionary.Count == 0) {
                        if (section.Properties.TryGetValue(2, out OfficeOlePropertyValue? category)) {
                            target.Category = category.AsString();
                            parsed = true;
                        }

                        if (section.Properties.TryGetValue(14, out OfficeOlePropertyValue? manager)) {
                            target.Manager = manager.AsString();
                            parsed = true;
                        }

                        if (section.Properties.TryGetValue(15, out OfficeOlePropertyValue? company)) {
                            target.Company = company.AsString();
                            parsed = true;
                        }
                    }

                    foreach (KeyValuePair<uint, string> name in section.Dictionary) {
                        if (name.Key == PropertyDictionaryId || name.Key == CodePagePropertyId) {
                            continue;
                        }

                        if (!section.Properties.TryGetValue(name.Key, out OfficeOlePropertyValue? value)) {
                            continue;
                        }

                        if (TryCreateCustomPropertyValue(value, out LegacyXlsDocumentPropertyValue? customValue)) {
                            target.SetCustomProperty(name.Value, customValue!);
                            parsed = true;
                        } else {
                            AddUnsupportedCustomProperty(workbook, options, name.Key, name.Value, value.Type);
                        }
                    }
                }

                return parsed;
            } catch (Exception ex) when (ex is IOException || ex is ArgumentException || ex is InvalidDataException || ex is OverflowException) {
                AddPropertyWarning(workbook, DocumentSummaryInformationStreamName, ex);
                return false;
            }
        }

        private static void AddUnsupportedCustomProperty(
            LegacyXlsWorkbook workbook,
            LegacyXlsImportOptions options,
            uint propertyId,
            string propertyName,
            ushort propertyType) {
            string detailCode = $"DocumentProperty:Custom:PropertyId:0x{propertyId:X4}:Type:0x{propertyType:X4}";
            string description = $"The OLE custom document property '{propertyName}' uses unsupported VARTYPE 0x{propertyType:X4}; the property was not projected.";
            var feature = new LegacyXlsUnsupportedFeature(
                LegacyXlsUnsupportedFeatureKind.DocumentProperty,
                UnsupportedCustomPropertyCode,
                description,
                detailCode: detailCode);
            workbook.MutableUnsupportedFeatures.Add(feature);
            if (options.ReportUnsupportedRecords) {
                workbook.MutableDiagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Info,
                    feature.Code,
                    feature.Description,
                    detailCode: feature.DetailCode));
            }
        }

        private static bool TryCreateCustomPropertyValue(OfficeOlePropertyValue value, out LegacyXlsDocumentPropertyValue? customValue) {
            customValue = null;
            object? rawValue = value.Value;
            if (rawValue == null) {
                return false;
            }

            switch (rawValue) {
                case sbyte signedByte:
                    customValue = new LegacyXlsDocumentPropertyValue((int)signedByte, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case byte unsignedByte:
                    customValue = new LegacyXlsDocumentPropertyValue(unsignedByte, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case short signedShort:
                    customValue = new LegacyXlsDocumentPropertyValue((int)signedShort, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case ushort unsignedShort:
                    customValue = new LegacyXlsDocumentPropertyValue(unsignedShort, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case string text:
                    customValue = new LegacyXlsDocumentPropertyValue(text, LegacyXlsDocumentPropertyValueKind.Text);
                    return true;
                case bool boolean:
                    customValue = new LegacyXlsDocumentPropertyValue(boolean, LegacyXlsDocumentPropertyValueKind.Boolean);
                    return true;
                case DateTime dateTime:
                    customValue = new LegacyXlsDocumentPropertyValue(dateTime, LegacyXlsDocumentPropertyValueKind.DateTime);
                    return true;
                case int integer:
                    customValue = new LegacyXlsDocumentPropertyValue(integer, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case long integer64 when integer64 >= int.MinValue && integer64 <= int.MaxValue:
                    customValue = new LegacyXlsDocumentPropertyValue((int)integer64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case long integer64:
                    customValue = new LegacyXlsDocumentPropertyValue(integer64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case uint unsignedInteger:
                    customValue = new LegacyXlsDocumentPropertyValue(unsignedInteger, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case ulong unsignedInteger64 when unsignedInteger64 <= int.MaxValue:
                    customValue = new LegacyXlsDocumentPropertyValue((int)unsignedInteger64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case ulong unsignedInteger64 when unsignedInteger64 <= long.MaxValue:
                    customValue = new LegacyXlsDocumentPropertyValue((long)unsignedInteger64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case ulong unsignedInteger64:
                    customValue = new LegacyXlsDocumentPropertyValue(unsignedInteger64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case double number:
                    customValue = new LegacyXlsDocumentPropertyValue(number, LegacyXlsDocumentPropertyValueKind.Number);
                    return true;
                case float number:
                    customValue = new LegacyXlsDocumentPropertyValue((double)number, LegacyXlsDocumentPropertyValueKind.Number);
                    return true;
                case byte[] bytes:
                    customValue = new LegacyXlsDocumentPropertyValue((byte[])bytes.Clone(), LegacyXlsDocumentPropertyValueKind.Binary);
                    return true;
            }

            return false;
        }

        private static void AddPropertyWarning(LegacyXlsWorkbook workbook, string streamName, Exception exception) {
            workbook.MutableDiagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Warning,
                "XLS-OLE-PROPERTIES-UNREADABLE",
                string.Format(CultureInfo.InvariantCulture, "The OLE document property stream '{0}' could not be read. {1}", streamName, exception.Message)));
        }
    }
}
