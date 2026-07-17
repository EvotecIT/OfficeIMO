using System.Globalization;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPropertySetCodec {
        private static byte[] CreateEmptyPropertySet(Guid formatId) {
            var properties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(
                    OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            return OfficeOlePropertySetWriter.CreatePropertySet((formatId,
                OfficeOlePropertySetWriter.CreateSection(properties)));
        }

        private static byte[]? CreateSummaryInformation(
            PowerPointPresentation presentation) {
            var properties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(
                    OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            PowerPointBuiltinDocumentProperties builtIn = presentation
                .BuiltinDocumentProperties;
            OfficeOlePropertySetWriter.AddString(properties, 2, builtIn.Title);
            OfficeOlePropertySetWriter.AddString(properties, 3, builtIn.Subject);
            OfficeOlePropertySetWriter.AddString(properties, 4, builtIn.Creator);
            OfficeOlePropertySetWriter.AddString(properties, 5, builtIn.Keywords);
            OfficeOlePropertySetWriter.AddString(properties, 6,
                builtIn.Description);
            OfficeOlePropertySetWriter.AddString(properties, 8,
                builtIn.LastModifiedBy);
            OfficeOlePropertySetWriter.AddString(properties, 9,
                builtIn.Revision);
            if (int.TryParse(presentation.ApplicationProperties.TotalTime,
                    NumberStyles.Integer, CultureInfo.InvariantCulture,
                    out int totalMinutes) && totalMinutes >= 0) {
                properties.Add(OfficeOleProperty.FileTimeDuration(10,
                    TimeSpan.FromMinutes(totalMinutes)));
            }
            OfficeOlePropertySetWriter.AddFileTime(properties, 11,
                builtIn.LastPrinted);
            OfficeOlePropertySetWriter.AddFileTime(properties, 12,
                builtIn.Created);
            OfficeOlePropertySetWriter.AddFileTime(properties, 13,
                builtIn.Modified);
            OfficeOlePropertySetWriter.AddString(properties, 18,
                presentation.ApplicationProperties.Application);
            return properties.Count == 1 ? null
                : OfficeOlePropertySetWriter.CreatePropertySet((
                    OfficeOlePropertySetWriter.SummaryInformationFormatId,
                    OfficeOlePropertySetWriter.CreateSection(properties)));
        }

        private static bool TryCreateDocumentSummaryInformation(
            PowerPointPresentation presentation, out byte[]? bytes,
            out string? reason) {
            var sections = new List<(Guid FormatId, byte[] Section)>();
            var properties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(
                    OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            OfficeOlePropertySetWriter.AddString(properties, 2,
                presentation.BuiltinDocumentProperties.Category);
            OfficeOlePropertySetWriter.AddString(properties, 3,
                presentation.ApplicationProperties.PresentationFormat);
            AddIntegerText(properties, 7,
                presentation.ApplicationProperties.Slides);
            AddIntegerText(properties, 8,
                presentation.ApplicationProperties.Notes);
            AddIntegerText(properties, 9,
                presentation.ApplicationProperties.HiddenSlides);
            OfficeOlePropertySetWriter.AddString(properties, 14,
                presentation.ApplicationProperties.Manager);
            OfficeOlePropertySetWriter.AddString(properties, 15,
                presentation.ApplicationProperties.Company);
            if (properties.Count > 1) {
                sections.Add((OfficeOlePropertySetWriter
                    .DocumentSummaryInformationFormatId,
                    OfficeOlePropertySetWriter.CreateSection(properties)));
            }
            if (!TryCreateCustomSection(presentation.OpenXmlDocument,
                    out byte[]? customSection, out reason)) {
                bytes = null;
                return false;
            }
            if (customSection != null) {
                sections.Add((OfficeOlePropertySetWriter
                    .UserDefinedPropertiesFormatId, customSection));
            }
            bytes = sections.Count == 0 ? null
                : OfficeOlePropertySetWriter.CreatePropertySet(
                    sections.ToArray());
            reason = null;
            return true;
        }

        private static void AddIntegerText(ICollection<OfficeOleProperty> target,
            uint id, string? text) {
            if (int.TryParse(text, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out int value)) {
                target.Add(OfficeOleProperty.Int32(id, value));
            }
        }

        private static bool TryCreateCustomSection(
            PresentationDocument document, out byte[]? section,
            out string? reason) {
            section = null;
            reason = null;
            CustomDocumentProperty[] source = document
                .CustomFilePropertiesPart?.Properties?
                .Elements<CustomDocumentProperty>().ToArray()
                ?? Array.Empty<CustomDocumentProperty>();
            if (source.Length == 0) return true;
            var properties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(
                    OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            var names = new Dictionary<uint, string>();
            uint propertyId = 2;
            foreach (CustomDocumentProperty item in source) {
                if (string.IsNullOrWhiteSpace(item.Name?.Value)
                    || !TryCreateOleCustomProperty(propertyId, item,
                        out OfficeOleProperty property)) {
                    reason = "A custom document property has no name or uses an unsupported Open XML variant type.";
                    return false;
                }
                names.Add(propertyId, item.Name!.Value!);
                properties.Add(property);
                propertyId++;
            }
            properties.Insert(1, OfficeOleProperty.Dictionary(0, names));
            section = OfficeOlePropertySetWriter.CreateSection(properties);
            return true;
        }

        private static bool TryCreateOleCustomProperty(uint id,
            CustomDocumentProperty source, out OfficeOleProperty property) {
            if (source.VTLPWSTR != null) {
                property = OfficeOleProperty.String(id,
                    source.VTLPWSTR.Text ?? string.Empty);
            } else if (source.VTBool != null
                       && TryParseBoolean(source.VTBool.Text,
                           out bool boolean)) {
                property = OfficeOleProperty.Boolean(id, boolean);
            } else if (source.VTFileTime != null
                       && DateTime.TryParse(source.VTFileTime.Text,
                           CultureInfo.InvariantCulture,
                           DateTimeStyles.AssumeUniversal
                           | DateTimeStyles.AdjustToUniversal,
                           out DateTime date)) {
                property = OfficeOleProperty.FileTime(id, date);
            } else if (source.VTDouble != null
                       && double.TryParse(source.VTDouble.Text,
                           NumberStyles.Float, CultureInfo.InvariantCulture,
                           out double number)) {
                property = OfficeOleProperty.Double(id, number);
            } else if (source.VTFloat != null
                       && float.TryParse(source.VTFloat.Text,
                           NumberStyles.Float, CultureInfo.InvariantCulture,
                           out float single)) {
                property = OfficeOleProperty.Float(id, single);
            } else if (source.VTInt32 != null
                       && int.TryParse(source.VTInt32.Text,
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out int integer)) {
                property = OfficeOleProperty.Int32(id, integer);
            } else if (source.VTInt64 != null
                       && long.TryParse(source.VTInt64.Text,
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out long integer64)) {
                property = OfficeOleProperty.Int64(id, integer64);
            } else if (source.VTUnsignedByte != null
                       && byte.TryParse(source.VTUnsignedByte.Text,
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out byte uint8)) {
                property = OfficeOleProperty.UInt8(id, uint8);
            } else if (source.VTUnsignedShort != null
                       && ushort.TryParse(source.VTUnsignedShort.Text,
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out ushort uint16)) {
                property = OfficeOleProperty.UInt16(id, uint16);
            } else if (source.VTUnsignedInt32 != null
                       && uint.TryParse(source.VTUnsignedInt32.Text,
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out uint uint32)) {
                property = OfficeOleProperty.UInt32(id, uint32);
            } else if (source.VTUnsignedInt64 != null
                       && ulong.TryParse(source.VTUnsignedInt64.Text,
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out ulong uint64)) {
                property = OfficeOleProperty.UInt64(id, uint64);
            } else if (source.VTBlob != null
                       && TryParseBase64(source.VTBlob.Text,
                           out byte[] blob)) {
                property = OfficeOleProperty.Blob(id, blob);
            } else {
                property = default;
                return false;
            }
            return true;
        }

        private static bool TryParseBoolean(string? text, out bool value) {
            if (bool.TryParse(text, out value)) return true;
            if (text == "1" || text == "-1") {
                value = true;
                return true;
            }
            if (text == "0") {
                value = false;
                return true;
            }
            return false;
        }

        private static bool TryParseBase64(string? text, out byte[] value) {
            try {
                value = string.IsNullOrEmpty(text)
                    ? Array.Empty<byte>()
                    : Convert.FromBase64String(text);
                return true;
            } catch (FormatException) {
                value = Array.Empty<byte>();
                return false;
            }
        }
    }
}
