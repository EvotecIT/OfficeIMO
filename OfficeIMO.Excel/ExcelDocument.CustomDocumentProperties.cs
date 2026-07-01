using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Sets or replaces a custom workbook property.
        /// </summary>
        public void SetCustomDocumentProperty(string name, ExcelCustomProperty property) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Custom property name is required.", nameof(name));
            }

            if (property == null) {
                throw new ArgumentNullException(nameof(property));
            }

            CustomDocumentProperties[name.Trim()] = property;
            _customDocumentPropertiesDirty = true;
            DisqualifyDirectDataSetFastSaveState();
            MarkPackageDirty();
        }

        /// <summary>
        /// Sets or replaces a custom workbook property, inferring the custom property type from the value.
        /// </summary>
        public void SetCustomDocumentProperty(string name, object? value) {
            SetCustomDocumentProperty(name, CreateCustomProperty(value));
        }

        /// <summary>
        /// Removes a custom workbook property.
        /// </summary>
        public bool RemoveCustomDocumentProperty(string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                return false;
            }

            bool removed = CustomDocumentProperties.Remove(name.Trim());
            if (removed) {
                _customDocumentPropertiesDirty = true;
                DisqualifyDirectDataSetFastSaveState();
                MarkPackageDirty();
            }

            return removed;
        }

        internal void LoadCustomDocumentProperties() {
            var loadedProperties = new Dictionary<string, ExcelCustomProperty>(StringComparer.OrdinalIgnoreCase);
            _customDocumentPropertiesDirty = false;
            CustomFilePropertiesPart? customPart = _spreadSheetDocument.CustomFilePropertiesPart;
            if (customPart?.Properties == null) {
                CustomDocumentProperties.ReplaceWith(loadedProperties);
                return;
            }

            foreach (CustomDocumentProperty property in customPart.Properties.Elements<CustomDocumentProperty>()) {
                if (string.IsNullOrWhiteSpace(property.Name?.Value)) {
                    continue;
                }

                loadedProperties[property.Name!.Value!] = new ExcelCustomProperty(property);
            }

            CustomDocumentProperties.ReplaceWith(loadedProperties);
        }

        internal void SaveCustomDocumentProperties() {
            CustomFilePropertiesPart? customPart = _spreadSheetDocument.CustomFilePropertiesPart;
            if (CustomDocumentProperties.Count == 0) {
                if (_customDocumentPropertiesDirty && customPart != null) {
                    _spreadSheetDocument.DeletePart(customPart);
                }

                _customDocumentPropertiesDirty = false;
                return;
            }

            if (customPart == null) {
                customPart = _spreadSheetDocument.AddCustomFilePropertiesPart();
            }

            customPart.Properties = new Properties();
            int propertyId = 2;
            foreach (var pair in CustomDocumentProperties.OrderBy(property => property.Key, StringComparer.OrdinalIgnoreCase)) {
                CustomDocumentProperty property = pair.Value.ToOpenXml(pair.Key);
                property.PropertyId = propertyId++;
                customPart.Properties.Append(property);
            }

            customPart.Properties.Save();
            _customDocumentPropertiesDirty = false;
        }

        private static ExcelCustomProperty CreateCustomProperty(object? value) {
            if (value == null) {
                return new ExcelCustomProperty(string.Empty);
            }

            if (value is ExcelCustomProperty customProperty) {
                return customProperty;
            }

            if (value is byte[] bytes) {
                return new ExcelCustomProperty(bytes);
            }

            if (value is bool boolean) {
                return new ExcelCustomProperty(boolean);
            }

            if (value is DateTime dateTime) {
                return new ExcelCustomProperty(dateTime);
            }

            if (value is DateTimeOffset dateTimeOffset) {
                return new ExcelCustomProperty(dateTimeOffset.UtcDateTime);
            }

            if (TryConvertToInt32(value, out int integer)) {
                return new ExcelCustomProperty(integer);
            }

            if (value is long or uint or ulong) {
                return new ExcelCustomProperty(value, ExcelCustomPropertyType.NumberInteger);
            }

            if (value is float or double or decimal) {
                return new ExcelCustomProperty(Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture));
            }

            return new ExcelCustomProperty(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
        }

        private static bool TryConvertToInt32(object value, out int result) {
            switch (value) {
                case byte byteValue:
                    result = byteValue;
                    return true;
                case sbyte sbyteValue:
                    result = sbyteValue;
                    return true;
                case short shortValue:
                    result = shortValue;
                    return true;
                case ushort ushortValue:
                    result = ushortValue;
                    return true;
                case int intValue:
                    result = intValue;
                    return true;
                case long longValue when longValue >= int.MinValue && longValue <= int.MaxValue:
                    result = (int)longValue;
                    return true;
                case uint uintValue when uintValue <= int.MaxValue:
                    result = (int)uintValue;
                    return true;
                case ulong ulongValue when ulongValue <= int.MaxValue:
                    result = (int)ulongValue;
                    return true;
                default:
                    result = default;
                    return false;
            }
        }

        private void DisqualifyDirectDataSetFastSaveState() {
            ClearDirectDataSetSaveCandidate();
            _materializedDirectDataSetFastSaveModel = null;
            _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = false;
        }

        private void MarkCustomDocumentPropertiesChanged() {
            _customDocumentPropertiesDirty = true;
            DisqualifyDirectDataSetFastSaveState();
            MarkPackageDirty();
        }
    }
}
