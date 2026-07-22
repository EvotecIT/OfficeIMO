using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    internal static class ExcelTemplateBindingHelper {
        internal static Dictionary<string, object?> Create(IDictionary<string, object?> values) {
            return new Dictionary<string, object?>(values, StringComparer.OrdinalIgnoreCase);
        }

        [RequiresUnreferencedCode("Object-model template binding walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        internal static Dictionary<string, object?> Create(object model) {
            if (model is IDictionary<string, object?> objectDictionary) {
                return Create(objectDictionary);
            }

            var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            AddBindings(values, prefix: null, model, depth: 0);
            return values;
        }

        [RequiresUnreferencedCode("Object-model template binding walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        private static void AddBindings(IDictionary<string, object?> values, string? prefix, object? source, int depth) {
            if (source == null || depth > 8) {
                return;
            }

            Type type = source.GetType();
            if (IsSimpleValue(type)) {
                if (prefix != null) {
                    values[prefix] = source;
                }
                return;
            }

            foreach (PropertyInfo property in type.GetProperties(BindingFlags.Instance | BindingFlags.Public)) {
                if (property.GetIndexParameters().Length != 0) {
                    continue;
                }

                object? value;
                try {
                    value = property.GetValue(source);
                } catch (TargetInvocationException) {
                    continue;
                }

                string key = prefix == null ? property.Name : prefix + "." + property.Name;
                values[key] = value;

                if (value != null && !IsSimpleValue(value.GetType())) {
                    AddBindings(values, key, value, depth + 1);
                }
            }
        }

        private static bool IsSimpleValue(Type type) {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(string)
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)
                || type == typeof(Guid)
                || type == typeof(ExcelTemplateImage);
        }
    }
}
