#nullable enable

using System.Diagnostics.CodeAnalysis;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// UTF-8 object-mapping readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private IEnumerable<T> ReadObjectsStreamUtf8OrXmlAdaptive<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            CancellationToken ct) where T : new() {
            if (!ShouldAttemptUtf8Range(r1, r2)
                || !RangeReachesDeclaredWorksheetEnd(r2)
                || !ExcelUtf8RangeRowSource.TryCreate(this, r1, r2, c1, cols, ct, out var source)) {
                foreach (T item in ReadObjectsStreamXmlAdaptive<T>(a1Range, r1, c1, r2, c2, cols, ct)) {
                    yield return item;
                }

                yield break;
            }

            bool useUtf8Source = false;
            using (source) {
                if (source!.SelectRow(r1)) {
                    var headerValues = new object?[cols];
                    for (int columnOffset = 0; columnOffset < cols; columnOffset++) {
                        source.ReadValue(
                            columnOffset,
                            XmlDataReaderTargetKind.None,
                            out _,
                            out _,
                            out _,
                            out _,
                            out headerValues[columnOffset]);
                    }

                    var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(
                        cols,
                        columnOffset => headerValues[columnOffset]?.ToString(),
                        _opt.NormalizeHeaders);
                    TypedPropertyBinding<T>?[] bindings = GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
                    useUtf8Source = CanUseUtf8TypedBindings(bindings);

                    if (useUtf8Source) {
                        bool canCancel = ct.CanBeCanceled;
                        for (int rowIndex = r1 + 1; rowIndex <= r2; rowIndex++) {
                            if (canCancel && ((rowIndex - r1) & 1023) == 0) {
                                ct.ThrowIfCancellationRequested();
                            }

                            var target = new T();
                            if (source.SelectRow(rowIndex)) {
                                for (int columnOffset = 0; columnOffset < bindings.Length; columnOffset++) {
                                    TypedPropertyBinding<T>? binding = bindings[columnOffset];
                                    if (binding != null) {
                                        ReadUtf8ValueIntoTypedObject(source, columnOffset, binding, target);
                                    }
                                }
                            }

                            yield return target;
                        }
                    }
                }
            }

            if (useUtf8Source) {
                yield break;
            }

            foreach (T item in ReadObjectsStreamXmlAdaptive<T>(a1Range, r1, c1, r2, c2, cols, ct)) {
                yield return item;
            }
        }

        private static bool CanUseUtf8TypedBindings<T>(TypedPropertyBinding<T>?[] bindings) {
            for (int i = 0; i < bindings.Length; i++) {
                TypedPropertyBinding<T>? binding = bindings[i];
                if (binding == null) {
                    continue;
                }

                switch (binding.BindingKind) {
                    case TypedBindingKind.String:
                    case TypedBindingKind.Int32:
                    case TypedBindingKind.Double:
                    case TypedBindingKind.Boolean:
                    case TypedBindingKind.DateTime:
                        continue;
                    default:
                        return false;
                }
            }

            return true;
        }

        private void ReadUtf8ValueIntoTypedObject<T>(
            ExcelUtf8RangeRowSource source,
            int columnOffset,
            TypedPropertyBinding<T> binding,
            T target) {
            source.ReadValue(
                columnOffset,
                GetUtf8TargetKind(binding.BindingKind),
                out XmlDataReaderPrimitiveKind primitiveKind,
                out double doubleValue,
                out DateTime dateTimeValue,
                out bool booleanValue,
                out object? objectValue);

            switch (primitiveKind) {
                case XmlDataReaderPrimitiveKind.Double:
                    if (binding.BindingKind == TypedBindingKind.Int32
                        && binding.SetInt32 != null
                        && doubleValue >= int.MinValue
                        && doubleValue <= int.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        binding.SetInt32(target, (int)doubleValue);
                    } else if (binding.BindingKind == TypedBindingKind.Double && binding.SetDouble != null) {
                        binding.SetDouble(target, doubleValue);
                    }

                    return;
                case XmlDataReaderPrimitiveKind.DateTime:
                    if (binding.BindingKind == TypedBindingKind.DateTime && binding.SetDateTime != null) {
                        binding.SetDateTime(target, dateTimeValue);
                    }

                    return;
                case XmlDataReaderPrimitiveKind.Boolean:
                    if (binding.BindingKind == TypedBindingKind.Boolean && binding.SetBoolean != null) {
                        binding.SetBoolean(target, booleanValue);
                    }

                    return;
            }

            if (objectValue == null) {
                return;
            }

            if (objectValue is string text && TrySetStringTextBinding(text, binding, target)) {
                return;
            }

            if (objectValue is double number) {
                if (binding.BindingKind == TypedBindingKind.Int32
                    && binding.SetInt32 != null
                    && number >= int.MinValue
                    && number <= int.MaxValue
                    && Math.Truncate(number) == number) {
                    binding.SetInt32(target, (int)number);
                    return;
                }

                if (binding.BindingKind == TypedBindingKind.Double && binding.SetDouble != null) {
                    binding.SetDouble(target, number);
                    return;
                }

                if (binding.BindingKind == TypedBindingKind.Boolean
                    && binding.SetBoolean != null
                    && (number == 0d || number == 1d)) {
                    binding.SetBoolean(target, number == 1d);
                    return;
                }
            } else if (objectValue is DateTime dateValue
                && binding.BindingKind == TypedBindingKind.DateTime
                && binding.SetDateTime != null) {
                binding.SetDateTime(target, dateValue);
                return;
            } else if (objectValue is bool boolValue
                && binding.BindingKind == TypedBindingKind.Boolean
                && binding.SetBoolean != null) {
                binding.SetBoolean(target, boolValue);
                return;
            }

            object? converted = binding.ConvertValue(objectValue, _opt.Culture);
            if (converted != null || binding.IsNullable) {
                binding.SetValue(target, converted);
            }
        }

        private static XmlDataReaderTargetKind GetUtf8TargetKind(TypedBindingKind bindingKind) =>
            bindingKind switch {
                TypedBindingKind.Int32 => XmlDataReaderTargetKind.Int32,
                TypedBindingKind.Double => XmlDataReaderTargetKind.Double,
                TypedBindingKind.Boolean => XmlDataReaderTargetKind.Boolean,
                TypedBindingKind.DateTime => XmlDataReaderTargetKind.DateTime,
                TypedBindingKind.String => XmlDataReaderTargetKind.String,
                _ => XmlDataReaderTargetKind.None
            };
    }
}
