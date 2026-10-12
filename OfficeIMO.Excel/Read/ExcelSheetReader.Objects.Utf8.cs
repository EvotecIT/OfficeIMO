#nullable enable

using System.Diagnostics.CodeAnalysis;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// UTF-8 object-mapping readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    internal sealed partial class ExcelSheetReader {
        private IEnumerable<T> ReadObjectsStreamUtf8OrXmlAdaptive<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range, int r1, int c1, int r2, int c2, int cols, CancellationToken ct) where T : new() {
            if (TryCreateTypedObjectsUtf8Source<T>(a1Range, r1, c1, r2, c2, cols, ct, out var source, out var bindings)) {
                using (source) {
                    foreach (T item in ReadTypedObjectsUtf8(source!, bindings!, r1, r2, ct)) {
                        yield return item;
                    }
                }
                yield break;
            }

            foreach (T item in ReadObjectsStreamXmlAdaptive<T>(a1Range, r1, c1, r2, c2, cols, ct)) {
                yield return item;
            }
        }

        private bool TryReadObjectsFromUtf8Materialized<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range, int r1, int c1, int r2, int c2, int cols, CancellationToken ct, out List<T> results) where T : new() {
            results = [];
            if (!TryCreateTypedObjectsUtf8Source<T>(a1Range, r1, c1, r2, c2, cols, ct, out var source, out var bindings, allowSmallRange: true)) {
                return false;
            }
            using (source) {
                results = ReadTypedObjectsUtf8(source!, bindings!, r1, r2, ct).ToList();
            }
            return true;
        }

        // Successful creation proves the complete worksheet's row order. Declined sources
        // are disposed here so materializers can select their own bounded XML fallback.
        private bool TryCreateTypedObjectsUtf8Source<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range, int r1, int c1, int r2, int c2, int cols, CancellationToken ct,
            out ExcelUtf8RangeRowSource? source, out TypedPropertyBinding<T>?[]? bindings,
            bool allowSmallRange = false) where T : new() {
            source = null;
            bindings = null;
            if ((!allowSmallRange && !ShouldAttemptUtf8Range(r1, r2))
                || !RangeReachesDeclaredWorksheetEnd(r2)) {
                return false;
            }

            try {
                if (!ExcelUtf8RangeRowSource.TryCreate(this, r1, r2, c1, cols, ct, out source)) return false;
            } catch (NotSupportedException) {
                // Indexing checks every cell, including shared formulas outside the
                // requested columns. Decline before invoking any user mapping code;
                // the typed XML reader can handle this narrower projection.
                return false;
            }

            try {
                if (source!.SelectRow(r1, ct, ct)) {
                    var headerValues = new object?[cols];
                    for (int columnOffset = 0; columnOffset < cols; columnOffset++) {
                        source.ReadValue(columnOffset, XmlDataReaderTargetKind.None,
                            out _, out _, out _, out _, out _, out _, out headerValues[columnOffset]);
                    }
                    var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(
                        cols, columnOffset => headerValues[columnOffset]?.ToString(), _opt.NormalizeHeaders);
                    bindings = GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
                    if (CanUseUtf8TypedBindings(bindings)) return true;
                }
                source.Dispose();
                source = null;
                return false;
            } catch {
                source?.Dispose();
                source = null;
                throw;
            }
        }

        private IEnumerable<T> ReadTypedObjectsUtf8<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            ExcelUtf8RangeRowSource source, TypedPropertyBinding<T>?[] bindings,
            int r1, int r2, CancellationToken ct) where T : new() {
            bool canCancel = ct.CanBeCanceled;
            for (int rowIndex = r1 + 1; rowIndex <= r2; rowIndex++) {
                if (canCancel && ((rowIndex - r1) & 1023) == 0) ct.ThrowIfCancellationRequested();
                var target = new T();
                if (source.SelectRow(rowIndex, ct, ct)) {
                    for (int columnOffset = 0; columnOffset < bindings.Length; columnOffset++) {
                        TypedPropertyBinding<T>? binding = bindings[columnOffset];
                        if (binding != null) ReadUtf8ValueIntoTypedObject(source, columnOffset, binding, target);
                    }
                }
                yield return target;
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
                out bool isFormulaText,
                out _,
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

            if (isFormulaText) {
                if (binding.BindingKind == TypedBindingKind.String && binding.SetString != null) {
                    binding.SetString(target, objectValue as string);
                }

                return;
            }

            if (objectValue == null) {
                if (binding.IsNullable && source.IsCellPresent(columnOffset)) binding.SetValue(target, null);
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
                TypedBindingKind.Int32 => XmlDataReaderTargetKind.Numeric,
                TypedBindingKind.Double => XmlDataReaderTargetKind.Numeric,
                TypedBindingKind.Boolean => XmlDataReaderTargetKind.Boolean,
                TypedBindingKind.DateTime => XmlDataReaderTargetKind.DateTime,
                TypedBindingKind.String => XmlDataReaderTargetKind.String,
                _ => XmlDataReaderTargetKind.None
            };
    }
}
