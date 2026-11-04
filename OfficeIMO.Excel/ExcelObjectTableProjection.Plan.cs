using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Data;
using System.Reflection;
using System.Text;
using OfficeIMO.Data;
using OfficeIMO.Spreadsheet;
using OfficeColor = OfficeIMO.OfficeColor;

namespace OfficeIMO.Excel;

internal sealed partial class ExcelObjectTableProjection {
        private sealed class RowsFromSimpleTypePlan {
            internal static readonly RowsFromSimpleTypePlan NotSupported = new(
                Array.Empty<RowsFromSimpleValueGetter>(),
                Array.Empty<string>(),
                Array.Empty<Type>(),
                Array.Empty<bool>(),
                Array.Empty<bool>(),
                Array.Empty<Type>(),
                Array.Empty<bool>(),
                Array.Empty<int>(),
                hasInferenceColumns: false,
                canUseDirectSave: false);

            internal RowsFromSimpleTypePlan(
                RowsFromSimpleValueGetter[] getters,
                string[] headers,
                Type[] staticColumnTypes,
                bool[] staticBlankAsEmptyString,
                bool[] staticRequiresBlankCoercion,
                Type[] inferenceFallbackTypes,
                bool[] inferColumns,
                int[] inferenceColumnIndexes,
                bool hasInferenceColumns,
                bool canUseDirectSave) {
                Getters = getters;
                Headers = headers;
                StaticColumnTypes = staticColumnTypes;
                StaticBlankAsEmptyString = staticBlankAsEmptyString;
                StaticRequiresBlankCoercion = staticRequiresBlankCoercion;
                InferenceFallbackTypes = inferenceFallbackTypes;
                InferColumns = inferColumns;
                InferenceColumnIndexes = inferenceColumnIndexes;
                HasInferenceColumns = hasInferenceColumns;
                CanUseDirectSave = canUseDirectSave;
            }

            internal RowsFromSimpleValueGetter[] Getters { get; }

            internal string[] Headers { get; }

            internal Type[] StaticColumnTypes { get; }

            internal bool[] StaticBlankAsEmptyString { get; }

            internal bool[] StaticRequiresBlankCoercion { get; }

            internal Type[] InferenceFallbackTypes { get; }

            internal bool[] InferColumns { get; }

            internal int[] InferenceColumnIndexes { get; }

            internal bool HasInferenceColumns { get; }

            internal bool CanUseDirectSave { get; }
        }

        private delegate object? RowsFromSimpleValueGetter(object? row);
}
