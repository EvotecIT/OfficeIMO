using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares workbook calculation properties.</summary>
    internal static class XlsbCalculationPropertiesProjector {
        internal static void Apply(ExcelDocument document, XlsbCalculationProperties source) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (source == null) throw new ArgumentNullException(nameof(source));
            CalculationProperties properties = Create(source);
            OpenXmlWorkbookElementOrder.InsertInOrder(document.WorkbookRoot, properties);
        }

        internal static bool Matches(CalculationProperties? actual, XlsbCalculationProperties? expected) {
            if (actual == null || expected == null) return actual == null && expected == null;
            if (actual.HasChildren || actual.GetAttributes().Any(attribute =>
                    !string.IsNullOrEmpty(attribute.NamespaceUri)
                    || !IsSupportedAttribute(attribute.LocalName))) {
                return false;
            }

            return (actual.CalculationId?.Value ?? 0U) == expected.CalculationId
                && ToCalculationMode(actual.CalculationMode?.Value ?? CalculateModeValues.Auto) == expected.CalculationMode
                && (actual.FullCalculationOnLoad?.Value ?? false) == expected.FullCalculationOnLoad
                && (actual.ReferenceMode?.Value != ReferenceModeValues.R1C1) == expected.UsesA1References
                && (actual.Iterate?.Value ?? false) == expected.IterationEnabled
                && (actual.IterateCount?.Value ?? 100U) == expected.IterationCount
                && (actual.IterateDelta?.Value ?? 0.001D).Equals(expected.IterationDelta)
                && (actual.FullPrecision?.Value ?? true) == expected.FullPrecision
                && (actual.CalculationCompleted?.Value ?? true) == expected.CalculationCompleted
                && (actual.CalculationOnSave?.Value ?? true) == expected.CalculationOnSave
                && (actual.ConcurrentCalculation?.Value ?? true) == expected.ConcurrentCalculation
                && MatchesConcurrentCount(actual.ConcurrentManualCount?.Value, expected)
                && (actual.ForceFullCalculation?.Value ?? false) == expected.ForceFullCalculation;
        }

        private static CalculationProperties Create(XlsbCalculationProperties source) {
            var properties = new CalculationProperties {
                CalculationId = source.CalculationId,
                CalculationMode = source.CalculationMode == 0U
                    ? CalculateModeValues.Manual
                    : source.CalculationMode == 2U ? CalculateModeValues.AutoNoTable : CalculateModeValues.Auto,
                FullCalculationOnLoad = source.FullCalculationOnLoad,
                ReferenceMode = source.UsesA1References ? ReferenceModeValues.A1 : ReferenceModeValues.R1C1,
                Iterate = source.IterationEnabled,
                IterateCount = source.IterationCount,
                IterateDelta = source.IterationDelta,
                FullPrecision = source.FullPrecision,
                CalculationCompleted = source.CalculationCompleted,
                CalculationOnSave = source.CalculationOnSave,
                ConcurrentCalculation = source.ConcurrentCalculation,
                ForceFullCalculation = source.ForceFullCalculation
            };
            if (source.HasManualConcurrentCount) {
                properties.ConcurrentManualCount = checked((uint)source.ConcurrentThreadCount);
            }
            return properties;
        }

        private static bool MatchesConcurrentCount(uint? actual, XlsbCalculationProperties expected) =>
            expected.HasManualConcurrentCount
                ? actual.HasValue && actual.Value == checked((uint)expected.ConcurrentThreadCount)
                : !actual.HasValue;

        private static uint ToCalculationMode(CalculateModeValues mode) =>
            mode == CalculateModeValues.Manual ? 0U : mode == CalculateModeValues.AutoNoTable ? 2U : 1U;

        private static bool IsSupportedAttribute(string localName) =>
            localName == "calcId"
            || localName == "calcMode"
            || localName == "fullCalcOnLoad"
            || localName == "refMode"
            || localName == "iterate"
            || localName == "iterateCount"
            || localName == "iterateDelta"
            || localName == "fullPrecision"
            || localName == "calcCompleted"
            || localName == "calcOnSave"
            || localName == "concurrentCalc"
            || localName == "concurrentManualCount"
            || localName == "forceFullCalc";
    }
}
