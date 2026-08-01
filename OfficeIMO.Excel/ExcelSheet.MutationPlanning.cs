using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private MutationPlanScanBudget CreateMutationPlanScanBudget(ExcelMutationPlanOptions options) {
            EnsureMutationPlanPartsFitWithinBudget(options);
            return new MutationPlanScanBudget(
                options.MaximumScannedElements,
                options.MaximumScannedCharacters);
        }

        private static IEnumerable<T> InspectMutationPlanElements<T>(
            IEnumerable<T> elements,
            MutationPlanScanBudget? budget) {
            foreach (T element in elements) {
                budget?.Consume();
                yield return element;
            }
        }

        private void ValidateA1MutationReferenceMode(string operation) {
            if (WorkbookRoot.GetFirstChild<CalculationProperties>()?.ReferenceMode?.Value == ReferenceModeValues.R1C1) {
                throw new InvalidOperationException(
                    $"{operation} are not supported while the workbook uses R1C1 reference mode. Switch to A1 reference mode first.");
            }
        }

        private void ValidatePackageMutationReferenceSafety(
            string operation,
            Action? consumeScannedElement = null,
            ExcelReference? rewriteBoundary = null,
            ExcelCellShiftDirection? cellShiftDirection = null,
            Func<ExcelReference, ExcelReference?>? capacityTransform = null) {
            ValidateA1MutationReferenceMode(operation);
            _excelDocument.ValidateMutationReferencesCanBeRewritten(
                this,
                operation,
                consumeScannedElement,
                rewriteBoundary,
                cellShiftDirection,
                capacityTransform);
        }
    }
}
