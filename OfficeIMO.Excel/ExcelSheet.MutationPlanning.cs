using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
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

        /// <summary>Rejects a stale plan whose captured worksheet relationship is no longer active in the workbook.</summary>
        internal void EnsureWorksheetCapturedByMutationPlanIsActive() {
            WorkbookPart workbookPart = _excelDocument.WorkbookPartRoot;
            if (!workbookPart.Parts.Any(pair => ReferenceEquals(pair.OpenXmlPart, _worksheetPart))) {
                throw new InvalidOperationException(
                    "The worksheet captured by this Excel mutation plan is no longer part of the workbook.");
            }
            string relationshipId = workbookPart.GetIdOfPart(_worksheetPart);
            bool relationshipIsActive = WorkbookRoot.Sheets?
                .Elements<Sheet>()
                .Any(sheet => string.Equals(
                    sheet.Id?.Value,
                    relationshipId,
                    StringComparison.Ordinal)) == true;
            if (!relationshipIsActive) {
                throw new InvalidOperationException(
                    "The worksheet captured by this Excel mutation plan is no longer part of the workbook.");
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
