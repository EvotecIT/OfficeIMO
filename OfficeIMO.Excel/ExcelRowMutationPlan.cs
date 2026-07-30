using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>Kind of structural row mutation represented by an <see cref="ExcelRowMutationPlan"/>.</summary>
    public enum ExcelRowMutationKind {
        /// <summary>Insert blank rows before the first affected row.</summary>
        Insert,

        /// <summary>Delete rows starting at the first affected row.</summary>
        Delete
    }

    /// <summary>Controls resource use while an Excel structural mutation plan is inspected.</summary>
    public sealed class ExcelMutationPlanOptions {
        /// <summary>
        /// Maximum Open XML elements inspected while building a dry-run impact summary.
        /// The bounded scan completes before semantic mutation preflight is allowed to run.
        /// </summary>
        public int MaximumScannedElements { get; set; } = 250_000;

        internal ExcelMutationPlanOptions CloneAndValidate() {
            if (MaximumScannedElements < 1) {
                throw new ArgumentOutOfRangeException(
                    nameof(MaximumScannedElements),
                    "The mutation-plan element budget must be positive.");
            }

            return new ExcelMutationPlanOptions {
                MaximumScannedElements = MaximumScannedElements
            };
        }
    }

    /// <summary>One bounded impact category discovered by a structural mutation dry run.</summary>
    public sealed class ExcelMutationImpact {
        internal ExcelMutationImpact(string category, int itemCount, string description) {
            Category = category;
            ItemCount = itemCount;
            Description = description;
        }

        /// <summary>Stable machine-readable impact category.</summary>
        public string Category { get; }

        /// <summary>Number of potentially affected items in this category.</summary>
        public int ItemCount { get; }

        /// <summary>Human-readable explanation of the affected structures.</summary>
        public string Description { get; }
    }

    /// <summary>
    /// Validated, non-mutating impact plan for inserting or deleting worksheet rows.
    /// </summary>
    /// <remarks>
    /// Applying a plan re-runs every safety check against the current workbook state. A plan therefore
    /// cannot be used to bypass a new array-formula, PivotTable, control, mapping, or capacity conflict.
    /// Planning rejects pending deferred worksheet writes rather than materializing them as a side effect.
    /// Each plan permits one application attempt. A failed attempt consumes the plan because mutation may
    /// already have started before a later validation, package, or save failure is observed.
    /// </remarks>
    public sealed class ExcelRowMutationPlan {
        private readonly ExcelSheet _owner;
        private int _applyState;

        internal ExcelRowMutationPlan(
            ExcelSheet owner,
            ExcelRowMutationKind kind,
            string sheetName,
            int firstRow,
            int count,
            int scannedElements,
            IReadOnlyList<ExcelMutationImpact> impacts) {
            _owner = owner;
            Kind = kind;
            SheetName = sheetName;
            FirstRow = firstRow;
            Count = count;
            ScannedElements = scannedElements;
            Impacts = new ReadOnlyCollection<ExcelMutationImpact>(
                new List<ExcelMutationImpact>(impacts));
        }

        /// <summary>Planned operation.</summary>
        public ExcelRowMutationKind Kind { get; }

        /// <summary>Worksheet name captured by the plan.</summary>
        public string SheetName { get; }

        /// <summary>First affected 1-based worksheet row.</summary>
        public int FirstRow { get; }

        /// <summary>Number of rows inserted or deleted.</summary>
        public int Count { get; }

        /// <summary>Workbook elements inspected while producing the bounded impact summary.</summary>
        public int ScannedElements { get; }

        /// <summary>Potentially affected workbook structures grouped by stable category.</summary>
        public IReadOnlyList<ExcelMutationImpact> Impacts { get; }

        /// <summary>Whether applying the mutation requests a complete workbook recalculation.</summary>
        public bool RequiresFullRecalculation => true;

        /// <summary>Whether this plan has already been applied successfully.</summary>
        public bool IsApplied => Volatile.Read(ref _applyState) == 2;

        /// <summary>Whether an application attempt has started, succeeded, or failed.</summary>
        public bool IsConsumed => Volatile.Read(ref _applyState) != 0;

        /// <summary>
        /// Revalidates and applies the planned operation exactly once.
        /// </summary>
        public void Apply() {
            if (Interlocked.CompareExchange(ref _applyState, 1, 0) != 0) {
                throw new InvalidOperationException(
                    "This Excel mutation plan is already being applied, was applied successfully, or a previous attempt failed.");
            }

            try {
                _owner.ApplyStructuralRowMutationPlan(Kind, FirstRow, Count);
                Volatile.Write(ref _applyState, 2);
            } catch {
                Volatile.Write(ref _applyState, 3);
                throw;
            }
        }
    }
}
