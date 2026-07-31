using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
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
        /// Maximum Open XML parts and elements admitted for one dry-run inspection phase.
        /// Unloaded package XML is bounded before materialization, and the impact scan applies
        /// the same limit independently before semantic mutation preflight is allowed to run.
        /// </summary>
        public int MaximumScannedElements { get; set; } = 250_000;

        /// <summary>
        /// Maximum XML characters admitted per lazily loaded part before semantic inspection. The same value is
        /// also applied as a conservative aggregate decompressed-byte ceiling across all lazily loaded XML parts,
        /// before any Open XML DOM is materialized.
        /// </summary>
        public long MaximumScannedCharacters { get; set; } = 64_000_000;

        /// <summary>Maximum cells that a planned grid operation may inspect or move.</summary>
        public int MaximumAffectedCells { get; set; } = 1_000_000;

        /// <summary>Maximum XML characters retained for transactional rollback.</summary>
        public long MaximumSnapshotCharacters { get; set; } = 128_000_000;

        /// <summary>Maximum post-edit package diagnostics returned to the caller.</summary>
        public int MaximumDiagnostics { get; set; } = 100;

        internal ExcelMutationPlanOptions CloneAndValidate() {
            if (MaximumScannedElements < 1) {
                throw new ArgumentOutOfRangeException(
                    nameof(MaximumScannedElements),
                    "The mutation-plan element budget must be positive.");
            }
            if (MaximumScannedCharacters < 1) {
                throw new ArgumentOutOfRangeException(
                    nameof(MaximumScannedCharacters),
                    "The mutation-plan character budget must be positive.");
            }
            if (MaximumAffectedCells < 1) throw new ArgumentOutOfRangeException(nameof(MaximumAffectedCells));
            if (MaximumSnapshotCharacters < 1) throw new ArgumentOutOfRangeException(nameof(MaximumSnapshotCharacters));
            if (MaximumDiagnostics < 1) throw new ArgumentOutOfRangeException(nameof(MaximumDiagnostics));

            return new ExcelMutationPlanOptions {
                MaximumScannedElements = MaximumScannedElements,
                MaximumScannedCharacters = MaximumScannedCharacters,
                MaximumAffectedCells = MaximumAffectedCells,
                MaximumSnapshotCharacters = MaximumSnapshotCharacters,
                MaximumDiagnostics = MaximumDiagnostics
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
        private readonly ExcelMutationPlanOptions _options;
        private int _applyState;

        internal ExcelRowMutationPlan(
            ExcelSheet owner,
            ExcelRowMutationKind kind,
            string sheetName,
            int firstRow,
            int count,
            int scannedElements,
            IReadOnlyList<ExcelMutationImpact> impacts,
            ExcelMutationPlanOptions options) {
            _owner = owner;
            Kind = kind;
            SheetName = sheetName;
            FirstRow = firstRow;
            Count = count;
            ScannedElements = scannedElements;
            Impacts = new ReadOnlyCollection<ExcelMutationImpact>(
                new List<ExcelMutationImpact>(impacts));
            _options = options;
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

        /// <summary>Post-edit result after a successful application.</summary>
        public ExcelMutationResult? Result { get; private set; }

        /// <summary>
        /// Revalidates and applies the planned operation exactly once.
        /// </summary>
        public void Apply() {
            ApplyWithDiagnostics();
        }

        /// <summary>Revalidates and transactionally applies the plan, returning package diagnostics.</summary>
        public ExcelMutationResult ApplyWithDiagnostics(CancellationToken cancellationToken = default) {
            if (Interlocked.CompareExchange(ref _applyState, 1, 0) != 0) {
                throw new InvalidOperationException(
                    "This Excel mutation plan is already being applied, was applied successfully, or a previous attempt failed.");
            }

            try {
                int affectedCells = Impacts.FirstOrDefault(item => item.Category == "cells")?.ItemCount ?? 0;
                Result = _owner.ApplyTransactionalMutation(
                    _ => _owner.ApplyStructuralRowMutationPlan(Kind, FirstRow, Count),
                    affectedCells,
                    _options,
                    cancellationToken);
                Volatile.Write(ref _applyState, 2);
                return Result;
            } catch {
                Volatile.Write(ref _applyState, 3);
                throw;
            }
        }
    }
}
