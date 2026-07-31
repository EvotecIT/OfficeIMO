using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>Direction used when inserting or deleting a rectangular block of cells.</summary>
    public enum ExcelCellShiftDirection {
        /// <summary>Shift cells to the right.</summary>
        Right,
        /// <summary>Shift cells down.</summary>
        Down,
        /// <summary>Shift cells left.</summary>
        Left,
        /// <summary>Shift cells up.</summary>
        Up
    }

    /// <summary>Supported transactional worksheet mutation.</summary>
    public enum ExcelStructuralMutationKind {
        /// <summary>Insert complete columns.</summary>
        InsertColumns,
        /// <summary>Delete complete columns.</summary>
        DeleteColumns,
        /// <summary>Insert cells and shift existing cells right.</summary>
        InsertCellsRight,
        /// <summary>Insert cells and shift existing cells down.</summary>
        InsertCellsDown,
        /// <summary>Delete cells and shift remaining cells left.</summary>
        DeleteCellsLeft,
        /// <summary>Delete cells and shift remaining cells up.</summary>
        DeleteCellsUp,
        /// <summary>Copy a cell range.</summary>
        Copy,
        /// <summary>Move a cell range.</summary>
        Move,
        /// <summary>Copy and transpose a cell range.</summary>
        Transpose
    }

    /// <summary>Severity of a post-edit package diagnostic.</summary>
    public enum ExcelMutationDiagnosticSeverity {
        /// <summary>Informational observation.</summary>
        Information,
        /// <summary>Package content that may need attention.</summary>
        Warning,
        /// <summary>Open XML validation error.</summary>
        Error
    }

    /// <summary>One post-edit package diagnostic.</summary>
    public sealed class ExcelMutationDiagnostic {
        internal ExcelMutationDiagnostic(string code, ExcelMutationDiagnosticSeverity severity, string message, string? partUri) {
            Code = code;
            Severity = severity;
            Message = message;
            PartUri = partUri;
        }
        /// <summary>Stable diagnostic code.</summary>
        public string Code { get; }
        /// <summary>Diagnostic severity.</summary>
        public ExcelMutationDiagnosticSeverity Severity { get; }
        /// <summary>Diagnostic message.</summary>
        public string Message { get; }
        /// <summary>Package part URI, when available.</summary>
        public string? PartUri { get; }
    }

    /// <summary>Result of one committed structural mutation.</summary>
    public sealed class ExcelMutationResult {
        internal ExcelMutationResult(int affectedCells, IReadOnlyList<ExcelMutationDiagnostic> diagnostics) {
            AffectedCells = affectedCells;
            Diagnostics = new ReadOnlyCollection<ExcelMutationDiagnostic>(new List<ExcelMutationDiagnostic>(diagnostics));
        }
        /// <summary>Number of cells moved, copied, inserted, or removed.</summary>
        public int AffectedCells { get; }
        /// <summary>Bounded Open XML diagnostics captured after the edit.</summary>
        public IReadOnlyList<ExcelMutationDiagnostic> Diagnostics { get; }
        /// <summary>Whether no error diagnostics were produced.</summary>
        public bool PackageIsValid => !System.Linq.Enumerable.Any(Diagnostics, item => item.Severity == ExcelMutationDiagnosticSeverity.Error);
    }

    /// <summary>Validated dry-run plan for a single transactional worksheet mutation.</summary>
    public sealed class ExcelStructuralMutationPlan {
        private readonly ExcelSheet _owner;
        private readonly Action<CancellationToken> _apply;
        private readonly ExcelMutationPlanOptions _options;
        private int _state;

        internal ExcelStructuralMutationPlan(
            ExcelSheet owner,
            ExcelStructuralMutationKind kind,
            string sourceRange,
            string? destination,
            int affectedCells,
            IReadOnlyList<ExcelMutationImpact> impacts,
            ExcelMutationPlanOptions options,
            Action<CancellationToken> apply) {
            _owner = owner;
            Kind = kind;
            SourceRange = sourceRange;
            Destination = destination;
            AffectedCells = affectedCells;
            Impacts = new ReadOnlyCollection<ExcelMutationImpact>(new List<ExcelMutationImpact>(impacts));
            _options = options;
            _apply = apply;
        }

        /// <summary>Planned operation.</summary>
        public ExcelStructuralMutationKind Kind { get; }
        /// <summary>Source or affected A1 range.</summary>
        public string SourceRange { get; }
        /// <summary>Destination top-left cell, when applicable.</summary>
        public string? Destination { get; }
        /// <summary>Bounded estimate of cells affected by the operation.</summary>
        public int AffectedCells { get; }
        /// <summary>Dry-run impact categories.</summary>
        public IReadOnlyList<ExcelMutationImpact> Impacts { get; }
        /// <summary>Whether the plan has committed successfully.</summary>
        public bool IsApplied => Volatile.Read(ref _state) == 2;
        /// <summary>Result after a successful commit.</summary>
        public ExcelMutationResult? Result { get; private set; }

        /// <summary>Commits the plan once, rolling package roots back if the operation throws.</summary>
        public ExcelMutationResult Apply(CancellationToken cancellationToken = default) {
            if (Interlocked.CompareExchange(ref _state, 1, 0) != 0) {
                throw new InvalidOperationException("This Excel mutation plan has already been consumed.");
            }
            try {
                Result = _owner.ApplyTransactionalMutation(_apply, AffectedCells, _options, cancellationToken);
                Volatile.Write(ref _state, 2);
                return Result;
            } catch {
                Volatile.Write(ref _state, 3);
                throw;
            }
        }
    }
}
