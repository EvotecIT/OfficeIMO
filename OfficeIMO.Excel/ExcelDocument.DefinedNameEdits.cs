namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const int PendingDefinedNameFormulaMarks = 1;
        private const int PendingWorkbookRootSave = 2;
        private int _pendingDefinedNameEdits;
        private int _definedNameEditThreadId;
        private readonly object _definedNameEditGate = new object();

        /// <summary>
        /// Defined-name edits change workbook.xml, and each one marks every formula cell for recalculation and may
        /// save the workbook part. Consecutive edits repeat that work for nothing, so the marks and the save wait
        /// until something else reaches the package. Every such route goes through <see cref="_spreadSheetDocument"/>,
        /// the workbook part field or the part fields of <see cref="ExcelSheet"/>, as for pending row deletes, and
        /// applies them first; only the thread inside a defined-name edit passes without applying them.
        /// </summary>
        internal readonly struct DefinedNameEditScope : IDisposable {
            private readonly ExcelDocument? _document;

            internal DefinedNameEditScope(ExcelDocument document) => _document = document;

            public void Dispose() {
                if (_document != null) Volatile.Write(ref _document._definedNameEditThreadId, 0);
            }
        }

        private DefinedNameEditScope BeginDefinedNameEdit() {
            int current = Environment.CurrentManagedThreadId;
            if (Volatile.Read(ref _definedNameEditThreadId) == current) return default;
            Volatile.Write(ref _definedNameEditThreadId, current);
            return new DefinedNameEditScope(this);
        }

        /// <summary>Applies deferred package edits before anything reads or changes the package.</summary>
        internal void FlushPendingPackageEditsIfAny() {
            if (_pendingRowDeletes != null) FlushPendingRowDeletes();
            int pending = Volatile.Read(ref _pendingDefinedNameEdits);
            if (pending != 0 && (pending != PendingDefinedNameFormulaMarks || !AreDefinedNameFormulaMarksDeferredOnThisThread)) {
                FlushPendingDefinedNameEdits();
            }
        }

        private int _formulaMarkDeferralThreadId;

        private bool AreDefinedNameFormulaMarksDeferredOnThisThread
            => Volatile.Read(ref _formulaMarkDeferralThreadId) == Environment.CurrentManagedThreadId;

        /// <summary>
        /// Lets a literal cell value write pass while defined-name edits still owe their formula marks. Such a write adds
        /// no formula and reads no recalculation flag, so marking every formula cell before it only repeats the walk for the
        /// next name; the marks are applied by the next access that is not such a write, at the latest by the save.
        /// </summary>
        internal readonly struct DefinedNameFormulaMarkDeferral : IDisposable {
            private readonly ExcelDocument? _document;

            internal DefinedNameFormulaMarkDeferral(ExcelDocument document) => _document = document;

            public void Dispose() {
                if (_document != null) Volatile.Write(ref _document._formulaMarkDeferralThreadId, 0);
            }
        }

        internal bool HasPendingDefinedNameFormulaMarks
            => (Volatile.Read(ref _pendingDefinedNameEdits) & PendingDefinedNameFormulaMarks) != 0;

        internal DefinedNameFormulaMarkDeferral DeferDefinedNameFormulaMarks() {
            if (!HasPendingDefinedNameFormulaMarks) return default;
            int current = Environment.CurrentManagedThreadId;
            return Interlocked.CompareExchange(ref _formulaMarkDeferralThreadId, current, 0) == 0
                ? new DefinedNameFormulaMarkDeferral(this)
                : default;
        }

        private void QueueDefinedNameEdit(int edit) {
            int current;
            do {
                current = Volatile.Read(ref _pendingDefinedNameEdits);
                if ((current & edit) == edit) return;
            } while (Interlocked.CompareExchange(ref _pendingDefinedNameEdits, current | edit, current) != current);
        }

        private void FlushPendingDefinedNameEdits() {
            if (Volatile.Read(ref _definedNameEditThreadId) == Environment.CurrentManagedThreadId) return;
            lock (_definedNameEditGate) {
                // Taken first: applying the edits reads the package through the same guarded fields.
                int edits = Interlocked.Exchange(ref _pendingDefinedNameEdits, 0);
                try {
                    if ((edits & PendingDefinedNameFormulaMarks) != 0 && !AreDefinedNameFormulaMarksDeferredOnThisThread) {
                        foreach (ExcelSheet sheet in GetSheetsForLockedOperation()) sheet.MarkFormulasForRecalculation();
                        edits &= ~PendingDefinedNameFormulaMarks;
                    }
                    if ((edits & PendingWorkbookRootSave) != 0) {
                        WorkbookRoot.Save();
                        edits &= ~PendingWorkbookRootSave;
                    }
                } finally {
                    if (edits != 0) QueueDefinedNameEdit(edits);
                }
            }
        }
    }
}
