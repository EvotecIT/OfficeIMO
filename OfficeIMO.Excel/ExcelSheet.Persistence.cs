using DocumentFormat.OpenXml.Packaging;
using System.Runtime.CompilerServices;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Sheet wrappers can be recreated when caching is disabled or a sheet is added.
        // Pending serialization belongs to the package part and must outlive any one wrapper.
        private static readonly ConditionalWeakTable<WorksheetPart, WorksheetPersistenceState> WorksheetPersistenceStates = new();
        private WorksheetPersistenceState? _worksheetPersistence;
        private WorksheetPersistenceState Persistence => _worksheetPersistence ??=
            WorksheetPersistenceStates.GetValue(_worksheetPart, static _ => new WorksheetPersistenceState());
        private HashSet<OpenXmlPart> PendingPartSaves => Persistence.PartSaves;
        private Dictionary<VmlDrawingPart, XDocument> PendingCommentVmlDocuments => Persistence.CommentVmlDocuments;
        private bool _requiresSavePreparation {
            get => Persistence.RequiresSavePreparation;
            set => Persistence.RequiresSavePreparation = value;
        }

        private sealed class WorksheetPersistenceState {
            internal readonly WorksheetReadCache ReadCaches = new();
            internal readonly HashSet<OpenXmlPart> PartSaves = new();
            internal readonly Dictionary<VmlDrawingPart, XDocument> CommentVmlDocuments = new();
            internal bool RequiresSavePreparation;
            /// <summary>Print settings on this sheet were chosen by the table defaults, not by the caller.</summary>
            internal bool OwnsTablePrintDefaults;
        }

        /// <summary>
        /// Persists pending changes on this worksheet to its underlying OpenXml part.
        /// </summary>
        internal void Commit() {
            CommitPendingPartChanges();
            _worksheetPart.Worksheet?.Save();
            _requiresSavePreparation = false;
        }

        private void CommitPendingPartChanges() {
            foreach (OpenXmlPart part in PendingPartSaves) {
                // A cleared comment or removed table can detach a queued part before Save.
                if (_worksheetPart.Parts.Any(pair => ReferenceEquals(pair.OpenXmlPart, part))) part.RootElement?.Save();
            }
            PendingPartSaves.Clear();
            CommitCommentVmlChanges();
        }

        /// <summary>Flushes deferred payloads before a transaction captures the package streams for rollback.</summary>
        private void FlushPendingPartsBeforeMutationSnapshot() {
            foreach (ExcelSheet sheet in _excelDocument.GetSheetsForLockedOperation()) sheet.CommitPendingPartChanges();
        }

        private void ResetPendingPartsAfterMutationRollback() {
            foreach (WorksheetPart part in _excelDocument.WorkbookPartRoot.WorksheetParts) {
                if (!WorksheetPersistenceStates.TryGetValue(part, out WorksheetPersistenceState? state)) continue;
                state.PartSaves.Clear();
                state.CommentVmlDocuments.Clear();
            }
        }

        internal bool RequiresSavePreparation => _requiresSavePreparation;

        internal void MarkRequiresSavePreparation() {
            if (_requiresSavePreparation && _excelDocument.IsPackageDirty) {
                return;
            }

            _requiresSavePreparation = true;
            _excelDocument.MarkRequiresSavePreflight();
        }

        internal void MarkRequiresWorksheetPreparation() {
            if (_requiresSavePreparation && _excelDocument.IsPackageDirty) {
                return;
            }

            _requiresSavePreparation = true;
            _excelDocument.MarkPackageDirty();
        }

        /// <summary>Queues a worksheet-owned XML part for the next workbook save.</summary>
        internal void DeferPartSave(OpenXmlPart part) {
            PendingPartSaves.Add(part);
            MarkRequiresSavePreparation();
        }
    }
}
