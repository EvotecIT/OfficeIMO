#if NET8_0_OR_GREATER
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Serializes and compresses the buffered worksheets of an extended save on worker threads, a bounded number
        /// ahead of the package writer, which appends each one at its usual position. Entry order, CRCs and the
        /// inflated bytes match a sequential save.
        /// </summary>
        private sealed class PreparedDirectWorksheets : IDisposable {
            // Larger sheets keep the streaming entry, whose memory does not grow with the part.
            private const long MaximumPreparedCells = 4_000_000;

            private readonly ExcelZipWriter _archive;
            private readonly DirectDataSetWorkbookWriter.ExtendedDirectWritePlan _plan;
            private readonly List<(DirectDataSetSheetModel Sheet, string? TableRelationshipId)> _sheets;
            private readonly Dictionary<WorksheetPart, int> _ordinals;
            private readonly Task<ExcelZipWriter.PreparedEntry>?[] _tasks;
            private readonly CancellationTokenSource _cancellation;
            private readonly int _window;
            private int _started;

            private PreparedDirectWorksheets(
                ExcelZipWriter archive,
                DirectDataSetWorkbookWriter.ExtendedDirectWritePlan plan,
                List<(DirectDataSetSheetModel Sheet, string? TableRelationshipId)> sheets,
                Dictionary<WorksheetPart, int> ordinals,
                int window,
                CancellationToken ct) {
                _archive = archive;
                _plan = plan;
                _sheets = sheets;
                _ordinals = ordinals;
                _window = window;
                _tasks = new Task<ExcelZipWriter.PreparedEntry>?[sheets.Count];
                _cancellation = CancellationTokenSource.CreateLinkedTokenSource(ct);
                while (_started < Math.Min(_window, _sheets.Count)) Start();
            }

            /// <summary>
            /// Starts preparing when at least two buffered worksheets qualify. Sheets backed by caller objects stay
            /// sequential so caller code never runs on a worker thread.
            /// </summary>
            internal static PreparedDirectWorksheets? TryCreate(
                ExcelPackageArchive archive,
                ExtendedWorkbookPackageModel model,
                DirectDataSetWorkbookWriter.ExtendedDirectWritePlan? plan,
                ExcelExecutionPolicy? execution,
                CancellationToken ct) {
                if (archive is not ExcelZipWriter zip || plan == null || execution?.AllowsBackgroundProcessing != true) {
                    return null;
                }

                int window = Math.Min(execution.MaxDegreeOfParallelism ?? Environment.ProcessorCount, Environment.ProcessorCount);
                if (window < 2) {
                    return null;
                }

                var sheets = new List<(DirectDataSetSheetModel, string?)>();
                var ordinals = new Dictionary<WorksheetPart, int>();
                foreach (var part in model.Parts) {
                    if (part.Part is not WorksheetPart worksheetPart
                        || !model.DirectWorksheetModels.TryGetValue(worksheetPart, out DirectDataSetSheetModel? sheet)
                        || sheet.Table.CellValueBuffer is not { } buffer
                        || buffer.CellCount > MaximumPreparedCells) {
                        continue;
                    }

                    string? tableRelationshipId = null;
                    if (sheet.HasTable && worksheetPart.TableDefinitionParts.FirstOrDefault() is { } tablePart) {
                        tableRelationshipId = worksheetPart.GetIdOfPart(tablePart);
                    }

                    ordinals.Add(worksheetPart, sheets.Count);
                    sheets.Add((sheet, tableRelationshipId));
                }

                return sheets.Count < 2 ? null : new PreparedDirectWorksheets(zip, plan, sheets, ordinals, window, ct);
            }

            /// <summary>Appends the prepared entry for <paramref name="part"/>, or returns false when it was not prepared.</summary>
            internal bool TryWrite(WorksheetPart part, string path) {
                if (!_ordinals.TryGetValue(part, out int ordinal) || _tasks[ordinal] is not { } task) {
                    return false;
                }

                ExcelZipWriter.PreparedEntry entry = task.GetAwaiter().GetResult();
                _tasks[ordinal] = null;
                _archive.WritePrepared(path, entry);
                if (_started < _sheets.Count) Start();
                return true;
            }

            private void Start() {
                int ordinal = _started++;
                var (sheet, tableRelationshipId) = _sheets[ordinal];
                CancellationToken token = _cancellation.Token;
                _tasks[ordinal] = Task.Run(() => ExcelZipWriter.Prepare(output =>
                    DirectDataSetWorkbookWriter.WriteExtendedWorksheetContents(output, _plan, sheet, tableRelationshipId, token)), token);
            }

            /// <summary>Stops sheets that were not written, as after a failed save, and waits for them to finish.</summary>
            public void Dispose() {
                _cancellation.Cancel();
                foreach (var task in _tasks) {
                    if (task == null) continue;
                    try {
                        task.Wait();
                    } catch (AggregateException) {
                        // The save already failed or completed; an abandoned sheet's error has nothing to report to.
                    }
                }
                _cancellation.Dispose();
            }
        }
    }
}
#endif
