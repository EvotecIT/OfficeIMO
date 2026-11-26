using System.Globalization;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Saving a workbook whose SDK package was never opened writes the loaded package again through OfficeIMO's own
    /// Open XML layer: parts nothing changed are copied as stored, edited worksheets are merged with their new values,
    /// the shared strings and styles gain what the values need, and docProps/core.xml gets the modification time a
    /// save through the model gives it. Anything this path does not cover opens the SDK and saves as before.
    /// </summary>
    /// <remarks>
    /// The model's save can also run a whole-workbook preflight that repairs damaged references and adds default views
    /// to every worksheet, when workbook state asks for it or <see cref="ExcelSaveOptions.SafePreflight"/> is set. This
    /// path leaves parts it did not edit as they were read, as Excel does; <see cref="ExcelSaveOptions.SafePreflight"/>
    /// and the other repair options send the save through the model.
    /// </remarks>
    public partial class ExcelDocument {
        private const string SharedStringsRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "sharedStrings";
        private const string StylesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "styles";
        private const string TableRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "table";

        private readonly object _deferredEditsGate = new();
        // Sheets with recorded changes, in the order their first change was made.
        private readonly List<ExcelSheet> _deferredEditedSheets = new();
        private const string CellMetadataRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "sheetMetadata";
        private bool? _deferredHasCellMetadata;

        /// <summary>Records a plain value for the owned save; false when the sheet needs the model for it.</summary>
        internal bool RecordDeferredCellEdit(ExcelSheet sheet, in ExcelDeferredCellEdit edit) {
            ExcelDeferredSheet? entry = sheet.DeferredEntry;
            // The checks read the bytes taken here; if the package opens meanwhile, the check under the gate below sends
            // the change through the model.
            byte[]? packageBytes = _deferredPackageBytes;
            ExcelDeferredWorkbook? workbook = _deferredWorkbook;
            // A value written to a table's header row renames its column, which the model does; styles are plain.
            if (entry == null || workbook == null || packageBytes == null) return false;
            if (edit.Kind != ExcelDeferredCellKind.Style) {
                if (IsTableHeaderCell(entry, packageBytes, edit.Row, edit.Column)) {
                    OpenDeferredSdk("Cell value in a table header row on worksheet '" + entry.Name + "'");
                    return false;
                }
                if (WorkbookHasCellMetadata(workbook, packageBytes)) {
                    OpenDeferredSdk("Cell value in a workbook with cell metadata (dynamic arrays or rich values)");
                    return false;
                }
            }
            lock (_deferredEditsGate) {
                // The package may have opened since the checks above; the change then goes through the model.
                if (_deferredWorkbook == null || !ReferenceEquals(sheet.DeferredEntry, entry)) return false;
                ExcelDeferredCellStore? store = entry.Cells;
                if (store == null) {
                    entry.Cells = store = new ExcelDeferredCellStore();
                    _deferredEditedSheets.Add(sheet);
                }
                store.Add(edit);
            }
            if (!_packageDirty) MarkPackageDirty();
            MarkFormulaInputMutation();
            return true;
        }

        // A value written to a table's header row renames its column, which the model does; the header rows of the
        // sheet's tables are read once.
        private static bool IsTableHeaderCell(ExcelDeferredSheet entry, byte[] packageBytes, int row, int column) {
            if (!entry.TableHeadersRead) {
                entry.TableHeaders = ReadTableHeaders(entry, packageBytes);
                entry.TableHeadersRead = true;
            }
            List<(int Row, int First, int Last)>? headers = entry.TableHeaders;
            // A table the owned reader cannot read counts as covering every cell.
            if (headers == null) return true;
            for (int index = 0; index < headers.Count; index++) {
                (int headerRow, int first, int last) = headers[index];
                if (row == headerRow && column >= first && column <= last) return true;
            }
            return false;
        }

        private static List<(int Row, int First, int Last)>? ReadTableHeaders(ExcelDeferredSheet entry, byte[] packageBytes) {
            var headers = new List<(int Row, int First, int Last)>();
            try {
                using var source = new MemoryStream(packageBytes, writable: false);
                using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
                foreach (OfficeOpcRelationship relationship in package.GetRelationships(entry.PartName)) {
                    if (!string.Equals(relationship.Type, TableRelationshipType, StringComparison.Ordinal)) continue;
                    if (relationship.IsExternal || relationship.TargetPartName is not { } target || !package.PartExists(target)) return null;
                    OfficeXmlPullReader reader = package.OpenXml(target, out OfficePooledBytes content, processMarkupCompatibility: false);
                    using (content) {
                        while (reader.Read()) {
                            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                            if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "table")) return null;
                            if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "headerRowCount", out string count) && count == "0") break;
                            if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "ref", out string reference)
                                || !ExcelA1.TryParseRange(reference, out int top, out int left, out _, out int right)) {
                                return null;
                            }
                            headers.Add((top, left, right));
                            break;
                        }
                    }
                }
                return headers;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException) {
                return null;
            }
        }

        // Dynamic arrays and rich values live in the cell metadata part, which the model reads before a value is written.
        private bool WorkbookHasCellMetadata(ExcelDeferredWorkbook workbook, byte[] packageBytes) {
            if (_deferredHasCellMetadata is { } known) return known;
            bool found = true;
            try {
                using var source = new MemoryStream(packageBytes, writable: false);
                using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
                found = package.GetRelatedPartName(workbook.WorkbookPartName, CellMetadataRelationshipType) != null;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException) {
                found = true;
            }
            _deferredHasCellMetadata = found;
            return found;
        }

        /// <summary>Writes the recorded values through the model once the SDK package is open, in their original order.</summary>
        private void ReplayDeferredCellEdits() {
            ExcelSheet[] sheets;
            lock (_deferredEditsGate) {
                sheets = _deferredEditedSheets.ToArray();
                _deferredEditedSheets.Clear();
            }
            // Every handle to a sheet is bound by now; the one that recorded first applies the sheet's changes. A sheet
            // whose changes fail does not stop the others from being applied; the first failure is raised after them.
            System.Runtime.ExceptionServices.ExceptionDispatchInfo? failure = null;
            foreach (ExcelSheet sheet in sheets) {
                ExcelDeferredCellStore? store = sheet.DeferredCellsForReplay;
                sheet.DeferredCellsForReplay = null;
                if (store == null) continue;
                try {
                    sheet.ReplayDeferredCells(store);
                } catch (Exception exception) {
                    failure ??= System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception);
                }
            }
            failure?.Throw();
        }

        /// <summary>Takes the recorded changes for a save; changes recorded while it runs start new stores.</summary>
        private (ExcelSheet Sheet, ExcelDeferredCellStore Store)[] TakeDeferredEdits() {
            lock (_deferredEditsGate) {
                var taken = new (ExcelSheet Sheet, ExcelDeferredCellStore Store)[_deferredEditedSheets.Count];
                for (int index = 0; index < taken.Length; index++) {
                    ExcelSheet sheet = _deferredEditedSheets[index];
                    ExcelDeferredSheet entry = sheet.DeferredEntry ?? throw new InvalidOperationException("A recorded worksheet is bound to the model.");
                    taken[index] = (sheet, entry.Cells!);
                    entry.Cells = null;
                }
                _deferredEditedSheets.Clear();
                return taken;
            }
        }

        /// <summary>
        /// Reads and edits through the model see recorded changes once the SDK has replayed them, so it opens before they
        /// look for unsaved cells or move buffered ones into the worksheet: the replay can leave values in the cell buffer.
        /// </summary>
        internal void OpenDeferredSdkBeforeReadingRecordedChanges() {
            if (_deferredWorkbook == null) return;
            bool recorded;
            lock (_deferredEditsGate) recorded = _deferredEditedSheets.Count > 0;
            if (recorded) OpenDeferredSdk("Read after recorded changes");
        }

        /// <summary>Puts back changes a save did not write, ahead of any recorded while it ran.</summary>
        private void RestoreDeferredEdits((ExcelSheet Sheet, ExcelDeferredCellStore Store)[] taken) {
            if (taken.Length == 0) return;
            List<(ExcelSheet Sheet, ExcelDeferredCellStore Store)>? bound = null;
            lock (_deferredEditsGate) {
                var later = new List<ExcelSheet>(_deferredEditedSheets);
                _deferredEditedSheets.Clear();
                foreach ((ExcelSheet sheet, ExcelDeferredCellStore store) in taken) {
                    if (sheet.DeferredEntry is not { } entry) {
                        // The package opened meanwhile; these changes go through the model below.
                        (bound ??= new()).Add((sheet, store));
                        continue;
                    }
                    if (entry.Cells is { } newer) {
                        for (int index = 0; index < newer.Count; index++) store.Add(newer[index]);
                    }
                    entry.Cells = store;
                    _deferredEditedSheets.Add(sheet);
                }
                foreach (ExcelSheet sheet in later) {
                    if (!_deferredEditedSheets.Contains(sheet)) _deferredEditedSheets.Add(sheet);
                }
            }
            if (bound != null) {
                foreach ((ExcelSheet sheet, ExcelDeferredCellStore store) in bound) sheet.ReplayDeferredCells(store);
            }
        }

        private bool CanSaveDeferredPackage(ExcelSaveOptions? options) =>
            _deferredWorkbook != null
            && !_packagePropertiesDirty
            && !_customDocumentPropertiesDirty
            && !HasCalculationSaveWork(options)
            && (options == null || !options.ValidateOpenXml && !options.SafeRepairDefinedNames && !options.SafePreflight
                && !options.DisableFastPackageWriter && options.Format == ExcelFileFormat.Xlsx);

        /// <summary>
        /// Builds the saved package and makes it the document's state before anything reaches the destination, so a
        /// destination that reads the document, edits it or fails leaves no taken changes to put back. Opening the SDK
        /// waits on the same gate, so the package stays deferred while the save builds it.
        /// </summary>
        private bool TryPrepareDeferredSave(ExcelSaveOptions? options, string? path, CancellationToken cancellationToken, out byte[] bytes, out int save) {
            byte[] built = Array.Empty<byte>();
            int completedSave = 0;
            string? declined = null;
            bool prepared = WithDeferredOpenGate(() => {
                if (!CanSaveDeferredPackage(options) || path != null && !DeferredPackageTypeMatches(path)) return false;
                cancellationToken.ThrowIfCancellationRequested();
                var taken = TakeDeferredEdits();
                bool completed = false;
                try {
                    // The package is built within the in-memory limit; past it the save throws before anything is written.
                    if (!TryCreateDeferredPackageBytes(taken, GetPackageMaterializationLimit(options), out built, out string? reason)) {
                        declined = reason ?? "the owned save declined";
                        return false;
                    }
                    completedSave = CompleteDeferredSave(built);
                    completed = true;
                    return true;
                } finally {
                    if (!completed) RestoreDeferredEdits(taken);
                }
            });
            bytes = built;
            save = completedSave;
            // The SDK opens outside the gate, taking the document lock first as every opener does.
            if (declined != null) OpenDeferredSdk("Save: " + declined);
            return prepared;
        }

        /// <summary>Saves through the owned path into a stream, or returns false having written nothing.</summary>
        private bool TryWriteDeferredPackageToStream(Stream destination, ExcelSaveOptions? options) {
            if (!TryPrepareDeferredSave(options, null, CancellationToken.None, out byte[] bytes, out int save)) return false;
            try {
                PrepareDestinationStreamForWrite(destination);
                destination.Write(bytes, 0, bytes.Length);
                try { destination.Flush(); } catch (NotSupportedException) { }
            } catch {
                // The document holds the saved state; the destination does not, so the next save writes it again.
                MarkPackageDirty();
                throw;
            }
            PublishDeferredSaveDiagnostics(save);
            return true;
        }

        /// <summary>Saves through the owned path into a stream with asynchronous writes, or returns false having written nothing.</summary>
        private async Task<bool> TryWriteDeferredPackageToStreamAsync(Stream destination, ExcelSaveOptions? options, CancellationToken cancellationToken) {
            if (!TryPrepareDeferredSave(options, null, cancellationToken, out byte[] bytes, out int save)) return false;
            try {
                cancellationToken.ThrowIfCancellationRequested();
                PrepareDestinationStreamForWrite(destination);
                await destination.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
                try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
            } catch {
                MarkPackageDirty();
                throw;
            }
            PublishDeferredSaveDiagnostics(save);
            return true;
        }

        /// <summary>Saves through the owned path into a file whose extension matches the package type.</summary>
        private bool TryWriteDeferredPackageToFile(string path, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (!TryPrepareDeferredSave(options, path, cancellationToken, out byte[] bytes, out int save)) return false;
            try {
                cancellationToken.ThrowIfCancellationRequested();
                CommitPreparedPackageToFile(path, bytes);
            } catch {
                MarkPackageDirty();
                throw;
            }
            FilePath = path;
            PublishDeferredSaveDiagnostics(save);
            return true;
        }

        // A save to another extension changes the package type, which the model does.
        private bool DeferredPackageTypeMatches(string path) {
            string extension = Path.GetExtension(path);
            string? type = ReadDeferredMainContentType();
            return type switch {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" => extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" => extension.Equals(".xltx", StringComparison.OrdinalIgnoreCase),
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml" => extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase),
                "application/vnd.ms-excel.template.macroEnabled.main+xml" => extension.Equals(".xltm", StringComparison.OrdinalIgnoreCase),
                _ => false
            };
        }

        private string? ReadDeferredMainContentType() {
            using var source = new MemoryStream(_deferredPackageBytes!, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            return package.GetContentType(_deferredWorkbook!.WorkbookPartName)?.ToLowerInvariant() switch {
                null => null,
                string lowered => WorkbookMainContentTypes.FirstOrDefault(type => type.Equals(lowered, StringComparison.OrdinalIgnoreCase))
            };
        }

        /// <summary>Reports the owned save once its destination holds the package, unless another save reported meanwhile.</summary>
        private void PublishDeferredSaveDiagnostics(int save) {
            if (Interlocked.CompareExchange(ref _saveDiagnosticsVersion, 0, 0) == save) LastSaveDiagnostics = ExcelSaveDiagnostics.OwnedPackage();
        }

        /// <summary>
        /// The document now holds what it saved: later reads and saves start from these bytes. Returns the diagnostics version to report against
        /// for <see cref="PublishDeferredSaveDiagnostics"/>.
        /// </summary>
        private int CompleteDeferredSave(byte[] bytes) {
            // The caller's stream or file has its own copy; the document keeps the array.
            _deferredPackageBytes = bytes;
            foreach (ExcelDeferredSheet entry in _deferredWorkbook?.Sheets ?? new List<ExcelDeferredSheet>()) entry.TableHeadersRead = false;
            _deferredHasCellMetadata = null;
            MarkPackageClean(bytes);
            bool pending;
            lock (_deferredEditsGate) pending = _deferredEditedSheets.Count > 0;
            // Changes recorded while the save ran are still to be saved.
            if (pending) MarkPackageDirty();
            ReleaseReadSnapshot();
            return Interlocked.CompareExchange(ref _saveDiagnosticsVersion, 0, 0);
        }

        /// <summary>Builds the saved package, or returns false when something needs the model.</summary>
        private bool TryCreateDeferredPackageBytes((ExcelSheet Sheet, ExcelDeferredCellStore Store)[] edits, long? limit, out byte[] bytes, out string? reason) {
            _deferredSaveDeclineReason = null;
            bool created = TryCreateDeferredPackageBytesCore(edits, limit, out bytes);
            reason = _deferredSaveDeclineReason;
            return created;
        }

        private string? _deferredSaveDeclineReason;

        private bool Decline(string reason) {
            _deferredSaveDeclineReason = reason;
            return false;
        }

        private bool TryCreateDeferredPackageBytesCore((ExcelSheet Sheet, ExcelDeferredCellStore Store)[] edits, long? limit, out byte[] bytes) {
            bytes = Array.Empty<byte>();
            ExcelDeferredWorkbook workbook = _deferredWorkbook!;

            try {
                string? sharedStringsPart, stylesPart, corePart;
                using (var source = new MemoryStream(_deferredPackageBytes!, writable: false))
                using (OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true)) {
                    sharedStringsPart = package.GetRelatedPartName(workbook.WorkbookPartName, SharedStringsRelationshipType);
                    stylesPart = package.GetRelatedPartName(workbook.WorkbookPartName, StylesRelationshipType);
                    corePart = package.GetRelatedPartName(null, OfficeOpcPackage.CorePropertiesRelationshipType);
                    if (sharedStringsPart != null && !package.PartExists(sharedStringsPart)) return Decline("the shared strings relationship has no part");
                    if (stylesPart != null && !package.PartExists(stylesPart)) return Decline("the styles relationship has no part");
                    // The model's save sets the modification time in the core properties, which it creates when missing.
                    if (corePart != null && !package.PartExists(corePart)) return Decline("the core properties relationship has no part");
                    if (corePart == null && package.PartExists(NewCorePropertiesPartName)) return Decline("the core properties part is not related to the package");
                }

                using var input = new MemoryStream(_deferredPackageBytes!, writable: false);
                using OfficeOpcEditablePackage editable = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
                ExcelDeferredSharedStrings? sharedStrings = ExcelDeferredSharedStrings.Load(editable, sharedStringsPart);
                if (sharedStrings == null) return Decline("the shared strings markup needs the model");
                ExcelOwnedStyles? styles = null;
                bool stylesLoaded = false;

                // Shared strings are added in the order values were written, as the model adds them; a cell's last value
                // wins and its style changes apply in turn when the worksheet is merged.
                foreach ((_, ExcelDeferredCellStore store) in edits) {
                    for (int index = 0; index < store.Count; index++) {
                        ref ExcelDeferredCellEdit edit = ref store[index];
                        if (edit.Kind == ExcelDeferredCellKind.Text) {
                            sharedStrings.TryGetOrAdd(edit.Text, ExcelSheet.CellValuePlainStringPromotionSharedStringCount, out _);
                        } else if (edit.Kind is ExcelDeferredCellKind.Date or ExcelDeferredCellKind.Style && !stylesLoaded) {
                            styles = ExcelOwnedStyles.Load(editable, stylesPart);
                            stylesLoaded = true;
                            if (styles == null) return Decline("the stylesheet needs the model");
                        }
                    }
                }

                ExcelDateSystem dateSystem = DateSystem;
                long changedBytes = 0;
                foreach ((ExcelSheet sheet, ExcelDeferredCellStore store) in edits) {
                    ExcelDeferredSheet entry = sheet.DeferredEntry ?? throw new InvalidOperationException("A recorded worksheet is bound to the model.");
                    var cells = new ExcelDeferredSheetCells(store, sharedStrings.IndexOf, date => ExcelDateSystemConverter.ToSerial(date, dateSystem));
                    byte[] merged;
                    int mergedLength;
                    OfficePooledBytes worksheet = editable.ReadPart(entry.PartName);
                    using (worksheet) {
                        byte[]? result = ExcelDeferredWorksheetWriter.TryMerge(worksheet.Array, worksheet.Length, cells,
                            (baseStyle, operation) => styles?.Apply(baseStyle, operation), out mergedLength,
                            out int addedSharedReferences, out int removedSharedReferences, out string? mergeReason);
                        if (result == null) return Decline(entry.Name + ": " + mergeReason);
                        merged = result;
                        changedBytes += mergedLength;
                        sharedStrings.AddedReferences += addedSharedReferences;
                        sharedStrings.RemovedReferences += removedSharedReferences;
                    }
                    editable.ReplacePart(entry.PartName, merged, mergedLength, pooled: true);
                }

                if (!sharedStrings.TryWrite(editable, workbook.WorkbookPartName)) return Decline("the shared strings part cannot take new items");
                styles?.Write();
                if (corePart == null) {
                    CreateCoreProperties(editable);
                } else if (!TrySetModifiedTime(editable, corePart)) {
                    return Decline("the core properties have no modification time to update");
                }

                // Deflated XML is a fraction of its size; the package rarely outgrows its source by more than its new parts.
                long capacity = Math.Min(int.MaxValue / 2, _deferredPackageBytes!.Length + changedBytes / 3 + 16 * 1024);
                using var output = new LimitedMemoryStream((int)Math.Min(capacity, limit ?? capacity), limit);
                editable.Save(output);
                bytes = output.ToArray();
                return true;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or FormatException or OverflowException
                && !exception.Data.Contains(PackageMaterializationLimitMarker)) {
                return Decline(exception.Message);
            }
        }

        private const string NewCorePropertiesPartName = "/docProps/core.xml";

        /// <summary>A memory stream that refuses to grow past the in-memory save limit.</summary>
        private sealed class LimitedMemoryStream : MemoryStream {
            private readonly long? _limit;

            internal LimitedMemoryStream(int capacity, long? limit) : base(capacity) => _limit = limit;

            private void Reserve(long count) {
                if (_limit is not long limit) return;
                long needed = Position + count;
                if (needed > limit) {
                    throw CreatePackageMaterializationLimitException($"The Excel package exceeds the {limit}-byte in-memory save limit.");
                }
                // The stream doubles its buffer when it grows; the buffer stops at the limit instead.
                if (needed > Capacity) Capacity = (int)Math.Min(limit, Math.Max(needed, Math.Min(int.MaxValue, 2L * Capacity)));
            }

            public override void Write(byte[] buffer, int offset, int count) {
                Reserve(count);
                base.Write(buffer, offset, count);
            }

            public override void WriteByte(byte value) {
                Reserve(1);
                base.WriteByte(value);
            }

#if NETCOREAPP2_1_OR_GREATER || NETSTANDARD2_1_OR_GREATER
            public override void Write(ReadOnlySpan<byte> buffer) {
                Reserve(buffer.Length);
                base.Write(buffer);
            }
#endif
        }

        private static string ModifiedTimeText() => DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture);

        /// <summary>Adds core properties holding the modification time, as a save through the model does for a package without them.</summary>
        private static void CreateCoreProperties(OfficeOpcEditablePackage package) {
            string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\""
                + " xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
                + "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + ModifiedTimeText() + "</dcterms:modified></cp:coreProperties>";
            package.CreateBinaryPart(NewCorePropertiesPartName, "application/vnd.openxmlformats-package.core-properties+xml", System.Text.Encoding.UTF8.GetBytes(xml));
            package.AddRelationship(null, OfficeOpcPackage.CorePropertiesRelationshipType, "docProps/core.xml");
        }

        // A save through the model writes package dates in UTC with seven fractional digits.
        private static bool TrySetModifiedTime(OfficeOpcEditablePackage package, string? corePart) {
            if (corePart == null) return false;
            OfficeXmlTree core = package.GetXml(corePart);
            int modified = core.FindChild(core.Root, OfficeXmlNamespaces.DublinCoreTerms, "modified");
            if (modified == OfficeXmlTree.NoNode) return false;
            core.SetInnerText(modified, ModifiedTimeText());
            return true;
        }
    }
}
