using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Drawing.Internal;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        internal SharedStringTablePart SharedStringTablePart {
            get {
                // Check if already initialized without lock first (double-check locking pattern)
                if (_sharedStringTablePart != null) {
                    return _sharedStringTablePart;
                }

                // Check if we're in a NoLock scope or already have a lock - if so, initialize without locking
                if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)) {
                    var existingPart = _workBookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (existingPart != null) {
                        _sharedStringTablePart = existingPart;
                    } else {
                        _sharedStringTablePart = _workBookPart.AddNewPart<SharedStringTablePart>();
                        _sharedStringTablePart.SharedStringTable = new SharedStringTable();
                        _sharedStringTableCount = 0;
                    }
                    return _sharedStringTablePart!;
                }

                // Use write lock for initialization when no lock is held
                return Locking.ExecuteWrite(EnsureLock(), () => {
                    // Double-check inside the lock
                    if (_sharedStringTablePart == null) {
                        var existingPart = _workBookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (existingPart != null) {
                            _sharedStringTablePart = existingPart;
                        } else {
                            _sharedStringTablePart = _workBookPart.AddNewPart<SharedStringTablePart>();
                            _sharedStringTablePart.SharedStringTable = new SharedStringTable();
                            _sharedStringTableCount = 0;
                        }
                    }
                    return _sharedStringTablePart;
                });
            }
        }

        internal int GetSharedStringIndex(string text) {
            return GetSharedStringIndex(text, validateNewString: false);
        }

        internal int GetSharedStringIndex(string text, bool validateNewString) {
            if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)) {
                return GetSharedStringIndexCore(text, validateNewString);
            }

            lock (_sharedStringLock) {
                return GetSharedStringIndexCore(text, validateNewString);
            }
        }

        internal int GetSharedStringIndex(string text, bool validateNewString, out bool containsLineBreak) {
            if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)) {
                return GetSharedStringIndexCore(text, validateNewString, out containsLineBreak);
            }

            lock (_sharedStringLock) {
                return GetSharedStringIndexCore(text, validateNewString, out containsLineBreak);
            }
        }

        internal bool TryGetExistingSharedStringIndex(string text, out int index, out bool containsLineBreak, out int sharedStringCount) {
            if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)) {
                return TryGetExistingSharedStringIndexCore(text, out index, out containsLineBreak, out sharedStringCount);
            }

            lock (_sharedStringLock) {
                return TryGetExistingSharedStringIndexCore(text, out index, out containsLineBreak, out sharedStringCount);
            }
        }

        internal bool TryGetOrAddSharedStringIndexBelowLimit(string text, int addLimit, bool validateNewString, out int index, out bool containsLineBreak) {
            if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)) {
                return TryGetOrAddSharedStringIndexBelowLimitCore(text, addLimit, validateNewString, out index, out containsLineBreak);
            }

            lock (_sharedStringLock) {
                return TryGetOrAddSharedStringIndexBelowLimitCore(text, addLimit, validateNewString, out index, out containsLineBreak);
            }
        }

        private int GetSharedStringIndexCore(string text, bool validateNewString) {
            // Check cache first
            if (_sharedStringCache.TryGetValue(text, out int cachedIndex)) {
                return cachedIndex;
            }

            var sharedStringTable = SharedStringTablePart.SharedStringTable ??= new SharedStringTable();
            int tableCount = EnsureSharedStringCacheAndCount(sharedStringTable);

            // Check again after rebuilding cache
            if (_sharedStringCache.TryGetValue(text, out int foundIndex)) {
                return foundIndex;
            }

            if (validateNewString) {
                CoerceValueHelper.ValidateSharedStringLength(text, nameof(text));
            }

            // Add new string
            int newIndex = tableCount;
            sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            _sharedStringTableCount = newIndex + 1;
            _sharedStringTableDirty = true;
            MarkPackageDirty();
            _sharedStringCache[text] = newIndex;

            return newIndex;
        }

        private int GetSharedStringIndexCore(string text, bool validateNewString, out bool containsLineBreak) {
            if (_sharedStringCache.TryGetValue(text, out int cachedIndex)) {
                containsLineBreak = GetCachedOrComputeSharedStringLineBreak(text);
                return cachedIndex;
            }

            var sharedStringTable = SharedStringTablePart.SharedStringTable ??= new SharedStringTable();
            int tableCount = EnsureSharedStringCacheAndCount(sharedStringTable);

            if (_sharedStringCache.TryGetValue(text, out int foundIndex)) {
                containsLineBreak = GetCachedOrComputeSharedStringLineBreak(text);
                return foundIndex;
            }

            if (validateNewString) {
                CoerceValueHelper.ValidateSharedStringLength(text, nameof(text));
            }

            containsLineBreak = ContainsLineBreak(text);

            int newIndex = tableCount;
            sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            _sharedStringTableCount = newIndex + 1;
            _sharedStringTableDirty = true;
            MarkPackageDirty();
            _sharedStringCache[text] = newIndex;

            return newIndex;
        }

        private bool TryGetExistingSharedStringIndexCore(string text, out int index, out bool containsLineBreak, out int sharedStringCount) {
            if (_sharedStringCache.TryGetValue(text, out index)) {
                containsLineBreak = GetCachedOrComputeSharedStringLineBreak(text);
                sharedStringCount = _sharedStringTableCount >= 0 ? _sharedStringTableCount : _sharedStringCache.Count;
                return true;
            }

            if (_sharedStringTablePart != null && _sharedStringTableCount >= 0 && _sharedStringCache.Count > 0) {
                index = -1;
                containsLineBreak = false;
                sharedStringCount = _sharedStringTableCount;
                return false;
            }

            SharedStringTablePart? sharedStringTablePart = _sharedStringTablePart
                ?? _workBookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var sharedStringTable = sharedStringTablePart?.SharedStringTable;
            if (sharedStringTable == null) {
                index = -1;
                containsLineBreak = false;
                sharedStringCount = 0;
                return false;
            }

            _sharedStringTablePart = sharedStringTablePart;
            sharedStringCount = EnsureSharedStringCacheAndCount(sharedStringTable);
            if (_sharedStringCache.TryGetValue(text, out index)) {
                containsLineBreak = GetCachedOrComputeSharedStringLineBreak(text);
                return true;
            }

            index = -1;
            containsLineBreak = false;
            return false;
        }

        private bool TryGetOrAddSharedStringIndexBelowLimitCore(string text, int addLimit, bool validateNewString, out int index, out bool containsLineBreak) {
            if (_sharedStringCache.TryGetValue(text, out index)) {
                containsLineBreak = GetCachedOrComputeSharedStringLineBreak(text);
                return true;
            }

            var sharedStringTable = SharedStringTablePart.SharedStringTable ??= new SharedStringTable();
            int tableCount = EnsureSharedStringCacheAndCount(sharedStringTable);
            if (_sharedStringCache.TryGetValue(text, out index)) {
                containsLineBreak = GetCachedOrComputeSharedStringLineBreak(text);
                return true;
            }

            if (tableCount >= addLimit) {
                index = -1;
                containsLineBreak = ContainsLineBreak(text);
                return false;
            }

            if (validateNewString) {
                CoerceValueHelper.ValidateSharedStringLength(text, nameof(text));
            }

            containsLineBreak = ContainsLineBreak(text);
            index = tableCount;
            sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            _sharedStringTableCount = index + 1;
            _sharedStringTableDirty = true;
            MarkPackageDirty();
            _sharedStringCache[text] = index;
            return true;
        }

        private bool GetCachedOrComputeSharedStringLineBreak(string text) {
            if (_sharedStringLineBreakCache != null
                && _sharedStringLineBreakCache.TryGetValue(text, out bool containsLineBreak)) {
                return containsLineBreak;
            }

            containsLineBreak = ContainsLineBreak(text);
            if (text.Length >= 16) {
                (_sharedStringLineBreakCache ??= new Dictionary<string, bool>(StringComparer.Ordinal))[text] = containsLineBreak;
            }

            return containsLineBreak;
        }

        private static bool ContainsLineBreak(string text) {
            return text.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0;
        }

        internal Dictionary<string, int> GetSharedStringIndices(IEnumerable<string> texts, bool assumeDistinct = false) {
            if (texts == null) {
                throw new ArgumentNullException(nameof(texts));
            }

            int capacity = texts is ICollection<string> collection ? collection.Count : 0;
            if (capacity == 0 && texts is ICollection<string>) {
                return new Dictionary<string, int>(0, StringComparer.Ordinal);
            }

            lock (_sharedStringLock) {
                var sharedStringTable = SharedStringTablePart.SharedStringTable ??= new SharedStringTable();
                int tableCount = EnsureSharedStringCacheAndCount(sharedStringTable);

                var result = capacity > 0
                    ? new Dictionary<string, int>(capacity, StringComparer.Ordinal)
                    : new Dictionary<string, int>(StringComparer.Ordinal);
                bool changed = false;

                foreach (string text in texts) {
                    if (!assumeDistinct && result.ContainsKey(text)) {
                        continue;
                    }

                    if (_sharedStringCache.TryGetValue(text, out int existingIndex)) {
                        result[text] = existingIndex;
                        continue;
                    }

                    int newIndex = tableCount++;
                    sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
                    _sharedStringCache[text] = newIndex;
                    result[text] = newIndex;
                    changed = true;
                }

                _sharedStringTableCount = tableCount;

                if (changed) {
                    _sharedStringTableDirty = true;
                    MarkPackageDirty();
                }

                return result;
            }
        }

        internal int[] GetSharedStringIndexArray(IReadOnlyList<string> texts, bool assumeDistinct = false) {
            if (texts == null) {
                throw new ArgumentNullException(nameof(texts));
            }

            if (texts.Count == 0) {
                return Array.Empty<int>();
            }

            lock (_sharedStringLock) {
                var sharedStringTable = SharedStringTablePart.SharedStringTable ??= new SharedStringTable();
                int tableCount = EnsureSharedStringCacheAndCount(sharedStringTable);
                var result = new int[texts.Count];
                Dictionary<string, int>? localIndexes = assumeDistinct
                    ? null
                    : new Dictionary<string, int>(texts.Count, StringComparer.Ordinal);
                bool changed = false;

                for (int i = 0; i < texts.Count; i++) {
                    string text = texts[i];
                    if (localIndexes != null && localIndexes.TryGetValue(text, out int duplicateIndex)) {
                        result[i] = duplicateIndex;
                        continue;
                    }

                    if (!_sharedStringCache.TryGetValue(text, out int sharedStringIndex)) {
                        sharedStringIndex = tableCount++;
                        sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
                        _sharedStringCache[text] = sharedStringIndex;
                        changed = true;
                    }

                    result[i] = sharedStringIndex;
                    localIndexes?.Add(text, sharedStringIndex);
                }

                _sharedStringTableCount = tableCount;

                if (changed) {
                    _sharedStringTableDirty = true;
                    MarkPackageDirty();
                }

                return result;
            }
        }

        private int EnsureSharedStringCacheAndCount(SharedStringTable sharedStringTable) {
            if (_sharedStringCache.Count == 0) {
                int idx = 0;
                foreach (SharedStringItem item in sharedStringTable.Elements<SharedStringItem>()) {
                    _sharedStringCache[item.InnerText] = idx;
                    idx++;
                }

                _sharedStringTableCount = idx;
            } else if (_sharedStringTableCount < 0) {
                _sharedStringTableCount = sharedStringTable.Elements<SharedStringItem>().Count();
            }

            return _sharedStringTableCount;
        }
    }
}
