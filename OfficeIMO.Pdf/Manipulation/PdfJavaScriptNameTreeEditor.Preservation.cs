using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Pdf;

internal static partial class PdfJavaScriptNameTreeEditor {
    internal static byte[] CreateUntouchedSnapshot(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        IReadOnlyList<PdfJavaScriptEditSession.EditCommand> commands,
        PdfReadLimits limits) {
        int lastClear = -1;
        var affectedNames = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < commands.Count; i++) {
            if (commands[i].Kind == PdfJavaScriptEditSession.EditKind.Clear) {
                lastClear = i;
                affectedNames.Clear();
            } else if (i > lastClear && commands[i].Name is string name) {
                affectedNames.Add(name);
            }
        }

        using var fingerprint = new ObjectGraphFingerprint(objects);
        if (lastClear >= 0 ||
            !security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) ||
            root.Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("Names", out PdfObject? namesObject) ||
            ResolveDictionary(objects, namesObject) is not PdfDictionary names ||
            !names.Items.TryGetValue("JavaScript", out PdfObject? treeObject)) {
            return fingerprint.Complete();
        }

        var entries = new List<NameTreeEntry>();
        int traversedNodes = 0;
        CollectEntries(
            objects,
            treeObject,
            entries,
            new HashSet<(int ObjectNumber, int Generation)>(),
            0,
            ref traversedNodes,
            limits);

        for (int i = 0; i < entries.Count; i++) {
            NameTreeEntry entry = entries[i];
            if (entry.Name is not null && affectedNames.Contains(entry.Name)) continue;
            fingerprint.AppendEntry(entry.KeyBytes, entry.Value);
        }
        return fingerprint.Complete();
    }

    private sealed class ObjectGraphFingerprint : IDisposable {
        private readonly Dictionary<int, PdfIndirectObject> _objects;
        private readonly IncrementalHash _hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        private readonly Dictionary<(int ObjectNumber, int Generation), int> _references = new();
        private readonly byte[] _int32Bytes = new byte[sizeof(int)];
        private readonly byte[] _singleByte = new byte[1];

        internal ObjectGraphFingerprint(Dictionary<int, PdfIndirectObject> objects) {
            _objects = objects;
        }

        internal void AppendEntry(byte[] keyBytes, PdfObject value) {
            AppendByte(1);
            AppendBytes(keyBytes);
            AppendObject(value);
        }

        internal byte[] Complete() => _hash.GetHashAndReset();

        public void Dispose() => _hash.Dispose();

        private void AppendObject(PdfObject value) {
            switch (value) {
                case PdfNumber number:
                    AppendByte(2);
                    AppendString(number.Value.ToString("R", CultureInfo.InvariantCulture));
                    break;
                case PdfBoolean boolean:
                    AppendByte(boolean.Value ? (byte)3 : (byte)4);
                    break;
                case PdfName name:
                    AppendByte(5);
                    AppendString(name.Name);
                    break;
                case PdfStringObj text:
                    AppendByte(6);
                    AppendBytes(text.RawBytes);
                    break;
                case PdfArray array:
                    AppendByte(7);
                    AppendInt32(array.Items.Count);
                    for (int i = 0; i < array.Items.Count; i++) AppendObject(array.Items[i]);
                    break;
                case PdfDictionary dictionary:
                    AppendDictionary(8, dictionary);
                    break;
                case PdfReference reference:
                    AppendReference(reference);
                    break;
                case PdfStream stream:
                    AppendDictionary(10, stream.Dictionary);
                    AppendBytes(stream.Data);
                    break;
                case PdfNull:
                    AppendByte(11);
                    break;
                default:
                    throw new InvalidDataException("The PDF JavaScript action graph contains an unsupported object type.");
            }
        }

        private void AppendReference(PdfReference reference) {
            var key = (reference.ObjectNumber, reference.Generation);
            if (_references.TryGetValue(key, out int existingId)) {
                AppendByte(12);
                AppendInt32(existingId);
                return;
            }

            int id = _references.Count;
            _references[key] = id;
            AppendByte(13);
            AppendInt32(id);
            if (PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject? indirect)) {
                AppendByte(14);
                AppendObject(indirect.Value);
            } else {
                AppendByte(15);
            }
        }

        private void AppendDictionary(byte marker, PdfDictionary dictionary) {
            AppendByte(marker);
            AppendInt32(dictionary.Items.Count);
            foreach (KeyValuePair<string, PdfObject> item in dictionary.Items.OrderBy(static item => item.Key, StringComparer.Ordinal)) {
                AppendString(item.Key);
                AppendObject(item.Value);
            }
        }

        private void AppendString(string value) => AppendBytes(Encoding.UTF8.GetBytes(value));

        private void AppendBytes(byte[] value) {
            AppendInt32(value.Length);
            _hash.AppendData(value);
        }

        private void AppendInt32(int value) {
            _int32Bytes[0] = (byte)value;
            _int32Bytes[1] = (byte)(value >> 8);
            _int32Bytes[2] = (byte)(value >> 16);
            _int32Bytes[3] = (byte)(value >> 24);
            _hash.AppendData(_int32Bytes);
        }

        private void AppendByte(byte value) {
            _singleByte[0] = value;
            _hash.AppendData(_singleByte);
        }
    }
}
