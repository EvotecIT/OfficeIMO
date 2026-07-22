using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Security.Cryptography;

namespace OfficeIMO.Email;

/// <summary>Fingerprints the mutable public email model so raw passthrough is limited to unchanged documents.</summary>
internal static class EmailDocumentStateFingerprint {
    internal static byte[]? TryCompute(EmailDocument document) {
        using (SHA256 hash = SHA256.Create())
        using (CryptoStream crypto = new CryptoStream(Stream.Null, hash, CryptoStreamMode.Write))
        using (BinaryWriter output = new BinaryWriter(crypto, Encoding.UTF8, leaveOpen: true)) {
            var writer = new StateWriter(output);
            if (!writer.TryWrite(document)) return null;
            output.Flush();
            crypto.FlushFinalBlock();
            return hash.Hash == null ? null : (byte[])hash.Hash.Clone();
        }
    }

    internal static bool Matches(EmailDocument document, byte[] baseline) {
        byte[]? current = TryCompute(document);
        return current != null && FixedTimeEquals(current, baseline);
    }

    private static bool FixedTimeEquals(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        int difference = 0;
        for (int index = 0; index < left.Length; index++) difference |= left[index] ^ right[index];
        return difference == 0;
    }

    private sealed class StateWriter {
        private readonly BinaryWriter _output;
        private readonly HashSet<object> _visited = new HashSet<object>(ReferenceEqualityComparer.Instance);

        internal StateWriter(BinaryWriter output) {
            _output = output;
        }

        [UnconditionalSuppressMessage(
            "Trimming",
            "IL2075",
            Justification = "The fingerprint walks only OfficeIMO.Email model instances created by the directly referenced readers and writers. Unsupported external object-bag values return false and disable raw passthrough instead of affecting document output.")]
        internal bool TryWrite(object? value) {
            if (value == null) {
                _output.Write((byte)0);
                return true;
            }

            Type type = value.GetType();
            _output.Write((byte)1);
            _output.Write(type.FullName ?? type.Name);
            if (!type.IsValueType && value is not string && value is not byte[]) {
                if (!_visited.Add(value)) {
                    _output.Write((byte)2);
                    return true;
                }
            }

            if (value is byte[] bytes) {
                _output.Write(bytes.Length);
                _output.Write(bytes);
                return true;
            }
            if (value is string text) {
                _output.Write(text);
                return true;
            }
            if (value is DateTimeOffset dateTimeOffset) {
                _output.Write(dateTimeOffset.Ticks);
                _output.Write(dateTimeOffset.Offset.Ticks);
                return true;
            }
            if (value is DateTime dateTime) {
                _output.Write(dateTime.ToBinary());
                return true;
            }
            if (value is TimeSpan timeSpan) {
                _output.Write(timeSpan.Ticks);
                return true;
            }
            if (value is Guid guid) {
                _output.Write(guid.ToByteArray());
                return true;
            }
            if (type.IsEnum || type.IsPrimitive || value is decimal) {
                _output.Write(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                return true;
            }
            if (value is IDictionary dictionary) return TryWriteDictionary(dictionary);
            if (value is IEnumerable sequence) return TryWriteSequence(sequence);
            if (type.Assembly != typeof(EmailDocument).Assembly) return false;

            PropertyInfo[] properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(property => property.CanRead && property.GetIndexParameters().Length == 0 &&
                    !(type == typeof(EmailDocument) && property.Name == nameof(EmailDocument.RawSource)))
                .OrderBy(property => property.Name, StringComparer.Ordinal)
                .ToArray();
            _output.Write(properties.Length);
            try {
                foreach (PropertyInfo property in properties) {
                    _output.Write(property.Name);
                    if (!TryWrite(property.GetValue(value, null))) return false;
                }
            } catch (TargetInvocationException) {
                return false;
            }
            return true;
        }

        private bool TryWriteDictionary(IDictionary dictionary) {
            var entries = new List<DictionaryEntry>();
            foreach (DictionaryEntry entry in dictionary) entries.Add(entry);
            entries.Sort((left, right) => StringComparer.Ordinal.Compare(
                Convert.ToString(left.Key, CultureInfo.InvariantCulture),
                Convert.ToString(right.Key, CultureInfo.InvariantCulture)));
            _output.Write(entries.Count);
            foreach (DictionaryEntry entry in entries) {
                if (!TryWrite(entry.Key) || !TryWrite(entry.Value)) return false;
            }
            return true;
        }

        private bool TryWriteSequence(IEnumerable sequence) {
            int count = 0;
            foreach (object? item in sequence) {
                count++;
                if (!TryWrite(item)) return false;
            }
            _output.Write(count);
            return true;
        }
    }

    private sealed class ReferenceEqualityComparer : IEqualityComparer<object> {
        internal static ReferenceEqualityComparer Instance { get; } = new ReferenceEqualityComparer();

        public new bool Equals(object? left, object? right) => ReferenceEquals(left, right);

        public int GetHashCode(object value) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }
}
