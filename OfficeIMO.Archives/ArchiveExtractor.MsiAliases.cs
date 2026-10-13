using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Archives;

public static partial class ArchiveExtractor {
    /// <summary>Compares aliases across cabinets without retaining representative byte arrays in memory.</summary>
    private sealed class MsiAliasVerifier : IDisposable {
        private readonly string _root;
        private readonly CancellationToken _token;
        private readonly Dictionary<MsiPayloadTarget, AliasState> _states = new();

        internal MsiAliasVerifier(string staging, CancellationToken token) {
            _root = Path.Combine(Path.GetDirectoryName(staging)!, ".officeimo-msi-alias-" + Guid.NewGuid().ToString("N"));
            _token = token;
        }

        internal bool Verify(MsiPayloadTarget target, byte[] bytes) {
            _token.ThrowIfCancellationRequested();
            if (!_states.TryGetValue(target, out AliasState? state)) {
                ArchivePaths.RejectLinks(_root);
                Directory.CreateDirectory(_root);
                string path = Path.Combine(_root, Guid.NewGuid().ToString("N"));
                ArchivePaths.RejectLinks(path, _root);
                using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                    for (int offset = 0; offset < bytes.Length;) {
                        _token.ThrowIfCancellationRequested();
                        int count = Math.Min(81920, bytes.Length - offset);
                        output.Write(bytes, offset, count); offset += count;
                    }
                }
                _states.Add(target, new AliasState(path));
                return true;
            }
            ArchivePaths.RejectLinks(state.Path, _root);
            using (var input = new FileStream(state.Path, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                if (input.Length != bytes.LongLength) throw Conflict(target);
                var buffer = new byte[81920];
                int offset = 0;
                while (offset < bytes.Length) {
                    _token.ThrowIfCancellationRequested();
                    int read = input.Read(buffer, 0, Math.Min(buffer.Length, bytes.Length - offset));
                    if (read == 0) throw Conflict(target);
                    for (int index = 0; index < read; index++) if (buffer[index] != bytes[offset + index]) throw Conflict(target);
                    offset += read;
                }
            }
            if (++state.Seen == target.AliasCount) {
                File.Delete(state.Path);
                _states.Remove(target);
            }
            return false;
        }

        private static InvalidDataException Conflict(MsiPayloadTarget target) =>
            new("MSI alias payload bytes conflict: " + target.Path);

        public void Dispose() {
            if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
        }

        private sealed class AliasState {
            internal AliasState(string path) => Path = path;
            internal string Path { get; }
            internal int Seen { get; set; } = 1;
        }
    }
}
