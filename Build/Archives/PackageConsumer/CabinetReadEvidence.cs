using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.Archives.Qualification {
    /// <summary>Opt-in selected CAB/MSI read allocations, outside correctness CI and peak-memory claims.</summary>
    internal static class CabinetReadEvidence {
        internal static int Run(string[] args) {
            if (args.Length != 2) throw new ArgumentException("Expected: <CAB/MSI input> <JSON report>");
            if (!File.Exists(args[0])) CreateInput(args[0]);
            var results = new List<object>();
            string[]? singleHashes = null;
            foreach (bool grouped in new[] { false, true }) {
                using var inspection = ArchiveInspector.Open(args[0], new() {
                    MaxExpandedBytes = 512L * 1024 * 1024, MaxDecompressionWorkBytes = 1024L * 1024 * 1024
                });
                if (inspection.Format == ArchiveFormat.Zip) throw new ArgumentException("Use a CAB or MSI input.");
                int[] ordinals = inspection.Entries.Where(e => e.CanRead && e.Length > 0)
                    .OrderBy(e => e.Length).ThenBy(e => e.Ordinal).Take(2).Select(e => e.Ordinal).ToArray();
                if (ordinals.Length == 0) throw new InvalidDataException("No supported nonempty members.");
                GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                long before = GC.GetAllocatedBytesForCurrentThread();
                var hashes = new List<string>();
                if (grouped) {
                    using var batch = inspection.ReadMembers(ordinals, 128L * 1024 * 1024);
                    foreach (var member in batch.Members) hashes.Add(Hash(member.Content));
                } else {
                    foreach (int ordinal in ordinals) {
                        using var member = inspection.OpenRead(ordinal);
                        hashes.Add(Hash(member));
                    }
                }
                long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
                if (singleHashes != null && !singleHashes.SequenceEqual(hashes)) throw new InvalidDataException("Read routes disagree.");
                singleHashes = hashes.ToArray();
                results.Add(new { operation = grouped ? "ReadMembers" : "OpenRead", ordinals, hashes,
                    allocatedBytes = allocated, decodedBytes = inspection.ReadBytes, workBytes = inspection.DecompressionWorkBytes });
            }
            using var input = File.OpenRead(args[0]);
            File.WriteAllText(args[1], JsonSerializer.Serialize(new {
                runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
                assembly = typeof(ArchiveInspector).Assembly.GetName().Version?.ToString(),
                inputSha256 = Hash(input), inputBytes = input.Length, results,
                boundary = "Per-thread read allocation excludes input inventory; not peak/RSS or timing qualification"
            }, new JsonSerializerOptions { WriteIndented = true }));
            return 0;
        }

        private static string Hash(Stream stream) => Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();

        // Reproducible synthetic selection workload, not an independent CAB producer or product writer.
        private static void CreateInput(string path) {
            const int frameBytes = 32768, frameCount = 1024, dataOffset = 95;
            using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write);
            using var writer = new BinaryWriter(file);
            writer.Write(new byte[] { 77, 83, 67, 70 }); writer.Write(0);
            writer.Write(dataOffset + frameCount * (24 + frameBytes)); writer.Write(0); writer.Write(44); writer.Write(0);
            writer.Write((byte)3); writer.Write((byte)1); writer.Write((ushort)1); writer.Write((ushort)2);
            writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0);
            writer.Write(dataOffset); writer.Write((ushort)frameCount); writer.Write((ushort)4099);
            Member("cross.bin", 32767, 2); Member("tail.bin", frameCount * frameBytes - 5, 5);
            var payload = new byte[frameBytes];
            for (int frame = 0; frame < frameCount; frame++) {
                for (int i = 0; i < payload.Length; i++) payload[i] = (byte)(((long)frame * frameBytes + i) % 251);
                writer.Write(0); writer.Write((ushort)(frameBytes + 16)); writer.Write((ushort)frameBytes);
                // LZX uncompressed block header and repeated offsets, followed by one complete frame.
                writer.Write((ushort)(frame == 0 ? 0x3008 : 0x6010)); writer.Write((ushort)0);
                writer.Write(1u); writer.Write(1u); writer.Write(1u); writer.Write(payload);
            }
            void Member(string name, int offset, int length) {
                writer.Write(length); writer.Write(offset); writer.Write((ushort)0); writer.Write((ushort)0);
                writer.Write((ushort)0); writer.Write((ushort)0); writer.Write(System.Text.Encoding.ASCII.GetBytes(name)); writer.Write((byte)0);
            }
        }
    }
}
