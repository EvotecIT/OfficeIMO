using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Security.Tests {
    public sealed class OfficeVbaInspectionTests {
        [Fact]
        public void InspectionAndExpansionHonorCancellation() {
            using CancellationTokenSource canceled = new CancellationTokenSource(); canceled.Cancel();
            Assert.Throws<OperationCanceledException>(() => OfficeVbaProjectInspector.Inspect(new Dictionary<string, byte[]>(), 100, canceled.Token));
            int remaining = 100;
            Assert.Throws<OperationCanceledException>(() => OfficeVbaCompression.TryDecompress(Literal(new byte[32]), ref remaining, out _, out _, canceled.Token));
            Assert.Equal(100, remaining);
        }

        [Fact]
        public void FailedExpansionDebitsItsActualBytesAndExposesNoPartialOutput() {
            int remaining = 40;
            Assert.False(OfficeVbaCompression.TryDecompress(MalformedAfterExpansion(32), ref remaining, out byte[] output, out _));
            Assert.Empty(output); Assert.Equal(8, remaining);
            Assert.True(OfficeVbaCompression.TryDecompress(Literal(new byte[8]), ref remaining, out byte[] valid, out _));
            Assert.Equal(8, valid.Length); Assert.Equal(0, remaining);
            Assert.False(OfficeVbaCompression.TryDecompress(Literal(new byte[1]), ref remaining, out output, out _));
            Assert.Empty(output); Assert.Equal(0, remaining);
        }

        [Fact]
        public void MalformedModuleExpansionConsumesTheProjectBudget() {
            byte[] directory = Directory("Broken", "Later");
            Dictionary<string, byte[]> streams = new Dictionary<string, byte[]> {
                ["VBA/dir"] = Literal(directory),
                ["VBA/Broken"] = MalformedAfterExpansion(32),
                ["VBA/Later"] = Literal(Encoding.ASCII.GetBytes("OK"))
            };
            OfficeVbaInspection result = OfficeVbaProjectInspector.Inspect(streams, directory.Length + 32);
            Assert.Null(result.Limitation);
            Assert.Equal(2, result.Modules.Count);
            Assert.Null(result.Modules[0].Source);
            Assert.Contains("chunk header", result.Modules[0].Limitation);
            Assert.Null(result.Modules[1].Source);
            Assert.Contains("byte limit", result.Modules[1].Limitation);
        }

        [Fact]
        public void RepeatedMalformedReferencesRespectBothAllowancesAndRetainLaterSource() {
            byte[] directory = Directory("Broken", "Broken", "Broken", "Later");
            Dictionary<string, byte[]> streams = new Dictionary<string, byte[]> {
                ["VBA/dir"] = Literal(directory),
                ["VBA/Broken"] = MalformedAfterExpansion(32),
                ["VBA/Later"] = Literal(Encoding.ASCII.GetBytes("OK"))
            };
            OfficeVbaInspection result = OfficeVbaProjectInspector.Inspect(streams, directory.Length + 64);
            Assert.Equal(4, result.Modules.Count);
            for (int index = 0; index < 3; index++) Assert.Null(result.Modules[index].Source);
            Assert.Contains("encoded input byte limit", result.Modules[1].Limitation);
            Assert.Contains("encoded input byte limit", result.Modules[2].Limitation);
            Assert.Equal("OK", result.Modules[3].Source);
        }

        [Fact]
        public void MalformedModuleRetainsMetadataAndLaterSourceWithinTheRemainingBudget() {
            byte[] directory = Directory("Broken", "Later");
            Dictionary<string, byte[]> streams = new Dictionary<string, byte[]> {
                ["VBA/dir"] = Literal(directory),
                ["VBA/Broken"] = MalformedAfterExpansion(32),
                ["VBA/Later"] = Literal(Encoding.ASCII.GetBytes("OK"))
            };
            OfficeVbaInspection result = OfficeVbaProjectInspector.Inspect(streams, directory.Length + 34);
            Assert.Equal("Broken", result.Modules[0].Name);
            Assert.Null(result.Modules[0].Source);
            Assert.Equal("OK", result.Modules[1].Source);
        }

        private static byte[] MalformedAfterExpansion(int count) {
            List<byte> bytes = new List<byte>(Literal(new byte[count]));
            bytes.Add(0x99); // A truncated second chunk, after a successful expansion.
            return bytes.ToArray();
        }

        [Fact]
        public void InvalidUtf8SourceRetainsProjectInventoryAndLaterValidSource() {
            byte[] directory = Directory(65001, "Broken", "Later");
            byte[] later = Encoding.UTF8.GetBytes("Zażółć");
            Dictionary<string, byte[]> streams = new Dictionary<string, byte[]> {
                ["VBA/dir"] = Literal(directory),
                ["VBA/Broken"] = Literal(new byte[] { 0xC3, 0x28 }),
                ["VBA/Later"] = Literal(later)
            };
            OfficeVbaInspection result = OfficeVbaProjectInspector.Inspect(streams, directory.Length + 2 + later.Length);
            Assert.Null(result.Limitation);
            Assert.Equal("Project", result.Name); Assert.Equal(65001, result.CodePage);
            Assert.Equal(2, result.Modules.Count);
            Assert.Equal("Broken", result.Modules[0].Name); Assert.Equal("Broken", result.Modules[0].StreamName);
            Assert.Null(result.Modules[0].Source); Assert.Contains("declared code page", result.Modules[0].Limitation);
            Assert.Equal("Zażółć", result.Modules[1].Source); Assert.Null(result.Modules[1].Limitation);
            OfficeVbaInspection bounded = OfficeVbaProjectInspector.Inspect(streams, directory.Length + 2);
            Assert.Equal(2, bounded.Modules.Count);
            Assert.Null(bounded.Modules[1].Source); Assert.Contains("byte limit", bounded.Modules[1].Limitation);
        }

        [Fact]
        public void ImmediateMalformedReferencesCannotReuseTheEncodedInputAllowance() {
            byte[] directory = Directory("Broken", "Broken", "Later");
            Dictionary<string, byte[]> streams = new Dictionary<string, byte[]> {
                ["VBA/dir"] = Literal(directory),
                ["VBA/Broken"] = new byte[1024], // Invalid signature; no bytes expand.
                ["VBA/Later"] = Literal(Encoding.ASCII.GetBytes("OK"))
            };
            OfficeVbaInspection result = OfficeVbaProjectInspector.Inspect(streams, directory.Length + 2);
            Assert.Null(result.Limitation); Assert.Equal(3, result.Modules.Count);
            Assert.Null(result.Modules[0].Source); Assert.Contains("signature", result.Modules[0].Limitation);
            Assert.Null(result.Modules[1].Source); Assert.Contains("encoded input byte limit", result.Modules[1].Limitation);
            Assert.Equal("OK", result.Modules[2].Source);
        }

        private static byte[] Directory(params string[] names) => Directory(1252, names);

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        public void MalformedDirectoryStreamIdentityCannotBindReplacementNamedSource(int malformedKind) {
            byte[] unicodeStream = malformedKind == 0 ? new byte[] { 0, 0xD8 }
                : malformedKind == 1 ? new byte[] { 0x42 }
                : malformedKind == 2 ? Array.Empty<byte>() : Encoding.Unicode.GetBytes("Broken\0");
            byte[] ansiStream = malformedKind == 2 ? new byte[] { 0xC9 } : Encoding.ASCII.GetBytes("Broken");
            byte[] directory = Directory(1252, unicodeStream, ansiStream, "Broken");
            Dictionary<string, byte[]> streams = new Dictionary<string, byte[]> {
                ["VBA/dir"] = Literal(directory),
                ["VBA/\uFFFD"] = Literal(Encoding.ASCII.GetBytes("Unrelated replacement-named source")),
                ["VBA/?"] = Literal(Encoding.ASCII.GetBytes("Unrelated ASCII-fallback source")),
                ["VBA/Broken"] = Literal(Encoding.ASCII.GetBytes("Unrelated null-trimmed source"))
            };
            OfficeVbaInspection result = OfficeVbaProjectInspector.Inspect(streams, directory.Length + 100);
            Assert.NotNull(result.Limitation); Assert.Empty(result.Modules); Assert.Null(result.Name);
            List<OfficeCompoundStream> compoundStreams = new List<OfficeCompoundStream> {
                new OfficeCompoundStream("PROJECT", Encoding.ASCII.GetBytes("Name=\"Project\"\r\nModule=Broken\r\n"))
            };
            foreach (KeyValuePair<string, byte[]> stream in streams) compoundStreams.Add(new OfficeCompoundStream(stream.Key, stream.Value));
            byte[] projectBytes = OfficeCompoundFileWriter.Write(compoundStreams);
            Assert.Contains("module record", Assert.Throws<System.IO.InvalidDataException>(() => OfficeVbaProject.Load(projectBytes)).Message);
            Assert.False(OfficeVbaProjectCanonicalizer.TryCreate(projectBytes, directory.Length + 100, out _, out string detail));
            Assert.Contains("module record", detail);
        }

        private static byte[] Directory(ushort codePage, params string[] names) => Directory(codePage, null, null, names);

        private static byte[] Directory(ushort codePage, byte[]? unicodeStreamName, byte[]? ansiStreamName, params string[] names) {
            List<byte> bytes = new List<byte>();
            Sized(bytes, 0x0001, new byte[4]); Sized(bytes, 0x0002, new byte[] { 9, 4, 0, 0 });
            Sized(bytes, 0x0003, new byte[] { (byte)codePage, (byte)(codePage >> 8) }); Sized(bytes, 0x0004, Encoding.ASCII.GetBytes("Project"));
            foreach (ushort id in new ushort[] { 5, 0x40, 6, 0x3D }) Sized(bytes, id, Array.Empty<byte>());
            Sized(bytes, 7, new byte[4]); Sized(bytes, 8, new byte[4]);
            U16(bytes, 9); U32(bytes, 4); U32(bytes, 1); U16(bytes, 0);
            Sized(bytes, 0xC, Array.Empty<byte>()); Sized(bytes, 0x3C, Array.Empty<byte>());
            Sized(bytes, 0xF, new byte[] { (byte)names.Length, 0 }); Sized(bytes, 0x13, new byte[2]);
            foreach (string name in names) {
                Sized(bytes, 0x19, Encoding.ASCII.GetBytes(name)); Sized(bytes, 0x47, Encoding.Unicode.GetBytes(name));
                Sized(bytes, 0x1A, ansiStreamName ?? Encoding.ASCII.GetBytes(name)); Sized(bytes, 0x32, unicodeStreamName ?? Encoding.Unicode.GetBytes(name));
                Sized(bytes, 0x1C, Array.Empty<byte>()); Sized(bytes, 0x48, Array.Empty<byte>());
                Sized(bytes, 0x31, new byte[4]); Sized(bytes, 0x1E, new byte[4]); Sized(bytes, 0x2C, new byte[2]);
                U16(bytes, 0x21); U32(bytes, 0); U16(bytes, 0x2B); U32(bytes, 0);
            }
            U16(bytes, 0x10); U32(bytes, 0); return bytes.ToArray();
        }

        private static byte[] Literal(byte[] source) {
            List<byte> payload = new List<byte>();
            for (int offset = 0; offset < source.Length; offset += 8) {
                payload.Add(0);
                for (int i = offset; i < Math.Min(offset + 8, source.Length); i++) payload.Add(source[i]);
            }
            List<byte> bytes = new List<byte> { 1 }; U16(bytes, (ushort)(0xB000 | (payload.Count - 1)));
            bytes.AddRange(payload); return bytes.ToArray();
        }
        private static void Sized(List<byte> bytes, ushort id, byte[] value) { U16(bytes, id); U32(bytes, (uint)value.Length); bytes.AddRange(value); }
        private static void U16(List<byte> bytes, ushort value) { bytes.Add((byte)value); bytes.Add((byte)(value >> 8)); }
        private static void U32(List<byte> bytes, uint value) { U16(bytes, (ushort)value); U16(bytes, (ushort)(value >> 16)); }
    }
}
