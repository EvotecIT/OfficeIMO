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
            Assert.Throws<OperationCanceledException>(() => OfficeVbaProjectCanonicalizer.TryDecompress(Literal(new byte[32]), ref remaining, out _, out _, canceled.Token));
            Assert.Equal(100, remaining);
        }

        [Fact]
        public void FailedExpansionDebitsItsActualBytesAndExposesNoPartialOutput() {
            int remaining = 40;
            Assert.False(OfficeVbaProjectCanonicalizer.TryDecompress(MalformedAfterExpansion(32), ref remaining, out byte[] output, out _));
            Assert.Empty(output); Assert.Equal(8, remaining);
            Assert.True(OfficeVbaProjectCanonicalizer.TryDecompress(Literal(new byte[8]), ref remaining, out byte[] valid, out _));
            Assert.Equal(8, valid.Length); Assert.Equal(0, remaining);
            Assert.False(OfficeVbaProjectCanonicalizer.TryDecompress(Literal(new byte[1]), ref remaining, out output, out _));
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
        public void RepeatedMalformedStreamReferencesCannotReuseTheExpansionAllowance() {
            byte[] directory = Directory("Broken", "Broken", "Broken", "Later");
            Dictionary<string, byte[]> streams = new Dictionary<string, byte[]> {
                ["VBA/dir"] = Literal(directory),
                ["VBA/Broken"] = MalformedAfterExpansion(32),
                ["VBA/Later"] = Literal(Encoding.ASCII.GetBytes("OK"))
            };
            OfficeVbaInspection result = OfficeVbaProjectInspector.Inspect(streams, directory.Length + 64);
            Assert.Equal(4, result.Modules.Count);
            Assert.All(result.Modules, module => Assert.Null(module.Source));
            Assert.Contains("byte limit", result.Modules[2].Limitation);
            Assert.Contains("byte limit", result.Modules[3].Limitation);
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

        private static byte[] Directory(params string[] names) {
            List<byte> bytes = new List<byte>();
            Sized(bytes, 0x0001, new byte[4]); Sized(bytes, 0x0002, new byte[] { 9, 4, 0, 0 });
            Sized(bytes, 0x0003, new byte[] { 0xE4, 4 }); Sized(bytes, 0x0004, Encoding.ASCII.GetBytes("Project"));
            foreach (ushort id in new ushort[] { 5, 0x40, 6, 0x3D }) Sized(bytes, id, Array.Empty<byte>());
            Sized(bytes, 7, new byte[4]); Sized(bytes, 8, new byte[4]);
            U16(bytes, 9); U32(bytes, 4); U32(bytes, 1); U16(bytes, 0);
            Sized(bytes, 0xC, Array.Empty<byte>()); Sized(bytes, 0x3C, Array.Empty<byte>());
            Sized(bytes, 0xF, new byte[] { (byte)names.Length, 0 }); Sized(bytes, 0x13, new byte[2]);
            foreach (string name in names) {
                Sized(bytes, 0x19, Encoding.ASCII.GetBytes(name)); Sized(bytes, 0x47, Encoding.Unicode.GetBytes(name));
                Sized(bytes, 0x1A, Encoding.ASCII.GetBytes(name)); Sized(bytes, 0x32, Encoding.Unicode.GetBytes(name));
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
