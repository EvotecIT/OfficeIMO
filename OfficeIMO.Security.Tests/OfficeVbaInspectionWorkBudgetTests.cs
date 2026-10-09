using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.Security.Tests {
    public sealed class OfficeVbaInspectionWorkBudgetTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void FailedExpansionConsumesTheBudgetBeforeInspectingAnotherModule(bool invalidCopy) {
            OfficeVbaProject project = OfficeVbaProject.Create();
            project.AddModule("First", "'original"); project.AddModule("Second", "'valid");
            Assert.True(OfficeCompoundFileReader.TryRead(project.Write().GetBytes(), out OfficeCompoundFile? compound, out _));
            Dictionary<string, byte[]> streams = compound!.Streams.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
            byte[] prefix = OfficeVbaCompression.Compress(Enumerable.Repeat((byte)'a', 4096).ToArray());
            streams["VBA/First"] = prefix.Concat(invalidCopy ? new byte[] { 2, 0xb0, 1, 0, 0 } : new byte[] { 0 }).ToArray();
            Assert.True(OfficeVbaCompression.TryDecompress(streams["VBA/dir"], 20000, out byte[] directory, out _));
            int limit = directory.Length + 4096 + Encoding.ASCII.GetByteCount(project.GetModule("Second").Source) - 1;

            OfficeVbaInspection inspection = OfficeVbaProjectInspector.Inspect(streams, limit);

            Assert.Null(inspection.Limitation); Assert.Equal(2, inspection.Modules.Count);
            Assert.Null(inspection.Modules[0].Source); Assert.NotNull(inspection.Modules[0].Limitation);
            Assert.Null(inspection.Modules[1].Source); Assert.NotNull(inspection.Modules[1].Limitation);
        }

        [Fact]
        public void StreamInspectionObservesCallerCancellation() {
            Assert.Throws<OperationCanceledException>(() => OfficeVbaProjectInspector.Inspect(
                new Dictionary<string, byte[]>(), 10000, new CancellationToken(true)));
        }
    }
}
