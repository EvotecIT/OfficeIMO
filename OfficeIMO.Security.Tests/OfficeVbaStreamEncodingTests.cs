using OfficeIMO.Core.Internal;

namespace OfficeIMO.Security.Tests {

    public sealed class OfficeVbaStreamEncodingTests {
        [Theory]
        [InlineData("module")]
        [InlineData("projectName")]
        [InlineData("referenceName")]
        [InlineData("referenceId")]
        [InlineData("unsupported")]
        public void MalformedOrUnavailableMetadataReturnsAnExplicitLimitation(string field) {
            OfficeVbaInspection inspection = OfficeVbaProjectInspector.Inspect(ReadStreams(OfficeVbaMalformedTextFixtures.Create(field)), 20000);
            Assert.NotNull(inspection.Limitation);
            Assert.Empty(inspection.Modules); Assert.Empty(inspection.References);
        }

        [Fact]
        public void InvalidSourceEncodingKeepsItsModuleInventory() {
            OfficeVbaInspection inspection = OfficeVbaProjectInspector.Inspect(ReadStreams(OfficeVbaMalformedTextFixtures.Create("source")), 20000);
            Assert.Null(inspection.Limitation);
            OfficeVbaModuleInspection module = Assert.Single(inspection.Modules);
            Assert.Equal("Helpers", module.Name); Assert.Null(module.Source); Assert.NotNull(module.Limitation);
            Assert.Single(inspection.References);
        }

        private static IReadOnlyDictionary<string, byte[]> ReadStreams(byte[] bytes) {
            Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out _));
            return compound!.Streams;
        }
    }
}
