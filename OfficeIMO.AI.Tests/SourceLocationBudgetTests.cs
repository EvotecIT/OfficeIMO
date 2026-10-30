using OfficeIMO.AI;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class SourceLocationBudgetTests {
    [Theory]
    [InlineData(2, true)]
    [InlineData(6, false)]
    public void RepeatedTableRowLocationsRespectSnapshotMetadataBudget(int rowCount, bool accepted) {
        string path = new string('a', 20_000) + "/rows.csv";
        var source = new OfficeDocumentReadResult {
            Tables = [new ReaderTable {
                Title = "Rows", Columns = ["Value"],
                Rows = Enumerable.Range(1, rowCount).Select(index => (IReadOnlyList<string>)[index.ToString()]).ToArray(),
                Location = new ReaderLocation { Path = path }
            }]
        };
        var limits = new OfficeAiLimits { MaxEvidenceMetadataCharacters = 100_000 };

        if (accepted) {
            OfficeAiDocument snapshot = OfficeAiDocument.FromReadResult([1], source, limits: limits);
            Assert.Equal(rowCount + 1, snapshot.Evidence.Count);
            Assert.All(snapshot.Evidence, item => Assert.Equal(path, item.SourceLocation?.Path));
        } else {
            InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
                OfficeAiDocument.FromReadResult([1], source, limits: limits));
            Assert.Contains("metadata", error.Message, StringComparison.OrdinalIgnoreCase);
        }
    }
}
