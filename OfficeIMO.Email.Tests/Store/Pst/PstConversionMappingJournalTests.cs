namespace OfficeIMO.Email.Store.Tests;

public sealed class PstConversionMappingJournalTests {
    [Fact]
    public void Verification_mappings_preserve_count_order_identifiers_and_source_flags() {
        string destination = Path.Combine(Path.GetTempPath(), "officeimo-mapping-" + Guid.NewGuid().ToString("N") + ".pst");
        using var journal = new PstConversionMappingJournal(destination);
        var sources = new[] {
            new EmailStoreItemReference("source-α", "folder-1", isAssociated: false, isOrphaned: false),
            new EmailStoreItemReference("source-2", "folder-β", isAssociated: true, isOrphaned: false),
            new EmailStoreItemReference("source-3", "folder-3", isAssociated: false, isOrphaned: true)
        };
        for (int index = 0; index < sources.Length; index++) {
            journal.Add(index + 1, sources[index], "destination-folder-" + index, "destination-item-" + index);
        }

        Assert.Equal(sources.Length, journal.Count);
        PstConversionItemMap[] mappings = journal.ReadAll().ToArray();
        Assert.Equal(sources.Length, mappings.Length);
        for (int index = 0; index < sources.Length; index++) {
            Assert.Equal(index + 1, mappings[index].Ordinal);
            Assert.Equal(sources[index].Id, mappings[index].Source.Id);
            Assert.Equal(sources[index].FolderId, mappings[index].Source.FolderId);
            Assert.Equal(sources[index].IsAssociated, mappings[index].Source.IsAssociated);
            Assert.Equal(sources[index].IsOrphaned, mappings[index].Source.IsOrphaned);
            Assert.Equal("destination-folder-" + index, mappings[index].DestinationFolderId);
            Assert.Equal("destination-item-" + index, mappings[index].DestinationItemId);
        }
    }
}
