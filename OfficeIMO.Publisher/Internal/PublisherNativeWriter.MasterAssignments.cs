namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private readonly Dictionary<uint, uint?> _masterAssignments = new();

    private void CompareMasterAssignments() {
        foreach (var pair in _baseline.Pages.Zip(_current.Pages, (before, after) => (Before: before, After: after))) {
            _token.ThrowIfCancellationRequested();
            PublisherPageState before = pair.Before, after = pair.After;
            if (before.MasterPageId == after.MasterPageId) continue;
            if (before.IsMaster || !before.HasNativeOrigin || !before.SupportsPlacementEditing
                || before.HasUnassessedBackground || !before.MasterAssignmentQualified) {
                Reject("Contents/page/" + before.Id, "Only qualified existing document-page master assignments can change.");
                continue;
            }
            if (after.MasterPageId is uint id) {
                PublisherPageState? master = _baseline.Pages.SingleOrDefault(page => page.Id == id);
                if (master == null || !master.IsMaster || !master.HasNativeOrigin || !master.SupportsPlacementEditing
                    || master.HasUnassessedBackground || master.MasterPageId.GetValueOrDefault() != 0) {
                    Reject("Contents/page/" + before.Id, "The selected master is not a qualified existing direct definition.");
                    continue;
                }
            }
            _masterAssignments.Add(before.Id, after.MasterPageId);
        }
    }
}
