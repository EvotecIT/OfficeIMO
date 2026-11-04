namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPublicationState {
    internal static PublisherPublicationState Create(double width, double height) {
        var page = new PublisherPageState(1, string.Empty, false, width, height, null, width / 2, height / 2,
            true, false, Array.Empty<uint>(), new PublisherFrameEditRange(double.MinValue, double.MinValue, double.MaxValue, double.MaxValue), false);
        var text = new PublisherTextState(new Dictionary<uint, PublisherTextStory>(), new Dictionary<uint, PublisherFrameState>(),
            new Dictionary<uint, IReadOnlyList<uint>>(), Array.Empty<uint>(), new Dictionary<uint, PublisherTableStory>(),
            new Dictionary<uint, PublisherWrapInfo>(), new Dictionary<uint, PublisherImportedWrapObject>());
        return new PublisherPublicationState(new[] { page }, new Dictionary<uint, uint?>(), new Dictionary<uint, PublisherPicture>(),
            new Dictionary<int, PublisherImage>(), text, new Dictionary<uint, uint?>(), nextObjectId: 2);
    }

    internal PublisherPublicationState AddPage(bool isMaster, string name, uint? masterId, out uint pageId) {
        if (masterId.HasValue) QualifiedMaster(masterId.Value);
        if (NextObjectId == 0) throw new InvalidOperationException("The publication object identifier space is exhausted.");
        pageId = NextObjectId;
        PublisherPageState basis = Pages.First(page => !page.IsMaster);
        var page = new PublisherPageState(pageId, name, isMaster, Width, Height, masterId, Width / 2, Height / 2,
            true, false, Array.Empty<uint>(), basis.FrameEditRange, false, isMaster ? PublisherMasterPageSide.Single : null);
        return WithPages(Pages.Concat(new[] { page }), pageId == uint.MaxValue ? 0 : pageId + 1);
    }

    internal PublisherPublicationState AddMasterPagePair(string name, out uint leftId, out uint rightId) {
        if (NextObjectId == 0 || NextObjectId == uint.MaxValue)
            throw new InvalidOperationException("Two publication page identifiers are required for a master pair.");
        leftId = NextObjectId;
        rightId = leftId + 1;
        PublisherPageState Member(uint id, PublisherMasterPageSide side, double origin) =>
            new(id, name, true, Width, Height, null, origin, Height / 2, true, false, Array.Empty<uint>(),
                new PublisherFrameEditRange(origin + int.MinValue / 12700D, Height / 2 + int.MinValue / 12700D,
                    origin + int.MaxValue / 12700D, Height / 2 + int.MaxValue / 12700D), false, side);
        return WithPages(Pages.Concat(new[] {
            Member(leftId, PublisherMasterPageSide.Left, 0), Member(rightId, PublisherMasterPageSide.Right, Width)
        }), rightId == uint.MaxValue ? 0 : rightId + 1);
    }

    internal PublisherPublicationState SetPageMaster(uint pageId, uint? masterId) {
        PublisherPageState? page = Pages.SingleOrDefault(item => item.Id == pageId);
        if (page == null) throw new ArgumentOutOfRangeException(nameof(pageId), "The publication has no such page.");
        if (page.IsMaster) throw new NotSupportedException("Master inheritance is not supported; assign a master to a document page.");
        if (!page.SupportsPlacementEditing || page.HasUnassessedBackground || !page.MasterAssignmentQualified)
            throw new NotSupportedException("The page's registration, background or native master relationship is not qualified for master assignment.");
        if (masterId.HasValue) QualifiedMaster(masterId.Value);
        return WithPages(Pages.Select(item => item.Id == pageId ? item.WithMaster(masterId) : item), NextObjectId);
    }

    private void QualifiedMaster(uint id) {
        PublisherPageState? master = Pages.SingleOrDefault(page => page.Id == id);
        if (master == null) throw new ArgumentOutOfRangeException("masterPageId", "The publication has no such master page.");
        if (!master.IsMaster || !master.SupportsPlacementEditing || master.HasUnassessedBackground || master.MasterPageId.GetValueOrDefault() != 0)
            throw new NotSupportedException("Only a qualified direct master definition without an unassessed background or inherited master can be assigned.");
    }

    private PublisherPublicationState WithPages(IEnumerable<PublisherPageState> pages, uint nextId) =>
        new(pages, ObjectOwners, Pictures, Images, Text, TableStoryReferences, _sourceImageIds, NextImageId, nextId, Shapes, Tables);
}
