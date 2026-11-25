namespace OfficeIMO.Scaling.Benchmarks;

internal static class ScalingLanes {
    public static IReadOnlyList<ScalingLane> Create() => [
        new OdtBuildLane(), new OdtReadLane(),
        new OdsBuildLane(), new OdsReadLane(),
        new OdpBuildLane(), new OdpReadLane(),
        new VisioBuildLane(), new VisioReadLane(),
        new MsgAttachmentWriteLane(), new MsgAttachmentReadLane(),
        new EmlAttachmentReadLane(), new TnefAttachmentReadLane(),
        new MsgRecipientWriteLane(),
        new RtfWriteLane(), new RtfParseLane(), new RtfCommentsToDocxLane(),
        new HtmlParagraphsToDocxLane(), new HtmlTableToDocxLane(), new MarkdownToDocxLane(),
        new ProjectBuildLane(), new ProjectReadLane(),
        new AccessBuildLane(), new AccessReadLane(),
        new OneNoteBuildLane(), new OneNoteReadLane(),
        new WordParagraphsLane(), new WordTableRowsLane(), new WordImagesLane(),
        new WordHyperlinksLane(), new WordBookmarksLane(), new WordCommentsLane(), new WordLoadLane(),
        new ExcelSheetsLane(), new ExcelNamedRangesLane(), new ExcelTablesLane(), new ExcelCommentsLane(),
    ];
}
