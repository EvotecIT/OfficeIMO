namespace OfficeIMO.OneNote;

internal static class OneNoteFormatConstants {
    public const int RevisionStoreHeaderLength = 1024;
    public const int PackageStoreFixedPrefixLength = 72;

    public static readonly Guid SectionFileType = new Guid("7B5C52E4-D88C-4DA7-AEB1-5378D02996D3");
    public static readonly Guid TableOfContentsFileType = new Guid("43FF2FA1-EFD9-4C76-9EE2-10EA5722765F");
    public static readonly Guid RevisionStoreFormat = new Guid("109ADD3F-911B-49F5-A5D0-1791EDC8AED8");
    public static readonly Guid PackageStoreFormat = new Guid("638DE92F-A6D4-4BC1-9A36-B3FC2511A5B7");
    public static readonly Guid SectionCellSchema = new Guid("1F937CB4-B26F-445F-B9F8-17E20160E461");
    public static readonly Guid TableOfContentsCellSchema = new Guid("E4DBFD38-E5C7-408B-A8A1-0E7B421E1F5F");
}
