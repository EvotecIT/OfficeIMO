namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteFileTable(StringBuilder builder, RtfDocument document, int unicodeSkipCount) {
        if (document.FileReferences.Count == 0) return;

        builder.Append(@"{\*\filetbl");
        foreach (RtfFileReference file in document.FileReferences.OrderBy(file => file.Id)) {
            builder.Append(@"{\file\fid");
            builder.Append(file.Id.ToString(CultureInfo.InvariantCulture));
            AppendOptionalTwips(builder, @"\frelative", file.RelativePathStart);
            AppendOptionalTwips(builder, @"\fosnum", file.OperatingSystemNumber);
            WriteFileSource(builder, file.Sources);
            builder.Append(' ');
            builder.Append(EscapeText(file.Path, unicodeSkipCount));
            builder.Append('}');
        }

        builder.Append('}');
    }

    private static void WriteFileSource(StringBuilder builder, RtfFileSource sources) {
        if ((sources & RtfFileSource.Mac) == RtfFileSource.Mac) {
            builder.Append(@"\fvalidmac");
        }

        if ((sources & RtfFileSource.Dos) == RtfFileSource.Dos) {
            builder.Append(@"\fvaliddos");
        }

        if ((sources & RtfFileSource.Ntfs) == RtfFileSource.Ntfs) {
            builder.Append(@"\fvalidntfs");
        }

        if ((sources & RtfFileSource.Hpfs) == RtfFileSource.Hpfs) {
            builder.Append(@"\fvalidhpfs");
        }

        if ((sources & RtfFileSource.Network) == RtfFileSource.Network) {
            builder.Append(@"\fnetwork");
        }
    }
}
