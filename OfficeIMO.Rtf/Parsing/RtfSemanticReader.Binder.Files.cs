using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static IReadOnlyList<RtfFileReference> ReadFileReferences(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            RtfGroup? fileTable = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "filetbl");
            if (fileTable == null) return Array.Empty<RtfFileReference>();

            var files = new List<RtfFileReference>();
            foreach (RtfGroup fileGroup in fileTable.Children.OfType<RtfGroup>().Where(group => group.Destination == "file")) {
                FileTableEntry entry = ReadFileReference(fileGroup, ansiCodePage, unicodeSkipCount);
                if (entry.Id.HasValue && !string.IsNullOrWhiteSpace(entry.Path)) {
                    files.Add(new RtfFileReference(entry.Id.Value, entry.Path) {
                        RelativePathStart = entry.RelativePathStart,
                        OperatingSystemNumber = entry.OperatingSystemNumber,
                        Sources = entry.Sources
                    });
                }
            }

            return files;
        }

        private static FileTableEntry ReadFileReference(RtfGroup fileGroup, int ansiCodePage, int unicodeSkipCount) {
            var entry = new FileTableEntry {
                Path = CollectPlainText(fileGroup, ansiCodePage, unicodeSkipCount).Trim()
            };

            foreach (RtfNode node in fileGroup.Children) {
                if (node is RtfControlWord control) {
                    ApplyFileReferenceControl(control, entry);
                }
            }

            return entry;
        }

        private static void ApplyFileReferenceControl(RtfControlWord control, FileTableEntry entry) {
            switch (control.Name) {
                case "fid":
                    entry.Id = control.Parameter;
                    break;
                case "frelative":
                    entry.RelativePathStart = control.Parameter;
                    break;
                case "fosnum":
                    entry.OperatingSystemNumber = control.Parameter;
                    break;
                case "fvalidmac":
                    entry.Sources |= RtfFileSource.Mac;
                    break;
                case "fvaliddos":
                    entry.Sources |= RtfFileSource.Dos;
                    break;
                case "fvalidntfs":
                    entry.Sources |= RtfFileSource.Ntfs;
                    break;
                case "fvalidhpfs":
                    entry.Sources |= RtfFileSource.Hpfs;
                    break;
                case "fnetwork":
                    entry.Sources |= RtfFileSource.Network;
                    break;
            }
        }

        private sealed class FileTableEntry {
            public int? Id { get; set; }
            public string Path { get; set; } = string.Empty;
            public int? RelativePathStart { get; set; }
            public int? OperatingSystemNumber { get; set; }
            public RtfFileSource Sources { get; set; }
        }
    }
}
