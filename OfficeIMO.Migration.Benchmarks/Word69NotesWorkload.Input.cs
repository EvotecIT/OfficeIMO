using System.Globalization;
using System.Text;
using static OfficeIMO.Migration.Benchmarks.Word69BenchmarkInput;

namespace OfficeIMO.Migration.Benchmarks;

internal sealed partial class Word69NotesWorkload {
    private static byte[] CreateInput(int units, ushort fibVersion) {
        if (units < 2 || units > 1000 || units % 2 != 0) throw new ArgumentOutOfRangeException(nameof(units));
        if (fibVersion is not (101 or 104)) throw new ArgumentOutOfRangeException(nameof(fibVersion));
        var body = new StringBuilder();
        var footnotes = new StringBuilder();
        var endnotes = new StringBuilder();
        var footnoteReferences = new List<int>();
        var endnoteReferences = new List<int>();
        var footnotePositions = new List<int>();
        var endnotePositions = new List<int>();
        for (int index = 0; index < units; index++) {
            bool endnote = IsEndnote(index);
            StringBuilder story = endnote ? endnotes : footnotes;
            List<int> references = endnote ? endnoteReferences : footnoteReferences;
            List<int> positions = endnote ? endnotePositions : footnotePositions;
            body.Append(BodyBefore(index));
            references.Add(body.Length);
            body.Append('\u0002').Append(BodyAfter(index)).Append('\r');
            positions.Add(story.Length);
            story.Append('\u0002').Append(NoteFirst(index)).Append('\r').Append(NoteSecond(index)).Append('\r');
        }
        footnotePositions.Add(footnotes.Length);
        endnotePositions.Add(endnotes.Length);

        const int textStart = 0x300;
        byte[] text = Encoding.ASCII.GetBytes(body.ToString() + footnotes + endnotes);
        int textEnd = checked(textStart + text.Length);
        int footnoteReferenceBytes = checked(4 + footnoteReferences.Count * 6);
        int endnoteReferenceBytes = checked(4 + endnoteReferences.Count * 6);
        int footnotePositionBytes = checked(footnotePositions.Count * 4);
        int endnotePositionBytes = checked(endnotePositions.Count * 4);
        const int documentPropertyBytes = 84;
        var main = new byte[checked(textEnd + footnoteReferenceBytes + footnotePositionBytes
            + endnoteReferenceBytes + endnotePositionBytes + documentPropertyBytes)];
        WriteWord(main, 0, 0xA5DC);
        WriteWord(main, 2, fibVersion);
        WriteWord(main, 6, 0x0409);
        WriteInt(main, 0x18, textStart);
        WriteInt(main, 0x1C, textEnd);
        WriteInt(main, 0x20, main.Length);
        WriteInt(main, 0x34, body.Length);
        WriteInt(main, 0x38, footnotes.Length);
        WriteInt(main, 0x48, endnotes.Length);
        text.CopyTo(main, textStart);

        int next = textEnd;
        WriteReferences(0x68, footnoteReferences);
        WritePositions(0x70, footnotePositions);
        WriteReferences(0x1D2, endnoteReferences);
        WritePositions(0x1DA, endnotePositions);
        WriteInt(main, 0x150, next);
        WriteInt(main, 0x154, documentPropertyBytes);
        WriteWord(main, next, 1 << 5); // Footnotes at page bottom.
        WriteWord(main, next + 2, 1 << 2); // Continuous Arabic numbering from one.
        WriteWord(main, next + 0x34, 1 << 2);
        WriteWord(main, next + 0x36, 3); // Endnotes at document end; Arabic formats.
        return WrapCompound(main);

        void WriteReferences(int directory, IReadOnlyList<int> references) {
            int length = checked(4 + references.Count * 6);
            WriteInt(main, directory, next);
            WriteInt(main, directory + 4, length);
            for (int index = 0; index < references.Count; index++) WriteInt(main, next + index * 4, references[index]);
            WriteInt(main, next + references.Count * 4, body.Length);
            int descriptors = next + (references.Count + 1) * 4;
            for (int index = 0; index < references.Count; index++) WriteWord(main, descriptors + index * 2, 1);
            next += length;
        }

        void WritePositions(int directory, IReadOnlyList<int> positions) {
            int length = checked(positions.Count * 4);
            WriteInt(main, directory, next);
            WriteInt(main, directory + 4, length);
            for (int index = 0; index < positions.Count; index++) WriteInt(main, next + index * 4, positions[index]);
            next += length;
        }
    }

    private static bool IsEndnote(int index) => index % 2 != 0;
    private static string Number(int index) => index.ToString("D4", CultureInfo.InvariantCulture);
    private static string BodyBefore(int index) => "Body " + Number(index) + " before ";
    private static string BodyAfter(int index) => " after " + Number(index);
    private static string NoteFirst(int index) => " " + (IsEndnote(index) ? "Endnote " : "Footnote ") + Number(index);
    private static string NoteSecond(int index) => "Continuation " + (IsEndnote(index) ? "endnote " : "footnote ") + Number(index);
}
