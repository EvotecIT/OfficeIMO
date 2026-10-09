using System.Text;

namespace OfficeIMO.Visio;

internal sealed partial class VisioLegacyBinaryCodec {
    private string ReadName(VisioBinaryContainer.Node node, int relativeOffset) {
        int offset = node.Shift + relativeOffset;
        VisioBinaryData.Require(node.Data, offset, 2);
        int end = offset;
        while (end < node.Data.Length - 1 && VisioBinaryData.U16(node.Data, end) != 0) end += 2;
        if (end >= node.Data.Length - 1) throw new InvalidDataException("Binary Visio name is unterminated.");
        return Text(new UnicodeEncoding(false, false, true).GetString(node.Data, offset, end - offset));
    }
}
