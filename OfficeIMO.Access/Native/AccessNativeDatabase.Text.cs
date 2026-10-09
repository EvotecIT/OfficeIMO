using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        private static bool QualifiedJet3CodePage(int codePage) => codePage == 1250 || codePage == 1252 || codePage == 437 || codePage == 850 || codePage == 10000;
        private string NativeName(OfficeByteView bytes, ref int position) {
            if (!Layout.IsJet3) return Name(bytes, ref position);
            int length = Slice(bytes, position, 1)[0]; position = checked(position + 1);
            if (length == 0 || length > 64) throw new InvalidDataException("Native Access name length is invalid.");
            string name = OfficeLegacySingleByteEncoding.Decode(Slice(bytes, position, length), _document.CodePage!.Value); position = checked(position + length);
            try { return AccessNamedObject.ValidateName(name); }
            catch (ArgumentException exception) { throw new InvalidDataException("Native Access name is invalid.", exception); }
        }
        private object NativeText(AccessNativeColumn column, OfficeByteView bytes, CancellationToken cancellation) {
            if (!Layout.IsJet3) {
                string unicode = Text(bytes);
                return column.RedactConnection ? RedactConnection(unicode)! : unicode;
            }
            int codePage = column.CodePage == 0 ? _document.CodePage!.Value : column.CodePage;
            if (!QualifiedJet3CodePage(codePage)) {
                if (column.RedactConnection) return "[redacted: unsupported connection encoding]";
                return new AccessOpaqueValue(column.Type, bytes.ToArray(), "Jet3 text outside the qualified code-page subset is retained without character conversion.");
            }
            string text = OfficeLegacySingleByteEncoding.Decode(bytes, codePage, cancellation);
            return column.RedactConnection ? RedactConnection(text)! : text;
        }
    }
}
