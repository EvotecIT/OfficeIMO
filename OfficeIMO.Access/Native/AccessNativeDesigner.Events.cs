using OfficeIMO.Core.Internal;
using System.Text;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    internal static partial class AccessNativeDesigner {
        /// <summary>Changes one qualified expanded-designer event record, copying all other records and terminators exactly.</summary>
        internal static byte[] ReplaceEvent(byte[] bytes, string? controlName, ushort code, uint id, string? expression,
            int maximumNodes, int maximumBytes, CancellationToken cancellation) {
            AccessDesignerNode root = Read(bytes, maximumNodes, cancellation)
                ?? throw new NotSupportedException("Event authoring requires the qualified expanded version-21 designer.");
            IEnumerable<AccessDesignerNode> Walk(AccessDesignerNode node) => new[] { node }.Concat(node.Children.SelectMany(Walk));
            AccessDesignerNode[] targets = controlName == null ? new[] { root }
                : Walk(root).Where(x => string.Equals(x.Name, controlName, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (targets.Length != 1) throw new ArgumentException("The designer control name is missing or ambiguous.", nameof(controlName));
            AccessDesignerNode target = targets[0];
            if (controlName != null && (target.NativeKind != 100 && target.NativeKind != 109 && target.NativeKind != 111
                || code == 86 && target.NativeKind == 100))
                throw new NotSupportedException("This control/event combination is outside the qualified label, text-box and combo-box layouts.");
            AccessDesignerProperty[] existing = target.Properties.Where(x => x.NativeCode == code).ToArray();
            if (existing.Length > 1 || existing.Any(x => x.NativeId != id || x.NativeType != 12 || x.NativeDefaultWidth != 4))
                throw new NotSupportedException("The event property is outside its qualified native record layout.");
            AccessDesignerProperty[] macro = target.Properties.Where(x => x.NativeCode == 491).ToArray();
            if (code == 126 && (macro.Length > 1 || macro.Any(x => x.NativeId != 275 || x.NativeType != 17 || x.NativeDefaultWidth != 4)))
                throw new NotSupportedException("The associated Click macro is outside its qualified carrier layout.");
            if (code != 126 && existing.Any(x => x.Value is string value && value.Equals("[Embedded Macro]", StringComparison.OrdinalIgnoreCase)))
                throw new NotSupportedException("The associated embedded event macro has no qualified replacement carrier.");
            byte[]? payload = string.IsNullOrEmpty(expression) ? null : new UnicodeEncoding(false, false, true).GetBytes(expression);
            using OfficeBoundedMemoryStream output = new OfficeBoundedMemoryStream(maximumBytes);
            using BinaryWriter writer = new BinaryWriter(output, Encoding.Unicode, true);
            output.Write(bytes, 0, 8); int offset = 8, count = 0;
            void Property(uint nativeId) {
                if (payload == null) return;
                writer.Write(nativeId); writer.Write(code); writer.Write(12U); writer.Write(4U);
                writer.Write((uint)payload.Length); writer.Write(payload);
            }
            void Node(AccessDesignerNode node) {
                cancellation.ThrowIfCancellationRequested(); count++;
                output.Write(bytes, offset, 2); offset += 2; bool changed = false;
                while (offset < bytes.Length) {
                    cancellation.ThrowIfCancellationRequested(); int start = offset;
                    uint record = U32(bytes, offset); offset += 4;
                    if (record >= 253 && record <= 255) {
                        if (ReferenceEquals(node, target) && !changed && payload != null) { Property(id); count++; }
                        output.Write(bytes, start, 4);
                        if (record == 254) Node(node.Children[0]);
                        if (record == 255) {
                            output.Write(bytes, offset, 2); int children = U16(bytes, offset); offset += 2;
                            for (int index = 0; index < children; index++) Node(node.Children[index]);
                        }
                        return;
                    }
                    ushort propertyCode = U16(bytes, offset); uint size = U32(bytes, offset + 10); offset += checked(14 + (int)size); count++;
                    if (ReferenceEquals(node, target) && propertyCode == code) { Property(record); changed = true; }
                    else if (ReferenceEquals(node, target) && code == 126 && propertyCode == 491) {
                        // Access prefers this active carrier over the text binding. Replacing
                        // Click explicitly replaces its associated macro, including unknown actions.
                    }
                    else output.Write(bytes, start, offset - start);
                }
                throw new NotSupportedException("Event authoring requires explicit native designer terminators.");
            }
            Node(root);
            if (count > maximumNodes || offset != bytes.Length) throw new InvalidDataException("The edited designer exceeds its record boundary or limit.");
            writer.Flush(); return output.ToArray();
        }
    }
}
