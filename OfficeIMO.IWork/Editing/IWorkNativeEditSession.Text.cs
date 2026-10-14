using System.Text;
using System.Threading;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

public sealed partial class IWorkNativeEditSession {
    private static readonly UTF8Encoding EditUtf8 = new(false, true);

    private static void ValidateParagraphReplacementText(string text) {
        if (!IWorkXmlText.IsRepresentable(text) || text.Any(character => char.IsControl(character) && character != '\t'
            || character is '\u2028' or '\u2029' or '\ufffb' or '\ufffc'))
            throw new ArgumentException("Replacement text must be one valid paragraph without inline object markers or line breaks.", nameof(text));
    }

    private void ReplaceNativeParagraphText(IWorkSourceDocument view, IWorkTextContent content, int paragraph,
        string text, bool allowDecimalLists, CancellationToken cancellationToken) {
        if (paragraph > content.Paragraphs.Count) throw new ArgumentOutOfRangeException(nameof(paragraph));
        var selected = content.Paragraphs[paragraph - 1];
        if (selected.ListMarkerKind != IWorkListMarkerKind.None
                && (!allowDecimalLists || !IWorkDecimalListLayout.TryResolve(selected, out _, out _))
            || selected.Runs.Count > 1
            || selected.Runs.Any(run => run.Hyperlink != null || run.InlineObject != null))
            throw new NotSupportedException("The selected paragraph is outside the supported text-edit profile.");
        if (text.Length != selected.Text.Length)
            throw new NotSupportedException("Replacement text must retain the original UTF-16 length.");

        var identity = content.SourceIdentity ?? throw new NotSupportedException("A native text-storage identity is required.");
        var record = view.Index.Find(identity.RecordIdentifier);
        if (record == null || record.MessageType != 2001 || record.PayloadIndex != 0 || !SameIdentity(new(record), identity))
            throw new ArgumentException("The identity must identify a primary text storage.", nameof(content));
        var message = view.Index.Message(record);
        if (message.FieldCount(3) != 1 || message.HasUnexpectedWireKind(3, IWorkWireKind.Bytes))
            throw new NotSupportedException("Paragraph editing requires exactly one native UTF-8 text field.");
        byte[] sourceBytes = message.GetBytes(3)!;
        string sourceText = EditUtf8.GetString(sourceBytes);
        var spans = IWorkTextReader.ParagraphSpans(sourceText).ToArray();
        if (spans.Length != content.Paragraphs.Count)
            throw new NotSupportedException("Native paragraph offsets must match the selected text projection.");
        var span = spans[paragraph - 1];
        string original = sourceText.Substring(span.Start, span.End - span.Start);
        if (original != selected.Text)
            throw new NotSupportedException("The paragraph must retain an unnormalized native text range.");
        byte[] replacement = EditUtf8.GetBytes(text);
        if (replacement.Length != EditUtf8.GetByteCount(original))
            throw new NotSupportedException("Replacement text must retain the original UTF-8 byte length.");
        int offset = EditUtf8.GetByteCount(sourceText.Substring(0, span.Start));
        byte[] current = _payloads.TryGetValue(record, out byte[]? pending) ? pending : record.Payload;
        byte[] updated = (byte[])current.Clone();
        IWorkPreservingWireEdit.ReplaceBytesRange(updated, new[] { (3, -1) }, offset, replacement, view.Options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (updated.SequenceEqual(record.Payload)) _payloads.Remove(record); else _payloads[record] = updated;
    }
}
