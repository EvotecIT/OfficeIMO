using System.Text;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Preserves the observed single-story Quill profile while replacing plain text. Native acceptance remains unqualified.</summary>
internal static class PublisherQuillTextWriter {
    internal static byte[]? TryRewrite(byte[] source, PublisherQuillStoryStyles binding,
        PublisherTextStory before, PublisherTextStory after, OfficeLegacyImportLimits limits, long maximumBytes, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (binding.StoryOrdinal != 0 || binding.NativeStart != 0 || binding.NullCharacters != 0 || binding.CarriageReturns != 1
            || binding.NativeLength != before.Text.Length || !before.HasFinalParagraphTerminator || !after.HasFinalParagraphTerminator
            || before.Text.Length < 2 || after.Text.Length < 2 || before.Text.Count(value => value == '\n') != 1
            || after.Text.Count(value => value == '\n') != 1
            || !after.Text.EndsWith("\n", StringComparison.Ordinal)) return null;
        var data = new PublisherBinaryData(source, "Quill/QuillSub/CONTENTS");
        var context = new PublisherParseContext(new PublisherReadOptions { Limits = limits }, token);
        IReadOnlyList<PublisherQuillChunk> chunks = PublisherQuillDirectory.Read(data, context);
        // This is one observed producer profile, not a generic relocation rule
        // for unknown Quill records or chained directories.
        string[] expected = { "TEXT", "STSH", "STSH", "STSH", "FDPP", "FDPC", "SYID", "SGP ", "INK ", "BTEP", "BTEC", "FONT", "STRS", "MCLD", "PL  " };
        if (chunks.Count != expected.Length || !chunks.Select(chunk => chunk.Tag).SequenceEqual(expected)
            || chunks.Where(chunk => chunk.Tag != "STSH").Any(chunk => chunk.Id != 0)
            || !chunks.Where(chunk => chunk.Tag == "STSH").Select(chunk => chunk.Id).SequenceEqual(new[] { 0, 1, 2 })
            || data.U16(8) != 4 || data.U16(10) != 7 || data.U16(12) != chunks.Count
            || data.U32(16) != 512 || data.U32(20) != source.Length || data.U16(26) != chunks.Count
            || data.U32(28) != uint.MaxValue) return null;
        PublisherQuillChunk text = chunks[0];
        if (text.Offset != 512 || text.Length != binding.NativeLength * 2 || binding.TextByteOffset != text.Offset
            || !data.Utf16(text.Offset, text.Length).Equals(before.Text.Replace('\n', '\r'), StringComparison.Ordinal)) return null;
        PublisherQuillChunk[] physical = chunks.OrderBy(chunk => chunk.Offset).ToArray();
        if (physical[0] != text || physical.Any(chunk => chunk.Length <= 0)) return null;
        for (int index = 1; index < physical.Length; index++)
            if (physical[index].Offset < physical[index - 1].End) return null;
        PublisherQuillChunk Find(string tag) => chunks.Single(chunk => chunk.Tag == tag);
        PublisherQuillChunk syid = Find("SYID"), strs = Find("STRS"), pc = Find("FDPC"), pp = Find("FDPP");
        if (syid.Length != 12 || data.U32(syid.Offset + 4) != 1 || data.U32(syid.Offset + 8) != before.Id
            || strs.Length != 30 || data.U32(strs.Offset) != 1 || data.U32(strs.Offset + 4) != 8
            || data.U32(strs.Offset + 12) != binding.NativeLength
            || pc.Length != 512 || pp.Length != 512 || pc.Offset % 512 != 0 || pp.Offset % 512 != 0) return null;
        if (!FormattingPage(pc, "FDPC") || !FormattingPage(pp, "FDPP")
            || !Index(Find("BTEC"), pc) || !Index(Find("BTEP"), pp)) return null;
        byte[] replacement = Encoding.Unicode.GetBytes(after.Text.Replace('\n', '\r'));
        int growth = Math.Max(0, replacement.Length - text.Length);
        int shift = checked((growth + 511) / 512 * 512);
        int length = checked(source.Length + shift);
        if (length > maximumBytes) throw new InvalidDataException("Publisher native output byte limit exceeded.");
        token.ThrowIfCancellationRequested();
        byte[] output = new byte[length];
        Buffer.BlockCopy(source, 0, output, 0, text.Offset);
        Buffer.BlockCopy(replacement, 0, output, text.Offset, replacement.Length);
        // Preserve opaque tail records verbatim. Whole-page movement preserves
        // FDPC/FDPP alignment; the newly freed text tail contains zeros.
        Buffer.BlockCopy(source, text.End, output, text.End + shift, source.Length - text.End);
        Write(20, (uint)length);
        for (int index = 0; index < chunks.Count; index++) {
            token.ThrowIfCancellationRequested();
            PublisherQuillChunk chunk = chunks[index];
            int descriptor = 32 + index * 24;
            Write(descriptor + 16, (uint)(chunk == text ? text.Offset : chunk.Offset + shift));
            if (chunk == text) Write(descriptor + 20, (uint)replacement.Length);
        }
        Write(strs.Offset + shift + 12, (uint)after.Text.Length);
        RemapPage(pc); RemapPage(pp);
        foreach (string tag in new[] { "BTEC", "BTEP" }) {
            PublisherQuillChunk index = Find(tag);
            Write(index.Offset + shift + 16, (uint)(text.Offset + replacement.Length));
            Write(index.Offset + shift + 20, checked(data.U32(index.Offset + 20) + (uint)shift));
        }
        token.ThrowIfCancellationRequested();
        return output;

        bool FormattingPage(PublisherQuillChunk page, string tag) {
            int count = data.U16(page.Offset);
            if (data.U16(page.Offset + 2) != 1 || count < 1 || count > 2 || tag == "FDPP" && count != 1) return false;
            PublisherQuillStyleUse[] uses = binding.Uses.Where(use => use.Tag == tag).ToArray();
            if (uses.Length != count) return false;
            for (int index = 0; index < count; index++) {
                int end = data.Offset(data.U32(page.Offset + 8 + index * 4));
                if (end != text.End && (tag != "FDPC" || end != text.End - 2)) return false;
                int opl = page.Offset + data.U16(page.Offset + 8 + count * 4 + index * 2);
                PublisherQuillStyleUse use = uses[index];
                if (use.ChunkId != 0 || use.RangeIndex != index || use.NativeStart + use.NativeLength != (end - text.Offset) / 2
                    || use.DirectDefinition.Offset != opl || use.DirectDefinition.Length != data.U32(opl)
                    || opl < page.Offset + 8 + count * 6 || use.DirectDefinition.Length < 4
                    || opl + (long)use.DirectDefinition.Length > page.End) return false;
            }
            return data.U32(page.Offset + 8 + (count - 1) * 4) == text.End;
        }
        bool Index(PublisherQuillChunk index, PublisherQuillChunk page) => index.Length == 24
            && data.U32(index.Offset) == 1 && data.U32(index.Offset + 4) == 4 && data.U32(index.Offset + 8) == 0
            && data.U32(index.Offset + 12) == 0 && data.U32(index.Offset + 16) == text.End
            && data.U32(index.Offset + 20) == page.Offset;
        void RemapPage(PublisherQuillChunk page) {
            int count = data.U16(page.Offset);
            for (int index = 0; index < count; index++) {
                int old = data.Offset(data.U32(page.Offset + 8 + index * 4));
                Write(page.Offset + shift + 8 + index * 4,
                    (uint)(text.Offset + replacement.Length - (old == text.End ? 0 : 2)));
            }
        }
        void Write(int offset, uint value) {
            for (int index = 0; index < 4; index++) output[offset + index] = (byte)(value >> (index * 8));
        }
    }
}
