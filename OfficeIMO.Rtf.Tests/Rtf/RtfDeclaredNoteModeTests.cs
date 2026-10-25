using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfDeclaredNoteModeTests {
    [Theory]
    [InlineData("", null)]
    [InlineData(@"\fet0", RtfNoteMode.Footnotes)]
    [InlineData(@"\fet1", RtfNoteMode.Endnotes)]
    [InlineData(@"\fet2", RtfNoteMode.FootnotesAndEndnotes)]
    public void Declared_Mode_Is_Exposed_And_Independent_In_Clones(string controls, RtfNoteMode? declared) {
        RtfDocument source = RtfDocument.Read(RtfNativeNoteModeTests.Source(controls)).Document;
        Assert.Equal(declared, source.Settings.NoteMode);
        RtfDocument clone = source.Clone();
        Assert.Equal(declared, clone.Settings.NoteMode);
        clone.Settings.NoteMode = declared == RtfNoteMode.Footnotes ? RtfNoteMode.Endnotes : RtfNoteMode.Footnotes;
        Assert.Equal(declared, source.Settings.NoteMode);
        Assert.NotEqual(source.Settings.NoteMode, clone.Settings.NoteMode);
    }
}
