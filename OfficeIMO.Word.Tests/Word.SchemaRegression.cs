using System;
using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Custom_list_presets_save_with_valid_numbering_metadata() {
        using var document = WordDocument.Create();
        foreach (WordListLevelKind kind in Enum.GetValues(typeof(WordListLevelKind))) {
            var list = document.AddCustomList();
            list.Numbering.AddLevel(new WordListLevel(kind));
            list.AddItem(kind.ToString());
        }
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Assert.Equal(Enum.GetNames(typeof(WordListLevelKind)), reopened.Paragraphs.Select(p => p.Text));
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void Document_settings_remain_valid_after_later_settings_are_present() {
        using var document = WordDocument.Create();
        document.AddParagraph("Body");
        document.Settings.UpdateFieldsOnOpen = true;
        document.Settings.MirrorMargins = true;
        document.Settings.GutterAtTop = true;
        document.Settings.DefaultTabStop = 720;
        document.Settings.CharacterSpacingControl = WordCharacterSpacing.DoNotCompress;
        document.Settings.TrackRevisions = true;
        document.Settings.TrackFormatting = false;
        document.Settings.TrackMoves = false;
        document.DifferentOddAndEvenPages = true;
        document.SetDocumentVariable("Qualification", "settings-order");
        document.Settings.ProtectionPassword = "schema-test-only";
        document.Settings.SetProtectionType(WordDocumentProtectionType.ReadOnly);
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Assert.True(reopened.DifferentOddAndEvenPages);
        Assert.True(reopened.Settings.UpdateFieldsOnOpen);
        Assert.True(reopened.Settings.MirrorMargins);
        Assert.True(reopened.Settings.GutterAtTop);
        Assert.Equal(720, reopened.Settings.DefaultTabStop);
        Assert.Equal(WordCharacterSpacing.DoNotCompress, reopened.Settings.CharacterSpacingControl);
        Assert.True(reopened.Settings.TrackRevisions);
        Assert.False(reopened.Settings.TrackFormatting);
        Assert.False(reopened.Settings.TrackMoves);
        Assert.Equal("settings-order", reopened.GetDocumentVariable("Qualification"));
        Assert.Equal(WordDocumentProtectionType.ReadOnly, reopened.Settings.ProtectionType);
        Assert.Empty(reopened.ValidateDocument());
    }
}
