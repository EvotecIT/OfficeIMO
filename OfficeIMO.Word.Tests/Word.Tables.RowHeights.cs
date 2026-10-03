using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Table_row_height_constraints_save_reopen_and_switch_without_losing_other_row_properties() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordTableRow row = document.AddTable(1, 1).Rows[0];
            row.AllowRowToBreakAcrossPages = false;
            row.Height = 200;
            row.MinimumHeight = null;
            Assert.Equal(200, row.Height);
            Assert.Null(row.MinimumHeight);
            row.MinimumHeight = 400;
            Assert.Equal(400, row.Height);
            Assert.Throws<ArgumentOutOfRangeException>(() => row.MinimumHeight = -1);
            Assert.Throws<ArgumentOutOfRangeException>(() => row.Height = -1);
            Assert.Equal(400, row.MinimumHeight);
            document.Save();
        }
        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        WordTableRow saved = reopened.Tables[0].Rows[0];
        Assert.Equal(400, saved.MinimumHeight);
        Assert.False(saved.AllowRowToBreakAcrossPages);
        Assert.Equal(HeightRuleValues.AtLeast, saved._tableRow.TableRowProperties!.GetFirstChild<TableRowHeight>()!.HeightType!.Value);
        saved.Height = 300;
        Assert.Null(saved.MinimumHeight);
        Assert.Equal(HeightRuleValues.Exact, saved._tableRow.TableRowProperties.GetFirstChild<TableRowHeight>()!.HeightType!.Value);
        saved.MinimumHeight = 500;
        saved.MinimumHeight = null;
        Assert.Null(saved.Height);
        Assert.False(saved.AllowRowToBreakAcrossPages);
    }
}
