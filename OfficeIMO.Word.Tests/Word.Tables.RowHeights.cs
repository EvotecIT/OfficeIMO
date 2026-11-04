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
            row.Height = global::OfficeIMO.OfficeLength.Points(10D);
            row.MinimumHeight= null;
            Assert.Equal(200, checked((int?)(row.Height?.ToTwips())));
            Assert.Null(checked((int?)(row.MinimumHeight?.ToTwips())));
            row.MinimumHeight = global::OfficeIMO.OfficeLength.Points(20D);
            Assert.Equal(400, checked((int?)(row.Height?.ToTwips())));
            Assert.Throws<ArgumentOutOfRangeException>(() => row.MinimumHeight = global::OfficeIMO.OfficeLength.Points((-1) / 20D));
            Assert.Throws<ArgumentOutOfRangeException>(() => row.Height = global::OfficeIMO.OfficeLength.Points((-1) / 20D));
            Assert.Equal(400, checked((int?)(row.MinimumHeight?.ToTwips())));
            document.Save();
        }
        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        WordTableRow saved = reopened.Tables[0].Rows[0];
        Assert.Equal(400, checked((int?)(saved.MinimumHeight?.ToTwips())));
        Assert.False(saved.AllowRowToBreakAcrossPages);
        Assert.Equal(HeightRuleValues.AtLeast, saved._tableRow.TableRowProperties!.GetFirstChild<TableRowHeight>()!.HeightType!.Value);
        saved.Height = global::OfficeIMO.OfficeLength.Points(15D);
        Assert.Null(checked((int?)(saved.MinimumHeight?.ToTwips())));
        Assert.Equal(HeightRuleValues.Exact, saved._tableRow.TableRowProperties.GetFirstChild<TableRowHeight>()!.HeightType!.Value);
        saved.MinimumHeight = global::OfficeIMO.OfficeLength.Points(25D);
        saved.MinimumHeight= null;
        Assert.Null(checked((int?)(saved.Height?.ToTwips())));
        Assert.False(saved.AllowRowToBreakAcrossPages);
    }
}
