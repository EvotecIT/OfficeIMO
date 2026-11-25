using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Word;

namespace OfficeIMO.Scaling.Benchmarks;

internal abstract class WordEditLane : ScalingLane {
    protected abstract void AddUnit(WordDocument document, int index);
    protected virtual void Begin(WordDocument document) {
    }

    public override object? Run() {
        var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            Begin(document);
            for (int i = 0; i < Units; i++) {
                AddUnit(document, i);
            }

            document.Save();
        }

        return stream.ToArray();
    }

    protected abstract int Count(WordDocument loaded);

    public override void Validate(object? result) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream((byte[])result!));
        Require(Count(loaded) >= Units, "unit count");
    }
}

internal sealed class WordParagraphsLane : WordEditLane {
    public override string Name => "Word.AddParagraph";
    public override string Unit => "paragraph";
    protected override void AddUnit(WordDocument document, int index) => document.AddParagraph($"Paragraph {index} with ordinary text.");
    protected override int Count(WordDocument loaded) => loaded.Paragraphs.Count;
}

internal sealed class WordTableRowsLane : WordEditLane {
    private WordTable _table = null!;
    public override string Name => "Word.Table.AddRow";
    public override string Unit => "row";
    protected override void Begin(WordDocument document) => _table = document.AddTable(1, 3);

    protected override void AddUnit(WordDocument document, int index) {
        WordTableRow row = _table.AddRow(3);
        row.Cells[0].Paragraphs[0].Text = index.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    protected override int Count(WordDocument loaded) => loaded.Tables[0].Rows.Count - 1;
}

internal sealed class WordImagesLane : WordEditLane {
    // 1x1 transparent PNG.
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
    public override string Name => "Word.AddImage";
    public override string Unit => "image";

    protected override void AddUnit(WordDocument document, int index) =>
        document.AddImage(new MemoryStream(Png), OfficeLength.Pixels(20), OfficeLength.Pixels(20));

    protected override int Count(WordDocument loaded) => loaded.Images.Count;
}

internal sealed class WordHyperlinksLane : WordEditLane {
    public override string Name => "Word.AddHyperLink";
    public override string Unit => "hyperlink";
    protected override void AddUnit(WordDocument document, int index) => document.AddHyperLink($"Link {index}", new Uri($"https://example.com/{index}"));
    protected override int Count(WordDocument loaded) => loaded.HyperLinks.Count;
}

internal sealed class WordBookmarksLane : WordEditLane {
    public override string Name => "Word.AddBookmark";
    public override string Unit => "bookmark";
    protected override void AddUnit(WordDocument document, int index) => document.AddParagraph($"Bookmarked {index}").AddBookmark($"bm{index}");
    protected override int Count(WordDocument loaded) => loaded.Bookmarks.Count;
}

internal sealed class WordCommentsLane : WordEditLane {
    public override string Name => "Word.AddComment";
    public override string Unit => "comment";
    protected override void AddUnit(WordDocument document, int index) => document.AddParagraph($"Commented {index}").AddComment("Author", "AU", $"Comment {index}");
    protected override int Count(WordDocument loaded) => loaded.Comments.Count;
}

internal sealed class WordLoadLane : ScalingLane {
    private byte[] _bytes = [];
    public override string Name => "Word.Load";
    public override string Unit => "paragraph";

    protected override void Prepare(int units) {
        var lane = new WordParagraphsLane();
        lane.Setup(units);
        _bytes = (byte[])lane.Run()!;
    }

    public override object? Run() {
        using WordDocument loaded = WordDocument.Load(new MemoryStream(_bytes, writable: false));
        return loaded.Paragraphs.Count;
    }

    public override void Validate(object? result) => Require((int)result! >= Units, "paragraph count");
}

internal abstract class ExcelEditLane : ScalingLane {
    protected abstract void Edit(ExcelDocument document);
    protected abstract int Count(ExcelDocument loaded);

    public override object? Run() {
        var stream = new MemoryStream();
        using (ExcelDocument document = ExcelDocument.Create(stream)) {
            Edit(document);
            document.Save();
        }

        return stream.ToArray();
    }

    public override void Validate(object? result) {
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream((byte[])result!));
        Require(Count(loaded) >= Units, "unit count");
    }
}

internal sealed class ExcelSheetsLane : ExcelEditLane {
    public override string Name => "Excel.AddSheet";
    public override string Unit => "sheet";

    protected override void Edit(ExcelDocument document) {
        for (int i = 0; i < Units; i++) {
            document.AddSheet($"Sheet{i}").Cell(1, 1).SetValue(i);
        }
    }

    protected override int Count(ExcelDocument loaded) => loaded.Sheets.Count;
}

internal sealed class ExcelNamedRangesLane : ExcelEditLane {
    public override string Name => "Excel.SetNamedRange";
    public override string Unit => "name";

    protected override void Edit(ExcelDocument document) {
        ExcelSheet sheet = document.AddSheet("Data");
        for (int i = 0; i < Units; i++) {
            sheet.Cell(i + 1, 1).SetValue(i);
            document.SetNamedRange($"Name{i}", $"'Data'!$A${i + 1}", scope: null, save: false);
        }
    }

    protected override int Count(ExcelDocument loaded) => loaded.GetAllNamedRanges().Count;
}

internal sealed class ExcelTablesLane : ExcelEditLane {
    public override string Name => "Excel.AddTable";
    public override string Unit => "table";

    protected override void Edit(ExcelDocument document) {
        ExcelSheet sheet = document.AddSheet("Data");
        for (int i = 0; i < Units; i++) {
            int top = 3 * i + 1;
            sheet.Cell(top, 1).SetValue("Key");
            sheet.Cell(top, 2).SetValue("Value");
            sheet.Cell(top + 1, 1).SetValue(i);
            sheet.Cell(top + 1, 2).SetValue("v");
            sheet.AddTable($"A{top}:B{top + 1}", true, $"Table{i}");
        }
    }

    protected override int Count(ExcelDocument loaded) => loaded.Sheets[0].GetTables().Count;
}

internal sealed class ExcelCommentsLane : ExcelEditLane {
    public override string Name => "Excel.SetComment";
    public override string Unit => "comment";

    protected override void Edit(ExcelDocument document) {
        ExcelSheet sheet = document.AddSheet("Data");
        for (int i = 0; i < Units; i++) {
            sheet.Cell(i + 1, 1).SetValue(i);
            sheet.SetComment(i + 1, 1, $"Comment {i}");
        }
    }

    protected override int Count(ExcelDocument loaded) => loaded.Sheets[0].GetComments().Count;
}
