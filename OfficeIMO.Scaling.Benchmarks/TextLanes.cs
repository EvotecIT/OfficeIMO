using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Rtf;

namespace OfficeIMO.Scaling.Benchmarks;

internal sealed class RtfWriteLane : ScalingLane {
    public override string Name => "Rtf.Write";
    public override string Unit => "paragraph";

    public override object? Run() {
        var document = RtfDocument.Create();
        for (int i = 0; i < Units; i++) {
            document.AddParagraph($"Paragraph {i} with some ordinary text.");
        }

        return document.ToRtf();
    }

    public override void Validate(object? result) =>
        Require(RtfDocument.Parse((string)result!).Paragraphs.Count >= Units, "paragraph count");
}

internal sealed class RtfParseLane : ScalingLane {
    private string _rtf = string.Empty;
    public override string Name => "Rtf.Parse";
    public override string Unit => "paragraph";

    protected override void Prepare(int units) {
        var lane = new RtfWriteLane();
        lane.Setup(units);
        _rtf = (string)lane.Run()!;
    }

    public override object? Run() => RtfDocument.Parse(_rtf).Paragraphs.Count;
    public override void Validate(object? result) => Require((int)result! >= Units, "paragraph count");
}

internal sealed class RtfCommentsToDocxLane : ScalingLane {
    private string _rtf = string.Empty;
    public override string Name => "Rtf.CommentsToDocx";
    public override string Unit => "comment";

    protected override void Prepare(int units) {
        var document = RtfDocument.Create();
        for (int i = 0; i < units; i++) {
            RtfRun run = document.AddParagraph().AddText($"Commented paragraph {i}.");
            var note = new RtfNote(RtfNoteKind.Annotation) { Author = $"Author {i % 5}" };
            note.AddParagraph($"Comment {i}");
            run.Note = note;
        }

        _rtf = document.ToRtf();
    }

    public override object? Run() {
        using WordDocument document = RtfDocument.Parse(_rtf).ToWord();
        var stream = new MemoryStream();
        document.Save(stream);
        return stream.ToArray();
    }

    public override void Validate(object? result) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream((byte[])result!));
        Require(loaded.Comments.Count == Units, "comment count");
    }
}

internal abstract class ToDocxLane : ScalingLane {
    protected string Source { get; private set; } = string.Empty;
    protected abstract string BuildSource(int units);
    protected abstract WordDocument Convert();
    protected override void Prepare(int units) => Source = BuildSource(units);

    public override object? Run() {
        using WordDocument document = Convert();
        var stream = new MemoryStream();
        document.Save(stream);
        return stream.ToArray();
    }

    public override void Validate(object? result) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream((byte[])result!));
        Require(loaded.Paragraphs.Count >= Units, "paragraph count");
    }
}

internal sealed class HtmlParagraphsToDocxLane : ToDocxLane {
    public override string Name => "Html.ToDocx.Paragraphs";
    public override string Unit => "paragraph";

    protected override string BuildSource(int units) {
        var builder = new StringBuilder("<html><body>");
        for (int i = 0; i < units; i++) {
            builder.Append("<p>Paragraph ").Append(i).Append(" with <b>bold</b> and <i>italic</i> text.</p>");
        }

        return builder.Append("</body></html>").ToString();
    }

    protected override WordDocument Convert() => HtmlConversionDocument.Parse(Source).ToWord();
}

internal sealed class HtmlTableToDocxLane : ToDocxLane {
    public override string Name => "Html.ToDocx.TableRows";
    public override string Unit => "row";

    protected override string BuildSource(int units) {
        var builder = new StringBuilder("<html><body><table>");
        for (int i = 0; i < units; i++) {
            builder.Append("<tr><td>").Append(i).Append("</td><td>Name ").Append(i).Append("</td><td>Value</td></tr>");
        }

        return builder.Append("</table></body></html>").ToString();
    }

    protected override WordDocument Convert() => HtmlConversionDocument.Parse(Source).ToWord();

    public override void Validate(object? result) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream((byte[])result!));
        Require(loaded.Tables.Count == 1 && loaded.Tables[0].Rows.Count == Units, "table row count");
    }
}

internal sealed class MarkdownToDocxLane : ToDocxLane {
    public override string Name => "Markdown.ToDocx";
    public override string Unit => "paragraph";

    protected override string BuildSource(int units) {
        var builder = new StringBuilder();
        for (int i = 0; i < units; i++) {
            if (i % 10 == 0) {
                builder.Append("## Section ").Append(i).Append("\n\n");
            }

            builder.Append("Paragraph ").Append(i).Append(" with **bold** and *italic* text.\n\n");
        }

        return builder.ToString();
    }

    protected override WordDocument Convert() => MarkdownDocument.Parse(Source).ToWord();
}
