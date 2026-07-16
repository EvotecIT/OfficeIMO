using OfficeIMO.OneNote.Html;
using OfficeIMO.OneNote.Markdown;
using OfficeIMO.OneNote.Pdf;
using OfficeIMO.Reader.OneNote;

namespace OfficeIMO.Examples.OneNote;

internal static class OfflineOneNoteExample {
    internal static void Example(string folderPath) {
        var section = new global::OfficeIMO.OneNote.OneNoteSection { Name = "Offline planning" };
        var page = new global::OfficeIMO.OneNote.OneNotePage { Title = "Release checklist" };

        var introduction = new global::OfficeIMO.OneNote.OneNoteParagraph();
        var bold = new global::OfficeIMO.OneNote.OneNoteTextRun { Text = "OfficeIMO OneNote" };
        bold.Style.Bold = true;
        introduction.Runs.Add(bold);
        introduction.Runs.Add(new global::OfficeIMO.OneNote.OneNoteTextRun { Text = " works entirely offline." });
        page.DirectContent.Add(introduction);

        var item = new global::OfficeIMO.OneNote.OneNoteParagraph {
            List = new global::OfficeIMO.OneNote.OneNoteListInfo { Ordered = false, Level = 0 }
        };
        item.Runs.Add(new global::OfficeIMO.OneNote.OneNoteTextRun { Text = "Validate the packed artifact" });
        item.Tags.Add(new global::OfficeIMO.OneNote.OneNoteTag {
            Label = "To Do",
            ActionItemType = 0,
            IsCheckable = true
        });
        page.DirectContent.Add(item);
        section.Pages.Add(page);

        string sectionPath = Path.Combine(folderPath, "OfficeIMO-OneNote.one");
        section.Save(sectionPath);
        global::OfficeIMO.OneNote.OneNoteSection reopened = global::OfficeIMO.OneNote.OneNoteSectionReader.Read(sectionPath);

        File.WriteAllText(Path.Combine(folderPath, "OfficeIMO-OneNote.md"), reopened.ToMarkdown());
        reopened.SaveAsHtml(Path.Combine(folderPath, "OfficeIMO-OneNote.html"));
        reopened.SaveAsPdf(Path.Combine(folderPath, "OfficeIMO-OneNote.pdf"));

        var notebook = new global::OfficeIMO.OneNote.OneNoteNotebook { Name = "OfficeIMO Offline Notebook" };
        notebook.Sections.Add(reopened);
        string packagePath = Path.Combine(folderPath, "OfficeIMO-OneNote.onepkg");
        global::OfficeIMO.OneNote.OneNotePackageWriter.Write(notebook, packagePath);

        global::OfficeIMO.Reader.OfficeDocumentReader readerFacade = new global::OfficeIMO.Reader.OfficeDocumentReaderBuilder()
            .AddOneNoteHandler()
            .Build();
        global::OfficeIMO.Reader.OfficeDocumentReadResult reader = readerFacade.ReadDocument(sectionPath);
        File.WriteAllLines(Path.Combine(folderPath, "OfficeIMO-OneNote-reader.txt"), reader.Chunks.Select(chunk => chunk.Markdown ?? chunk.Text));

        Console.WriteLine("    OneNote: wrote .one, .onepkg, Markdown, HTML, PDF, and Reader projection artifacts.");
    }
}
