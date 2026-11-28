#region QuickStart
using OfficeIMO.OneNote;

using var section = OneNoteSection.Create("Planning.one");
section.Name = "Planning";
var page = new OneNotePage { Title = "Release checklist" };
var item = new OneNoteParagraph {
    List = new OneNoteListInfo { Ordered = false, Level = 0 }
};
item.Runs.Add(new OneNoteTextRun { Text = "Validate the packaged artifact" });
page.DirectContent.Add(item);
section.Pages.Add(page);

section.Save();
#endregion
