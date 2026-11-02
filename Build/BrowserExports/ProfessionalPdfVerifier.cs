using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;

/// <summary>Checks real composed reports through the independent .NET reader, including page navigation and semantic ownership.</summary>
internal static class ProfessionalPdfVerifier {
    internal static object Verify(string path) {
        using var expected = JsonDocument.Parse(File.ReadAllText(path + ".json"));
        JsonElement contract = expected.RootElement;
        using var document = PdfDocument.Load(path);
        using var painted = PdfDocument.Load(path, new PdfLoadOptions { IncludeArtifactText = true });
        Require(!document.RepairReport.HasRepairs, "The report required repairs.");
        string[] pageTexts = document.Pages.Select(page => string.Concat(page.GetTextSpans().Select(span => span.Text))).ToArray();
        string text = string.Concat(pageTexts);
        string normalized = Regex.Replace(string.Join(" ", document.Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text)), @"\s+", " ");
        Require(!text.Contains("UNSELECTED-", StringComparison.Ordinal) && !text.Contains("INTERNAL-ONLY", StringComparison.Ordinal), "Excluded services/internal fields leaked.");
        string[][] wanted = contract.GetProperty("rows").EnumerateArray().Select(row => row.EnumerateArray().Select(cell => cell.GetString()!).ToArray()).ToArray();
        string[] ids = wanted.Select(row => row[1]).ToArray();
        Require(Regex.Matches(text, @"WEB-\d{6}").Select(match => match.Value).SequenceEqual(ids), "Selected service IDs were lost, duplicated or reordered.");
        Require(document.Outlines.Count == 1 && document.Outlines[0].Title == "Service health report" && document.Outlines[0].Level == 1, "Root bookmark differs.");
        PdfOutlineItem[] outlines = [document.Outlines[0], .. document.Outlines[0].Children];
        Require(outlines.Select(item => item.Title).SequenceEqual(new[] { "Service health report", "Inventory appendix", "Next actions" }) &&
            outlines.Skip(1).All(item => item.Level == 2), "Section bookmark hierarchy/order differs.");
        Require(outlines.All(item => item.PageNumber is int target && target >= 1 && target <= document.Pages.Count && item.DestinationTop.HasValue), "A bookmark has no resolved destination.");
        int inventory = outlines[1].PageNumber!.Value - 1, next = outlines[2].PageNumber!.Value - 1;
        Require(inventory > 0 && next > inventory && pageTexts[inventory].Contains("Inventory appendix", StringComparison.Ordinal) && pageTexts[next].Contains("Next actions", StringComparison.Ordinal), "Section destinations do not point at their content.");
        string[] headers = contract.GetProperty("headers").EnumerateArray().Select(value => value.GetString()!).ToArray();
        int rowNumber = 0;
        var pageProof = new List<object>();
        for (int index = 0; index < document.Pages.Count; index++) {
            PdfReadPage page = document.Pages[index]; (double width, double height) = page.GetPageSize();
            bool appendix = index >= inventory && index < next;
            Require(appendix ? width > height : height > width, "Section page orientation differs.");
            Require(page.GetTextSpans().Count > 0, "A transition emitted an empty page.");
            if (appendix) {
                var spans = page.GetTextSpans();
                foreach (string header in headers) Require(painted.Pages[index].GetTextSpans().Any(span => span.Text == header), "An appendix page lost its leaf heading: " + header);
                foreach (PdfTextSpan id in spans.Where(span => Regex.IsMatch(span.Text, @"^WEB-\d{6}$"))) {
                    Require(rowNumber < wanted.Length, "Extra selected row.");
                    string[] cells = spans.Where(span => Math.Abs(span.Y - id.Y) < .01).OrderBy(span => span.X).Select(span => span.Text).ToArray();
                    Require(cells.SequenceEqual(wanted[rowNumber]), "Appendix row differs at " + rowNumber + ": " + string.Join(" | ", cells));
                    rowNumber++;
                }
            }
            foreach (PdfTextSpan span in page.GetTextSpans())
                Require(span.X >= 35.5 && span.X + span.Advance <= width - 35.5 && span.Y >= 35.5 && span.Y <= height - 35.5, "Report text escaped the printable bounds.");
            pageProof.Add(new { page = index + 1, width, height, appendix });
        }
        Require(rowNumber == wanted.Length, "Not every selected row was independently read.");
        var tagged = document.TaggedContent ?? throw new InvalidDataException("Professional report has no logical structure.");
        foreach (string role in new[] { "Document", "H1", "H2", "L", "LI", "Lbl", "LBody", "Table", "TH", "TD", "Figure", "Caption", "Link" })
            Require(tagged.StructureTypes.Contains(role), "Missing semantic role: " + role);
        Require(tagged.StructureTypeCounts["LI"] == 3 && tagged.StructureTypeCounts["TD"] == wanted.Length * headers.Length, "List/table semantic counts differ.");
        var byId = tagged.StructureElements.ToDictionary(element => element.ObjectNumber);
        PdfStructureElementInfo[] captions = tagged.StructureElements.Where(element => element.StructureType == "Caption").ToArray();
        Require(captions.Length == 2 && captions.All(caption => caption.ParentObjectNumber is int parent && byId.TryGetValue(parent, out var owner) && owner.StructureType is "Figure" or "Table"), "Caption is not owned by its real figure/table.");
        foreach (PdfStructureElementInfo caption in captions) {
            PdfStructureElementInfo owner = byId[caption.ParentObjectNumber!.Value];
            var captionPages = caption.MarkedContentReferences.Select(reference => reference.PageObjectNumber).Distinct().ToArray();
            PdfStructureElementInfo content = owner.StructureType == "Figure" ? owner : byId[byId[owner.ChildElementObjectNumbers.First(id => byId[id].StructureType == "TR")].ChildElementObjectNumbers[0]];
            Require(captionPages.Length == 1 && content.MarkedContentReferences.Any(reference => reference.PageObjectNumber == captionPages[0]), "Caption was separated from its first real content.");
        }
        Require(tagged.StructureElements.Any(element => element.StructureType == "Figure" && element.AlternateText == "Selected service status: 84 healthy, 8 requiring attention and 3 offline; 95 services in total."), "The figure lost its selected-service alternative.");
        Require(document.Pages.SelectMany(page => page.GetLinkAnnotations()).Any(link => link.Uri == "https://example.com/operations/recovery"), "Finding link target was lost.");
        Require(!text.Contains("OFFICEIMO / OPERATIONS", StringComparison.Ordinal) && !text.Contains("Browser-local reference report", StringComparison.Ordinal), "Running decorations entered logical reading order.");
        for (int index = 0; index < painted.Pages.Count; index++) {
            var paintedSpans = painted.Pages[index].GetTextSpans();
            string all = Regex.Replace(string.Join(" ", paintedSpans.Select(span => span.Text)), @"\s+", " ");
            int section = index < inventory ? 1 : index < next ? 2 : 3, start = section == 1 ? 0 : section == 2 ? inventory : next;
            string title = outlines[section - 1].Title;
            Require(all.Contains("OFFICEIMO / OPERATIONS · " + title, StringComparison.Ordinal) && all.Contains("Section " + section + " / " + (index - start + 1), StringComparison.Ordinal), "Running section context differs.");
            // Form-XObject totals are returned as separate spans after page content, rather than in label order.
            Require(all.Contains("Page " + (index + 1) + " of ", StringComparison.Ordinal) && paintedSpans.Any(span =>
                span.Text == painted.Pages.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) && span.Y < 36 && span.X > painted.Pages[index].GetPageSize().Width / 2), "Global page number/total footer components differ.");
        }
        Require(document.Pages[0].GetTextSpans().Any(span => span.Text == "Service health report" && Math.Abs(span.FontSize - 22) < .01), "Named title style was not applied.");
        Require(normalized.Contains(contract.GetProperty("figureCaption").GetString()!, StringComparison.Ordinal) && normalized.Contains("Table 1. Selected service inventory; check times are UTC.", StringComparison.Ordinal), "A report caption was lost.");
        return new { file = Path.GetFileName(path), rows = rowNumber, pages = pageProof, outlines = outlines.Select(item => new { item.Title, item.Level, item.PageNumber, item.DestinationTop }),
            tagged.StructureTypeCounts, allSelectedCells = true, globalAndSectionPageNumbers = true, externalFindingLink = true, captionOwners = true, repairs = 0, bytes = new FileInfo(path).Length };
    }
    private static void Require(bool condition, string message) => WorkbookVerifier.Require(condition, "Professional PDF: " + message);
}
