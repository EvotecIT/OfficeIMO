namespace OfficeIMO.IWork;

public sealed partial class IWorkPagesProjection {
    private IEnumerable<IWorkPagesSection> ReconstructedSections(int? count) => Sections.Take(count ?? Sections.Count);

    private IEnumerable<IWorkObjectIdentity?> ReconstructedUnits(int? sectionCount) {
        yield return SourceIdentity;
        if (Body.IsTextComplete || Body.Paragraphs.Count > 0) yield return Body.SourceIdentity;
        foreach (IWorkTextContent content in ReconstructedSections(sectionCount)
                     .SelectMany(section => section.HeaderContents.Concat(section.FooterContents)))
            yield return content.SourceIdentity;
        foreach (IWorkTextBox text in TextBoxObjects.Where(text => text.Content.IsTextComplete || text.Content.Paragraphs.Count > 0))
            yield return text.Content.SourceIdentity;
        foreach (IWorkImageAsset image in Images) yield return image.SourceIdentity;
        foreach (IWorkTable table in Tables.Where(table => table.RowCount > 0 && table.ColumnCount > 0)) {
            yield return table.SourceIdentity;
            foreach (IWorkTableCell cell in table.Cells) yield return cell.RichText?.SourceIdentity;
        }
    }

    private IEnumerable<IWorkObjectIdentity?> OmittedUnits(int? sectionCount) {
        if (!Body.IsTextComplete && Body.Paragraphs.Count == 0) yield return Body.SourceIdentity;
        foreach (IWorkObjectIdentity identity in OmittedSourceUnits) yield return identity;
        foreach (IWorkTextContent content in Sections.Skip(sectionCount ?? Sections.Count)
                     .SelectMany(section => section.HeaderContents.Concat(section.FooterContents)))
            yield return content.SourceIdentity;
        foreach (IWorkObjectIdentity identity in Tables.SelectMany(table => table.OmittedTextUnits)) yield return identity;
        foreach (IWorkTable table in Tables.Where(table => table.RowCount == 0 || table.ColumnCount == 0)) {
            yield return table.SourceIdentity;
            foreach (IWorkTableCell cell in table.Cells) yield return cell.RichText?.SourceIdentity;
        }
    }
}

public sealed partial class IWorkNumbersProjection {
    private IEnumerable<IWorkObjectIdentity?> OmittedUnits() {
        foreach (IWorkObjectIdentity identity in OmittedSourceUnits) yield return identity;
        foreach (IWorkObjectIdentity identity in Sheets.SelectMany(sheet => sheet.Tables)
                     .SelectMany(table => table.OmittedTextUnits)) yield return identity;
    }

    private IEnumerable<IWorkObjectIdentity?> ReconstructedUnits() {
        yield return SourceIdentity;
        foreach (IWorkNumbersSheet sheet in Sheets) {
            yield return sheet.SourceIdentity;
            foreach (IWorkNumbersDrawable drawable in sheet.Drawables) yield return drawable.SourceIdentity;
            foreach (IWorkTableCell cell in sheet.Tables.SelectMany(table => table.Cells))
                yield return cell.RichText?.SourceIdentity;
        }
    }
}

public sealed partial class IWorkKeynoteProjection {
    private IEnumerable<IWorkObjectIdentity?> ReconstructedUnits() {
        yield return SourceIdentity;
        foreach (IWorkKeynoteSlide slide in Slides) {
            yield return slide.SourceIdentity;
            if (slide.TitleBox?.Content is { } title && (title.IsTextComplete || title.Paragraphs.Count > 0))
                yield return title.SourceIdentity;
            if (slide.PresenterNoteContent.Paragraphs.Count > 0) yield return slide.PresenterNoteContent.SourceIdentity;
            foreach (IWorkTextBox text in slide.TextBoxes.Where(text => text.Content.IsTextComplete || text.Content.Paragraphs.Count > 0))
                yield return text.Content.SourceIdentity;
            foreach (IWorkImageAsset image in slide.Images) yield return image.SourceIdentity;
            foreach (IWorkTable table in slide.Tables.Where(table => table.RowCount > 0 && table.ColumnCount > 0)) {
                yield return table.SourceIdentity;
                foreach (IWorkTableCell cell in table.Cells) yield return cell.RichText?.SourceIdentity;
            }
        }
    }

    private IEnumerable<IWorkObjectIdentity?> OmittedUnits() {
        foreach (IWorkObjectIdentity identity in OmittedSourceUnits) yield return identity;
        foreach (IWorkKeynoteSlide slide in Slides)
            if (!slide.PresenterNoteContent.IsTextComplete && slide.PresenterNoteContent.Paragraphs.Count == 0)
                yield return slide.PresenterNoteContent.SourceIdentity;
        foreach (IWorkObjectIdentity identity in Slides.SelectMany(slide => slide.Tables)
                     .SelectMany(table => table.OmittedTextUnits)) yield return identity;
        foreach (IWorkTable table in Slides.SelectMany(slide => slide.Tables)
                     .Where(table => table.RowCount == 0 || table.ColumnCount == 0)) {
            yield return table.SourceIdentity;
            foreach (IWorkTableCell cell in table.Cells) yield return cell.RichText?.SourceIdentity;
        }
    }
}
