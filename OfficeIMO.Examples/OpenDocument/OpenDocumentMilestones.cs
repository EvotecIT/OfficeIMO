using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;

namespace OfficeIMO.Examples.OpenDocument;

internal static class OpenDocumentMilestones {
    internal static void Example(string folderPath) {
        string output = Path.Combine(folderPath, "OpenDocument");
        Directory.CreateDirectory(output);

        using (OdtDocument text = OdtDocument.Create()) {
            text.AddHeading("Native OpenDocument", 1);
            text.AddTrackedParagraphInsertion("This paragraph is tracked.", "OfficeIMO").Accept();
            text.AddParagraph("The same document can be written as packaged ODT or flat FODT XML.");
            text.Save(Path.Combine(output, "native-text.odt"));
            text.SaveFlatXml(Path.Combine(output, "native-text.fodt"));
        }

        using (OdsDocument spreadsheet = OdsDocument.Create()) {
            OdsSheet data = spreadsheet.AddSheet("Data");
            data.Cell(0, 0).SetNumber(20D);
            data.Cell(1, 0).SetNumber(22D);
            data.Cell(0, 1).Formula = "of:=SUM([.A1:.A2])";
            OdsRecalculationReport calculation = spreadsheet.Recalculate();
            if (calculation.FailedCells != 0) throw new InvalidOperationException("The OpenFormula example did not recalculate.");
            spreadsheet.Save(Path.Combine(output, "native-sheet.ods"));
        }

        using (OdpPresentation presentation = OdpPresentation.Create()) {
            OdpSlide slide = presentation.AddSlide("Animation");
            OdpRectangle panel = slide.AddRectangle(OdfRect.FromCentimeters(2, 2, 8, 4));
            slide.AddFadeInAnimation(panel, TimeSpan.FromSeconds(1));
            presentation.Save(Path.Combine(output, "native-slides.odp"));
        }

        using (WordDocument word = WordDocument.Create()) {
            word.AddParagraph("Explicit Word to ODT conversion").Style = WordParagraphStyles.Heading1;
            OdfConversionResult<OdtDocument> conversion = word.ToOpenDocumentResult();
            using OdtDocument converted = conversion.Value;
            converted.Save(Path.Combine(output, "converted-from-word.odt"));
            File.WriteAllLines(Path.Combine(output, "converted-from-word.mapping.txt"), conversion.Report.Mappings
                .Select(mapping => $"{mapping.Feature}: {mapping.Status} ({mapping.Count})"));
        }
    }
}
