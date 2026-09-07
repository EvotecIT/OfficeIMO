namespace OfficeIMO.Examples.Showcase;

internal static partial class DocumentFeatureShowcase {
    private static void CreatePdfExamples(string output) {
        Pdf.ShowcaseStatementPdf.Example_Pdf_ShowcaseStatement(CreateExampleFolder(output, "pdf-statement"));
        Pdf.RowColumnsPdf.Example_Pdf_RowColumns(CreateExampleFolder(output, "pdf-columns"));
        Pdf.DrawingGalleryPdf.Example_Pdf_DrawingGallery(CreateExampleFolder(output, "pdf-drawing"));
        Pdf.TableStyleGalleryPdf.Example_Pdf_TableStyleGallery(CreateExampleFolder(output, "pdf-table-styles"));
        Pdf.ShowcaseManipulationPdf.Example_Pdf_ShowcaseManipulation(CreateExampleFolder(output, "pdf-page-operations"));
    }
}
