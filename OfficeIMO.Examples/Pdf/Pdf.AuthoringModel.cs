using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Pdf {
    internal static class AuthoringModelPdf {
        public static void Example_Pdf_AuthoringModel(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.AuthoringModel.pdf");

            PdfDocument.Create(document => document
                .Typography(OfficeRenderingProfile.Managed)
                .Settings(options => {
                    options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
                    options.PageSize = PdfPageSizes.A4;
                    options.Margins = PdfPageMargins.Uniform(OfficeLength.Points(64));
                    options.Language = "en-GB";
                })
                .Content(content => content
                    .Element(element => element
                        .Semantic(PdfSemanticRole.Article)
                        .Background(OfficeColor.FromRgb(248, 250, 252))
                        .Border(OfficeColor.FromRgb(148, 163, 184))
                        .Padding(vertical: OfficeLength.Points(16), horizontal: OfficeLength.Points(18))
                        .KeepTogether()
                        .Content(article => article
                            .H1("Operational summary")
                            .Text("The same content receiver powers documents, pages, elements, components, and row columns.")
                            .Spacer(global::OfficeIMO.OfficeLength.Points(10))
                            .Row(row => row
                                .Gap(global::OfficeIMO.OfficeLength.Points(12))
                                .FixedColumn(global::OfficeIMO.OfficeLength.Points(54), cell => cell
                                    .H3("ID")
                                    .Text("A-104"))
                                .AutoColumn(cell => cell
                                    .H3("Owner")
                                    .Text("Operations"), maximum: global::OfficeIMO.OfficeLength.Points(100))
                                .RelativeColumn(cell => cell
                                    .H3("Status")
                                    .Text("Validated through explicit fixed, automatic, and relative sizing.")))))
                    .Spacer(global::OfficeIMO.OfficeLength.Points(12))
                    .Semantic(PdfSemanticRole.Section, section => section
                        .H2("Notes")
                        .Text("Semantic groups feed the tagged-PDF structure tree without changing visual flow."))))
                .Meta(title: "OfficeIMO.Pdf authoring model", author: "OfficeIMO")
                .Save(path);

            if (open) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo {
                    FileName = path,
                    UseShellExecute = true
                });
            }
        }
    }
}
