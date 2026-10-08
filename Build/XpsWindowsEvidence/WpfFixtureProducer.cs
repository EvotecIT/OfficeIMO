using System.IO;
using System.IO.Packaging;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using NativeXps = System.Windows.Xps.Packaging.XpsDocument;

namespace OfficeIMO.XpsWindowsEvidence;

internal static class WpfFixtureProducer {
    internal static byte[] Create(string repository, string output) {
        string font = Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-Regular.ttf");
        var sequence = new FixedDocumentSequence();
        int pageNumber = 0;
        for (int documentNumber = 0; documentNumber < 2; documentNumber++) {
            var document = new FixedDocument();
            for (int index = 0; index <= documentNumber; index++) {
                pageNumber++;
                var page = new FixedPage { Width = 200, Height = 160 };
                page.Children.Add(new System.Windows.Shapes.Path {
                    Data = Geometry.Parse("M0,0H200V160H0Z"),
                    Fill = new RadialGradientBrush(Colors.Red, Colors.Blue) {
                        MappingMode = BrushMappingMode.Absolute, Center = new Point(100, 80),
                        GradientOrigin = new Point(100, 80), RadiusX = 60, RadiusY = 25,
                        ColorInterpolationMode = ColorInterpolationMode.SRgbLinearInterpolation
                    }
                });
                page.Children.Add(new Glyphs {
                    FontUri = new Uri(font, UriKind.Absolute), FontRenderingEmSize = 18,
                    OriginX = 12, OriginY = 28, Fill = Brushes.Black,
                    UnicodeString = "Native WPF page " + pageNumber
                });
                page.Measure(new Size(page.Width, page.Height));
                page.Arrange(new Rect(0, 0, page.Width, page.Height));
                document.Pages.Add(new PageContent { Child = page });
            }
            var reference = new DocumentReference();
            reference.SetDocument(document);
            sequence.References.Add(reference);
        }
        string path = Path.Combine(output, "independent-wpf.xps");
        using (var package = Package.Open(path, FileMode.Create, FileAccess.ReadWrite))
        using (var document = new NativeXps(package, CompressionOption.Maximum))
            NativeXps.CreateXpsDocumentWriter(document).Write(sequence);
        return File.ReadAllBytes(path);
    }
}
