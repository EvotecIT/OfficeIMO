using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void DynamicFields_FreshAddChangeAndRemove_RoundTripNatively() {
            const string rtfFormat = "{\\rtf1\\ansi dddd, MMMM d}";
            string encodedRtf = Convert.ToBase64String(
                System.Text.Encoding.Unicode.GetBytes(rtfFormat));
            byte[] sourceBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointTextBox textBox = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBox(string.Empty);
                Assert.IsType<P.Shape>(textBox.Element).TextBody =
                    new P.TextBody(new A.BodyProperties(),
                        new A.ListStyle(), new A.Paragraph(
                            CreateField("{00000000-0000-0000-0000-000000000001}",
                                "slidenum", "7", bold: true),
                            new A.Run(new A.Text("|")),
                            CreateField("{00000000-0000-0000-0000-000000000002}",
                                "datetime4", "July 16, 2026"),
                            new A.Run(new A.Text("|")),
                            CreateField("{00000000-0000-0000-0000-000000000003}",
                                "datetimeFigureOut", "16.07.2026"),
                            new A.Run(new A.Text("|")),
                            CreateField("{00000000-0000-0000-0000-000000000004}",
                                "header", "Header"),
                            new A.Run(new A.Text("|")),
                            CreateField("{00000000-0000-0000-0000-000000000005}",
                                "footer", "Footer"),
                            new A.Run(new A.Text("|")),
                            CreateField("{00000000-0000-0000-0000-000000000006}",
                                "datetimeRtf:" + encodedRtf,
                                "Thursday, July 16"),
                            new A.EndParagraphRunProperties()));

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation sourceLegacy = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptTextBody sourceText = Assert.Single(Assert.Single(
                    sourceLegacy.Slides).Shapes).TextBody;
            Assert.Equal("*|*|*|*|*|*", sourceText.Text);
            Assert.Collection(sourceText.Fields,
                field => Assert.Equal(LegacyPptTextFieldKind.SlideNumber,
                    field.Kind),
                field => {
                    Assert.Equal(LegacyPptTextFieldKind.DateTime,
                        field.Kind);
                    Assert.Equal((byte)3, field.DateTimeFormatIndex);
                },
                field => Assert.Equal(LegacyPptTextFieldKind.GenericDate,
                    field.Kind),
                field => Assert.Equal(LegacyPptTextFieldKind.Header,
                    field.Kind),
                field => Assert.Equal(LegacyPptTextFieldKind.Footer,
                    field.Kind),
                field => {
                    Assert.Equal(LegacyPptTextFieldKind.RtfDateTime,
                        field.Kind);
                    Assert.Equal(rtfFormat, field.RtfFormat);
                });
            Assert.False(sourceText.IsFieldDataMalformed);
            Assert.Equal(6, sourceLegacy.CreateImportReport()
                .TextFieldCount);

            byte[] editedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                A.Paragraph paragraph = Assert.Single(imported.Slides[0]
                    .SlidePart.Slide!.Descendants<A.Paragraph>());
                A.Field[] fields = paragraph.Elements<A.Field>().ToArray();
                Assert.Equal(6, fields.Length);
                Assert.Equal("slidenum", fields[0].Type!.Value);
                Assert.Equal("datetime4", fields[1].Type!.Value);
                Assert.Equal("datetimeFigureOut", fields[2].Type!.Value);
                Assert.Equal("header", fields[3].Type!.Value);
                Assert.Equal("footer", fields[4].Type!.Value);
                Assert.Equal("datetimeRtf:" + encodedRtf,
                    fields[5].Type!.Value);
                fields[1].Type = "datetime2";
                fields[3].InsertAfterSelf(new A.Run(new A.Text("*")));
                fields[3].Remove();
                paragraph.InsertBefore(CreateField(
                        "{00000000-0000-0000-0000-000000000007}",
                        "slidenum", "8"),
                    paragraph.GetFirstChild<A.EndParagraphRunProperties>());

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation edited = LegacyPptPresentation.Load(
                editedBytes);
            LegacyPptTextBody editedText = Assert.Single(Assert.Single(
                    edited.Slides).Shapes).TextBody;
            Assert.Equal(6, editedText.Fields.Count);
            Assert.DoesNotContain(editedText.Fields,
                field => field.Kind == LegacyPptTextFieldKind.Header);
            Assert.Equal(2, editedText.Fields.Count(field =>
                field.Kind == LegacyPptTextFieldKind.SlideNumber));
            Assert.Equal((byte)1, Assert.Single(editedText.Fields,
                field => field.Kind == LegacyPptTextFieldKind.DateTime)
                .DateTimeFormatIndex);
            Assert.Equal(sourceLegacy.Package.UserEdits.Count + 1,
                edited.Package.UserEdits.Count);
            Assert.True(edited.Package.DocumentStream.AsSpan(0,
                    sourceLegacy.Package.DocumentStream.Length)
                .SequenceEqual(sourceLegacy.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(editedBytes,
                writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(reopenedInput);
            Assert.Equal(6, reopened.Slides[0].SlidePart.Slide!
                .Descendants<A.Field>().Count());
            Assert.Empty(reopened.ValidateDocument());
        }

        private static A.Field CreateField(string id, string type,
            string displayText, bool bold = false) => new(
                new A.RunProperties { Bold = bold },
                new A.Text(displayText)) {
                Id = id,
                Type = type
            };

    }
}
