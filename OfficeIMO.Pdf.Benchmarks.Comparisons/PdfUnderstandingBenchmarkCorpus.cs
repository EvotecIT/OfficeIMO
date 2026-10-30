using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed record PdfUnderstandingBenchmarkExpectation(
    int PageNumber,
    IReadOnlyList<string> ReadingOrder,
    string TableMarker,
    IReadOnlyList<IReadOnlyList<string>> ExpectedTableRows,
    IReadOnlyDictionary<string, string> ExpectedRegionText,
    IReadOnlyDictionary<string, PdfUnderstandingSemanticKind> SemanticKinds,
    IReadOnlyDictionary<string, int> HeadingLevels) {
    internal IReadOnlyList<string> ExpectedTableCells =>
        ExpectedTableRows.SelectMany(static row => row).ToArray();
}

internal sealed record PdfUnderstandingBenchmarkCorpus(
    byte[] Pdf,
    IReadOnlyList<PdfUnderstandingBenchmarkExpectation> Pages,
    byte[] ContinuationPdf,
    IReadOnlyList<(int PreviousPage, int CurrentPage)> ExpectedContinuationPairs);

internal static class PdfUnderstandingBenchmarkCorpusFactory {
    internal static PdfUnderstandingBenchmarkCorpus Create(PdfBenchmarkScale scale) {
        // Running-header/footer recovery is intentionally document-wide, so this labelled
        // semantic corpus must contain at least two pages even for the Easy scale.
        int pageCount = Math.Max(2, PdfBenchmarkScenario.Get(scale).PageCount);
        var expectations = new List<PdfUnderstandingBenchmarkExpectation>(pageCount);
        PdfDocument document = PdfDocument.Create(pdf => pdf.Content(content => {
            for (int pageNumber = 1; pageNumber <= pageCount; pageNumber++) {
                if (pageNumber > 1) {
                    content.PageBreak();
                }

                string suffix = AlphabeticPageToken(pageNumber);
                string header = "HEADER-P" + pageNumber.ToString("D3", System.Globalization.CultureInfo.InvariantCulture);
                string title = "TITLE-P" + suffix;
                string leftOne = "LEFT-ONE-P" + suffix;
                string leftTwo = "LEFT-TWO-P" + suffix;
                string leftThree = "LEFT-THREE-P" + suffix;
                string rightOne = "RIGHT-ONE-P" + suffix;
                string rightTwo = "RIGHT-TWO-P" + suffix;
                string rightThree = "RIGHT-THREE-P" + suffix;
                string list = "LIST-P" + suffix;
                string caption = "CAPTION-P" + suffix;
                string table = "TABLE-P" + suffix;
                string footer = "FOOTER-P" + pageNumber.ToString("D3", System.Globalization.CultureInfo.InvariantCulture);
                string ownerHeader = "OH-P" + suffix;
                string amountHeader = "AH-P" + suffix;
                string statusHeader = "SH-P" + suffix;
                string accountOne = "C1-P" + suffix;
                string accountTwo = "C2-P" + suffix;
                string accountThree = accountOne;
                string ownerOne = "O1-P" + suffix;
                string ownerTwo = "O2-P" + suffix;
                string ownerThree = ownerOne;
                string amountOne = "A1-P" + suffix + " 1037.25";
                string amountTwo = "A2-P" + suffix + " 1074.50";
                string amountThree = amountOne;
                string statusOne = "S1-P" + suffix + " Approved";
                string statusTwo = "S2-P" + suffix + " Review";
                string statusThree = statusOne;
                string tableRegionText = string.Join(" ", new[] {
                    table, ownerHeader, amountHeader, statusHeader,
                    accountOne, ownerOne, amountOne, statusOne,
                    accountTwo, ownerTwo, amountTwo, statusTwo,
                    accountThree, ownerThree, amountThree, statusThree
                });

                content.Canvas(canvas => canvas
                    .Text(header + " semantic benchmark running header", global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(8D), global::OfficeIMO.OfficeLength.Points(520D), global::OfficeIMO.OfficeLength.Points(22D), fontSize: global::OfficeIMO.OfficeLength.Points(10D))
                    .Text(title + " Structured understanding spanning heading", global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(48D), global::OfficeIMO.OfficeLength.Points(520D), global::OfficeIMO.OfficeLength.Points(28D), fontSize: global::OfficeIMO.OfficeLength.Points(18D))
                    .Text(leftOne + " left opening", global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(112D), global::OfficeIMO.OfficeLength.Points(200D), global::OfficeIMO.OfficeLength.Points(24D), fontSize: global::OfficeIMO.OfficeLength.Points(11D))
                    .Text(leftTwo + " left indented", global::OfficeIMO.OfficeLength.Points(66D), global::OfficeIMO.OfficeLength.Points(164D), global::OfficeIMO.OfficeLength.Points(170D), global::OfficeIMO.OfficeLength.Points(24D), fontSize: global::OfficeIMO.OfficeLength.Points(11D))
                    .Text(leftThree + " left closing", global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(216D), global::OfficeIMO.OfficeLength.Points(200D), global::OfficeIMO.OfficeLength.Points(24D), fontSize: global::OfficeIMO.OfficeLength.Points(11D))
                    .Text(rightOne + " right opening", global::OfficeIMO.OfficeLength.Points(340D), global::OfficeIMO.OfficeLength.Points(112D), global::OfficeIMO.OfficeLength.Points(206D), global::OfficeIMO.OfficeLength.Points(24D), fontSize: global::OfficeIMO.OfficeLength.Points(11D))
                    .Text(rightTwo + " right middle", global::OfficeIMO.OfficeLength.Points(340D), global::OfficeIMO.OfficeLength.Points(164D), global::OfficeIMO.OfficeLength.Points(206D), global::OfficeIMO.OfficeLength.Points(24D), fontSize: global::OfficeIMO.OfficeLength.Points(11D))
                    .Text(rightThree + " right closing", global::OfficeIMO.OfficeLength.Points(340D), global::OfficeIMO.OfficeLength.Points(216D), global::OfficeIMO.OfficeLength.Points(206D), global::OfficeIMO.OfficeLength.Points(24D), fontSize: global::OfficeIMO.OfficeLength.Points(11D))
                    .Text("1. " + list + " classified list item", global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(292D), global::OfficeIMO.OfficeLength.Points(510D), global::OfficeIMO.OfficeLength.Points(24D), fontSize: global::OfficeIMO.OfficeLength.Points(11D))
                    .Text(CaptionLabel(pageNumber) + " 1. " + caption + " classified table caption", global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(370D), global::OfficeIMO.OfficeLength.Points(510D), global::OfficeIMO.OfficeLength.Points(18D), fontSize: global::OfficeIMO.OfficeLength.Points(9D))
                    .Table(new[] {
                        new[] { table, ownerHeader, amountHeader, statusHeader },
                        new[] { accountOne, ownerOne, amountOne, statusOne },
                        new[] { accountTwo, ownerTwo, amountTwo, statusTwo },
                        new[] { accountThree, ownerThree, amountThree, statusThree }
                    }, global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(390D), global::OfficeIMO.OfficeLength.Points(510D), global::OfficeIMO.OfficeLength.Points(190D), style: new PdfTableStyle { HeaderRowCount = 1 })
                    .Text(footer + " semantic benchmark running footer", global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(756D), global::OfficeIMO.OfficeLength.Points(510D), global::OfficeIMO.OfficeLength.Points(18D), fontSize: global::OfficeIMO.OfficeLength.Points(12D)));

                expectations.Add(new PdfUnderstandingBenchmarkExpectation(
                    pageNumber,
                    new[] {
                        header, title,
                        leftOne, leftTwo, leftThree,
                        rightOne, rightTwo, rightThree,
                        list, caption, table, footer
                    },
                    table,
                    new IReadOnlyList<string>[] {
                        new[] { table, ownerHeader, amountHeader, statusHeader },
                        new[] { accountOne, ownerOne, amountOne, statusOne },
                        new[] { accountTwo, ownerTwo, amountTwo, statusTwo },
                        new[] { accountThree, ownerThree, amountThree, statusThree }
                    },
                    new Dictionary<string, string>(StringComparer.Ordinal) {
                        [header] = header + " semantic benchmark running header",
                        [title] = title + " Structured understanding spanning heading",
                        [leftOne] = leftOne + " left opening",
                        [leftTwo] = leftTwo + " left indented",
                        [leftThree] = leftThree + " left closing",
                        [rightOne] = rightOne + " right opening",
                        [rightTwo] = rightTwo + " right middle",
                        [rightThree] = rightThree + " right closing",
                        [list] = "1. " + list + " classified list item",
                        [caption] = CaptionLabel(pageNumber) + " 1. " + caption + " classified table caption",
                        [table] = tableRegionText,
                        [footer] = footer + " semantic benchmark running footer"
                    },
                    new Dictionary<string, PdfUnderstandingSemanticKind>(StringComparer.Ordinal) {
                        [header] = PdfUnderstandingSemanticKind.Header,
                        [title] = PdfUnderstandingSemanticKind.Heading,
                        [leftOne] = PdfUnderstandingSemanticKind.Paragraph,
                        [leftTwo] = PdfUnderstandingSemanticKind.Paragraph,
                        [leftThree] = PdfUnderstandingSemanticKind.Paragraph,
                        [rightOne] = PdfUnderstandingSemanticKind.Paragraph,
                        [rightTwo] = PdfUnderstandingSemanticKind.Paragraph,
                        [rightThree] = PdfUnderstandingSemanticKind.Paragraph,
                        [list] = PdfUnderstandingSemanticKind.ListItem,
                        [caption] = PdfUnderstandingSemanticKind.Caption,
                        [table] = PdfUnderstandingSemanticKind.Table,
                        [footer] = PdfUnderstandingSemanticKind.Footer
                    },
                    new Dictionary<string, int>(StringComparer.Ordinal) {
                        [title] = 1
                    }));
            }
        }));

        byte[] continuationPdf = CreateContinuationPdf();
        using var continuationInput = new MemoryStream(continuationPdf, writable: false);
        using PdfDocument continuationDocument = PdfDocument.Load(continuationInput);
        int continuationPageCount = continuationDocument.Pages.Count;
        var continuationPairs = Enumerable.Range(1, Math.Max(0, continuationPageCount - 1))
            .Select(static page => (page, page + 1))
            .ToArray();
        return new PdfUnderstandingBenchmarkCorpus(
            document.ToBytes(),
            expectations.AsReadOnly(),
            continuationPdf,
            continuationPairs);
    }

    private static byte[] CreateContinuationPdf() {
        var rows = new List<string[]> {
            new[] { "Account", "Owner", "Amount", "Status" }
        };
        for (int index = 1; index <= 60; index++) {
            rows.Add(new[] {
                $"CONT-{index:D3}",
                $"Owner-{index % 11:D2}",
                (index * 17.25M).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture),
                index % 5 == 0 ? "Review" : "Approved"
            });
        }

        return PdfDocument.Create(pdf => pdf.Content(content => content.Table(rows, style: new PdfTableStyle {
                HeaderRowCount = 1,
                RepeatHeaderRowCount = 1,
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(82), global::OfficeIMO.OfficeLength.Points(82), global::OfficeIMO.OfficeLength.Points(82), global::OfficeIMO.OfficeLength.Points(82) },
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(3),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(2)
            })), new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(260),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(24),
                MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(24),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(8)
            })
            .ToBytes();
    }

    private static string AlphabeticPageToken(int pageNumber) {
        if (pageNumber < 1) throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "The benchmark page number must be positive.");
        int value = pageNumber - 1;
        var characters = new char[3];
        for (int index = characters.Length - 1; index >= 0; index--) {
            characters[index] = (char)('A' + (value % 26));
            value /= 26;
        }
        if (value != 0) throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "The benchmark page token supports at most 17,576 pages.");
        return new string(characters);
    }

    private static string CaptionLabel(int pageNumber) => ((pageNumber - 1) % 4) switch {
        0 => "Tabell",
        1 => "Tabela",
        2 => "Tabelle",
        _ => "Cuadro"
    };
}
