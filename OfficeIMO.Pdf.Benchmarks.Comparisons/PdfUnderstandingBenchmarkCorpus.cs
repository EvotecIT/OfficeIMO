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
                string header = "HEADER-P" + suffix;
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
                string footer = "FOOTER-P" + suffix;
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
                    .Text(header + " semantic benchmark running header", 36D, 8D, 520D, 22D, fontSize: 10D)
                    .Text(title + " Structured understanding spanning heading", 36D, 48D, 520D, 28D, fontSize: 18D)
                    .Text(leftOne + " left opening", 36D, 112D, 200D, 24D, fontSize: 11D)
                    .Text(leftTwo + " left indented", 66D, 164D, 170D, 24D, fontSize: 11D)
                    .Text(leftThree + " left closing", 36D, 216D, 200D, 24D, fontSize: 11D)
                    .Text(rightOne + " right opening", 340D, 112D, 206D, 24D, fontSize: 11D)
                    .Text(rightTwo + " right middle", 340D, 164D, 206D, 24D, fontSize: 11D)
                    .Text(rightThree + " right closing", 340D, 216D, 206D, 24D, fontSize: 11D)
                    .Text("1. " + list + " classified list item", 36D, 292D, 510D, 24D, fontSize: 11D)
                    .Text(CaptionLabel(pageNumber) + " 1. " + caption + " classified table caption", 36D, 370D, 510D, 18D, fontSize: 9D)
                    .Table(new[] {
                        new[] { table, ownerHeader, amountHeader, statusHeader },
                        new[] { accountOne, ownerOne, amountOne, statusOne },
                        new[] { accountTwo, ownerTwo, amountTwo, statusTwo },
                        new[] { accountThree, ownerThree, amountThree, statusThree }
                    }, 36D, 390D, 510D, 190D, style: new PdfTableStyle { HeaderRowCount = 1 })
                    .Text(footer + " semantic benchmark running footer", 36D, 756D, 510D, 18D, fontSize: 12D));

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
        int continuationPageCount = PdfReadDocument.Open(continuationPdf).Pages.Count;
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
                ColumnWidthPoints = new List<double?> { 82, 82, 82, 82 },
                CellPaddingX = 3,
                CellPaddingY = 2
            })), new PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24,
                DefaultFontSize = 8
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
