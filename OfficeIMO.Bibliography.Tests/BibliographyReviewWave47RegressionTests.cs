using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave47RegressionTests {
    [Theory]
    [InlineData("abc")]
    [InlineData("")]
    [InlineData("  ")]
    public void Named_EndNote_types_with_nonnumeric_codes_block_strict_canonical_output(string code) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">" + code + "</ref-type></record></records></xml>";
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBEND004" && diagnostic.Field == "ref-type");
        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            read.Document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV222" && diagnostic.Field == "ref-type");
    }

    [Theory]
    [InlineData("secondary-authors")]
    [InlineData("tertiary-authors")]
    [InlineData("subsidiary-authors")]
    public void Empty_EndNote_contributor_groups_mixed_with_typed_names_block_strict_canonical_output(string emptyGroup) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><contributors><authors><author>Doe, Jane</author></authors><" + emptyGroup + "/></contributors></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Publisher = "After";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(document.Items[0].NativeFields, field => field.Name == "contributors" && field.RawValue!.Contains("<" + emptyGroup, StringComparison.Ordinal));
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV123" && diagnostic.Field == "contributors");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex, "")]
    [InlineData(BibliographyFormat.BibLatex, " ")]
    [InlineData(BibliographyFormat.Ris, "")]
    [InlineData(BibliographyFormat.Nbib, " ")]
    [InlineData(BibliographyFormat.EndNoteXml, "")]
    public void Numeric_dates_with_blank_literals_block_strict_non_CSL_output(BibliographyFormat format, string literal) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2026, Literal = literal });
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "x"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV221" && diagnostic.Field == "dates.Issued.literal");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex, "author")]
    [InlineData(BibliographyFormat.BibTex, "editor")]
    [InlineData(BibliographyFormat.BibTex, "translator")]
    [InlineData(BibliographyFormat.BibLatex, "author")]
    public void Blank_native_Bib_name_fields_cannot_promote_into_typed_contributors(BibliographyFormat format, string fieldName) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book };
        item.NativeFields.Add(new BibliographyNativeField(format, fieldName, " "));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV119" && diagnostic.Field == fieldName);
    }

    [Fact]
    public void Duplicate_output_key_generation_scales_to_large_collision_sets() {
        const int count = 20_000;
        var items = new List<BibliographyItem>(count);
        for (int index = 0; index < count; index++) items.Add(new BibliographyItem { Key = "duplicate" });

        string[] keys = CodecMappings.OutputKeys(items, BibliographyFormat.BibTex, CancellationToken.None);

        Assert.Equal(count, keys.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.Equal("duplicate", keys[0]);
        Assert.Equal("duplicate-20000", keys[count - 1]);
    }

    [Fact]
    public void Output_key_suffixes_skip_preexisting_collisions_deterministically() {
        var items = new[] {
            new BibliographyItem { Key = "x" },
            new BibliographyItem { Key = "x-2" },
            new BibliographyItem { Key = "x" },
            new BibliographyItem { Key = "X" }
        };

        string[] keys = CodecMappings.OutputKeys(items, BibliographyFormat.Ris, CancellationToken.None);

        Assert.Equal(new[] { "x", "x-2", "x-3", "X-4" }, keys);
    }

    [Theory]
    [InlineData(BibliographyFormat.CslJson)]
    [InlineData(BibliographyFormat.BibTex)]
    public void Document_native_entry_writers_observe_cancellation(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        for (int index = 0; index < 500_000; index++)
            document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.Ris, "comment", "value", "entry"));
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() => {
                if (format == BibliographyFormat.CslJson)
                    CslJsonCodec.Write(document, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), cancellation.Token);
                else
                    BibCodec.Write(document, format, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), cancellation.Token);
            });
        } finally {
            cancellationThread.Join();
        }
    }
}
