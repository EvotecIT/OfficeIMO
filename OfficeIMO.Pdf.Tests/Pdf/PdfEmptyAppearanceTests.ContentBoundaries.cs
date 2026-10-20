using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfEmptyAppearanceTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void FlatteningPreservesPagePaintAndAppearanceState(bool form, bool leaveChangedGraphicsState) {
        string pageContent = "0 g 10 160 50 20 re f" + (leaveChangedGraphicsState ? "\n1 0 0 1 40 0 cm\n" : "");
        byte[] source = BuildPdf(form ? "[0 0 20 20]" : null, "[0 0 20 20]", pageContent, widgetOn: form);
        byte[] flattened = form
            ? PdfDocument.Load(source).Forms.Flatten().ToBytes()
            : PdfDocument.Load(source).Annotations.Flatten().Bytes;

        Assert.Equal(Render(source), Render(flattened));
        AssertOriginalPagePaintBytes(source, flattened);
    }

    [Fact]
    public void TextReplacementStampPreservesPagePaintWhenTheOriginalStreamHasNoTrailingWhitespace() {
        byte[] source = BuildPdf(null, "[0 0 0 0]");
        var request = new PdfStamper.TextStampRequest(1, "boundary audit", 20D, 20D,
            PdfStandardFont.Helvetica, 10D, new PdfColor(0D, 0D, 0D), 0D, textRenderingMode: 3);

        byte[] stamped = PdfStamper.StampTextBatch(source, new[] { request });

        Assert.Equal(Render(source), Render(stamped));
        AssertOriginalPagePaintBytes(source, stamped);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(stamped);
        Assert.Contains("boundary audit", independent.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ImportedPagePreservesTheAdmittedSplitNumericOperandAndItsPaint() {
        byte[] unsplit = BuildPdf(null, "[0 0 0 0]");
        byte[] source = PdfDocumentObjectGraphRewriter.Rewrite(unsplit, null, null, (objects, security) => {
            var first = (PdfStream)objects[8].Value;
            objects[8] = new PdfIndirectObject(8, 0, new PdfStream(first.Dictionary, Encoding.ASCII.GetBytes("0 g 10 1")));
            int next = objects.Keys.Max() + 1;
            objects[next] = new PdfIndirectObject(next, 0, new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes("60 50 20 re f")));
            var contents = new PdfArray();
            contents.Items.Add(new PdfReference(8, 0));
            contents.Items.Add(new PdfReference(next, 0));
            ((PdfDictionary)objects[3].Value).Items["Contents"] = contents;
            return security.InfoObjectNumber;
        });
        PdfDocument input = PdfDocument.Load(source);
        PdfDocument target = PdfDocument.Load(BuildPdf(null, "[0 0 0 0]", pageContent: ""));

        PdfDocument imported = target.Stamp.OverlayPage(source, new PdfPageOverlayOptions {
            X = 0D, Y = 0D, Fit = PdfPageOverlayFit.None
        });

        Assert.Equal(Render(unsplit), Render(source));
        Assert.Equal(Render(source), Render(imported.ToBytes()));
        Assert.Equal(source, input.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlatteningAdmitsOnlyTheKnownAddedContentBeyondTightSourceLimits(bool form) =>
        AssertFlattenReadback(form, selected: false, empty: false, nested: false);

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void FullAndSelectedEmptyFormAppearancesKeepBoundedReadback(bool selected, bool empty) =>
        AssertFlattenReadback(form: true, selected, empty, nested: false);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlattenedNestedSharedAppearancesReserveEveryInvocation(bool form) =>
        AssertFlattenReadback(form, selected: form, empty: false, nested: true);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FillAndFlattenKeepsTheGeneratedReadbackGrowth(bool typed) =>
        AssertFlattenReadback(form: true, selected: false, empty: false, nested: false, compose: true, typed);

    private static void AssertFlattenReadback(bool form, bool selected, bool empty, bool nested,
        bool compose = false, bool typed = false) {
        byte[] source = BuildPdf(form ? (empty ? "[0 0 0 0]" : "[0 0 20 20]") : null,
            "[0 0 20 20]", widgetOn: form && !empty);
        int repeatedAppearanceBytes = 0;
        if (nested) source = AddNestedSharedAppearance(source, form ? 5 : 9, out repeatedAppearanceBytes);
        else source = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null,
            (objects, security) => security.InfoObjectNumber);
        PdfStream[] streams = PdfSyntax.ParseObjects(source).Map.Values.Select(item => item.Value)
            .OfType<PdfStream>().ToArray();
        // Source Read also inspects annotation appearances. Shared child invocations are charged each time.
        int sourceContentBytes = streams.Sum(stream => stream.Data.Length) + repeatedAppearanceBytes;
        var limits = new PdfReadLimits {
            MaxInputBytes = source.LongLength,
            MaxRawStreamBytes = streams.Max(stream => stream.Data.Length),
            MaxDecodedStreamBytes = streams.Max(stream => stream.Data.Length),
            MaxTotalDecodedStreamBytes = streams.Sum(stream => (long)stream.Data.Length),
            MaxPageContentBytes = sourceContentBytes,
            MaxRetainedContentBytes = sourceContentBytes,
            MaxContentOperations = sourceContentBytes,
            MaxContentOperands = sourceContentBytes
        };

        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Limits = limits });
        _ = document.Read();
        PdfDocument flattened = compose
            ? (typed ? document.Forms.FillAndFlatten(new Dictionary<string, PdfFormFieldValue> { ["agree"] = PdfFormFieldValue.From("Yes") })
                : document.Forms.FillAndFlatten(new Dictionary<string, string> { ["agree"] = "Yes" }))
            : form ? (selected ? document.Forms.Flatten("agree") : document.Forms.Flatten())
                : document.Annotations.Flatten().ToDocument();

        Assert.Equal(Render(source), Render(flattened.ToBytes()));
        Assert.DoesNotContain(flattened.Reader.Annotations(), annotation => annotation.Subtype == (form ? "Widget" : "Square"));
        _ = flattened.Read();
        Assert.True(flattened.ReadOptions.Limits.MaxPageContentBytes < new PdfReadLimits().MaxPageContentBytes);
        Assert.True(flattened.ReadOptions.Limits.MaxContentOperations < new PdfReadLimits().MaxContentOperations);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PromotionInspectsExistingAppearanceBytesWithinTheCallerBudget(bool form) {
        byte[] source = BuildPdf(form ? "[0 0 20 20]" : null, "[0 0 20 20]", pageContent: "", widgetOn: form);
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxDecodedStreamBytes = 1,
            MaxPageContentBytes = 1,
            MaxTotalDecodedStreamBytes = 1
        }});

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => {
            if (form) document.Forms.Flatten(); else document.Annotations.Flatten();
        });

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(source, document.ToBytes());
    }

    private static byte[] AddNestedSharedAppearance(byte[] source, int appearanceObject, out int repeatedBytes) {
        repeatedBytes = ((PdfStream)PdfSyntax.ParseObjects(source).Map[appearanceObject].Value).Data.Length;
        return PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            var original = (PdfStream)objects[appearanceObject].Value;
            var childDictionary = new PdfDictionary();
            foreach (var entry in original.Dictionary.Items) childDictionary.Items[entry.Key] = entry.Value;
            int child = objects.Keys.Max() + 1;
            objects[child] = new PdfIndirectObject(child, 0, new PdfStream(childDictionary, original.Data));
            var xObjects = new PdfDictionary();
            xObjects.Items["Nested"] = new PdfReference(child, 0);
            var resources = new PdfDictionary(); resources.Items["XObject"] = xObjects;
            var parentDictionary = new PdfDictionary();
            foreach (var entry in original.Dictionary.Items) parentDictionary.Items[entry.Key] = entry.Value;
            parentDictionary.Items["Resources"] = resources;
            objects[appearanceObject] = new PdfIndirectObject(appearanceObject, 0,
                new PdfStream(parentDictionary, Encoding.ASCII.GetBytes("/Nested Do\n/Nested Do\n")));
            return security.InfoObjectNumber;
        });
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LargerGeneratedTextAppearanceKeepsTheOwnedReadPolicy(bool flatten, bool typed) {
        PdfDocument document = CreateBlankTextDocument(out byte[] source, out PdfLoadOptions options);
        _ = document.Read();
        const string value = "Delivery contact: service@example.test";
        var textValues = new Dictionary<string, string> { ["notes"] = value };
        var typedValues = new Dictionary<string, PdfFormFieldValue> { ["notes"] = PdfFormFieldValue.From(value) };

        PdfDocument result = flatten
            ? (typed ? document.Forms.FillAndFlatten(typedValues) : document.Forms.FillAndFlatten(textValues))
            : (typed ? document.Forms.Fill(typedValues) : document.Forms.Fill(textValues));

        _ = result.Read();
        if (!flatten) Assert.Equal(value, result.Inspect().FormFieldsByName["notes"].Value);
        else {
            using var independent = UglyToad.PdfPig.PdfDocument.Open(result.ToBytes());
            Assert.Contains(value, independent.GetPage(1).Text, StringComparison.Ordinal);
        }
        Assert.True(result.ReadOptions.Limits.MaxDecodedStreamBytes > options.Limits.MaxDecodedStreamBytes);
        Assert.True(result.ReadOptions.Limits.MaxDecodedStreamBytes < new PdfReadLimits().MaxDecodedStreamBytes);
        Assert.Equal(source, document.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DirectAnnotationFlattenFacadesKeepTheKnownReadbackGrowth(bool resultRoute) {
        byte[] source = BuildPdf(null, "[0 0 20 20]");
        PdfStream[] streams = PdfSyntax.ParseObjects(source).Map.Values.Select(item => item.Value).OfType<PdfStream>().ToArray();
        int bytes = streams.Sum(stream => stream.Data.Length);
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxInputBytes = source.LongLength,
            MaxRawStreamBytes = streams.Max(stream => stream.Data.Length),
            MaxDecodedStreamBytes = streams.Max(stream => stream.Data.Length),
            MaxPageContentBytes = bytes, MaxRetainedContentBytes = bytes,
            MaxTotalDecodedStreamBytes = bytes, MaxContentOperations = bytes, MaxContentOperands = bytes
        }});
        _ = document.Read();

        PdfDocument flattened = resultRoute
            ? document.FlattenVisualAnnotationsResult().Value!
            : document.FlattenVisualAnnotations();

        _ = flattened.Read();
        Assert.Equal(Render(source), Render(flattened.ToBytes()));
        Assert.Empty(flattened.Reader.Annotations());
    }

    [Theory]
    [InlineData("fill-result")]
    [InlineData("fill-flatten-result")]
    [InlineData("data")]
    [InlineData("xfdf")]
    [InlineData("transaction")]
    public void FillAliasesPreserveTheirCanonicalOutputPolicy(string route) {
        PdfDocument source = CreateBlankTextDocument(out byte[] original, out _);
        _ = source.Read();
        const string value = "Delivery contact: service@example.test";
        var values = new Dictionary<string, string> { ["notes"] = value };
        var data = new PdfFormDataSet(new[] { new PdfFormDataField("notes", new[] { value }) });

        PdfDocument result = route switch {
            "fill-result" => source.Forms.FillResult(values).Value!,
            "fill-flatten-result" => source.Forms.FillAndFlattenResult(values).Value!,
            "data" => source.Forms.ImportData(data),
            "xfdf" => source.Forms.ImportXfdf(data.ToXfdf()),
            _ => source.Forms.Edit(edit => edit.Move("notes", 1, 20, 20, 160.5, 24).Flatten("notes")).ToDocument()
        };

        _ = result.Read();
        if (route == "transaction") Assert.Empty(result.Inspect().FormFields);
        else if (route != "fill-flatten-result") Assert.Equal(value, result.Inspect().FormFieldsByName["notes"].Value);
        Assert.Equal(original, source.ToBytes());
    }

    private static PdfDocument CreateBlankTextDocument(out byte[] source, out PdfLoadOptions options) {
        source = PdfDocument.Load(BuildPdf(null, "[0 0 0 0]")).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "notes", Kind = PdfFormFieldCreationKind.Text, X = 20, Y = 20, Width = 160, Height = 24, Value = ""
        })).ToDocument().ToBytes();
        PdfStream[] streams = PdfSyntax.ParseObjects(source).Map.Values.Select(item => item.Value).OfType<PdfStream>().ToArray();
        int bytes = streams.Sum(stream => stream.Data.Length);
        options = new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxInputBytes = source.LongLength,
            MaxRawStreamBytes = streams.Max(stream => stream.Data.Length),
            MaxDecodedStreamBytes = streams.Max(stream => stream.Data.Length),
            MaxPageContentBytes = bytes, MaxRetainedContentBytes = bytes,
            MaxTotalDecodedStreamBytes = bytes, MaxContentOperations = bytes, MaxContentOperands = bytes
        }};
        return PdfDocument.Load(source, options);
    }

    private static byte[] Render(byte[] pdf) {
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf,
            options: new PdfPageRenderOptions { ContinueOnError = false }));
        Assert.True(result.Succeeded);
        return Assert.IsType<byte[]>(result.Bytes);
    }

    private static void AssertOriginalPagePaintBytes(byte[] source, byte[] output) {
        PdfStream original = Assert.IsType<PdfStream>(PdfSyntax.ParseObjects(source).Map[8].Value);
        Assert.Contains(PdfSyntax.ParseObjects(output).Map.Values, item => item.Value is PdfStream stream
            && stream.Data.SequenceEqual(original.Data));
    }
}
