using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfJavaScriptEditorTests {
    [Fact]
    public void Edit_AddsListsAndReplacesNamedScriptsWithPlannerProof() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Document JavaScript preservation marker"))
            .ToBytes();

        PdfJavaScriptEditResult created = PdfDocument.Open(source).JavaScript.Edit(scripts => scripts
            .AddOrReplace("Initialize", "this.zoom = 100;")
            .AddOrReplace("Calculate", "var total = 40 + 2;"));

        Assert.Equal(PdfMutationOperation.ModifyJavaScript, created.MutationPlan.Operation);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, created.MutationPlan.ExecutionMode);
        Assert.Contains(PdfMutationStructure.ActiveContent, created.MutationPlan.AffectedStructures);
        Assert.Contains(PdfMutationProof.JavaScriptReadback, created.MutationPlan.RequiredProofs);
        Assert.Contains(created.MutationPlan.CapabilityRecords, static capability =>
            capability.Kind == PdfMutationCapabilityKind.ActiveContentChanges);
        Assert.True(created.PreservationReport.IsPreserved,
            string.Join(" ", created.PreservationReport.Issues.Select(static issue => issue.Message)));
        Assert.Equal(new[] { "Calculate", "Initialize" }, created.JavaScripts.Select(static script => script.Name));
        Assert.Equal("this.zoom = 100;", Assert.Single(created.JavaScripts, static script => script.Name == "Initialize").Script);

        PdfDocument opened = created.ToDocument();
        Assert.Equal(created.JavaScripts.Select(static script => script.Script), opened.JavaScript.List().Select(static script => script.Script));
        Assert.True(opened.Inspect().HasActiveContent);
        Assert.Contains("Document JavaScript preservation marker", opened.Read.Text(), StringComparison.Ordinal);

        PdfJavaScriptEditResult replaced = opened.JavaScript.AddOrReplace("Initialize", "this.zoom = 125;");
        Assert.Equal(2, replaced.JavaScripts.Count);
        Assert.Equal("this.zoom = 125;", Assert.Single(replaced.JavaScripts, static script => script.Name == "Initialize").Script);
        Assert.Equal("var total = 40 + 2;", Assert.Single(replaced.JavaScripts, static script => script.Name == "Calculate").Script);
    }

    [Fact]
    public void RemoveAndClear_PruneNameTreeAndRemainCompatibleWithSanitizer() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("JavaScript cleanup"))
            .ToBytes();
        PdfJavaScriptEditResult authored = PdfDocument.Open(source).JavaScript.Edit(scripts => scripts
            .AddOrReplace("One", "app.alert('one');")
            .AddOrReplace("Two", "app.alert('two');"));

        Assert.Contains(PdfSanitizer.Analyze(authored.ToBytes()), static finding =>
            finding.Kind == PdfSanitizationFindingKind.ActiveAction && finding.Detail == "JavaScript");

        PdfJavaScriptEditResult removed = authored.ToDocument().JavaScript.Remove("One");
        Assert.Equal("Two", Assert.Single(removed.JavaScripts).Name);

        PdfJavaScriptEditResult cleared = removed.ToDocument().JavaScript.Clear();
        Assert.Empty(cleared.JavaScripts);
        Assert.Empty(cleared.ToDocument().Read.JavaScripts());
        Assert.False(PdfInspector.Inspect(cleared.ToBytes()).HasActiveContent);
        Assert.DoesNotContain("app.alert", PdfEncoding.Latin1GetString(cleared.ToBytes()), StringComparison.Ordinal);

        PdfRewritePreservationReport strictPreservation = PdfRewritePreservation.Assess(
            authored.ToBytes(),
            cleared.ToBytes(),
            new PdfRewritePreservationOptions { PreserveRevisionStructure = false });
        Assert.False(strictPreservation.IsPreserved);
        Assert.Contains(strictPreservation.Issues, static issue => issue.Feature == "CatalogActions.Count");

        PdfSanitizationResult sanitized = authored.ToDocument().Sanitize();
        Assert.Empty(sanitized.ToDocument().JavaScript.List());
        Assert.False(sanitized.ToDocument().Inspect().HasActiveContent);
    }

    [Fact]
    public void Reader_DecodesUtf16JavaScriptStreamAndHonorsConfiguredBudget() {
        const string script = "app.alert('Zażółć');";
        byte[] source = BuildJavaScriptStreamPdf("Startup", script);

        PdfJavaScript saved = Assert.Single(PdfDocument.Open(source).Read.JavaScripts());
        Assert.Equal("Startup", saved.Name);
        Assert.Equal(script, saved.Script);

        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxJavaScriptBytes = 8 }
        };
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Open(source, options).Read.JavaScripts());
        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(8, exception.Limit);
    }

    [Fact]
    public void Edit_RejectsMissingCommandsAndEmptyScriptSource() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("JavaScript validation"))
            .ToBytes();
        PdfDocument document = PdfDocument.Open(source);

        Assert.Throws<ArgumentException>(() => document.JavaScript.Edit(static _ => { }));
        Assert.Throws<ArgumentException>(() => document.JavaScript.AddOrReplace("Startup", string.Empty));
        Assert.Throws<ArgumentException>(() => document.JavaScript.AddOrReplace(string.Empty, "app.alert('x');"));
    }

    [Fact]
    public void Edit_FailsClosedWhenCatalogNamesDictionaryIsUnreadable() {
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names 99 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 4 >>",
            "%%EOF",
            string.Empty
        }));

        Assert.Throws<InvalidDataException>(() => PdfDocument.Open(source).JavaScript.AddOrReplace("Startup", "app.alert('x');"));
    }

    [Fact]
    public void Edit_MatchesWhitespaceKeysExactlyAndPreservesUntouchedScripts() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Exact key matching"))
            .ToBytes();
        PdfJavaScriptEditResult authored = PdfDocument.Open(source).JavaScript.Edit(scripts => scripts
            .AddOrReplace("Startup", "app.alert('plain');")
            .AddOrReplace(" Startup ", "app.alert('spaced');"));

        PdfJavaScriptEditResult removed = authored.ToDocument().JavaScript.Remove(" Startup ");

        PdfJavaScript remaining = Assert.Single(removed.JavaScripts);
        Assert.Equal("Startup", remaining.Name);
        Assert.Equal("app.alert('plain');", remaining.Script);
    }

    [Fact]
    public void Edit_ReplacesSourceWithoutDroppingNextActionsOrExtensionFields() {
        byte[] source = BuildJavaScriptActionChainPdf();

        PdfJavaScriptEditResult result = PdfDocument.Open(source).JavaScript
            .AddOrReplace("Startup", "app.alert('updated');");
        byte[] output = result.ToBytes();
        PdfReadDocument saved = PdfReadDocument.Open(output);

        Assert.True(result.PreservationReport.IsPreserved);
        Assert.Equal("app.alert('updated');", Assert.Single(saved.JavaScripts).Script);
        Assert.Contains(saved.CatalogActions, static action =>
            action.Name == "Startup.Next" && action.ActionType == "Launch" && action.IsChainedAction);
        Assert.Equal("extension-marker", ReadFirstJavaScriptAction(output).Get<PdfStringObj>("OfficeIMOExtension")?.Value);
        Assert.Equal(Encoding.ASCII.GetBytes("Startup"), ReadFirstJavaScriptKeyBytes(output));
    }

    [Fact]
    public void Reader_UsesPdfDocEncodingAndEnforcesCountAndAggregateBudgets() {
        byte[] source = BuildPdfDocEncodedJavaScriptPdf();
        PdfJavaScript script = Assert.Single(PdfReadDocument.Open(source).JavaScripts);
        Assert.Equal("\u2022", script.Script);

        PdfJavaScript controls = Assert.Single(PdfReadDocument.Open(BuildPdfDocControlEncodedJavaScriptPdf()).JavaScripts);
        Assert.Equal("\0\f\u0017", controls.Script);

        var countOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxJavaScripts = 1 }
        };
        PdfReadLimitException countException = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(BuildTwoScriptPdf(), countOptions));
        Assert.Equal(PdfReadLimitKind.JavaScripts, countException.Kind);

        var totalOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxTotalJavaScriptBytes = 1 }
        };
        PdfReadLimitException totalException = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(BuildTwoScriptPdf(), totalOptions));
        Assert.Equal(PdfReadLimitKind.JavaScriptBytes, totalException.Kind);

        PdfReadLimitException invalidTextException = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(BuildTwoInvalidTextStreamPdf(), totalOptions));
        Assert.Equal(PdfReadLimitKind.JavaScriptBytes, invalidTextException.Kind);
    }

    [Fact]
    public void Edit_EmitsByteSortedUnicodeKeysAndRetainsCustomReadLimitsForResult() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Unicode JavaScript keys"))
            .ToBytes();
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxJavaScriptBytes = 5_000_000,
                MaxTotalJavaScriptBytes = 6_000_000,
                MaxObjectCharacters = 10_000_000
            }
        };
        string largeScript = new string('x', 2_100_000);

        PdfJavaScriptEditResult result = PdfDocument.Open(source, options).JavaScript.Edit(scripts => scripts
            .AddOrReplace("Ā", largeScript)
            .AddOrReplace("ÿ", "app.alert('small');"));

        Assert.Equal(largeScript.Length, Assert.Single(result.JavaScripts, static script => script.Name == "Ā").Script.Length);
        Assert.Equal(2, result.ToDocument().JavaScript.List().Count);
        AssertNameTreeKeysAreByteSorted(result.ToBytes(), options);
    }

    [Fact]
    public void Edit_ExplicitlyBlocksEncryptedInputsEvenWithOwnerAuthorization() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted JavaScript edit"))
            .ToBytes();
        PdfDocument document = PdfDocument.Open(source, new PdfReadOptions { Password = "owner" });

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            document.JavaScript.AddOrReplace("Startup", "app.alert('x');"));

        Assert.Equal(PdfMutationOperation.ModifyJavaScript, exception.Plan.Operation);
        Assert.Contains(exception.Plan.BlockerCodes, static blocker => blocker.Contains("Encryption", StringComparison.Ordinal));

        PdfMutationBlockedException signedException = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfDocument.Open(PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf())
                .JavaScript.AddOrReplace("Startup", "app.alert('x');"));
        Assert.Contains("FullRewrite.AppendOnlyRequired", signedException.Plan.BlockerCodes);
    }

    [Fact]
    public void Edit_PreservesUntouchedOpaqueActionsAndFailsClosedForAmbiguousReplacement() {
        byte[] opaqueSource = BuildOpaqueJavaScriptPdf();

        byte[] output = PdfDocument.Open(opaqueSource).JavaScript
            .AddOrReplace("New", "app.alert('new');")
            .ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(output);

        Assert.Contains(objects.Values, static item =>
            item.Value is PdfStream stream && Encoding.ASCII.GetString(stream.Data) == "OPAQUE-JAVASCRIPT-BYTES");
        Assert.Equal("app.alert('new');", Assert.Single(PdfReadDocument.Open(output).JavaScripts).Script);

        Assert.Throws<InvalidDataException>(() => PdfDocument.Open(BuildDuplicateJavaScriptNamePdf())
            .JavaScript.AddOrReplace("Duplicate", "app.alert('replacement');"));
    }

    [Fact]
    public void EditResult_SnapshotsOperationsWhenTheCallbackRetainsItsSession() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Immutable JavaScript edit result"))
            .ToBytes();
        PdfJavaScriptEditSession? retained = null;

        PdfJavaScriptEditResult result = PdfDocument.Open(source).JavaScript.Edit(session => {
            retained = session;
            session.AddOrReplace("Startup", "app.alert('saved');");
        });
        retained!.Clear();

        Assert.Equal(new[] { "AddOrReplace:Startup" }, result.Operations);
        PdfJavaScript saved = Assert.Single(result.JavaScripts);
        Assert.Equal("Startup", saved.Name);
        Assert.Equal("app.alert('saved');", saved.Script);
    }

    [Fact]
    public void Edit_PreservesDuplicateRawKeyOrderWhenAddingAnUnrelatedScript() {
        byte[] output = PdfDocument.Open(BuildDuplicateJavaScriptNamePdf()).JavaScript
            .AddOrReplace("Unrelated", "app.alert('new');")
            .ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(output);
        PdfArray values = ReadJavaScriptNameTree(output);
        string[] duplicateSources = Enumerable.Range(0, values.Items.Count / 2)
            .Where(index => Assert.IsType<PdfStringObj>(PdfObjectLookup.Resolve(objects, values.Items[index * 2])).Value == "Duplicate")
            .Select(index => Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, values.Items[(index * 2) + 1]))
                .Get<PdfStringObj>("JS")?.Value)
            .OfType<string>()
            .ToArray();

        Assert.Equal(new[] { "one", "two" }, duplicateSources);
    }

    [Fact]
    public void Edit_UsesReaderNodeCountingForDirectRootAndIndirectLeaf() {
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 1 }
        };
        byte[] source = BuildIndirectJavaScriptLeafPdf();

        Assert.Equal("Existing", Assert.Single(PdfDocument.Open(source, options).Read.JavaScripts()).Name);
        PdfJavaScriptEditResult result = PdfDocument.Open(source, options).JavaScript
            .AddOrReplace("Added", "app.alert('added');");

        Assert.Equal(new[] { "Added", "Existing" }, result.JavaScripts.Select(static script => script.Name));
    }

    private static byte[] BuildJavaScriptStreamPdf(string name, string script) {
        byte[] encodedScript = Encoding.BigEndianUnicode.GetPreamble()
            .Concat(Encoding.BigEndianUnicode.GetBytes(script))
            .ToArray();
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(" + name + ") 5 0 R] >> >> >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /S /JavaScript /JS 6 0 R >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Length " + encodedScript.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(encodedScript, 0, encodedScript.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R /Size 7 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildJavaScriptActionChainPdf() => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.7",
        "1 0 obj",
        "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Startup) 5 0 R] >> >> >>",
        "endobj",
        "2 0 obj",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "endobj",
        "3 0 obj",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
        "endobj",
        "5 0 obj",
        "<< /S /JavaScript /JS (app.alert('old')) /Next 6 0 R /OfficeIMOExtension (extension-marker) >>",
        "endobj",
        "6 0 obj",
        "<< /S /Launch /F (tool.exe) >>",
        "endobj",
        "trailer",
        "<< /Root 1 0 R /Size 7 >>",
        "%%EOF",
        string.Empty
    }));

    private static byte[] BuildPdfDocEncodedJavaScriptPdf() => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.7",
        "1 0 obj",
        "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Startup) 5 0 R] >> >> >>",
        "endobj",
        "2 0 obj",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "endobj",
        "3 0 obj",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
        "endobj",
        "5 0 obj",
        "<< /S /JavaScript /JS <80> >>",
        "endobj",
        "trailer",
        "<< /Root 1 0 R /Size 6 >>",
        "%%EOF",
        string.Empty
    }));

    private static byte[] BuildPdfDocControlEncodedJavaScriptPdf() => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.7",
        "1 0 obj",
        "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Startup) 5 0 R] >> >> >>",
        "endobj",
        "2 0 obj",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "endobj",
        "3 0 obj",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
        "endobj",
        "5 0 obj",
        "<< /S /JavaScript /JS <000C17> >>",
        "endobj",
        "trailer",
        "<< /Root 1 0 R /Size 6 >>",
        "%%EOF",
        string.Empty
    }));

    private static byte[] BuildTwoScriptPdf() => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.7",
        "1 0 obj",
        "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(One) 5 0 R (Two) 6 0 R] >> >> >>",
        "endobj",
        "2 0 obj",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "endobj",
        "3 0 obj",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
        "endobj",
        "5 0 obj",
        "<< /S /JavaScript /JS (a) >>",
        "endobj",
        "6 0 obj",
        "<< /S /JavaScript /JS (b) >>",
        "endobj",
        "trailer",
        "<< /Root 1 0 R /Size 7 >>",
        "%%EOF",
        string.Empty
    }));

    private static byte[] BuildTwoInvalidTextStreamPdf() {
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(One) 5 0 R (Two) 6 0 R] >> >> >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /S /JavaScript /JS 7 0 R >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /S /JavaScript /JS 8 0 R >>\nendobj\n");
        WriteAscii(output, "7 0 obj\n<< /Length 1 >>\nstream\n");
        output.WriteByte(0x7F);
        WriteAscii(output, "\nendstream\nendobj\n8 0 obj\n<< /Length 1 >>\nstream\n");
        output.WriteByte(0x7F);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R /Size 9 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildIndirectJavaScriptLeafPdf() => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.7",
        "1 0 obj",
        "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Kids [5 0 R] >> >> >>",
        "endobj",
        "2 0 obj",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "endobj",
        "3 0 obj",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
        "endobj",
        "5 0 obj",
        "<< /Names [(Existing) 6 0 R] >>",
        "endobj",
        "6 0 obj",
        "<< /S /JavaScript /JS (app.alert('existing')) >>",
        "endobj",
        "trailer",
        "<< /Root 1 0 R /Size 7 >>",
        "%%EOF",
        string.Empty
    }));

    private static byte[] BuildOpaqueJavaScriptPdf() => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.7",
        "1 0 obj",
        "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Opaque) 5 0 R] >> >> >>",
        "endobj",
        "2 0 obj",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "endobj",
        "3 0 obj",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
        "endobj",
        "5 0 obj",
        "<< /S /JavaScript /JS 6 0 R >>",
        "endobj",
        "6 0 obj",
        "<< /Length 23 /Filter /DCTDecode >>",
        "stream",
        "OPAQUE-JAVASCRIPT-BYTES",
        "endstream",
        "endobj",
        "trailer",
        "<< /Root 1 0 R /Size 7 >>",
        "%%EOF",
        string.Empty
    }));

    private static byte[] BuildDuplicateJavaScriptNamePdf() => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.7",
        "1 0 obj",
        "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Duplicate) 5 0 R (Duplicate) 6 0 R] >> >> >>",
        "endobj",
        "2 0 obj",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "endobj",
        "3 0 obj",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
        "endobj",
        "5 0 obj",
        "<< /S /JavaScript /JS (one) >>",
        "endobj",
        "6 0 obj",
        "<< /S /JavaScript /JS (two) >>",
        "endobj",
        "trailer",
        "<< /Root 1 0 R /Size 7 >>",
        "%%EOF",
        string.Empty
    }));

    private static void AssertNameTreeKeysAreByteSorted(byte[] pdf, PdfReadOptions? options = null) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf, options);
        PdfDictionary catalog = FindCatalog(objects);
        PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
        PdfDictionary tree = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["JavaScript"]));
        PdfArray values = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, tree.Items["Names"]));
        byte[][] keys = Enumerable.Range(0, values.Items.Count / 2)
            .Select(index => Assert.IsType<PdfStringObj>(PdfObjectLookup.Resolve(objects, values.Items[index * 2])).RawBytes)
            .ToArray();
        for (int i = 1; i < keys.Length; i++) Assert.True(CompareBytes(keys[i - 1], keys[i]) <= 0);
    }

    private static PdfDictionary ReadFirstJavaScriptAction(byte[] pdf) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfArray values = ReadJavaScriptNameTree(pdf);
        return Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, values.Items[1]));
    }

    private static byte[] ReadFirstJavaScriptKeyBytes(byte[] pdf) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfArray values = ReadJavaScriptNameTree(pdf);
        return Assert.IsType<PdfStringObj>(PdfObjectLookup.Resolve(objects, values.Items[0])).RawBytes;
    }

    private static PdfArray ReadJavaScriptNameTree(byte[] pdf) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfDictionary catalog = FindCatalog(objects);
        PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
        PdfDictionary tree = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["JavaScript"]));
        return Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, tree.Items["Names"]));
    }

    private static PdfDictionary FindCatalog(Dictionary<int, PdfIndirectObject> objects) => Assert.IsType<PdfDictionary>(Assert.Single(
        objects.Values,
        static item => (item.Value as PdfDictionary)?.Get<PdfName>("Type")?.Name == "Catalog").Value);

    private static int CompareBytes(byte[] left, byte[] right) {
        int count = Math.Min(left.Length, right.Length);
        for (int i = 0; i < count; i++) {
            int comparison = left[i].CompareTo(right[i]);
            if (comparison != 0) return comparison;
        }
        return left.Length.CompareTo(right.Length);
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
