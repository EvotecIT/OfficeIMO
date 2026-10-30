using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRotationNormalizationTests {
    [Fact]
    public void NormalizationPlanReportsResourcesAndContentChanges() {
        var plan = PdfMutationPlanner.Plan(CreateGeometrySource(90), PdfMutationOperation.NormalizePageRotation);
        Assert.Contains(PdfMutationStructure.PageResources, plan.AffectedStructures);
        Assert.Contains(PdfMutationStructure.PageContent, plan.AffectedStructures);
    }

    [Theory]
    [InlineData(90)] [InlineData(180)] [InlineData(270)]
    public void NavigationRetainsFitFamiliesNullAxesAndPageActionChains(int rotation) {
        byte[] source = CreateGeometrySource(rotation);
        var read = PdfReadDocument.Open(source);
        var (objects, trailer) = PdfSyntax.ParseObjects(source);
        int first = read.Pages[0].ObjectNumber, second = read.Pages[1].ObjectNumber;
        var destinations = new PdfDictionary();
        foreach (string mode in new[] { "FitH", "FitBH", "FitV", "FitBV" }) {
            var destination = new PdfArray(); destination.Items.Add(new PdfReference(first, objects[first].Generation));
            destination.Items.Add(new PdfName(mode)); destination.Items.Add(new PdfNumber(mode.EndsWith("H") ? 400 : 100));
            destinations.Items[mode] = destination;
            var retained = new PdfArray(); retained.Items.Add(destination.Items[0]); retained.Items.Add(destination.Items[1]); retained.Items.Add(PdfNull.Instance);
            destinations.Items[mode + "Null"] = retained;
        }
        foreach (bool x in new[] { true, false }) {
            var destination = new PdfArray(); destination.Items.Add(new PdfReference(first, objects[first].Generation));
            destination.Items.Add(new PdfName("XYZ")); destination.Items.Add(x ? new PdfNumber(100) : PdfNull.Instance);
            destination.Items.Add(x ? PdfNull.Instance : new PdfNumber(400)); destination.Items.Add(new PdfNumber(2));
            destinations.Items[x ? "PartialX" : "PartialY"] = destination;
        }
        var action = new PdfDictionary(); action.Items["S"] = new PdfName("GoTo"); action.Items["D"] = CloneDestination(100, 400);
        var next = new PdfDictionary(); next.Items["S"] = new PdfName("GoTo"); next.Items["D"] = destinations.Items["PartialX"];
        action.Items["Next"] = next;
        var actions = new PdfDictionary(); actions.Items["O"] = action;
        ((PdfDictionary)objects[second].Value).Items["AA"] = actions;
        int root = read.Security.RootObjectNumber!.Value;
        ((PdfDictionary)objects[root].Value).Items["Dests"] = destinations;
        source = PdfIncrementalObjectWriter.Append(source, objects, read.Security, trailer, new[] { root, second });
        var before = PdfDocument.Load(source);
        var after = before.Pages.NormalizeRotation(PdfPageSelection.From(1));
        var old = before.GetPageLayouts()[0]; var current = after.GetPageLayouts()[0];
        var visual = PdfVisualCoordinateMapper.TransformBounds(old.Geometry.EffectiveBox!, rotation, 100, 400, 100, 400, 2);
        var point = current.MapVisualPointToUserSpace(visual.Left, visual.Top);
        var output = PdfReadDocument.Open(after.ToBytes());
        var catalog = Assert.IsType<PdfDictionary>(output.Objects[output.Security.RootObjectNumber!.Value].Value);
        var names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, catalog.Items["Dests"]));
        bool swap = rotation != 180;
        foreach (string mode in new[] { "FitH", "FitBH", "FitV", "FitBV" }) {
            var destination = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(output.Objects, names.Items[mode]));
            bool horizontal = mode.EndsWith("H");
            Assert.Equal(swap ? mode.Substring(0, mode.Length - 1) + (horizontal ? "V" : "H") : mode, ((PdfName)destination.Items[1]).Name);
            Assert.Equal(horizontal != swap ? point.Y : point.X, ((PdfNumber)destination.Items[2]).Value, 6);
            Assert.Equal(3, destination.Items.Count);
            var retained = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(output.Objects, names.Items[mode + "Null"]));
            Assert.Equal(((PdfName)destination.Items[1]).Name, ((PdfName)retained.Items[1]).Name);
            Assert.IsType<PdfNull>(retained.Items[2]);
        }
        foreach (bool x in new[] { true, false }) {
            var destination = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(output.Objects, names.Items[x ? "PartialX" : "PartialY"]));
            int known = x != swap ? 2 : 3;
            Assert.Equal(known == 2 ? point.X : point.Y, ((PdfNumber)destination.Items[known]).Value, 6);
            Assert.IsType<PdfNull>(destination.Items[known == 2 ? 3 : 2]);
            Assert.Equal(2, ((PdfNumber)destination.Items[4]).Value);
        }
        var page = Assert.IsType<PdfDictionary>(output.Objects[output.Pages[1].ObjectNumber].Value);
        var aa = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, page.Items["AA"]));
        var open = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, aa.Items["O"]));
        var target = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(output.Objects, open.Items["D"]));
        Assert.Equal(point.X, ((PdfNumber)target.Items[2]).Value, 6);
        Assert.Equal(point.Y, ((PdfNumber)target.Items[3]).Value, 6);
        var chained = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, open.Items["Next"]));
        var chainedTarget = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(output.Objects, chained.Items["D"]));
        Assert.IsType<PdfNull>(chainedTarget.Items[swap ? 2 : 3]);
        Assert.Equal(swap ? point.Y : point.X, ((PdfNumber)chainedTarget.Items[swap ? 3 : 2]).Value, 6);

        PdfArray CloneDestination(double x, double y) {
            var value = new PdfArray(); value.Items.Add(new PdfReference(first, objects[first].Generation));
            value.Items.Add(new PdfName("XYZ")); value.Items.Add(new PdfNumber(x)); value.Items.Add(new PdfNumber(y)); value.Items.Add(PdfNull.Instance);
            return value;
        }
    }

    [Theory]
    [InlineData(false, 1)] [InlineData(true, 1)] [InlineData(false, 3)]
    public void UnsignedSignaturePlaceholdersRetainFieldsAndExplicitAppendOnlyRestrictions(bool invisible, int signatureFlags) {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Sign later")).Forms.Edit(edit => edit.Create(
            new PdfFormFieldCreateOptions { Name = "Approval", Kind = PdfFormFieldCreationKind.Signature,
                PageNumber = 1, X = 100, Y = 400, Width = 160, Height = 40 })).ToBytes();
        var read = PdfReadDocument.Open(source);
        var (objects, trailer) = PdfSyntax.ParseObjects(source);
        int pageNumber = read.Pages[0].ObjectNumber, fieldNumber = read.Security.SignatureFieldObjectNumbers.Single();
        ((PdfDictionary)objects[pageNumber].Value).Items["Rotate"] = new PdfNumber(90);
        var field = (PdfDictionary)objects[fieldNumber].Value;
        if (invisible) { field.Items["Rect"] = Box(0, 0, 0, 0); field.Items["F"] = new PdfNumber(2); field.Items.Remove("AP"); }
        else {
            int appearanceNumber = objects.Keys.Max() + 1;
            var appearance = new PdfDictionary(); appearance.Items["Type"] = new PdfName("XObject"); appearance.Items["Subtype"] = new PdfName("Form");
            appearance.Items["BBox"] = Box(0, 0, 160, 40); appearance.Items["Resources"] = new PdfDictionary();
            objects[appearanceNumber] = new PdfIndirectObject(appearanceNumber, 0, new PdfStream(appearance,
                System.Text.Encoding.ASCII.GetBytes("0 G 1 w 0.5 0.5 159 39 re S")));
            var ap = new PdfDictionary(); ap.Items["N"] = new PdfReference(appearanceNumber, 0); field.Items["AP"] = ap;
        }
        var catalog = (PdfDictionary)objects[read.Security.RootObjectNumber!.Value].Value;
        var form = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["AcroForm"]));
        form.Items["SigFlags"] = new PdfNumber(signatureFlags);
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (map, security) => {
            map.Clear(); foreach (var item in objects) map.Add(item.Key, item.Value); return security.InfoObjectNumber;
        });
        PdfDocument document = PdfDocument.Load(source);
        Assert.Equal(0, document.Inspect().Security.SignatureValueCount);
        if ((signatureFlags & 2) != 0) {
            Assert.False(document.PlanMutation(PdfMutationOperation.NormalizePageRotation).CanExecute);
            Assert.Throws<PdfMutationBlockedException>(() => document.Pages.NormalizeRotation());
        } else {
            var normalized = document.Pages.NormalizeRotationResult().RequireValue();
            var output = normalized.Inspect();
            Assert.Equal(0, output.Pages[0].RotationDegrees);
            Assert.Equal(PdfFormFieldKind.Signature, Assert.Single(output.FormFields).Kind);
            Assert.Equal("Approval", output.FormFields[0].Name);
            Assert.Equal(0, output.Security.SignatureValueCount);
            if (invisible) {
                var parsed = PdfReadDocument.Open(normalized.ToBytes());
                var signature = parsed.Objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
                    .Single(item => item.Get<PdfName>("FT")?.Name == "Sig");
                var rect = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Objects, signature.Items["Rect"]));
                Assert.Equal(((PdfNumber)rect.Items[0]).Value, ((PdfNumber)rect.Items[2]).Value);
                Assert.Equal(((PdfNumber)rect.Items[1]).Value, ((PdfNumber)rect.Items[3]).Value);
            }
        }
        Assert.Equal(source, document.ToBytes());
    }

    [Theory]
    [InlineData(90)] [InlineData(180)] [InlineData(270)]
    public void InheritedPatternResourcesAreTransformedOnlyForSelectedPages(int rotation) {
        byte[] source = PdfDocument.Create(compose => {
            for (int index = 0; index < 2; index++) compose.Page(page => page.Size(300, 500));
        }).ToBytes();
        var read = PdfReadDocument.Open(source);
        var (objects, trailer) = PdfSyntax.ParseObjects(source);
        int next = objects.Keys.Max() + 1;
        int patternNumber = next++, contentsNumber = next++;
        var pattern = new PdfDictionary(); pattern.Items["Type"] = new PdfName("Pattern");
        pattern.Items["PatternType"] = new PdfNumber(1); pattern.Items["PaintType"] = new PdfNumber(1); pattern.Items["TilingType"] = new PdfNumber(1);
        pattern.Items["BBox"] = Box(0, 0, 16, 28); pattern.Items["XStep"] = new PdfNumber(16); pattern.Items["YStep"] = new PdfNumber(28);
        pattern.Items["Matrix"] = Box(1, 0, 0, 1, 7, 9); pattern.Items["Resources"] = new PdfDictionary();
        byte[] tile = System.Text.Encoding.ASCII.GetBytes("1 0 0 rg 0 0 6 13 re f 0 0 1 rg 9 18 7 10 re f");
        objects[patternNumber] = new PdfIndirectObject(patternNumber, 0, new PdfStream(pattern, tile));
        objects[contentsNumber] = new PdfIndirectObject(contentsNumber, 0, new PdfStream(new PdfDictionary(),
            System.Text.Encoding.ASCII.GetBytes("/Pattern cs /P1 scn 20 50 240 360 re f")));
        var patterns = new PdfDictionary(); patterns.Items["P1"] = new PdfReference(patternNumber, 0);
        var resources = new PdfDictionary(); resources.Items["Pattern"] = patterns;
        var changed = new HashSet<int> { patternNumber, contentsNumber };
        foreach (var readPage in read.Pages) {
            var page = (PdfDictionary)objects[readPage.ObjectNumber].Value;
            int parentNumber = ((PdfReference)page.Items["Parent"]).ObjectNumber;
            var parent = (PdfDictionary)objects[parentNumber].Value;
            parent.Items["Resources"] = resources; parent.Items["Rotate"] = new PdfNumber(rotation);
            page.Items.Remove("Resources"); page.Items.Remove("Rotate"); page.Items["Contents"] = new PdfReference(contentsNumber, 0);
            changed.Add(parentNumber); changed.Add(readPage.ObjectNumber);
        }
        source = PdfIncrementalObjectWriter.Append(source, objects, read.Security, trailer, changed);
        var output = PdfReadDocument.Open(PdfDocument.Load(source).Pages.NormalizeRotation(PdfPageSelection.From(1)).ToBytes());
        var selected = (PdfDictionary)output.Objects[output.Pages[0].ObjectNumber].Value;
        var localResources = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, selected.Items["Resources"]));
        var localPatterns = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, localResources.Items["Pattern"]));
        var selectedPattern = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(output.Objects, localPatterns.Items["P1"]));
        double[] expected = rotation == 90 ? new[] { 0D, -1, 1, 0, 9, 293 } : rotation == 180
            ? new[] { -1D, 0, 0, -1, 293, 491 } : new[] { 0D, 1, -1, 0, 491, 7 };
        Assert.Equal(expected, ((PdfArray)selectedPattern.Dictionary.Items["Matrix"]).Items.Cast<PdfNumber>().Select(item => item.Value));
        Assert.Equal(tile, selectedPattern.Data);
        Assert.Equal(rotation, output.Pages[1].GetRotationDegrees());
        var sibling = (PdfDictionary)output.Objects[output.Pages[1].ObjectNumber].Value;
        Assert.False(sibling.Items.ContainsKey("Resources"));
        var inherited = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, sibling.Items["Parent"]));
        var inheritedResources = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, inherited.Items["Resources"]));
        var inheritedPatterns = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, inheritedResources.Items["Pattern"]));
        var originalPattern = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(output.Objects, inheritedPatterns.Items["P1"]));
        Assert.Equal(new[] { 1D, 0, 0, 1, 7, 9 }, ((PdfArray)originalPattern.Dictionary.Items["Matrix"]).Items.Cast<PdfNumber>().Select(item => item.Value));
    }

    [Theory]
    [InlineData(90)] [InlineData(180)] [InlineData(270)]
    public void FormValuesStatesAndEditingSurviveNormalization(int rotation) {
        PdfDocument source = PdfDocument.Create().TaggedPdfCatalogMarkers()
            .TextField("Person.Name", width: 180, height: 24, value: "Ada")
            .CheckBox("Accept", isChecked: true, size: 16).Pages.RotateBy(rotation);
        PdfDocument normalized = source.Pages.NormalizeRotation();
        Assert.True(normalized.Inspect().HasTaggedContent);
        Assert.Equal("Ada", normalized.Inspect().FormFields.Single(field => field.Name == "Person.Name").Value);
        Assert.Equal(source.Inspect().FormFields.Single(field => field.Name == "Accept").Value,
            normalized.Inspect().FormFields.Single(field => field.Name == "Accept").Value);
        var fields = normalized.Forms.FillResult(new Dictionary<string, string> { ["Person.Name"] = "Grace", ["Accept"] = "Off" }).RequireValue().Inspect().FormFields;
        Assert.Equal("Grace", fields.Single(field => field.Name == "Person.Name").Value);
        Assert.Equal("Off", fields.Single(field => field.Name == "Accept").Value);
        Assert.Equal(rotation, source.Inspect().Pages[0].RotationDegrees);
    }

    [Fact]
    public void UnsupportedVisibleAppearanceFailsWithoutChangingInput() {
        PdfDocument source = PdfDocument.Create().Paragraph(p => p.Text("Keep original"));
        source = source.Annotations.Add(new PdfAnnotationCreateOptions { PageNumber = 1, Subtype = "Text", Name = "note",
            Rectangle = new[] { 100D, 400D, 160D, 460D }, GenerateAppearance = false }).ToDocument().Pages.RotateBy(90);
        byte[] bytes = source.ToBytes();
        var result = source.Pages.NormalizeRotationResult();
        Assert.False(result.Succeeded);
        Assert.Equal(bytes, source.ToBytes());
        Assert.Throws<NotSupportedException>(() => source.Pages.NormalizeRotation());
    }

    [Fact]
    public void SharedTagClassesKeepUnselectedBoundsAndCancellationLeavesInputUnchanged() {
        byte[] source = PdfDocument.Create().TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph.Text("Tagged first")).PageBreak()
            .Paragraph(paragraph => paragraph.Text("Tagged second")).ToBytes();
        PdfReadDocument read = PdfReadDocument.Open(source);
        var (objects, trailer) = PdfSyntax.ParseObjects(source);
        int first = read.Pages[0].ObjectNumber, second = read.Pages[1].ObjectNumber;
        var elements = objects.Values.Where(item => item.Value is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("Type")?.Name == "StructElem" && dictionary.Get<PdfReference>("Pg") is not null).ToArray();
        Assert.Contains(elements, item => ((PdfDictionary)item.Value).Get<PdfReference>("Pg")!.ObjectNumber == first);
        var root = objects.Values.Single(item => item.Value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "StructTreeRoot");
        var attributes = new PdfDictionary(); attributes.Items["O"] = new PdfName("Layout"); attributes.Items["BBox"] = Box(100, 400, 160, 460);
        var classes = new PdfDictionary(); classes.Items["Bounds"] = attributes;
        ((PdfDictionary)root.Value).Items["ClassMap"] = classes;
        foreach (var item in elements) ((PdfDictionary)item.Value).Items["C"] = new PdfName("Bounds");
        ((PdfDictionary)objects[first].Value).Items["Rotate"] = new PdfNumber(90);
        source = PdfIncrementalObjectWriter.Append(source, objects, read.Security, trailer,
            elements.Select(item => item.ObjectNumber).Concat(new[] { first, root.ObjectNumber }).ToArray());
        PdfDocument document = PdfDocument.Load(source);
        Assert.ThrowsAny<OperationCanceledException>(() => document.Pages.NormalizeRotationResult(cancellationToken: new CancellationToken(true)));
        Assert.Equal(source, document.ToBytes());
        PdfDocument output = document.Pages.NormalizeRotation(PdfPageSelection.From(1));
        var normalized = PdfReadDocument.Open(output.ToBytes());
        var normalizedElements = normalized.Objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
            .Where(item => item.Get<PdfName>("Type")?.Name == "StructElem" && item.Get<PdfName>("C") is not null).ToArray();
        int normalizedFirst = normalized.Pages[0].ObjectNumber, normalizedSecond = normalized.Pages[1].ObjectNumber;
        Assert.All(normalizedElements.Where(item => item.Get<PdfReference>("Pg")!.ObjectNumber == normalizedSecond),
            item => Assert.Equal("Bounds", item.Get<PdfName>("C")!.Name));
        Assert.All(normalizedElements.Where(item => item.Get<PdfReference>("Pg")!.ObjectNumber == normalizedFirst),
            item => Assert.NotEqual("Bounds", item.Get<PdfName>("C")!.Name));
        var normalizedRoot = normalized.Objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
            .Single(item => item.Get<PdfName>("Type")?.Name == "StructTreeRoot");
        var normalizedClasses = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(normalized.Objects, normalizedRoot.Items["ClassMap"]));
        var originalBounds = Assert.IsType<PdfDictionary>(normalizedClasses.Items["Bounds"]);
        Assert.Equal(new[] { 100D, 400D, 160D, 460D }, Assert.IsType<PdfArray>(originalBounds.Items["BBox"]).Items.Cast<PdfNumber>().Select(item => item.Value));
    }

    [Theory]
    [InlineData("OfficeIMO.TestAssets/MultilingualLayout/latin-0-native.pdf")]
    [InlineData("OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/openpreserve-pdfa1b-text.pdf")]
    public void IndependentProducerRetainsTextTagsAndPhysicalGeometry(string fixture) {
        byte[] source = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), fixture));
        PdfDocument rotated = PdfDocument.Load(source).Pages.RotateBy(90);
        var before = rotated.Inspect();
        var result = rotated.Pages.NormalizeRotationResult();
        PdfDocument normalized = result.RequireValue();
        var after = normalized.Inspect();
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.MutationPlan!.ExecutionMode);
        Assert.Equal(before.PageCount, after.PageCount);
        Assert.Equal(before.HasTaggedContent, after.HasTaggedContent);
        using var independentBefore = UglyToad.PdfPig.PdfDocument.Open(rotated.ToBytes());
        using var independent = UglyToad.PdfPig.PdfDocument.Open(normalized.ToBytes());
        for (int index = 0; index < before.PageCount; index++) {
            Assert.Equal(0, after.Pages[index].RotationDegrees);
            Assert.Equal(rotated.GetPageLayouts()[index].VisualWidth, normalized.GetPageLayouts()[index].VisualWidth, 6);
            Assert.Equal(rotated.GetPageLayouts()[index].VisualHeight, normalized.GetPageLayouts()[index].VisualHeight, 6);
            Assert.Equal(independentBefore.GetPage(index + 1).Letters.Select(letter => letter.Value),
                independent.GetPage(index + 1).Letters.Select(letter => letter.Value));
        }
        Assert.Equal(after.PageCount, independent.NumberOfPages);
        Assert.Equal(0, independent.GetPage(1).Rotation.Value);
    }

    [Theory]
    [InlineData(90)] [InlineData(180)] [InlineData(270)]
    public void InheritedRotationTransformsSelectedPageBoxesAnnotationsAndIncomingLinks(int rotation) {
        byte[] source = CreateGeometrySource(rotation);
        PdfDocument before = PdfDocument.Load(source);
        PdfDocument after = before.Pages.NormalizeRotation(PdfPageSelection.From(1));
        Assert.Equal(source, before.ToBytes());
        var inspection = after.Inspect();
        Assert.Equal(0, inspection.Pages[0].RotationDegrees);
        Assert.Equal(rotation, inspection.Pages[1].RotationDegrees);
        Assert.Equal(2, inspection.Pages[0].UserUnit);
        Assert.Equal(before.GetPageLayouts()[0].VisualWidth, after.GetPageLayouts()[0].VisualWidth, 6);
        Assert.Equal(before.GetPageLayouts()[0].VisualHeight, after.GetPageLayouts()[0].VisualHeight, 6);
        var oldLayout = before.GetPageLayouts()[0];
        var newLayout = after.GetPageLayouts()[0];
        var originalAnnotation = before.Inspect().Annotations.Single(item => item.Name == "rectangle");
        var annotation = inspection.Annotations.Single(item => item.Name == "rectangle");
        var oldVisual = oldLayout.MapUserSpaceRectangleToVisual(originalAnnotation.X1, originalAnnotation.Y1, originalAnnotation.X2, originalAnnotation.Y2);
        var newVisual = newLayout.MapUserSpaceRectangleToVisual(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2);
        Assert.Equal(oldVisual.Left, newVisual.Left, 6); Assert.Equal(oldVisual.Top, newVisual.Top, 6);
        Assert.Equal(oldVisual.Width, newVisual.Width, 6); Assert.Equal(oldVisual.Height, newVisual.Height, 6);
        var read = PdfReadDocument.Open(after.ToBytes());
        var objects = read.Objects;
        var second = (PdfDictionary)objects[read.Pages[1].ObjectNumber].Value;
        var annotations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, second.Items["Annots"]));
        var link = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, annotations.Items[0]));
        var destination = Assert.IsType<PdfArray>(link.Items["Dest"]);
        double x = ((PdfNumber)destination.Items[2]).Value, y = ((PdfNumber)destination.Items[3]).Value;
        // Destination XYZ coordinates denote a point, not a one-unit rectangle rotated about that point.
        var originalPoint = PdfVisualCoordinateMapper.TransformBounds(oldLayout.Geometry.EffectiveBox!, rotation, 100, 400, 100, 400, 2);
        var movedPoint = PdfVisualCoordinateMapper.TransformBounds(newLayout.Geometry.EffectiveBox!, 0, x, y, x, y, 2);
        Assert.Equal(originalPoint.Left, movedPoint.Left, 6); Assert.Equal(originalPoint.Top, movedPoint.Top, 6);
        Assert.Contains("Page geometry", PdfTextExtractor.ExtractAllText(after.ToBytes()));
    }

    private static byte[] CreateGeometrySource(int rotation) {
        PdfDocument document = PdfDocument.Create(compose => {
            for (int i = 0; i < 2; i++) compose.Page(page => page.Size(600, 800).Content(content => content.Text("Page geometry")));
        });
        document = document.Annotations.Add(new PdfAnnotationCreateOptions { PageNumber = 1, Subtype = "Square", Name = "rectangle",
            Rectangle = new[] { 100D, 400D, 160D, 460D }, GenerateAppearance = true }).ToDocument();
        byte[] source = document.ToBytes();
        PdfReadDocument read = PdfReadDocument.Open(source);
        var (objects, trailer) = PdfSyntax.ParseObjects(source);
        int first = read.Pages[0].ObjectNumber, second = read.Pages[1].ObjectNumber;
        var page = (PdfDictionary)objects[first].Value;
        page.Items["MediaBox"] = Box(-100, 200, 500, 1000);
        page.Items["CropBox"] = Box(-50, 240, 450, 960);
        page.Items["TrimBox"] = Box(0, 300, 400, 900);
        page.Items["BleedBox"] = Box(-20, 280, 420, 920);
        page.Items["UserUnit"] = new PdfNumber(2);
        page.Items.Remove("Rotate");
        int parent = ((PdfReference)page.Items["Parent"]).ObjectNumber;
        ((PdfDictionary)objects[parent].Value).Items["Rotate"] = new PdfNumber(rotation);
        var link = new PdfDictionary(); link.Items["Type"] = new PdfName("Annot"); link.Items["Subtype"] = new PdfName("Link");
        link.Items["NM"] = new PdfStringObj("incoming"); link.Items["Rect"] = Box(10, 10, 30, 30);
        var destination = new PdfArray(); destination.Items.Add(new PdfReference(first, objects[first].Generation));
        destination.Items.Add(new PdfName("XYZ")); destination.Items.Add(new PdfNumber(100)); destination.Items.Add(new PdfNumber(400)); destination.Items.Add(PdfNull.Instance);
        link.Items["Dest"] = destination;
        var annotations = new PdfArray(); annotations.Items.Add(link);
        ((PdfDictionary)objects[second].Value).Items["Annots"] = annotations;
        return PdfIncrementalObjectWriter.Append(source, objects, read.Security, trailer, new[] { first, second, parent });
    }

    private static PdfArray Box(params double[] coordinates) {
        var box = new PdfArray();
        foreach (double coordinate in coordinates) box.Items.Add(new PdfNumber(coordinate));
        return box;
    }
}
