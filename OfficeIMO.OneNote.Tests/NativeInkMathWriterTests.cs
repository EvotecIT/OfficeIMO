using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote.Tests;

public sealed class NativeInkMathWriterTests {
    [Fact]
    public void RoundTripsTypedInkAndNativeRecognitionTree() {
        var section = new OneNoteSection { Name = "Ink" };
        var page = new OneNotePage { Title = "Ink" };
        var ink = new OneNoteInk { Layout = new OneNoteLayout { X = 1.25, Y = 2.5 } };
        var stroke = new OfficeInkStroke {
            Color = OfficeColor.SteelBlue,
            Width = 0.04,
            Height = 0.06,
            Opacity = 0.75,
            Bias = OfficeInkBias.Handwriting,
            FitToCurve = true,
            LanguageId = 1033,
            RecognizedText = "hello"
        };
        stroke.RecognitionAlternatives.Add("hello");
        stroke.RecognitionAlternatives.Add("hullo");
        stroke.RecognitionAlternatives.Add("hello");
        stroke.AddPoint(0.1, 0.2, 0.25).AddPoint(0.4, 0.8, 1.0).AddPoint(1.1, 0.6, 0.5);
        ink.Ink.Add(stroke);
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNotePage resultPage = Assert.Single(roundTrip.Pages);
        OneNoteInk result = Assert.IsType<OneNoteInk>(Assert.Single(resultPage.Outlines).Children.Single());
        OfficeInkStroke resultStroke = Assert.Single(result.Strokes);

        Assert.NotNull(result.PreservedInkBoundingBox);
        OneNoteOutline resultOutline = Assert.Single(resultPage.Outlines);
        Assert.Equal(1.25, resultOutline.Layout!.X!.Value + result.Layout!.X!.Value, 5);
        Assert.Equal(2.5, resultOutline.Layout.Y!.Value + result.Layout.Y!.Value, 5);
        Assert.Equal(3, resultStroke.Points.Count);
        Assert.Equal(0.1, resultStroke.Points[0].X, 3);
        Assert.Equal(0.8, resultStroke.Points[1].Y, 3);
        Assert.Equal(0.25, resultStroke.Points[0].Pressure!.Value, 3);
        Assert.Equal(OfficeColor.SteelBlue, resultStroke.Color);
        Assert.Equal(0.04, resultStroke.Width, 3);
        Assert.Equal(0.06, resultStroke.Height, 3);
        Assert.Equal(0.75, resultStroke.Opacity, 2);
        Assert.Equal(OfficeInkBias.Handwriting, resultStroke.Bias);
        Assert.True(resultStroke.FitToCurve);
        Assert.Equal(1033U, resultStroke.LanguageId);
        Assert.Equal("hello", resultStroke.RecognizedText);
        Assert.Equal(new[] { "hello", "hullo" }, resultStroke.RecognitionAlternatives);

        OneNoteMaterializedObjectSpace resultSpace = roundTrip.PreservationState!.GetPageSpace(resultPage)!;
        OneNoteRevisionStoreObject pageNode = resultSpace.GetObject(resultPage.PreservationIds.PageNodeId!)!;
        OneNoteRevisionStoreObject recognitionRoot = Referenced(resultSpace, pageNode, OneNoteSchema.PageRecognizedTextContainer);
        OneNoteRevisionStoreObject recognitionLine = Referenced(resultSpace, recognitionRoot, OneNoteSchema.RecognizedTextChildNodes);
        OneNoteRevisionStoreObject recognitionBlock = Referenced(resultSpace, recognitionLine, OneNoteSchema.RecognizedTextChildNodes);
        OneNoteRevisionStoreObject recognitionWord = Referenced(resultSpace, recognitionBlock, OneNoteSchema.RecognizedTextChildNodes);
        Assert.Equal(OneNoteSchema.JcidRecognizedTextRoot, recognitionRoot.Jcid.Value);
        Assert.Equal(OneNoteSchema.JcidRecognizedTextLine, recognitionLine.Jcid.Value);
        Assert.Equal(OneNoteSchema.JcidRecognizedTextBlock, recognitionBlock.Jcid.Value);
        Assert.Equal(OneNoteSchema.JcidRecognizedTextWord, recognitionWord.Jcid.Value);
    }

    [Fact]
    public void WritesNativeTransparencyAndPreservesHighlighterEffectiveOpacity() {
        var section = new OneNoteSection { Name = "Ink opacity" };
        var page = new OneNotePage { Title = "Ink" };
        var ink = new OneNoteInk();
        ink.Ink.Add(new OfficeInkStroke { Opacity = 1D }
            .AddPoint(0.1, 0.2).AddPoint(0.3, 0.4));
        var highlighter = new OfficeInkStroke {
            Color = OfficeColor.Yellow,
            Opacity = 0.5D,
            IsHighlighter = true,
            TipShape = OfficeInkTipShape.Rectangle
        }.AddPoint(0.2, 0.5).AddPoint(0.8, 0.5);
        ink.Ink.Add(highlighter);
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteWriteObjectSpace graph = new OneNoteWriteGraphBuilder().BuildSection(section).ObjectSpaces[1];
        ulong[] transparency = graph.Objects
            .Where(item => item.Jcid == OneNoteSchema.JcidStrokePropertiesNode)
            .Select(item => Assert.Single(item.Properties, property =>
                (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.InkTransparency).Scalar!.Value)
            .ToArray();
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OfficeInkStroke[] actual = Assert.IsType<OneNoteInk>(
            Assert.Single(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children)).Strokes.ToArray();

        Assert.Equal(new ulong[] { 0UL, 204UL }, transparency);
        Assert.Equal(1D, actual[0].Opacity, 6);
        Assert.False(actual[1].IsHighlighter);
        Assert.Equal(OfficeInkRenderer.GetEffectiveOpacity(highlighter), actual[1].Opacity, 2);
        Assert.Equal(OfficeInkRenderer.GetEffectiveOpacity(highlighter), OfficeInkRenderer.GetEffectiveOpacity(actual[1]), 2);
    }

    [Fact]
    public void RoundTripsStructuredMathMlAsNativeOneNoteMath() {
        var section = new OneNoteSection { Name = "Structured math" };
        var page = new OneNotePage { Title = "Math" };
        OfficeMathExpression expression = OfficeMath.Fraction(
            OfficeMath.Row(OfficeMath.Identifier("x"), OfficeMath.Operator("+"), OfficeMath.Number("1")),
            OfficeMath.Radical(OfficeMath.Identifier("y")));
        page.DirectContent.Add(new OneNoteMath { MathMl = OfficeMathMarkup.ToMathMl(expression) });
        section.Pages.Add(page);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNoteParagraph result = Assert.IsType<OneNoteParagraph>(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children.Single());
        OfficeMathExpression roundTripExpression = Assert.Single(result.Runs).MathExpression!;

        Assert.Equal(expression, roundTripExpression);
        Assert.Contains("<mfrac>", OfficeMathMarkup.ToMathMl(roundTripExpression), StringComparison.Ordinal);
        Assert.Contains("\\frac", OfficeMathMarkup.ToLatex(roundTripExpression), StringComparison.Ordinal);
    }

    [Fact]
    public void RoundTripsIndexedRadicalInNativeChildOrder() {
        OfficeMathExpression expression = OfficeMath.Radical(OfficeMath.Identifier("x"), OfficeMath.Number("3"));
        var section = new OneNoteSection { Name = "Indexed radical" };
        var page = new OneNotePage { Title = "Math" };
        var paragraph = new OneNoteParagraph();
        paragraph.AddMath(expression);
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNoteParagraph result = Assert.IsType<OneNoteParagraph>(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children.Single());

        Assert.Equal(expression, Assert.Single(result.Runs).MathExpression);
    }

    [Fact]
    public void RetainsUndecodableNativeStrokeReferencesDuringPreservationWrites() {
        var section = new OneNoteSection { Name = "Preserved ink" };
        var page = new OneNotePage { Title = "Ink" };
        var ink = new OneNoteInk();
        var preservedStrokeId = new OneNoteExtendedGuid(Guid.NewGuid(), 7, 17);
        ink.PreservedStrokeObjectIds.Add(preservedStrokeId);
        ink.PreservedInkBoundingBox = InkBounds(-10, -20, 30, 40);
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteWriteObjectSpace pageSpace = new OneNoteWriteGraphBuilder().BuildSection(section).ObjectSpaces[1];
        OneNoteWriteObject inkData = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidInkDataNode);
        OneNoteWriteProperty strokes = Assert.Single(inkData.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.InkStrokes);

        Assert.Contains(preservedStrokeId, strokes.References);
        Assert.Equal(ink.PreservedInkBoundingBox, Assert.Single(inkData.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.InkBoundingBox).Data);
    }

    [Fact]
    public void PreservesRecognitionTreeForOpaqueStrokeDuringUnrelatedEdit() {
        var section = new OneNoteSection { Name = "Opaque recognition" };
        var page = new OneNotePage { Title = "Before" };
        var ink = new OneNoteInk();
        var stroke = new OfficeInkStroke { RecognizedText = "hello", LanguageId = 1033 }
            .AddPoint(0.1, 0.2)
            .AddPoint(0.3, 0.4);
        stroke.RecognitionAlternatives.Add("hello");
        stroke.RecognitionAlternatives.Add("hullo");
        ink.Ink.Add(stroke);
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNotePage loadedPage = Assert.Single(loaded.Pages);
        OneNoteInk loadedInk = Assert.IsType<OneNoteInk>(Assert.Single(Assert.Single(loadedPage.Outlines).Children));
        OfficeInkStroke opaqueStroke = Assert.Single(loadedInk.Strokes);
        OneNoteExtendedGuid opaqueStrokeId = loadedInk.StrokeObjectIds[opaqueStroke];
        OneNoteMaterializedObjectSpace sourceSpace = loaded.PreservationState!.GetPageSpace(loadedPage)!;
        OneNoteRevisionStoreObject sourcePageNode = sourceSpace.GetObject(loadedPage.PreservationIds.PageNodeId!)!;
        OneNoteRevisionStoreObject sourceRecognitionRoot = Referenced(sourceSpace, sourcePageNode, OneNoteSchema.PageRecognizedTextContainer);
        OneNoteRevisionStoreObject sourceRecognitionLine = Referenced(sourceSpace, sourceRecognitionRoot, OneNoteSchema.RecognizedTextChildNodes);
        OneNoteRevisionStoreObject sourceRecognitionBlock = Referenced(sourceSpace, sourceRecognitionLine, OneNoteSchema.RecognizedTextChildNodes);
        OneNoteRevisionStoreObject sourceRecognitionWord = Referenced(sourceSpace, sourceRecognitionBlock, OneNoteSchema.RecognizedTextChildNodes);
        Assert.True(loadedInk.Ink.Remove(opaqueStroke));
        loadedInk.PreservedStrokeObjectIds.Add(opaqueStrokeId);
        loadedPage.Title = "After";

        OneNoteWriteObjectSpace output = new OneNoteWriteGraphBuilder().BuildSection(loaded).ObjectSpaces[1];
        OneNoteWriteObject outputPageNode = Assert.Single(output.Objects, item => item.Id.Equals(loadedPage.PreservationIds.PageNodeId));
        OneNoteWriteObject outputRecognitionRoot = Referenced(output, outputPageNode, OneNoteSchema.PageRecognizedTextContainer);
        OneNoteWriteObject outputRecognitionLine = Referenced(output, outputRecognitionRoot, OneNoteSchema.RecognizedTextChildNodes);
        OneNoteWriteObject outputRecognitionBlock = Referenced(output, outputRecognitionLine, OneNoteSchema.RecognizedTextChildNodes);
        OneNoteWriteObject retainedRecognitionWord = Referenced(output, outputRecognitionBlock, OneNoteSchema.RecognizedTextChildNodes);

        Assert.Equal(sourceRecognitionRoot.Id, outputRecognitionRoot.Id);
        Assert.Equal(sourceRecognitionLine.Id, outputRecognitionLine.Id);
        Assert.Equal(sourceRecognitionBlock.Id, outputRecognitionBlock.Id);
        Assert.Equal(sourceRecognitionWord.Id, retainedRecognitionWord.Id);
        Assert.Equal(OneNoteSchema.JcidRecognizedTextWord, retainedRecognitionWord.Jcid);
        Assert.Equal(
            OneNoteSemanticMapper.ReadData(sourceRecognitionWord, OneNoteSchema.RecognizedText),
            Assert.Single(retainedRecognitionWord.Properties, property =>
                (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.RecognizedText).Data);
    }

    [Fact]
    public void RoundTripsNonUniformlyTransformedInkTipDimensions() {
        var section = new OneNoteSection { Name = "Transformed ink" };
        var page = new OneNotePage { Title = "Ink" };
        var ink = new OneNoteInk();
        ink.Ink.Add(new OfficeInkStroke {
            Width = 0.04,
            Height = 0.06,
            Transform = OfficeTransform.Scale(2, 3)
        }.AddPoint(0.1, 0.2).AddPoint(0.3, 0.4));
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OfficeInkStroke actual = Assert.Single(Assert.IsType<OneNoteInk>(
            Assert.Single(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children)).Strokes);

        Assert.Equal(0.08, actual.Width, 3);
        Assert.Equal(0.18, actual.Height, 3);
        Assert.Equal(0.2, actual.Points[0].X, 3);
        Assert.Equal(0.6, actual.Points[0].Y, 3);
    }

    [Fact]
    public void RejectsAffineInkTipsThatNativeOneNoteCannotRepresentLosslessly() {
        var section = new OneNoteSection { Name = "Affine ink" };
        var page = new OneNotePage { Title = "Ink" };
        var ink = new OneNoteInk();
        ink.Ink.Add(new OfficeInkStroke {
            Width = 0.04D,
            Height = 0.06D,
            Transform = new OfficeTransform(1D, 0D, 0.5D, 1D, 0D, 0D)
        }.AddPoint(0.1D, 0.2D).AddPoint(0.3D, 0.4D));
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteFormatException error = Assert.Throws<OneNoteFormatException>(() => OneNoteSectionWriter.Write(section));

        Assert.Equal("ONENOTE_WRITE_INK_AFFINE_TIP", error.Code);
    }

    [Fact]
    public void UnionsPreservedOpaqueInkBoundsWithNewlyAuthoredStrokes() {
        var section = new OneNoteSection { Name = "Opaque ink bounds" };
        var page = new OneNotePage { Title = "Ink" };
        var ink = new OneNoteInk {
            PreservedInkBoundingBox = InkBounds(-100, -200, 300, 400)
        };
        ink.PreservedStrokeObjectIds.Add(new OneNoteExtendedGuid(Guid.NewGuid(), 7, 17));
        ink.Ink.Add(new OfficeInkStroke().AddPoint(2, 3).AddPoint(4, 5));
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteWriteObjectSpace pageSpace = new OneNoteWriteGraphBuilder().BuildSection(section).ObjectSpaces[1];
        OneNoteWriteObject inkData = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidInkDataNode);
        byte[] bounds = Assert.Single(inkData.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.InkBoundingBox).Data!;

        Assert.Equal(-100, ReadInt32(bounds, 0));
        Assert.Equal(-200, ReadInt32(bounds, 4));
        Assert.True(ReadInt32(bounds, 8) >= 4 * OneNoteInkCodec.NativeUnitsPerHalfInch);
        Assert.True(ReadInt32(bounds, 12) >= 5 * OneNoteInkCodec.NativeUnitsPerHalfInch);
    }

    [Fact]
    public void RejectsOpaqueInkWhenItsCompleteNativeBoundsAreUnavailable() {
        var section = new OneNoteSection { Name = "Missing opaque bounds" };
        var page = new OneNotePage { Title = "Ink" };
        var ink = new OneNoteInk();
        ink.PreservedStrokeObjectIds.Add(new OneNoteExtendedGuid(Guid.NewGuid(), 7, 17));
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteFormatException error = Assert.Throws<OneNoteFormatException>(() => OneNoteSectionWriter.Write(section));

        Assert.Equal("ONENOTE_WRITE_OPAQUE_INK_BOUNDS", error.Code);
    }

    [Fact]
    public void ReusesNativeStrokeWithUnsupportedPacketDimensionsForRecognitionOnlyEdits() {
        var section = new OneNoteSection { Name = "Preserved packet dimensions" };
        var page = new OneNotePage { Title = "Ink" };
        var ink = new OneNoteInk();
        var stroke = new OfficeInkStroke().AddPoint(0.1, 0.2).AddPoint(0.3, 0.4);
        ink.Ink.Add(stroke);
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNotePage loadedPage = Assert.Single(loaded.Pages);
        OneNoteInk loadedInk = Assert.IsType<OneNoteInk>(Assert.Single(Assert.Single(loadedPage.Outlines).Children));
        OfficeInkStroke loadedStroke = Assert.Single(loadedInk.Strokes);
        OneNoteExtendedGuid nativeStrokeId = loadedInk.StrokeObjectIds[loadedStroke];
        OneNoteMaterializedObjectSpace sourceSpace = loaded.PreservationState!.GetPageSpace(loadedPage)!;
        OneNoteRevisionStoreObject nativeStroke = sourceSpace.GetObject(nativeStrokeId)!;
        OneNoteExtendedGuid propertyId = Assert.Single(OneNoteSemanticMapper.GetReferences(nativeStroke, OneNoteSchema.InkStrokeProperties));
        OneNoteRevisionStoreObject nativeProperties = sourceSpace.GetObject(propertyId)!;
        OneNotePropertyValue dimensions = OneNoteSemanticMapper.FindProperty(nativeProperties.PropertySet, OneNoteSchema.InkDimensions)!;
        byte[] extendedDimensions = dimensions.Data!.ToArray(long.MaxValue).Concat(ExtraInkDimension()).ToArray();
        dimensions.Data = OneNoteBinaryPayload.FromBytes(extendedDimensions);
        OneNotePropertyValue path = OneNoteSemanticMapper.FindProperty(nativeStroke.PropertySet, OneNoteSchema.InkPath)!;
        IReadOnlyList<long> originalPath = OneNoteInkCodec.DecodeSignedVector(path.Data!.ToArray(long.MaxValue), int.MaxValue);
        byte[] extendedPath = OneNoteInkCodec.EncodeSignedVector(originalPath
            .Concat(OneNoteInkCodec.EncodePacketValues(new long[] { 7, 9 })).ToArray());
        path.Data = OneNoteBinaryPayload.FromBytes(extendedPath);
        loadedInk.PreservedNativeStrokeSnapshots[loadedStroke] = loadedStroke.Clone();
        byte[] sourceBounds = InkBounds(-100, -200, 1000, 1200);
        loadedInk.PreservedInkBoundingBox = sourceBounds;
        loadedStroke.RecognizedText = "updated";
        loadedStroke.RecognitionAlternatives.Clear();
        loadedStroke.RecognitionAlternatives.Add("updated");
        loadedStroke.RecognitionAlternatives.Add("alternate");

        OneNoteWriteObjectSpace pageSpace = new OneNoteWriteGraphBuilder().BuildSection(loaded).ObjectSpaces[1];
        OneNoteWriteObject retainedProperties = Assert.Single(pageSpace.Objects, item => item.Id.Equals(propertyId));
        byte[] retainedDimensions = Assert.Single(retainedProperties.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.InkDimensions).Data!;

        Assert.Equal(extendedDimensions, retainedDimensions);
        OneNoteWriteObject retainedStroke = Assert.Single(pageSpace.Objects, item => item.Id.Equals(nativeStrokeId));
        Assert.Equal(extendedPath, Assert.Single(retainedStroke.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.InkPath).Data);
        OneNoteWriteObject inkData = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidInkDataNode);
        Assert.Equal(sourceBounds, Assert.Single(inkData.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.InkBoundingBox).Data);
        OneNoteWriteObject recognitionWord = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidRecognizedTextWord);
        byte[] recognitionData = Assert.Single(recognitionWord.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.RecognizedText).Data!;
        Assert.Contains("updated", System.Text.Encoding.Unicode.GetString(recognitionData), StringComparison.Ordinal);

        loadedStroke.Width = 2;
        OneNoteWriteObjectSpace editedPageSpace = new OneNoteWriteGraphBuilder().BuildSection(loaded).ObjectSpaces[1];
        OneNoteWriteObject editedProperties = Assert.Single(editedPageSpace.Objects, item => item.Id.Equals(propertyId));
        byte[] canonicalDimensions = Assert.Single(editedProperties.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.InkDimensions).Data!;
        Assert.NotEqual(extendedDimensions, canonicalDimensions);
    }

    [Fact]
    public void ReencodedSupportedInkClearsRetainedContainerScaling() {
        var section = new OneNoteSection { Name = "Scaled ink" };
        var page = new OneNotePage { Title = "Ink" };
        var ink = new OneNoteInk();
        ink.Ink.Add(new OfficeInkStroke { Width = 0.04, Height = 0.06 }
            .AddPoint(0.25, 0.5).AddPoint(1.5, 2.25));
        page.DirectContent.Add(ink);
        section.Pages.Add(page);

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNotePage loadedPage = Assert.Single(loaded.Pages);
        OneNoteInk loadedInk = Assert.IsType<OneNoteInk>(Assert.Single(Assert.Single(loadedPage.Outlines).Children));
        OneNoteMaterializedObjectSpace sourceSpace = loaded.PreservationState!.GetPageSpace(loadedPage)!;
        OneNoteRevisionStoreObject container = sourceSpace.GetObject(loadedInk.Id!)!;
        OneNoteSemanticMapper.FindProperty(container.PropertySet, OneNoteSchema.InkScalingX)!.ScalarValue = FloatBits(2F);
        OneNoteSemanticMapper.FindProperty(container.PropertySet, OneNoteSchema.InkScalingY)!.ScalarValue = FloatBits(3F);
        loadedInk.PreservedInkScaleX = 2D;
        loadedInk.PreservedInkScaleY = 3D;

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(loaded)));
        OfficeInkStroke actual = Assert.Single(Assert.IsType<OneNoteInk>(
            Assert.Single(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children)).Strokes);

        Assert.Equal(0.25, actual.Points[0].X, 3);
        Assert.Equal(2.25, actual.Points[1].Y, 3);
        Assert.Equal(0.04, actual.Width, 3);
        Assert.Equal(0.06, actual.Height, 3);
    }

    [Fact]
    public void PreservesUnchangedOpaqueOnlyNestedInkContainer() {
        var section = new OneNoteSection { Name = "Opaque nested ink" };
        var page = new OneNotePage { Title = "Ink" };
        page.DirectContent.Add(new OneNoteInk());
        var child = new OneNoteInk();
        child.Ink.Add(new OfficeInkStroke().AddPoint(0.1, 0.2).AddPoint(0.3, 0.4));
        page.DirectContent.Add(child);
        section.Pages.Add(page);

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNotePage loadedPage = Assert.Single(loaded.Pages);
        OneNoteInk[] inks = Assert.Single(loadedPage.Outlines).Children.OfType<OneNoteInk>().ToArray();
        Assert.Equal(2, inks.Length);
        OneNoteInk parent = inks[0];
        OneNoteInk retainedChild = inks[1];
        OneNoteExtendedGuid opaqueStrokeId = retainedChild.StrokeObjectIds[Assert.Single(retainedChild.Strokes)];
        parent.PreservedChildContainerIds.Add(retainedChild.Id!);
        parent.PreservedStrokeObjectIds.Add(opaqueStrokeId);

        OneNoteWriteObjectSpace output = new OneNoteWriteGraphBuilder().BuildSection(loaded).ObjectSpaces[1];
        OneNoteWriteObject parentObject = Assert.Single(output.Objects, item => item.Id.Equals(parent.Id));
        OneNoteWriteProperty children = Assert.Single(parentObject.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.ContentChildNodes);

        Assert.Contains(retainedChild.Id!, children.References);
        Assert.DoesNotContain(parentObject.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.InkData);
    }

    [Fact]
    public void RoundTripsNativeRecordingIdentityDurationAndPageIndex() {
        Guid recordingId = Guid.NewGuid();
        var section = new OneNoteSection { Name = "Media" };
        var page = new OneNotePage { Title = "Recording" };
        page.DirectContent.Add(new OneNoteMedia {
            FileName = "meeting.mp3",
            MediaType = "audio/mpeg",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 }),
            RecordingKind = OneNoteMediaKind.Audio,
            RecordingId = recordingId,
            Duration = TimeSpan.FromMilliseconds(12_345)
        });
        section.Pages.Add(page);

        byte[] bytes = OneNoteSectionWriter.Write(section);
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(bytes));
        OneNotePage actualPage = Assert.Single(roundTrip.Pages);
        OneNoteMedia actual = Assert.IsType<OneNoteMedia>(Assert.Single(Assert.Single(actualPage.Outlines).Children));

        Assert.Equal(recordingId, actual.RecordingId);
        Assert.Equal(TimeSpan.FromMilliseconds(12_345), actual.Duration);
        Assert.Equal(OneNoteMediaKind.Audio, actual.RecordingKind);
        OneNoteMaterializedObjectSpace pageSpace = roundTrip.PreservationState!.GetPageSpace(actualPage)!;
        OneNoteRevisionStoreObject pageNode = pageSpace.GetObject(actualPage.PreservationIds.PageNodeId!)!;
        Assert.Equal(recordingId.ToByteArray(), OneNoteSemanticMapper.ReadData(pageNode, OneNoteSchema.AudioRecordingGuids));
    }

    [Theory]
    [InlineData("meeting.mpeg", "video/mpeg")]
    [InlineData("meeting.mp4", "video/mp4")]
    public void RoundTripsSupportedVideoRecordingExtensions(string fileName, string mediaType) {
        var section = new OneNoteSection { Name = "Video" };
        var page = new OneNotePage { Title = "Recording" };
        page.DirectContent.Add(new OneNoteMedia {
            FileName = fileName,
            MediaType = mediaType,
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 }),
            RecordingKind = OneNoteMediaKind.Video,
            RecordingId = Guid.NewGuid()
        });
        section.Pages.Add(page);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNoteMedia actual = Assert.IsType<OneNoteMedia>(
            Assert.Single(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children));

        Assert.Equal(OneNoteMediaKind.Video, actual.RecordingKind);
        Assert.Equal(fileName, actual.FileName);
        Assert.Equal(mediaType, actual.MediaType);
    }

    [Fact]
    public void RoundTripsMixedTextAndInlineMathWithoutExposingNativeMarkers() {
        var section = new OneNoteSection { Name = "Inline math" };
        var page = new OneNotePage { Title = "Inline" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Area is " });
        OfficeMathExpression expression = OfficeMath.Row(
            OfficeMath.Identifier("π"),
            OfficeMath.Superscript(OfficeMath.Identifier("r"), OfficeMath.Number("2")));
        paragraph.AddMath(expression);
        paragraph.Runs.Add(new OneNoteTextRun { Text = "." });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNoteParagraph result = Assert.IsType<OneNoteParagraph>(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children.Single());

        Assert.Equal(3, result.Runs.Count);
        Assert.Equal("Area is ", result.Runs[0].Text);
        Assert.Equal(expression, result.Runs[1].MathExpression);
        Assert.Equal(expression.ToPlainText(), result.Runs[1].Text);
        Assert.DoesNotContain(result.Runs[1].Text, character => character == '\uFDD0' || character == '\uFDEE' || character == '\uFDEF');
        Assert.Equal(".", result.Runs[2].Text);
    }

    [Fact]
    public void RoundTripsAdjacentMathRunsAsDistinctStyledInlineExpressions() {
        var section = new OneNoteSection { Name = "Adjacent math" };
        var page = new OneNotePage { Title = "Math" };
        var paragraph = new OneNoteParagraph();
        OneNoteTextRun first = paragraph.AddMath(OfficeMath.Identifier("a"));
        first.Style.ColorArgb = 0xFFCC0000;
        OneNoteTextRun second = paragraph.AddMath(OfficeMath.Superscript(OfficeMath.Identifier("b"), OfficeMath.Number("2")));
        second.Style.Bold = true;
        second.Hyperlink = "https://example.com/math";
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNoteParagraph actual = Assert.IsType<OneNoteParagraph>(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children.Single());

        Assert.Equal(2, actual.Runs.Count);
        Assert.Equal(OfficeMath.Identifier("a"), actual.Runs[0].MathExpression);
        Assert.Equal(0xFFCC0000U, actual.Runs[0].Style.ColorArgb);
        Assert.Equal(OfficeMath.Superscript(OfficeMath.Identifier("b"), OfficeMath.Number("2")), actual.Runs[1].MathExpression);
        Assert.True(actual.Runs[1].Style.Bold);
        Assert.Equal("https://example.com/math", actual.Runs[1].Hyperlink);
    }

    [Fact]
    public void NativeMathCodecRoundTripsAdvancedDrawingOwnedStructures() {
        OfficeMathExpression[] expressions = {
            OfficeMath.Row(OfficeMath.Identifier("x"), OfficeMath.Number("2"), OfficeMath.Text("word"), OfficeMath.Operator("+")),
            OfficeMath.LeftSubSuperscript(OfficeMath.Identifier("T"), OfficeMath.Identifier("i"), OfficeMath.Identifier("j")),
            OfficeMath.LowerLimit(OfficeMath.Identifier("lim"), OfficeMath.Identifier("x")),
            OfficeMath.UpperLimit(OfficeMath.Identifier("max"), OfficeMath.Identifier("n")),
            OfficeMath.SlashedFraction(OfficeMath.Identifier("a"), OfficeMath.Identifier("b")),
            OfficeMath.Stack(OfficeMath.Identifier("a"), OfficeMath.Identifier("b")),
            OfficeMath.StretchStack(OfficeMath.Identifier("x"), OfficeMath.Identifier("y")),
            OfficeMath.DelimiterList("[", "]", ";", OfficeMath.Identifier("a"), OfficeMath.Identifier("b"))
        };

        Assert.All(expressions, expression => Assert.Equal(expression, OneNoteMathNativeCodec.Canonicalize(expression)));
    }

    [Fact]
    public void NativeMathWriterPreservesAdjacentTokenKindsWithoutMutatingTheCaller() {
        OfficeMathExpression expression = OfficeMath.Row(
            OfficeMath.Identifier("x"),
            OfficeMath.Number("2"),
            OfficeMath.Identifier("y"));
        var section = new OneNoteSection { Name = "Token boundaries" };
        var page = new OneNotePage { Title = "Math" };
        var paragraph = new OneNoteParagraph();
        OneNoteTextRun authoredRun = paragraph.AddMath(expression);
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        byte[] bytes = OneNoteSectionWriter.Write(section);
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(bytes));
        OneNoteParagraph actual = Assert.IsType<OneNoteParagraph>(
            Assert.Single(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children));

        Assert.Same(expression, authoredRun.MathExpression);
        Assert.Equal(expression, Assert.Single(actual.Runs).MathExpression);
    }

    [Fact]
    public void NativeMathReaderRejectsExcessiveObjectDepth() {
        var runs = new List<OneNoteTextRun>();
        for (int index = 0; index < 12; index++) {
            var run = new OneNoteTextRun { Text = OneNoteMathNativeCodec.ObjectStart.ToString(), MathDescriptor = new OneNoteMathInlineDescriptor { Type = 12, Count = 1 } };
            run.Style.IsMath = true;
            runs.Add(run);
        }
        var terminal = new OneNoteTextRun { Text = "x" + new string(OneNoteMathNativeCodec.ObjectEnd, 12), MathDescriptor = new OneNoteMathInlineDescriptor { Type = 12 } };
        terminal.Style.IsMath = true;
        runs.Add(terminal);

        OneNoteFormatException error = Assert.Throws<OneNoteFormatException>(() => OneNoteMathNativeCodec.Decode(runs, 8));

        Assert.Equal("ONENOTE_MATH_DEPTH", error.Code);
    }

    [Fact]
    public void WriterRejectsNativeMathArraysWiderThanDescriptorCapacity() {
        var section = new OneNoteSection { Name = "Wide math" };
        var page = new OneNotePage { Title = "Math" };
        var paragraph = new OneNoteParagraph();
        paragraph.AddMath(OfficeMath.Matrix(1, 256, Enumerable.Range(0, 256).Select(index => OfficeMath.Number(index.ToString())).ToArray()));
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        OneNoteFormatException error = Assert.Throws<OneNoteFormatException>(() => OneNoteSectionWriter.Write(section));

        Assert.Equal("ONENOTE_WRITE_MATH_COLUMNS", error.Code);
    }

    [Fact]
    public void WriterRejectsMathCharactersThatCannotFitNativeDescriptorsWithoutTruncation() {
        OfficeMathExpression[] expressions = {
            OfficeMath.Delimited(OfficeMath.Identifier("x"), "||", ")"),
            OfficeMath.Nary("😀", OfficeMath.Identifier("x")),
            OfficeMath.Delimited(OfficeMath.Identifier("x"), string.Empty, ")")
        };

        Assert.All(expressions, expression => {
            var section = new OneNoteSection { Name = "Invalid native math character" };
            var page = new OneNotePage { Title = "Math" };
            var paragraph = new OneNoteParagraph();
            paragraph.AddMath(expression);
            page.DirectContent.Add(paragraph);
            section.Pages.Add(page);

            OneNoteFormatException error = Assert.Throws<OneNoteFormatException>(() => OneNoteSectionWriter.Write(section));

            Assert.Equal("ONENOTE_WRITE_MATH_CHARACTER", error.Code);
        });
        Assert.Equal("||x)", expressions[0].ToPlainText());
        Assert.Equal("😀x", expressions[1].ToPlainText());
    }

    [Fact]
    public void PreservesUnsupportedNativeMathDescriptorsWhenPresentationChanges() {
        var paragraph = new OneNoteParagraph();
        var root = new OneNoteTextRun {
            Text = OneNoteMathNativeCodec.ObjectStart.ToString(),
            MathDescriptor = new OneNoteMathInlineDescriptor {
                Type = 99,
                Count = 3,
                Column = 2,
                Alignment = 7,
                Character = '(',
                Character1 = ')',
                Character2 = '|'
            }
        };
        root.Style.IsMath = true;
        var child = new OneNoteTextRun {
            Text = "x" + OneNoteMathNativeCodec.ObjectEnd,
            MathDescriptor = new OneNoteMathInlineDescriptor { Type = 98, Count = 5, Alignment = 4, Character2 = ';' }
        };
        child.Style.IsMath = true;
        paragraph.Runs.Add(root);
        paragraph.Runs.Add(child);
        OneNoteSemanticMapper.CollapseInlineMathRuns(paragraph);
        OneNoteTextRun semantic = Assert.Single(paragraph.Runs);
        semantic.Style.Bold = true;
        semantic.Style.ColorArgb = 0xFFCC2200U;
        semantic.Hyperlink = "https://example.com/preserved-math";
        var section = new OneNoteSection { Name = "Preserved math" };
        var page = new OneNotePage { Title = "Math" };
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        OneNoteWriteObjectSpace pageSpace = new OneNoteWriteGraphBuilder().BuildSection(section).ObjectSpaces[1];
        OneNoteWriteObject richText = Assert.Single(pageSpace.Objects, item =>
            item.Jcid == OneNoteSchema.JcidRichTextNode && item.Properties.Any(property =>
                (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.MathInlineObjects));
        OneNoteWriteProperty descriptors = Assert.Single(richText.Properties, property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.MathInlineObjects);

        Assert.Equal(2, descriptors.ChildPropertySets.Count);
        Assert.Equal(99UL, Scalar(descriptors.ChildPropertySets[0], OneNoteSchema.MathInlineObjectType));
        Assert.Equal(7UL, Scalar(descriptors.ChildPropertySets[0], OneNoteSchema.MathInlineObjectAlignment));
        Assert.Equal((ulong)'|', Scalar(descriptors.ChildPropertySets[0], OneNoteSchema.MathInlineObjectCharacter2));
        Assert.Equal(98UL, Scalar(descriptors.ChildPropertySets[1], OneNoteSchema.MathInlineObjectType));
        Assert.Equal(5UL, Scalar(descriptors.ChildPropertySets[1], OneNoteSchema.MathInlineObjectCount));
        Assert.Equal((ulong)';', Scalar(descriptors.ChildPropertySets[1], OneNoteSchema.MathInlineObjectCharacter2));
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNoteTextRun actual = Assert.Single(Assert.IsType<OneNoteParagraph>(
            Assert.Single(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children)).Runs);
        Assert.True(actual.Style.Bold);
        Assert.Equal(0xFFCC2200U, actual.Style.ColorArgb);
        Assert.Equal("https://example.com/preserved-math", actual.Hyperlink);
    }

    private static ulong Scalar(IReadOnlyList<OneNoteWriteProperty> properties, uint id) =>
        Assert.Single(properties, property => (property.RawId & 0x7FFFFFFFU) == id).Scalar!.Value;

    private static uint FloatBits(float value) => BitConverter.ToUInt32(BitConverter.GetBytes(value), 0);

    private static byte[] InkBounds(int left, int top, int right, int bottom) {
        var data = new byte[16];
        Buffer.BlockCopy(BitConverter.GetBytes(left), 0, data, 0, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(top), 0, data, 4, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(right), 0, data, 8, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(bottom), 0, data, 12, 4);
        return data;
    }

    private static int ReadInt32(byte[] data, int offset) => unchecked((int)OneNoteBinary.ReadUInt32(data, offset));

    private static byte[] ExtraInkDimension() {
        var bytes = new List<byte>();
        bytes.AddRange(Guid.NewGuid().ToByteArray());
        bytes.AddRange(BitConverter.GetBytes(int.MinValue));
        bytes.AddRange(BitConverter.GetBytes(int.MaxValue));
        bytes.AddRange(BitConverter.GetBytes(0U));
        bytes.AddRange(BitConverter.GetBytes(1F));
        return bytes.ToArray();
    }

    private static OneNoteRevisionStoreObject Referenced(
        OneNoteMaterializedObjectSpace space,
        OneNoteRevisionStoreObject owner,
        uint propertyId) => space.GetObject(Assert.Single(OneNoteSemanticMapper.GetReferences(owner, propertyId)))!;

    private static OneNoteWriteObject Referenced(
        OneNoteWriteObjectSpace space,
        OneNoteWriteObject owner,
        uint propertyId) {
        OneNoteWriteProperty property = Assert.Single(owner.Properties, item =>
            (item.RawId & 0x7FFFFFFFU) == propertyId);
        return Assert.Single(space.Objects, item => item.Id.Equals(Assert.Single(property.References)));
    }
}
