using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRotationNormalizationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NormalizationPreservesIndirectScalarDestinationOperands(bool nullZoom) {
        byte[] bytes = CreateGeometrySource(90);
        var read = PdfReadDocument.Open(bytes);
        var (objects, trailer) = PdfSyntax.ParseObjects(bytes);
        int number = objects.Keys.Max() + 1;
        objects[number] = new PdfIndirectObject(number, 0, nullZoom ? PdfNull.Instance : new PdfNumber(2));
        var destination = new PdfArray();
        destination.Items.Add(new PdfReference(read.Pages[0].ObjectNumber, 0));
        destination.Items.Add(new PdfName("XYZ"));
        destination.Items.Add(new PdfNumber(100));
        destination.Items.Add(new PdfNumber(400));
        destination.Items.Add(new PdfReference(number, 0));
        int root = read.Security.RootObjectNumber!.Value;
        var names = new PdfDictionary(); names.Items["valid"] = destination;
        ((PdfDictionary)objects[root].Value).Items["Dests"] = names;
        bytes = PdfIncrementalObjectWriter.Append(bytes, objects, read.Security, trailer, new[] { root, number });
        var source = PdfDocument.Load(bytes);

        var output = PdfReadDocument.Open(source.Pages.NormalizeRotation(PdfPageSelection.From(1)).ToBytes());

        var catalog = Assert.IsType<PdfDictionary>(output.Objects[output.Security.RootObjectNumber!.Value].Value);
        var destinations = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, catalog.Items["Dests"]));
        var result = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(output.Objects, destinations.Items["valid"]));
        if (nullZoom) Assert.IsType<PdfNull>(result.Items[4]);
        else Assert.Equal(2, Assert.IsType<PdfNumber>(result.Items[4]).Value);
        Assert.Equal(bytes, source.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NormalizationRejectsNonNumericDestinationOperands(bool cyclic) {
        byte[] bytes = CreateGeometrySource(90);
        var read = PdfReadDocument.Open(bytes);
        var (objects, trailer) = PdfSyntax.ParseObjects(bytes);
        int number = objects.Keys.Max() + 1;
        var destination = new PdfArray();
        destination.Items.Add(new PdfReference(read.Pages[0].ObjectNumber, 0));
        destination.Items.Add(new PdfName("XYZ"));
        destination.Items.Add(new PdfNumber(0));
        destination.Items.Add(new PdfNumber(0));
        destination.Items.Add(new PdfReference(number, 0));
        var invalidOperand = new PdfArray();
        invalidOperand.Items.Add(cyclic ? new PdfReference(number, 0) : new PdfNumber(1));
        objects[number] = new PdfIndirectObject(number, 0, cyclic ? destination : invalidOperand);
        int root = read.Security.RootObjectNumber!.Value;
        var names = new PdfDictionary(); names.Items["invalid"] = destination;
        ((PdfDictionary)objects[root].Value).Items["Dests"] = names;
        bytes = PdfIncrementalObjectWriter.Append(bytes, objects, read.Security, trailer, new[] { root, number });
        var source = PdfDocument.Load(bytes);

        Assert.Throws<InvalidOperationException>(() => source.Pages.NormalizeRotation(PdfPageSelection.From(1)));

        Assert.Equal(bytes, source.ToBytes());
    }
}
