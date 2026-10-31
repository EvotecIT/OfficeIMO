using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfEmptyAppearanceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CompressedAppearanceDecodingStopsAtTheRemainingTotalBudget(bool form) {
        byte[] source = BuildPdf(form ? "[0 0 20 20]" : null, "[0 0 20 20]", pageContent: "", widgetOn: form);
        source = ReplaceCompressedAppearance(source, form ? 5 : 9, new string(' ', 4096) + "0 g 0 0 20 20 re f");
        const int totalLimit = 32;
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxDecodedStreamBytes = 8192, MaxTotalDecodedStreamBytes = totalLimit
        }});

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => {
            if (form) document.Forms.Flatten(); else document.Annotations.Flatten();
        });

        Assert.Equal(PdfReadLimitKind.TotalDecodedStreamBytes, exception.Kind);
        Assert.Equal(totalLimit, exception.Limit);
        // A bounded decoder reports the first byte beyond its allowance, without inflating the whole AP.
        Assert.Equal(totalLimit + 1, exception.Actual);
        Assert.Equal(source, document.ToBytes());
    }

    [Theory]
    [InlineData(PdfFormFieldCreationKind.CheckBox)]
    [InlineData(PdfFormFieldCreationKind.PushButton)]
    [InlineData(PdfFormFieldCreationKind.Choice)]
    [InlineData(PdfFormFieldCreationKind.RadioButtonGroup)]
    public void TransactionalFieldCreationAdoptsItsImmediateAppearanceGrowth(PdfFormFieldCreationKind kind) {
        byte[] source = BuildPdf(null, "[0 0 0 0]");
        PdfStream[] streams = PdfSyntax.ParseObjects(source).Map.Values.Select(item => item.Value).OfType<PdfStream>().ToArray();
        int maximum = streams.Max(stream => stream.Data.Length);
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxInputBytes = source.LongLength, MaxRawStreamBytes = maximum, MaxDecodedStreamBytes = maximum
        }});
        _ = document.Read();
        var options = new PdfFormFieldCreateOptions {
            Name = "new-field", Kind = kind, X = 20, Y = 20, Width = 150, Height = 50,
            Caption = "Continue delivery", ChoiceOptions = new[] { "First", "Second" }
        };

        PdfDocument created = document.Forms.Edit(edit => edit.Create(options)).ToDocument();

        _ = created.Read();
        Assert.Single(created.Inspect().FormFields);
        Assert.True(created.ReadOptions.Limits.MaxRawStreamBytes > maximum);
        Assert.True(created.ReadOptions.Limits.MaxRawStreamBytes < new PdfReadLimits().MaxRawStreamBytes);
        Assert.Equal(source, document.ToBytes());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void FillGrowthDoesNotAdmitAnUnchangedOversizedAppearance(bool annotation, bool typed) {
        byte[] source = BuildPdf(annotation ? null : "[0 0 20 20]", "[0 0 20 20]", widgetOn: !annotation);
        source = ReplaceCompressedAppearance(source, annotation ? 9 : 5, new string(' ', 120) + "0 g 0 0 20 20 re f");
        source = PdfDocument.Load(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "notes", Kind = PdfFormFieldCreationKind.Text, X = 20, Y = 60, Width = 160, Height = 24, Value = ""
        })).ToDocument().ToBytes();
        int maximum = PdfSyntax.ParseObjects(source).Map.Values.Select(item => item.Value).OfType<PdfStream>()
            .Max(stream => stream.Data.Length);
        var document = PdfDocument.Load(source, new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxRawStreamBytes = maximum, MaxDecodedStreamBytes = maximum
        }});
        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, Assert.Throws<PdfReadLimitException>(() => document.Read()).Kind);
        const string value = "Delivery contact: service@example.test";

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => {
            PdfDocument filled = typed
                ? document.Forms.Fill(new Dictionary<string, PdfFormFieldValue> { ["notes"] = PdfFormFieldValue.From(value) })
                : document.Forms.Fill(new Dictionary<string, string> { ["notes"] = value });
            _ = filled.Read();
        });

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(maximum, exception.Limit);
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public void ImmediateAppearanceGrowthPreservesEncryptedAuthoringWithoutCopyPermission() {
        byte[] source = PdfDocument.Load(BuildPdf(null, "[0 0 0 0]")).Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.ModifyContents | PdfStandardPermissions.ModifyAnnotations
        }).Pdf;
        var originalOptions = new PdfLoadOptions { Password = "reader", Limits = new PdfReadLimits {
            MaxInputBytes = source.LongLength, MaxRawStreamBytes = 64, MaxDecodedStreamBytes = 64
        }};
        PdfDocument document = PdfDocument.Load(source, originalOptions);
        Assert.Throws<PdfPermissionDeniedException>(() => document.Read());

        PdfDocument created = document.Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "continue", Kind = PdfFormFieldCreationKind.PushButton, X = 20, Y = 20, Width = 150, Height = 24,
            Caption = "Continue delivery"
        })).ToDocument();

        Assert.Equal("reader", created.ReadOptions.Password);
        Assert.Equal(PdfPermissionPolicy.Enforce, created.ReadOptions.PermissionPolicy);
        Assert.Single(created.Forms.GetForInteraction().Fields);
        Assert.Throws<PdfPermissionDeniedException>(() => created.Read());
        _ = PdfDocument.Load(created.ToBytes(), new PdfLoadOptions {
            Password = "owner", Limits = created.ReadOptions.Limits
        }).Read();
        Assert.True(created.ReadOptions.Limits.MaxRawStreamBytes > originalOptions.Limits.MaxRawStreamBytes);
        Assert.True(created.ReadOptions.Limits.MaxRawStreamBytes < new PdfReadLimits().MaxRawStreamBytes);
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public void AppearanceAdmissionDoesNotRequireCopyPermissionForAuthorizedRefill() {
        PdfDocument original = PdfDocument.Create(pdf => pdf.Content(content => content.Canvas(canvas =>
            canvas.TextField("Name", "Ada", global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(160), global::OfficeIMO.OfficeLength.Points(24)))));
        byte[] source = original.Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.ModifyContents | PdfStandardPermissions.ModifyAnnotations
        }).Pdf;
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Password = "reader" });

        PdfDocument filled = document.Forms.Edit(edit => edit.Move("Name", 1, 20, 20, 160.5, 24)).ToDocument();

        Assert.Equal("Ada", Assert.Single(filled.Forms.GetForInteraction().Fields).Value);
        Assert.Throws<PdfPermissionDeniedException>(() => filled.Read());
        Assert.Equal(source, document.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptyOffAppearanceDoesNotConsumeAnExhaustedTotalBudget(bool fill) {
        byte[] source = PdfDocumentObjectGraphRewriter.Rewrite(BuildPdf("[0 0 0 0]", "[0 0 20 20]", pageContent: ""), null, null,
            (objects, security) => {
                var annotations = new PdfArray();
                annotations.Items.Add(new PdfReference(7, 0));
                annotations.Items.Add(new PdfReference(4, 0));
                ((PdfDictionary)objects[3].Value).Items["Annots"] = annotations;
                return security.InfoObjectNumber;
            });
        int visibleBytes = Encoding.ASCII.GetByteCount("1 0 0 rg 0 0 20 20 re f");
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxTotalDecodedStreamBytes = visibleBytes
        }});
        _ = document.Read();

        PdfDocument result = fill
            ? document.Forms.Fill(new Dictionary<string, string> { ["agree"] = "Off" })
            : document.Forms.Flatten();

        _ = result.Read();
        Assert.Equal(Render(source), Render(result.ToBytes()));
        Assert.Equal(source, document.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NestedEmptyFormDoesNotConsumeAnExhaustedAppearanceBudget(bool fill) {
        const string parentContent = "/Empty Do\n";
        byte[] source = PdfDocumentObjectGraphRewriter.Rewrite(BuildPdf("[0 0 20 20]", "[0 0 20 20]", pageContent: "", widgetOn: true), null, null,
            (objects, security) => {
                int child = objects.Keys.Max() + 1;
                var box = new PdfArray();
                for (int i = 0; i < 4; i++) box.Items.Add(new PdfNumber(0));
                var childDictionary = new PdfDictionary();
                childDictionary.Items["Type"] = new PdfName("XObject");
                childDictionary.Items["Subtype"] = new PdfName("Form");
                childDictionary.Items["BBox"] = box;
                objects[child] = new PdfIndirectObject(child, 0, new PdfStream(childDictionary, Array.Empty<byte>()));
                var original = (PdfStream)objects[5].Value;
                var resources = new PdfDictionary();
                var xObjects = new PdfDictionary();
                xObjects.Items["Empty"] = new PdfReference(child, 0);
                resources.Items["XObject"] = xObjects;
                original.Dictionary.Items["Resources"] = resources;
                objects[5] = new PdfIndirectObject(5, 0, new PdfStream(original.Dictionary, Encoding.ASCII.GetBytes(parentContent)));
                var annotations = new PdfArray();
                annotations.Items.Add(new PdfReference(7, 0)); annotations.Items.Add(new PdfReference(4, 0));
                ((PdfDictionary)objects[3].Value).Items["Annots"] = annotations;
                return security.InfoObjectNumber;
            });
        int visibleBytes = Encoding.ASCII.GetByteCount("1 0 0 rg 0 0 20 20 re f" + parentContent);
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxTotalDecodedStreamBytes = visibleBytes
        }});
        _ = document.Read();

        PdfDocument result = fill
            ? document.Forms.Fill(new Dictionary<string, string> { ["agree"] = "Yes" })
            : document.Forms.Flatten();

        _ = result.Read();
        Assert.Equal(Render(source), Render(result.ToBytes()));
        Assert.Equal(source, document.ToBytes());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void UnknownAppearancePlacementStillRequiresOriginalDecodeAdmission(bool nested, bool missingBox) {
        byte[] source = ReplaceCompressedAppearance(BuildPdf("[0 0 20 20]", "[0 0 20 20]", widgetOn: true), 5,
            new string(' ', 4096) + "0 g 0 0 20 20 re f");
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            PdfDictionary widget = Assert.IsType<PdfDictionary>(Assert.Single(objects.Values,
                item => item.Value is PdfDictionary dictionary && dictionary.Get<PdfStringObj>("T")?.Value == "agree").Value);
            var appearances = (PdfDictionary)widget.Items["AP"];
            var states = (PdfDictionary)appearances.Items["N"];
            int number = ((PdfReference)states.Items["Yes"]).ObjectNumber;
            var compressed = (PdfStream)objects[number].Value;
            var originalDictionary = new PdfDictionary();
            foreach (var item in compressed.Dictionary.Items) originalDictionary.Items[item.Key] = item.Value;
            if (missingBox) compressed.Dictionary.Items.Remove("BBox");
            else compressed.Dictionary.Items["BBox"] = new PdfName("Malformed");
            if (nested) {
                int child = objects.Keys.Max() + 1;
                objects[child] = new PdfIndirectObject(child, 0, compressed);
                originalDictionary.Items.Remove("Filter");
                var xObjects = new PdfDictionary(); xObjects.Items["Child"] = new PdfReference(child, 0);
                var resources = new PdfDictionary(); resources.Items["XObject"] = xObjects;
                originalDictionary.Items["Resources"] = resources;
                objects[number] = new PdfIndirectObject(number, 0,
                    new PdfStream(originalDictionary, Encoding.ASCII.GetBytes("/Child Do\n")));
            }
            return security.InfoObjectNumber;
        });
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxDecodedStreamBytes = 64
        }});

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.Forms.Edit(edit => edit.SetDefaultValue("agree", "Yes")));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(64, exception.Limit);
        Assert.Equal(source, document.ToBytes());
    }

    private static byte[] ReplaceCompressedAppearance(byte[] source, int number, string content) =>
        PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            var original = (PdfStream)objects[number].Value;
            var dictionary = new PdfDictionary();
            foreach (var item in original.Dictionary.Items) dictionary.Items[item.Key] = item.Value;
            dictionary.Items["Filter"] = new PdfName("FlateDecode");
            using var buffer = new MemoryStream();
            using (var compressor = new DeflateStream(buffer, CompressionLevel.Optimal, leaveOpen: true)) {
                byte[] bytes = Encoding.ASCII.GetBytes(content);
                compressor.Write(bytes, 0, bytes.Length);
            }
            objects[number] = new PdfIndirectObject(number, 0, new PdfStream(dictionary, buffer.ToArray()));
            return security.InfoObjectNumber;
        });
}
