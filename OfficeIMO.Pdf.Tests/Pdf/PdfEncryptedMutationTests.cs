using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEncryptedMutationTests {
    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4, false)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128, false)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256, false)]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4, true)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128, true)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256, true)]
    public void ContentPageAndFormRewritesRetainPasswordSecurity(PdfStandardEncryptionAlgorithm algorithm, bool useUserAuthorization) {
        byte[] source = CreateEncryptedSource(algorithm);
        var options = new PdfLoadOptions {
            Password = useUserAuthorization ? "open" : "owner",
            PermissionPolicy = useUserAuthorization ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce
        };
        PdfDocument current = PdfDocument.Load(source, options);
        current = current.Text.ReplaceAll("ORIGINAL", "CHANGED").Document;
        AssertProtected(source, current.ToBytes());
        Assert.Contains("CHANGED", current.Reader.Text());
        current = current.Stamp.Text("Added stamp");
        AssertProtected(source, current.ToBytes());
        Assert.Contains("Added stamp", current.Reader.Text());
        current = current.Pages.Rotate(90, 1);
        AssertProtected(source, current.ToBytes());
        Assert.Equal(90, PdfReadDocument.Open(current.ToBytes(), options).Pages[0].GetRotationDegrees());
        current = current.Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "Entry", Kind = PdfFormFieldCreationKind.Text, PageNumber = 1,
            X = 20, Y = 20, Width = 120, Height = 30
        })).ToDocument();
        AssertProtected(source, current.ToBytes());
        current = current.Forms.Fill(new Dictionary<string, string> { ["Entry"] = "Protected value" });
        AssertProtected(source, current.ToBytes());
        Assert.Equal("Protected value", Assert.Single(PdfReadDocument.Open(current.ToBytes(), options).FormFields).Value);
        current = current.Forms.Flatten();
        AssertProtected(source, current.ToBytes());
        Assert.Empty(PdfReadDocument.Open(current.ToBytes(), options).FormFields);
        Assert.Contains("Protected value", current.Reader.Text());
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void MergeRetainsPrimarySecurityAndReportsActualOutput(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] source = CreateEncryptedSource(algorithm);
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Password = "owner" });
        PdfDocument extra = PdfDocument.Create(pdf => pdf.Content(c => c.Paragraph(p => p.Text("Imported page"))));
        PdfDocument merged = document.MergeWith(extra);
        AssertProtected(source, merged.ToBytes());
        Assert.Equal(2, PdfReadDocument.Open(merged.ToBytes(), new PdfLoadOptions { Password = "owner" }).Pages.Count);
        Assert.Contains("Imported page", merged.Reader.Text());
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void RedactionRetainsEncryptionWithoutRetainingOriginalText(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] source = CreateEncryptedSource(algorithm);
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Password = "owner" });
        PdfTextMatch match = Assert.Single(document.Text.Find("ORIGINAL"));
        PdfDocument output = document.Redactions.Apply(new[] { new PdfRedactionArea(1, match.X, match.Y, match.Width, match.Height) });
        AssertProtected(source, output.ToBytes());
        Assert.DoesNotContain("ORIGINAL", output.Reader.Text());
        Assert.Contains("neighbor", output.Reader.Text());
        Assert.False(output.ToBytes().Take(source.Length).SequenceEqual(source));
        PdfSecurityMutationResult plaintext = output.Security.Decrypt("owner");
        Assert.False(plaintext.IsEncrypted);
        Assert.DoesNotContain("ORIGINAL", plaintext.ToDocument().Reader.Text());
        Assert.DoesNotContain("ORIGINAL", System.Text.Encoding.ASCII.GetString(plaintext.ToDocument().ToBytes()));
    }

    [Fact]
    public void UserRestrictionsStillBlockOrdinaryEditing() {
        byte[] source = CreateEncryptedSource(PdfStandardEncryptionAlgorithm.Aes256);
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Password = "open" });
        Assert.Throws<PdfPermissionDeniedException>(() => document.Text.ReplaceAll("ORIGINAL", "CHANGED"));
        Assert.Equal(source, document.ToBytes());
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void InterleaveRetainsPrimarySecurityAndAuthenticatedReadback(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] source = CreateEncryptedSource(algorithm);
        byte[] incoming = PdfDocument.Create().Paragraph(p => p.Text("Incoming")).ToBytes();
        PdfInterleaveResult output = PdfPageInterleaver.Interleave(new[] {
            new PdfInterleaveSource(source, "Protected") { ReadOptions = new PdfLoadOptions { Password = "owner" } },
            new PdfInterleaveSource(incoming, "Incoming")
        });
        AssertProtected(source, output.ToBytes());
        Assert.Equal(2, output.Pages.Count);
        Assert.Contains("Incoming", output.ToDocument().Reader.Text());
        Assert.True(output.MergeReport.OutputHasEncryption);
    }

    [Fact]
    public void EncryptedInterleaveUsesComposedObjectAndPageBudgets() {
        byte[] source = CreateEncryptedSource(PdfStandardEncryptionAlgorithm.Aes256);
        byte[] incoming = PdfDocument.Create().Paragraph(p => p.Text("Incoming")).ToBytes();
        var primaryOptions = new PdfLoadOptions {
            Password = "owner",
            Limits = new PdfReadLimits { MaxPages = 1, MaxIndirectObjects = PdfReadDocument.Open(source, new PdfLoadOptions { Password = "owner" }).RawStructure().TotalObjectCount }
        };
        var incomingOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPages = 1, MaxIndirectObjects = PdfReadDocument.Open(incoming).RawStructure().TotalObjectCount }
        };
        PdfInterleaveResult output = PdfPageInterleaver.Interleave(new[] {
            new PdfInterleaveSource(source) { ReadOptions = primaryOptions },
            new PdfInterleaveSource(incoming) { ReadOptions = incomingOptions }
        });
        AssertProtected(source, output.ToBytes());
        Assert.Equal(2, output.ToDocument().Reader.Pages().Count);
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void MetadataExemptionAndRewriteBudgetsRemainEffective(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] source = PdfDocument.Create(pdf => pdf.Content(c => c.Paragraph(p => p.Text("Protected text"))),
            new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", Algorithm = algorithm, EncryptMetadata = false
            })).ToBytes();
        var options = new PdfLoadOptions { Password = "owner" };
        PdfDocument output = PdfDocument.Load(source, options).SynchronizeMetadata(title: "Visible XMP", createXmpMetadata: true).Pages.Rotate(90, 1);
        Assert.False(PdfReadDocument.Open(output.ToBytes(), options).Security.EncryptMetadata);
        Assert.Contains("Visible XMP", System.Text.Encoding.ASCII.GetString(output.ToBytes()));
        AssertProtected(source, output.ToBytes());
        Assert.Throws<InvalidDataException>(() => PdfDocumentObjectGraphRewriter.Rewrite(source, options, null, maximumOutputBytes: 100));
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => PdfDocumentObjectGraphRewriter.Rewrite(source, options, null, cancellationToken: cancellation.Token));
        Assert.True(PdfReadDocument.Open(source, options).Security.HasEncryption);
    }

    [Theory]
    [InlineData(PdfStandardPermissions.FillForms, false)]
    [InlineData(PdfStandardPermissions.ModifyContents, false)]
    [InlineData(PdfStandardPermissions.ModifyAnnotations, false)]
    [InlineData(PdfStandardPermissions.ModifyContents | PdfStandardPermissions.ModifyAnnotations, true)]
    public void StructuralFormEditsRequireBothUserPermissions(PdfStandardPermissions permissions, bool allowed) {
        PdfDocument source = PdfDocument.Create().Paragraph(p => p.Text("Form document"));
        source = source.Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "Original", Kind = PdfFormFieldCreationKind.Text, PageNumber = 1,
            X = 20, Y = 20, Width = 100, Height = 20
        })).ToDocument();
        byte[] protectedSource = source.Security.Encrypt(new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes256, AllowedPermissions = permissions | PdfStandardPermissions.CopyContents
        }).ToDocument().ToBytes();
        PdfDocument user = PdfDocument.Load(protectedSource, new PdfLoadOptions { Password = "open" });
        Action<PdfAcroFormEditSession>[] edits = {
            edit => edit.Create(new PdfFormFieldCreateOptions { Name = "New", Kind = PdfFormFieldCreationKind.Text, PageNumber = 1, X = 20, Y = 50, Width = 100, Height = 20 }),
            edit => edit.Rename("Original", "Renamed"),
            edit => edit.Remove("Original")
        };
        foreach (Action<PdfAcroFormEditSession> edit in edits) {
            if (allowed) AssertProtected(protectedSource, user.Forms.Edit(edit).ToDocument().ToBytes());
            else Assert.Throws<PdfMutationBlockedException>(() => user.Forms.Edit(edit));
            Assert.Equal(protectedSource, user.ToBytes());
        }
        PdfDocument owner = PdfDocument.Load(protectedSource, new PdfLoadOptions { Password = "owner" });
        Assert.Single(owner.Forms.Edit(edit => edit.Rename("Original", "OwnerRenamed")).ToDocument().Reader.FormFields());
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void EncryptionProtectionAcceptsMeasuredGeneratedObjectGrowth(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] source = CreateEncryptedSource(algorithm);
        int count = PdfReadDocument.Open(source, new PdfLoadOptions { Password = "owner" }).RawStructure().TotalObjectCount;
        var options = new PdfLoadOptions { Password = "owner", Limits = new PdfReadLimits { MaxIndirectObjects = count } };
        PdfDocument output = PdfDocument.Load(source, options).Stamp.Text("Generated stamp");
        AssertProtected(source, output.ToBytes());
        Assert.Contains("Generated stamp", output.Reader.Text());
        Assert.Equal(count, options.Limits.MaxIndirectObjects);
        Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(output.ToBytes(), options));
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void EncryptedMergeAccountsForCiphertextStreamGrowth(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] source = CreateEncryptedSource(algorithm);
        byte[] incoming = PdfDocument.Create().Paragraph(p => p.Text(string.Join(" ", Enumerable.Range(0, 150).Select(index => "item" + index.ToString("X4"))))).ToBytes();
        int largestIncomingStream = PdfReadDocument.Open(incoming).Objects.Values
            .Where(item => item.Value is PdfStream).Max(item => ((PdfStream)item.Value).Data.Length);
        int largestSourceStream = PdfReadDocument.Open(source, new PdfLoadOptions { Password = "owner" }).Objects.Values
            .Where(item => item.Value is PdfStream).Max(item => (int)((PdfStream)item.Value).Dictionary.Get<PdfNumber>("Length")!.Value);
        Assert.True(largestIncomingStream > largestSourceStream);
        PdfInterleaveResult output = PdfPageInterleaver.Interleave(new[] {
            new PdfInterleaveSource(source) { ReadOptions = new PdfLoadOptions { Password = "owner", Limits = new PdfReadLimits { MaxRawStreamBytes = largestSourceStream } } },
            new PdfInterleaveSource(incoming) { ReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxRawStreamBytes = largestIncomingStream } } }
        });
        AssertProtected(source, output.ToBytes());
        Assert.Equal(2, output.ToDocument().Reader.Pages().Count);
    }

    private static byte[] CreateEncryptedSource(PdfStandardEncryptionAlgorithm algorithm) =>
        PdfDocument.Create(pdf => pdf.Content(c => c.Paragraph(p => p.Text("ORIGINAL neighbor"))),
            new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", Algorithm = algorithm, AllowedPermissions = PdfStandardPermissions.Print
            })).ToBytes();

    private static void AssertProtected(byte[] source, byte[] output) {
        PdfReadDocument before = PdfReadDocument.Open(source, new PdfLoadOptions { Password = "owner" });
        PdfReadDocument user = PdfReadDocument.Open(output, new PdfLoadOptions { Password = "open" });
        Assert.True(user.Security.HasEncryption);
        Assert.Equal(before.Security.EncryptionRevision, user.Security.EncryptionRevision);
        Assert.Equal(before.Security.EncryptionLengthBits, user.Security.EncryptionLengthBits);
        Assert.Equal(before.Security.AllowedStandardPermissions, user.Security.AllowedStandardPermissions);
        Assert.Equal(PdfSyntax.ReadPermanentTrailerIdentifier(before.TrailerRaw), PdfSyntax.ReadPermanentTrailerIdentifier(user.TrailerRaw));
        Assert.Equal(PdfPasswordAuthenticationRole.Owner,
            PdfReadDocument.Open(output, new PdfLoadOptions { Password = "owner" }).Security.PasswordAuthenticationRole);
        Assert.Throws<PdfPasswordRequiredException>(() => PdfReadDocument.Open(output));
        Assert.Throws<PdfInvalidPasswordException>(() => PdfReadDocument.Open(output, new PdfLoadOptions { Password = "wrong" }));
    }
}
