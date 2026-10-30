using System.IO.Compression;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ProtectDocumentWorkflowTests {
    [Theory]
    [InlineData(".docx", false)]
    [InlineData(".xlsx", false)]
    [InlineData(".pptx", false)]
    [InlineData(".docx", true)]
    [InlineData(".xlsx", true)]
    [InlineData(".pptx", true)]
    public void ImplicitPersistenceCannotOverwriteInPlaceEncryptionIncludingDirectoryAliases(string extension, bool alias) {
        using var files = new ProtectionFiles();
        string realDirectory = files.PathFor("real");
        Directory.CreateDirectory(realDirectory);
        string source = files.Create("real/source" + extension);
        string output = source;
        byte[] before = File.ReadAllBytes(source);
        using var password = Password("open-secret");
        using var runspace = OpenRunspace();
        string aliasDirectory = files.PathFor("alias");
        using var directoryAlias = alias
            ? PhysicalPathTestSupport.CreateDirectoryAlias(runspace, aliasDirectory, realDirectory)
            : null;
        if (alias) {
            output = Path.Combine(aliasDirectory, "source" + extension);
        }
        using (IDisposable document = extension switch {
            ".docx" => WordDocument.Load(source, new WordLoadOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose }),
            ".xlsx" => ExcelDocument.Load(source, new ExcelLoadOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose }),
            _ => PowerPointPresentation.Load(source, new PowerPointLoadOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose })
        }) {
            Set(runspace, ("document", document), ("output", output), ("password", password));
            Assert.Contains("Explicit persistence", Fail(runspace,
                "Protect-OfficeDocument -Document $document -Encrypt -Password $password -OutputPath $output -ErrorAction Stop"));
            Assert.Equal(before, File.ReadAllBytes(source));
        }
        // Disposal must still leave the original document intact, rather than reverting an encrypted result.
        AssertContent(source, null);
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256, "\u00e9", "e\u0301", true)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128, "12345678901234567890123456789012-open", "12345678901234567890123456789012-owner", true)]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4, "12345678901234567890123456789012-open", "12345678901234567890123456789012-owner", true)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256, "open-secret", "owner-secret", false)]
    public void PdfOwnerWarningMatchesEffectiveCredentialAccess(PdfStandardEncryptionAlgorithm algorithm,
        string userText, string ownerText, bool equivalent) {
        using var files = new ProtectionFiles();
        string source = files.Create("source.pdf"), output = files.PathFor("protected.pdf");
        using var user = Password(userText);
        using var owner = Password(ownerText);
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("user", user), ("owner", owner), ("algorithm", algorithm));
        using var shell = Shell(runspace,
            "Protect-OfficeDocument -Path $source -OutputPath $output -Encrypt -Password $user -Restriction Copy -RestrictionPassword $owner -Options @{ Algorithm = $algorithm } -ErrorAction Stop");
        shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(equivalent, shell.Streams.Warning.Any(warning => warning.Message.Contains("owner authorization")));
        using var opened = PdfDocument.Load(output, new PdfLoadOptions { Password = userText });
        Assert.Equal(equivalent ? PdfPasswordAuthenticationRole.Owner : PdfPasswordAuthenticationRole.User,
            opened.InspectGeometryForViewing().Security.PasswordAuthenticationRole);
    }

    [Theory]
    [InlineData(".docx")]
    [InlineData(".xlsx")]
    [InlineData(".pptx")]
    [InlineData(".pdf")]
    public void NativeEncryptionReopensWithPasswordRejectsMissingPasswordAndPreservesSource(string extension) {
        using var files = new ProtectionFiles();
        string source = files.Create("source" + extension), output = files.PathFor("protected" + extension);
        byte[] before = File.ReadAllBytes(source);
        using var password = Password("open-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("password", password));
        var result = Invoke(runspace, "Protect-OfficeDocument -LiteralPath $source -Encrypt -Password $password -OutputPath $output -PassThru -ErrorAction Stop");
        Assert.Equal(output, Assert.IsType<FileInfo>(Assert.Single(result).BaseObject).FullName);
        Assert.True(File.Exists(output));
        Assert.ThrowsAny<Exception>(() => Read(output, null));
        AssertContent(output, "open-secret");
        Assert.Equal(before, File.ReadAllBytes(source));
    }

    [Theory]
    [InlineData(".docm")]
    [InlineData(".dotx")]
    [InlineData(".dotm")]
    [InlineData(".xlsm")]
    [InlineData(".xltx")]
    [InlineData(".xltm")]
    [InlineData(".xlam")]
    [InlineData(".pptm")]
    [InlineData(".potx")]
    [InlineData(".potm")]
    [InlineData(".ppsx")]
    [InlineData(".ppsm")]
    [InlineData(".ppam")]
    public void EncryptionPreservesNativePackageKind(string extension) {
        using var files = new ProtectionFiles();
        string source = files.Create("source" + extension), output = files.PathFor("protected" + extension);
        using var password = Password("open-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("password", password));
        Invoke(runspace, "Protect-OfficeDocument -LiteralPath $source -Encrypt -Password $password -OutputPath $output -ErrorAction Stop");
        using var saved = Load(output, "open-secret");
        string actual = saved switch {
            WordDocument word => word.SourceFormatDescriptor.Extension,
            ExcelDocument excel => excel.SourceFormatDescriptor.Extension,
            PowerPointPresentation presentation => presentation.SourceFormatDescriptor.Extension,
            _ => throw new InvalidOperationException()
        };
        Assert.Equal(extension, actual);
    }

    [Theory]
    [InlineData("Get-Item -LiteralPath $inputs | Protect-OfficeDocument")]
    [InlineData("$inputs | Protect-OfficeDocument")]
    [InlineData("Protect-OfficeDocument -LiteralPath $inputs")]
    public void FileInfoAndStringPathsPreserveLiteralFullNamesAndInputOrder(string command) {
        using var files = new ProtectionFiles();
        bool literal = command.Contains("LiteralPath", StringComparison.Ordinal);
        string first = files.Create(literal ? "first[2].xlsx" : "first.xlsx");
        string second = files.Create(literal ? "second[1].xlsx" : "second.xlsx");
        using var password = Password("open-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("inputs", new[] { second, first }), ("password", password));
        var result = Invoke(runspace, command + " -Encrypt -Password $password -PassThru -ErrorAction Stop");
        Assert.Equal(new[] { second, first }, result.Select(item => Assert.IsType<FileInfo>(item.BaseObject).FullName));
        foreach (string path in new[] { second, first }) {
            Assert.ThrowsAny<Exception>(() => Read(path, null));
            AssertContent(path, "open-secret");
        }
    }

    [Theory]
    [InlineData(WordDocumentProtectionType.ReadOnly)]
    [InlineData(WordDocumentProtectionType.Comments)]
    [InlineData(WordDocumentProtectionType.TrackedChanges)]
    [InlineData(WordDocumentProtectionType.Forms)]
    public void WordRestrictionsRoundTripWithoutEncryptingAndMixedFamiliesWarn(WordDocumentProtectionType mode) {
        using var files = new ProtectionFiles();
        string source = files.Create("source.docx"), output = files.PathFor("restricted.docx");
        byte[] before = File.ReadAllBytes(source);
        using var password = Password("edit-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("password", password), ("mode", mode.ToString()));
        using var shell = Shell(runspace, "Protect-OfficeDocument -Path $source -Restriction $mode, Copy -Password $password -OutputPath $output -ErrorAction Stop");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Contains(shell.Streams.Warning, warning => warning.Message.Contains("Copy"));
        using var saved = WordDocument.Load(output);
        Assert.Equal(mode, saved.Settings.ProtectionType);
        Assert.Equal(before, File.ReadAllBytes(source));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExcelStructureWindowsAndReservationUseNativeOptionsAndReadBack(bool nativeReservation) {
        using var files = new ProtectionFiles();
        string source = files.Create("source.xlsx"), output = files.PathFor("protected.xlsx");
        var options = new ExcelWorkbookProtectionOptions { ProtectStructure = true, LegacyPasswordHash = "abcd" };
        var reservationOptions = new ExcelWorkbookWriteReservationOptions { UserName = "Reporting", LegacyPasswordHash = "1234" };
        using var password = Password("edit-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("options", options), ("password", password), ("reservation", reservationOptions));
        Assert.Empty(Invoke(runspace, "Protect-OfficeDocument -Path $source -Restriction Structure, Windows, ReadOnlyRecommended -RestrictionPassword $password -Options $options, " +
            (nativeReservation ? "$reservation" : "@{ ReservedBy = 'Reporting'; ModifyPasswordHash = '1234' }") + " -OutputPath $output -ErrorAction Stop"));
        using var saved = ExcelDocument.Load(output);
        Assert.True(saved.IsWorkbookProtected);
        var reservation = saved.GetWriteReservation();
        Assert.True(reservation.ReadOnlyRecommended);
        Assert.Equal("Reporting", reservation.UserName);
        Assert.Equal("1234", reservation.LegacyPasswordHash);
        using var archive = ZipFile.OpenRead(output);
        using var xml = archive.GetEntry("xl/workbook.xml")!.Open();
        var protection = XDocument.Load(xml).Descendants().Single(element => element.Name.LocalName == "workbookProtection");
        Assert.Contains(protection.Attribute("lockStructure")!.Value, new[] { "1", "true" });
        Assert.Contains(protection.Attribute("lockWindows")!.Value, new[] { "1", "true" });
        Assert.Equal("ABCD", protection.Attribute("workbookPassword")!.Value);
        Assert.Equal("abcd", options.LegacyPasswordHash);
        Assert.False(options.ProtectWindows); // Option overrides do not mutate the caller's object.
        Assert.False(reservationOptions.ReadOnlyRecommended);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfPermissionsUseSeparateOwnerPasswordAndPreserveNativeSettingsAndContent(bool explicitOwner) {
        using var files = new ProtectionFiles();
        string source = files.Create("source.pdf"), output = files.PathFor("protected.pdf");
        byte[] before = File.ReadAllBytes(source);
        var options = new PdfStandardEncryptionOptions("unused-original-open") {
            OwnerPassword = "unused-original-owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes128,
            EncryptMetadata = false, AllowedPermissions = PdfStandardPermissions.All
        };
        using var open = Password("open-secret");
        using var owner = Password("owner-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("open", open), ("owner", owner), ("options", options));
        Assert.Empty(Invoke(runspace, "Protect-OfficeDocument -Path $source -Encrypt -Password $open -Restriction Print, Copy, Modify, Annotate, FillForms, Assemble " +
            (explicitOwner ? "-RestrictionPassword $owner " : "") +
            "-Options $options, @{ EncryptMetadata = $false } -OutputPath $output -ErrorAction Stop"));
        using var userDocument = PdfDocument.Load(output, new PdfLoadOptions { Password = "open-secret" });
        var security = userDocument.InspectGeometryForViewing().Security;
        Assert.Equal(PdfPasswordAuthenticationRole.User, security.PasswordAuthenticationRole);
        Assert.Equal(4, security.EncryptionRevision);
        Assert.False(security.EncryptMetadata);
        Assert.Equal(PdfStandardPermissions.Accessibility, security.AllowedStandardPermissions);
        Assert.Throws<PdfPermissionDeniedException>(() => userDocument.Read());
        using var ownerDocument = PdfDocument.Load(output, new PdfLoadOptions { Password = explicitOwner ? "owner-secret" : "unused-original-owner" });
        Assert.True(ownerDocument.Inspect().Security.HasOwnerAuthorization);
        Assert.Contains("MARKER", ownerDocument.Read().Text);
        Assert.Equal(before, File.ReadAllBytes(source));
        Assert.Equal("unused-original-open", options.UserPassword);
        Assert.Equal("unused-original-owner", options.OwnerPassword);
        Assert.Equal(PdfStandardPermissions.All, options.AllowedPermissions);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void PermissionsOnlyPdfOpensWithoutPasswordAndUsesSuppliedOwnerAndOptionOverrides(bool restrictionPassword, bool typedOptions) {
        using var files = new ProtectionFiles();
        string source = files.Create("source.pdf"), output = files.PathFor("absent/restricted.pdf");
        byte[] before = File.ReadAllBytes(source);
        var options = new PdfStandardEncryptionOptions("typed-open-password") {
            OwnerPassword = "typed-owner-password", Algorithm = PdfStandardEncryptionAlgorithm.Aes256,
            EncryptMetadata = true, AllowedPermissions = PdfStandardPermissions.All & ~PdfStandardPermissions.AssembleDocument
        };
        using var owner = Password("owner-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("owner", owner), ("options", options));
        string arguments = " -Restriction Copy, Modify " + (restrictionPassword ? "-RestrictionPassword" : "-Password") + " $owner -Options " +
            (typedOptions ? "$options, @{ Algorithm = 'Aes128'; EncryptMetadata = $false }" : "@{ AllowedPermissions = 'Print,CopyContents' }") +
            " -OutputPath $output -PassThru -ErrorAction Stop";
        Assert.Empty(Invoke(runspace, "Protect-OfficeDocument -Path $source -WhatIf" + arguments));
        Assert.Empty(Directory.GetDirectories(files.Root));
        var result = Invoke(runspace, "Protect-OfficeDocument -Path $source" + arguments);
        Assert.Equal(output, Assert.IsType<FileInfo>(Assert.Single(result).BaseObject).FullName);

        // Anonymous opening must succeed; no invented open password or owner authorization is involved.
        using var reader = PdfDocument.Load(output);
        var security = reader.InspectGeometryForViewing().Security;
        Assert.True(security.HasEncryption);
        Assert.False(security.HasOwnerAuthorization);
        Assert.Equal(typedOptions ? 4 : 6, security.EncryptionRevision);
        Assert.Equal(!typedOptions, security.EncryptMetadata);
        var expected = typedOptions
            ? options.AllowedPermissions & ~(PdfStandardPermissions.CopyContents | PdfStandardPermissions.ModifyContents)
            : PdfStandardPermissions.Print;
        Assert.Equal(expected, security.AllowedStandardPermissions);
        Assert.Throws<PdfPermissionDeniedException>(() => reader.Read());
        using var authorized = PdfDocument.Load(output, new PdfLoadOptions { Password = "owner-secret" });
        Assert.True(authorized.InspectGeometryForViewing().Security.HasOwnerAuthorization);
        Assert.Contains("MARKER", authorized.Read().Text);
        Assert.Equal(before, File.ReadAllBytes(source));
        Assert.Equal("typed-open-password", options.UserPassword);
        Assert.Equal("typed-owner-password", options.OwnerPassword);
        Assert.Equal(PdfStandardEncryptionAlgorithm.Aes256, options.Algorithm);
        Assert.True(options.EncryptMetadata);

        // Even an anonymously readable secured PDF requires current owner authentication to change security.
        string rejected = files.PathFor("rejected/restricted.pdf");
        Set(runspace, ("secured", output), ("rejected", rejected));
        Assert.Contains("input-password path", Fail(runspace,
            "Protect-OfficeDocument -Path $secured -Restriction Print -Password $owner -OutputPath $rejected -ErrorAction Stop"));
        Assert.False(Directory.Exists(System.IO.Path.GetDirectoryName(rejected)));
    }

    [Fact]
    public void LivePdfPermissionsReturnCallerOwnedReplacementWithoutRequiringAFileDestination() {
        using var files = new ProtectionFiles();
        using var live = PdfDocument.Create();
        live.Content.Paragraph(paragraph => paragraph.Text("MARKER"));
        using var owner = Password("owner-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("live", live), ("owner", owner));
        const string command = "$live | Protect-OfficeDocument -Restriction Print -RestrictionPassword $owner";
        Assert.Empty(Invoke(runspace, command + " -PassThru -WhatIf -ErrorAction Stop"));
        Assert.Empty(Invoke(runspace, command + " -ErrorAction Stop"));
        var result = Invoke(runspace, command + " -PassThru -ErrorAction Stop");
        using var protectedPdf = Assert.IsType<PdfDocument>(Assert.Single(result).BaseObject);
        Assert.NotSame(live, protectedPdf);
        Assert.True(protectedPdf.InspectGeometryForViewing().Security.HasEncryption);
        Assert.False(protectedPdf.InspectGeometryForViewing().Security.AllowsPrinting);
        Assert.False(live.InspectGeometryForViewing().Security.HasEncryption);
        Assert.Empty(Directory.GetFileSystemEntries(files.Root));
        string output = files.PathFor("restricted.pdf");
        protectedPdf.Save(output);
        using var anonymous = PdfDocument.Load(output);
        Assert.False(anonymous.InspectGeometryForViewing().Security.HasOwnerAuthorization);
        Assert.False(anonymous.InspectGeometryForViewing().Security.AllowsPrinting);
        Assert.Contains("MARKER", anonymous.Read().Text);
    }

    [Fact]
    public void PasswordTransformationAcceptsCredentialsAndWarnsForPlaintextWithoutEmittingSecrets() {
        using var files = new ProtectionFiles();
        string source = files.Create("source.pdf"), output = files.PathFor("protected.pdf");
        using var open = Password("open-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("credential", new PSCredential("Payroll", open)));
        using var shell = Shell(runspace, "Protect-OfficeDocument -Path $source -Encrypt -Password $credential -Restriction Copy -RestrictionPassword 'owner-secret' -Options @{ Algorithm = 'Aes128'; EncryptMetadata = $false } -OutputPath $output -ErrorAction Stop");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Contains(shell.Streams.Warning, warning => warning.Message.Contains("plain-text"));
        Assert.DoesNotContain(shell.Streams.Warning, warning => warning.Message.Contains("owner-secret"));
        using var saved = PdfDocument.Load(output, new PdfLoadOptions { Password = "owner-secret" });
        Assert.True(saved.Inspect().Security.HasOwnerAuthorization);
        Assert.False(saved.Inspect().Security.AllowsCopying);
        Assert.Contains("MARKER", saved.Read().Text);
    }

    [Fact]
    public void LiveRestrictionsMutateOnlyMemoryAndPassThruIsQuietUnlessRequested() {
        using var files = new ProtectionFiles();
        using var word = WordDocument.Create();
        word.AddParagraph("MARKER");
        using var excel = ExcelDocument.Create();
        excel.AddWorksheet("Data");
        using var password = Password("edit-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("word", word), ("excel", excel), ("password", password));
        Assert.Empty(Invoke(runspace, "$word | Protect-OfficeDocument -Restriction Comments -Password $password -ErrorAction Stop"));
        var result = Invoke(runspace, "$excel | Protect-OfficeDocument -Restriction Structure, ReadOnlyRecommended -PassThru -ErrorAction Stop");
        Assert.Same(excel, Assert.Single(result).BaseObject);
        Assert.Equal(WordDocumentProtectionType.Comments, word.Settings.ProtectionType);
        Assert.True(excel.IsWorkbookProtected);
        Assert.True(excel.GetWriteReservation().ReadOnlyRecommended);
        Assert.Empty(Directory.GetFileSystemEntries(files.Root));
        word.AddParagraph("STILL OPEN");
        excel.AddWorksheet("Still open");
    }

    [Theory]
    [InlineData(".docx", false)]
    [InlineData(".docx", true)]
    [InlineData(".xlsx", false)]
    [InlineData(".xlsx", true)]
    [InlineData(".pptx", false)]
    [InlineData(".pptx", true)]
    [InlineData(".pdf", false)]
    [InlineData(".pdf", true)]
    public void LiveEncryptionRequiresARealDestinationAndReopensSavedOutput(string extension, bool explicitOutput) {
        using var files = new ProtectionFiles();
        string source = files.Create("source" + extension), output = files.PathFor("protected" + extension);
        using var live = Load(source, null);
        using var detached = NewDocument(extension);
        using var password = Password("open-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("live", live), ("detached", detached), ("output", output), ("password", password));
        Assert.Contains("requires OutputPath", Fail(runspace, "$detached | Protect-OfficeDocument -Encrypt -Password $password -ErrorAction Stop"));
        var result = Invoke(runspace, "$live | Protect-OfficeDocument -Encrypt -Password $password " +
            (explicitOutput ? "-OutputPath $output " : "") + "-PassThru -ErrorAction Stop");
        object protectedDocument = Assert.Single(result).BaseObject;
        if (extension == ".pdf") {
            using var protectedPdf = Assert.IsType<PdfDocument>(protectedDocument);
            Assert.NotSame(live, protectedPdf);
            Assert.True(protectedPdf.Inspect().Security.HasEncryption);
        } else Assert.Same(live, protectedDocument);
        AssertContent(explicitOutput ? output : source, "open-secret");
    }

    [Theory]
    [InlineData("-Encrypt", "Password")]
    [InlineData("-Restriction ReadOnly", "password")]
    [InlineData("-Restriction ReadOnly, Comments -Password $password", "exactly one")]
    [InlineData("-Restriction ReadOnly -Password $password -Options @{ Unknown = 1 }", "Options")]
    [InlineData("-Restriction ReadOnly -Password $password -Options $wrongOptions", "format")]
    [InlineData("-Restriction 999 -Password $password", "Restriction")]
    public void InvalidArgumentsDoNotMutateLiveWordOrCreateOutputDirectories(string arguments, string expected) {
        using var files = new ProtectionFiles();
        using var word = WordDocument.Create();
        word.AddParagraph("MARKER");
        using var password = Password("edit-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("live", word), ("output", files.PathFor("absent/protected.docx")), ("password", password),
            ("wrongOptions", new ExcelWorkbookProtectionOptions()));
        Assert.Contains(expected, Fail(runspace, "Protect-OfficeDocument -Document $live -OutputPath $output " + arguments + " -ErrorAction Stop"), StringComparison.OrdinalIgnoreCase);
        Assert.Null(word.Settings.ProtectionType);
        Assert.Empty(Directory.GetFileSystemEntries(files.Root));
    }

    [Fact]
    public void WhatIfUnsupportedOnlyAndInvalidFormatLeaveFilesAndLiveStateUnchanged() {
        using var files = new ProtectionFiles();
        string source = files.Create("source.docx"), output = files.PathFor("absent/protected.docx");
        byte[] before = File.ReadAllBytes(source);
        using var word = WordDocument.Create();
        using var password = Password("open-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("wrong", files.PathFor("absent/protected.xlsx")),
            ("live", word), ("password", password));
        Assert.Empty(Invoke(runspace, "Protect-OfficeDocument -Path $source -Encrypt -Password $password -OutputPath $output -WhatIf -PassThru -ErrorAction Stop"));
        Assert.Empty(Invoke(runspace, "Protect-OfficeDocument -Document $live -Restriction Forms -Password $password -OutputPath $output -WhatIf -PassThru -ErrorAction Stop"));
        Assert.Empty(Invoke(runspace, "Protect-OfficeDocument -Path $source -Restriction Copy -OutputPath $output -PassThru -ErrorAction Stop"));
        Assert.Contains("filename format", Fail(runspace, "Protect-OfficeDocument -Path $source -Restriction ReadOnly -Password $password -OutputPath $wrong -ErrorAction Stop"));
        Assert.Null(word.Settings.ProtectionType);
        Assert.Equal(before, File.ReadAllBytes(source));
        Assert.Empty(Directory.GetDirectories(files.Root));
        Assert.Single(Directory.GetFiles(files.Root));
    }

    [Fact]
    public void InvalidReservationHashIsRejectedBeforeWorkbookOrReservationMutation() {
        using var files = new ProtectionFiles();
        using var live = ExcelDocument.Create();
        live.AddWorksheet("Data");
        using var runspace = OpenRunspace();
        Set(runspace, ("live", live), ("output", files.PathFor("absent/protected.xlsx")));
        Assert.Contains("hexadecimal", Fail(runspace, "$live | Protect-OfficeDocument -Restriction Structure, ReadOnlyRecommended -Options @{ ModifyPasswordHash = 'not-a-hash' } -OutputPath $output -ErrorAction Stop"));
        Assert.False(live.IsWorkbookProtected);
        Assert.False(live.GetWriteReservation().Exists);
        Assert.Empty(Directory.GetFileSystemEntries(files.Root));
    }

    [Fact]
    public void MislabeledLaterNativeInputDoesNotProtectAnEarlierFile() {
        using var files = new ProtectionFiles();
        string source = files.Create("first.docx"), mislabeled = files.PathFor("second.docx");
        using (var template = WordDocument.Create(new WordCreateOptions { DocumentType = WordDocumentType.Template })) {
            template.AddParagraph("TEMPLATE");
            using var stream = File.Create(mislabeled);
            template.Save(stream);
        }
        byte[] before = File.ReadAllBytes(source), other = File.ReadAllBytes(mislabeled);
        using var password = Password("open-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("inputs", new[] { source, mislabeled }), ("password", password));
        Assert.Contains("format", Fail(runspace, "Protect-OfficeDocument -LiteralPath $inputs -Encrypt -Password $password -ErrorAction Stop"));
        Assert.Equal(before, File.ReadAllBytes(source));
        Assert.Equal(other, File.ReadAllBytes(mislabeled));
    }

    [Fact]
    public void PdfInvalidOptionsAndMissingPermissionsOwnerCreateNoDestination() {
        using var files = new ProtectionFiles();
        string source = files.Create("source.pdf"), output = files.PathFor("absent/protected.pdf");
        byte[] before = File.ReadAllBytes(source);
        using var password = Password("open-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("output", output), ("password", password));
        Assert.Contains("owner password", Fail(runspace,
            "Protect-OfficeDocument -Path $source -Restriction Copy -OutputPath $output -ErrorAction Stop"));
        foreach (string arguments in new[] {
            "-Restriction Copy -Password ''",
            "-Restriction Copy -RestrictionPassword ''",
            "-Restriction Copy -Password $password -Options @{ Algorithm = 999 }",
            "-Restriction Copy -Password $password -Options @{ Unknown = 1 }",
            "-Encrypt -Password $password -Options @{ Algorithm = 999 }",
            "-Encrypt -Password $password -Options @{ Unknown = 1 }"
        }) Fail(runspace, "Protect-OfficeDocument -Path $source -OutputPath $output " + arguments + " -ErrorAction Stop");
        Assert.Equal(before, File.ReadAllBytes(source));
        Assert.Empty(Directory.GetDirectories(files.Root));
    }

    [Theory]
    [InlineData(".docx")]
    [InlineData(".xlsx")]
    [InlineData(".pptx")]
    [InlineData(".pdf")]
    public void AlreadyEncryptedInputsAreNeverRewrittenToPlaintext(string extension) {
        using var files = new ProtectionFiles();
        string source = files.Create("source" + extension), encrypted = files.PathFor("encrypted" + extension);
        string rejected = files.PathFor("absent/rejected" + extension);
        using var password = Password("open-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("encrypted", encrypted), ("rejected", rejected), ("password", password));
        Invoke(runspace, "Protect-OfficeDocument -Path $source -Encrypt -Password $password -OutputPath $encrypted -ErrorAction Stop");
        byte[] before = File.ReadAllBytes(encrypted);
        Fail(runspace, "Protect-OfficeDocument -Path $encrypted -Encrypt -Password $password -OutputPath $rejected -ErrorAction Stop");
        Assert.Equal(before, File.ReadAllBytes(encrypted));
        Assert.Empty(Directory.GetDirectories(files.Root));
        AssertContent(encrypted, "open-secret");
    }

    [Fact]
    public void NativeSignatureMetadataBlocksBeforeLiveMutationAndDestinationCreation() {
        using var files = new ProtectionFiles();
        string source = files.Create("signed.docx"), output = files.PathFor("absent/protected.docx");
        // Structural signature fixture only: no real certificate, key or signing provider is acquired.
        using (var package = WordprocessingDocument.Open(source, true)) package.AddDigitalSignatureOriginPart();
        byte[] before = File.ReadAllBytes(source);
        using var live = WordDocument.Load(source);
        using var password = Password("edit-secret");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("live", live), ("output", output), ("password", password));
        Assert.Contains("signed", Fail(runspace, "Protect-OfficeDocument -Path $source -Encrypt -Password $password -OutputPath $output -ErrorAction Stop"));
        Assert.Contains("signed", Fail(runspace, "$live | Protect-OfficeDocument -Restriction Forms -Password $password -OutputPath $output -ErrorAction Stop"));
        Assert.Null(live.Settings.ProtectionType);
        Assert.Equal(before, File.ReadAllBytes(source));
        Assert.Empty(Directory.GetDirectories(files.Root));
    }

    private static IDisposable NewDocument(string extension) => extension switch {
        ".docx" or ".docm" or ".dotx" or ".dotm" => WordDocument.Create(),
        ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlam" => ExcelDocument.Create(),
        ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" or ".ppam" => PowerPointPresentation.Create(),
        _ => PdfDocument.Create()
    };

    private static IDisposable Load(string path, string? password) => System.IO.Path.GetExtension(path) switch {
        ".docx" or ".docm" or ".dotx" or ".dotm" => WordDocument.Load(path, new WordLoadOptions { Password = password }),
        ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlam" => ExcelDocument.Load(path, new ExcelLoadOptions { Password = password }),
        ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" or ".ppam" => PowerPointPresentation.Load(path, new PowerPointLoadOptions { Password = password }),
        _ => PdfDocument.Load(path, new PdfLoadOptions { Password = password })
    };

    private static string Read(string path, string? password) {
        using var document = Load(path, password);
        return document switch {
            WordDocument word => string.Join("\n", word.Paragraphs.Select(paragraph => paragraph.Text)),
            ExcelDocument excel => excel.Sheets.Single().TryGetCellText(1, 1, out string text) ? text : string.Empty,
            PowerPointPresentation presentation => presentation.Slides.Count.ToString(),
            PdfDocument pdf => pdf.Read().Text,
            _ => throw new InvalidOperationException()
        };
    }

    private static void AssertContent(string path, string? password) => Assert.Contains(
        System.IO.Path.GetExtension(path) == ".pptx" ? "1" : "MARKER", Read(path, password));

    private static Runspace OpenRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(ProtectOfficeDocumentCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
    private static void Set(Runspace runspace, params (string Name, object Value)[] values) {
        foreach (var value in values) runspace.SessionStateProxy.SetVariable(value.Name, value.Value);
    }
    private static System.Management.Automation.PowerShell Shell(Runspace runspace, string script) {
        var shell = System.Management.Automation.PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script);
        return shell;
    }
    private static PSObject[] Invoke(Runspace runspace, string script) {
        using var shell = Shell(runspace, script);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        return result.ToArray();
    }
    private static string Fail(Runspace runspace, string script) {
        using var shell = Shell(runspace, "try { " + script + "; throw 'PROTECT_UNEXPECTED_SUCCESS' } catch { $_.Exception.Message }");
        var result = shell.Invoke();
        string message = Assert.IsType<string>(Assert.Single(result).BaseObject);
        Assert.DoesNotContain("PROTECT_UNEXPECTED_SUCCESS", message);
        return message;
    }
    private static SecureString Password(string text) {
        var secure = new SecureString();
        foreach (char character in text) secure.AppendChar(character);
        secure.MakeReadOnly();
        return secure;
    }

    private sealed class ProtectionFiles : IDisposable {
        internal string Root { get; } = Directory.CreateTempSubdirectory("officeimo-protection-command-").FullName;
        internal string PathFor(string name) => System.IO.Path.GetFullPath(System.IO.Path.Combine(Root, name));
        internal string Create(string name) {
            string path = PathFor(name);
            using var document = NewDocument(System.IO.Path.GetExtension(path));
            switch (document) {
                case WordDocument word: word.AddParagraph("MARKER"); word.Save(path); break;
                case ExcelDocument excel: excel.AddWorksheet("Data").CellValue(1, 1, "MARKER"); excel.Save(path); break;
                case PowerPointPresentation presentation: presentation.AddSlide(); presentation.Save(path); break;
                case PdfDocument pdf: pdf.Content.Paragraph(paragraph => paragraph.Text("MARKER")); pdf.Save(path); break;
            }
            return path;
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
}
