using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf
{
    public sealed class PdfDocumentInputAdmissionTests {
        [Theory]
        [InlineData("")]
        [InlineData("This is not a PDF document.")]
        [InlineData("%PDF-1.7\nbroken")]
        public void LenientReadRejectsInputWithNoRecoverableObjects(string input) {
            byte[] bytes = Encoding.ASCII.GetBytes(input);

            AssertNoObjects(() => PdfReadDocument.Open(bytes));
            PdfDocument document = PdfDocument.Load(bytes);
            AssertNoObjects(() => document.InspectForViewing());
            AssertNoObjects(() => document.InspectGeometryForViewing());
            AssertNoObjects(() => document.Read());

            PdfDocumentPreflight preflight = document.Preflight();
            Assert.False(preflight.CanRead);
            Assert.True(preflight.HasReadBlocker(input.StartsWith("%PDF-", StringComparison.Ordinal)
                ? PdfReadBlockerKind.ParserUnsupported : PdfReadBlockerKind.MissingHeader));
        }

        [Fact]
        public async Task FileAndStreamSnapshotsUseTheSameReadAdmission() {
            byte[] bytes = Encoding.ASCII.GetBytes("This is not a PDF document.");
            string path = Path.Combine(Path.GetTempPath(), "officeimo-invalid-pdf-" + Guid.NewGuid().ToString("N") + ".pdf");
            File.WriteAllBytes(path, bytes);
            try {
                AssertNoObjects(() => PdfReadDocument.Open(path));
                AssertNoObjects(() => PdfDocument.Load(path).InspectForViewing());
                PdfDocument asyncFile = await PdfDocument.LoadAsync(path);
                AssertNoObjects(() => asyncFile.InspectForViewing());

                using MemoryStream stream = new MemoryStream(bytes);
                stream.Position = 5;
                AssertNoObjects(() => PdfReadDocument.Open(stream));
                Assert.Equal(5, stream.Position);
                AssertNoObjects(() => PdfDocument.Load(stream).InspectForViewing());
                PdfDocument asyncStream = await PdfDocument.LoadAsync(stream);
                AssertNoObjects(() => asyncStream.InspectGeometryForViewing());
                Assert.True(stream.CanRead);
                Assert.Equal(5, stream.Position);
            } finally { File.Delete(path); }
        }

        [Fact]
        public void ZeroPageCatalogAndForensicObjectsRemainReadable() {
            byte[] zeroPages = Encoding.ASCII.GetBytes(
                "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
                "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
                "trailer\n<< /Root 1 0 R /Size 3 >>\n%%EOF\n");
            byte[] fragment = Encoding.ASCII.GetBytes(
                "1 0 obj\n<< /Producer (Forensic fragment) >>\nendobj\n");

            Assert.Empty(PdfReadDocument.Open(zeroPages).Pages);
            PdfReadDocument forensic = PdfReadDocument.Open(fragment);
            Assert.Single(forensic.Objects);
            Assert.Empty(forensic.Pages);
            Assert.Contains(forensic.RepairReport.Diagnostics, diagnostic => diagnostic.Code == "MissingStartXref");
        }

        [Fact]
        public void GeneratedDiagnosticReportsUseOneSerializedSnapshot() {
            int compositionCalls = 0;
            PdfDocument CreateDocument() => PdfDocument.Create(new PdfOptions {
                TextLineBreakCallback = text => {
                    compositionCalls++;
                    return new[] { text.Length };
                }
            }).Paragraph(paragraph => paragraph.Text(new string('W', 600)));

            CreateDocument().ToBytes();
            int callsForOneSerialization = compositionCalls;
            Assert.True(callsForOneSerialization > 0);
            Action<PdfDocument>[] reports = {
                document => Assert.True(document.Preflight().CanRead),
                document => Assert.True(document.PlanMutation(PdfMutationOperation.ExtractPages).CanExecute),
                document => Assert.True(document.AssessMutations(new[] { PdfMutationOperation.ExtractPages }).CanExecuteAll)
            };
            foreach (Action<PdfDocument> report in reports) {
                compositionCalls = 0;
                report(CreateDocument());
                Assert.Equal(callsForOneSerialization, compositionCalls);
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("This is not a PDF document.")]
        [InlineData("%PDF-1.7\nbroken")]
        public void MutationReportsBlockUnreadableInputWithoutThrowingOrChangingBytes(string input) {
            byte[] bytes = Encoding.ASCII.GetBytes(input);
            PdfDocument document = PdfDocument.Load(bytes);

            PdfMutationPlan plan = document.PlanMutation(PdfMutationOperation.ExtractPages);
            Assert.False(plan.CanExecute);
            Assert.False(plan.Preflight.CanRead);
            Assert.NotEmpty(plan.Diagnostics);

            PdfMutationPortfolioReport portfolio = document.AssessMutations();
            Assert.False(portfolio.Preflight.CanRead);
            Assert.Empty(portfolio.ExecutablePlans);
            Assert.NotEmpty(portfolio.BlockedPlans);
            Assert.All(portfolio.Plans, blocked => {
                Assert.False(blocked.CanExecute);
                Assert.NotEmpty(blocked.Diagnostics);
                Assert.Same(portfolio.Preflight, blocked.Preflight);
            });
            Assert.Equal(bytes, document.ToBytes());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void MutationReportsCaptureWrongPasswordsWithoutAuthorizingOperations(bool withOutline) {
            byte[] bytes = PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Protected report input")),
                new PdfOptions { CreateOutlineFromHeadings = withOutline }.SetEncryption("open", "owner")).ToBytes();
            PdfLoadOptions incorrect = new PdfLoadOptions { Password = "incorrect" };
            PdfDocument document = PdfDocument.Load(bytes, incorrect);

            Assert.True(document.Preflight().HasReadBlocker(PdfReadBlockerKind.Encryption));
            Assert.True(PdfDocument.Preflight(bytes, incorrect).HasReadBlocker(PdfReadBlockerKind.Encryption));
            PdfMutationPlan plan = document.PlanMutation(PdfMutationOperation.ExtractPages);
            Assert.False(plan.CanExecute);
            Assert.Contains("Read.Encryption", plan.BlockerCodes);
            PdfMutationPortfolioReport portfolio = document.AssessMutations();
            Assert.Empty(portfolio.ExecutablePlans);
            Assert.Equal(PdfPasswordAuthenticationRole.None, portfolio.Preflight.Probe.Security.PasswordAuthenticationRole);
            PdfOperationResult<PdfDocument> result = document.Pages.ExtractResult("1");
            Assert.False(result.CanAttempt);
            Assert.Null(result.Value);
            Assert.NotEmpty(result.Diagnostics);
            Assert.Equal(bytes, document.ToBytes());

            Assert.Throws<PdfInvalidPasswordException>(() => document.Read());
            Assert.Throws<PdfInvalidPasswordException>(() => PdfInspector.Probe(bytes, incorrect));
            PdfLoadOptions authorized = new PdfLoadOptions { Password = "open" };
            Assert.True(document.PlanMutation(PdfMutationOperation.ExtractPages, options: authorized).CanExecute);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void MetadataReportsRetainStrictParserBlockersWithoutReparsingInvalidInput(bool withStartXref) {
            byte[] bytes = Encoding.ASCII.GetBytes("%PDF-1.7\ntrailer\n<< /Root 1 0 R /Size 2 >>\n" +
                (withStartXref ? "startxref\n0\n" : string.Empty) + "%%EOF\n");
            PdfDocument document = PdfDocument.Load(bytes, new PdfLoadOptions { ParsingMode = PdfParsingMode.Strict });

            PdfDocumentPreflight preflight = document.Preflight();
            Assert.False(preflight.CanRead);
            Assert.False(preflight.CanAppendMetadataRevision);
            Assert.False(preflight.CanAppendFormFieldRevision);
            Assert.False(preflight.CanPrepareExternalSignatureRevision);
            Assert.False(preflight.Can(PdfPreflightCapability.AppendMetadataRevision));
            Assert.False(preflight.Can(PdfPreflightCapability.AppendFormFieldRevision));
            Assert.False(preflight.Can(PdfPreflightCapability.PrepareExternalSignatureRevision));
            PdfMutationOperation[] operations = { PdfMutationOperation.UpdateMetadata, PdfMutationOperation.SynchronizeMetadata };
            foreach (PdfMutationOperation operation in operations) {
                PdfMutationPlan plan = document.PlanMutation(operation);
                Assert.False(plan.CanExecute);
                Assert.Contains("Read.ParserUnsupported", plan.BlockerCodes);
            }
            PdfMutationPortfolioReport portfolio = document.AssessMutations();
            Assert.Empty(portfolio.ExecutablePlans);
            PdfOperationResult<PdfDocument>[] results = {
                document.UpdateMetadataResult(title: "Blocked update"),
                document.SynchronizeMetadataResult(title: "Blocked synchronization")
            };
            foreach (PdfOperationResult<PdfDocument> result in results) {
                Assert.False(result.CanAttempt);
                Assert.False(result.Succeeded);
                Assert.Null(result.Value);
                Assert.NotEmpty(result.Diagnostics);
                Assert.True(result.Preflight.HasReadBlocker(PdfReadBlockerKind.ParserUnsupported));
            }
            Assert.Equal(bytes, document.ToBytes());
        }

        [Fact]
        public void RawProbeRetainsEvidenceWhenThereAreNoReadableObjects() {
            byte[] bytes = Encoding.ASCII.GetBytes("%PDF-1.7\nbroken");

            var (objects, _) = PdfSyntax.ParseObjects(bytes);
            Assert.Empty(objects);
            Assert.Equal("1.7", PdfInspector.Probe(bytes).HeaderVersion);
            PdfArtifactSnapshot artifact = PdfArtifactSnapshot.Capture(bytes);
            Assert.Null(artifact.PageCount);
            Assert.Equal(bytes.LongLength, artifact.ByteCount);
            Assert.False(string.IsNullOrEmpty(artifact.Sha256));
        }

        private static void AssertNoObjects(Action read) {
            PdfParseException error = Assert.Throws<PdfParseException>(read);
            Assert.Equal("NoIndirectObjects", error.Code);
            Assert.Null(error.ObjectNumber);
        }
    }
}
