using ModelContextProtocol.Protocol;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Mcp;
using Xunit;

namespace OfficeIMO.Tool.Tests {
    public sealed class AgentPdfPrintPlanPermissionTests {
        [Theory]
        [InlineData(null, new[] { 1, 2, 3 })]
        [InlineData("last,1,last", new[] { 3, 1, 3 })]
        public async Task McpPlansPrintPermittedUserDocumentsWithoutCopyPermission(string? pages, int[] expectedPages) {
            using PrintFixture fixture = new(PdfStandardPermissions.Print);
            PdfDocument authenticated = PdfDocument.Load(fixture.Bytes, new PdfLoadOptions { Password = PrintFixture.UserPassword });
            Assert.Equal(3, authenticated.GetPrintablePageLayouts().Count);
            PdfPermissionDeniedException copying = Assert.Throws<PdfPermissionDeniedException>(() => authenticated.Inspect());
            Assert.Equal(PdfStandardPermissions.CopyContents, copying.Permission);

            OfficeImoMcpTools tools = new(fixture.Service);
            CallToolResult result = await tools.PdfPrintPlanAsync(fixture.Path, pages: pages, pagesPerSheet: 2, maxOutputCharacters: 512,
                passwordEnvironmentVariable: fixture.PasswordVariable);

            Assert.False(result.IsError, Text(result));
            Assert.Equal(3, result.StructuredContent!.Value.GetProperty("sourcePageCount").GetInt32());
            Assert.Equal(expectedPages, result.StructuredContent.Value.GetProperty("selectedPages").EnumerateArray().Select(page => page.GetInt32()));
            Assert.Equal(3, result.StructuredContent.Value.GetProperty("selectedPageCount").GetInt32());
            Assert.Equal(2, result.StructuredContent.Value.GetProperty("sheetCount").GetInt32());
            Assert.True(result.StructuredContent.Value.GetRawText().Length <= 512);
            Assert.DoesNotContain(PrintFixture.UserPassword, Text(result));
            Assert.DoesNotContain(PrintFixture.OwnerPassword, Text(result));
            Assert.Equal(fixture.Bytes, await File.ReadAllBytesAsync(fixture.Path));
            Assert.Single(Directory.GetFiles(fixture.Root));
        }

        [Fact]
        public async Task AgentPlansOwnerAuthenticatedDocumentsWithUserPrintingDenied() {
            using PrintFixture fixture = new(PdfStandardPermissions.None, PrintFixture.OwnerPassword);
            AgentPdfPrintPlanResult result = await Plan(fixture);
            Assert.Equal(3, result.SourcePageCount);
            Assert.Equal(new[] { 1, 2, 3 }, result.SelectedPages);
            Assert.Equal(fixture.Bytes, await File.ReadAllBytesAsync(fixture.Path));
        }

        [Theory]
        [InlineData(PdfStandardPermissions.None)]
        [InlineData(PdfStandardPermissions.CopyContents)]
        public async Task AgentEnforcesPrintingPermissionIndependentlyOfCopyPermission(PdfStandardPermissions permissions) {
            using PrintFixture fixture = new(permissions);
            PdfPermissionDeniedException denied = await Assert.ThrowsAsync<PdfPermissionDeniedException>(() => Plan(fixture, "1"));
            Assert.Equal(PdfStandardPermissions.Print, denied.Permission);
            Assert.Equal(PdfPasswordAuthenticationRole.User, denied.AuthenticationRole);
            Assert.Equal(fixture.Bytes, await File.ReadAllBytesAsync(fixture.Path));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("incorrect-synthetic-password")]
        public async Task AgentStillRequiresValidAuthenticationBeforePlanning(string? password) {
            using PrintFixture fixture = new(PdfStandardPermissions.Print, password);
            if (password is null) {
                await Assert.ThrowsAsync<PdfPasswordRequiredException>(() => Plan(fixture, "1", supplyPassword: false));
            } else {
                await Assert.ThrowsAsync<PdfInvalidPasswordException>(() => Plan(fixture, "1"));
            }
            Assert.Equal(fixture.Bytes, await File.ReadAllBytesAsync(fixture.Path));
        }

        [Fact]
        public async Task AgentRejectsMalformedInputBeforeResolvingPrintSelection() {
            using PrintFixture fixture = new(PdfStandardPermissions.Print);
            byte[] malformed = Encoding.ASCII.GetBytes("%PDF-1.7\nnot a PDF object graph\n%%EOF\n");
            await File.WriteAllBytesAsync(fixture.Path, malformed);
            await Assert.ThrowsAsync<PdfParseException>(() => Plan(fixture, "1"));
            Assert.Equal(malformed, await File.ReadAllBytesAsync(fixture.Path));
        }

        [Fact]
        public async Task AgentRetainsTheSelectionLimitIncludingRepeatedPrintablePages() {
            using PrintFixture fixture = new(PdfStandardPermissions.Print, pageCount: 32);
            string pages = string.Join(",", Enumerable.Repeat("1-32", 313));
            await Assert.ThrowsAsync<InvalidOperationException>(() => Plan(fixture, pages));
            Assert.Equal(fixture.Bytes, await File.ReadAllBytesAsync(fixture.Path));
        }

        [Fact]
        public async Task AgentHonorsCancellationBeforeReadingPrintInput() {
            using PrintFixture fixture = new(PdfStandardPermissions.Print);
            using CancellationTokenSource cancellation = new();
            cancellation.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => Plan(fixture, "1", cancellation.Token));
            Assert.Equal(fixture.Bytes, await File.ReadAllBytesAsync(fixture.Path));
        }

        private static Task<AgentPdfPrintPlanResult> Plan(PrintFixture fixture, string? pages = null, CancellationToken cancellationToken = default,
            bool supplyPassword = true) =>
            fixture.Service.PdfPrintPlanAsync(fixture.Path, pages, "A4", "auto", 1, "fit", 18,
                supplyPassword ? fixture.PasswordVariable : null, 512, cancellationToken);

        private static string Text(CallToolResult result) => string.Join(" | ", result.Content.OfType<TextContentBlock>().Select(item => item.Text));

        private sealed class PrintFixture : IDisposable {
            internal const string UserPassword = "synthetic-print-user";
            internal const string OwnerPassword = "synthetic-print-owner";
            internal string Root { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-agent-print-" + Guid.NewGuid().ToString("N"));
            internal string PasswordVariable { get; } = "OFFICEIMO_PRINT_TEST_" + Guid.NewGuid().ToString("N");
            internal string Path { get; }
            internal byte[] Bytes { get; }
            internal OfficeImoAgentService Service { get; }

            internal PrintFixture(PdfStandardPermissions permissions, string? password = UserPassword, int pageCount = 3) {
                Directory.CreateDirectory(Root);
                Path = System.IO.Path.Combine(Root, "protected.pdf");
                Bytes = PdfDocument.Create(compose => {
                    for (int number = 1; number <= pageCount; number++) {
                        int pageNumber = number;
                        compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Printable page " + pageNumber)))));
                    }
                }, new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions(UserPassword) {
                    OwnerPassword = OwnerPassword, AllowedPermissions = permissions
                })).ToBytes();
                File.WriteAllBytes(Path, Bytes);
                Environment.SetEnvironmentVariable(PasswordVariable, password);
                Service = new(new AgentPathPolicy([Root]), pdfPasswordEnvironmentVariables: [PasswordVariable]);
            }

            public void Dispose() {
                Environment.SetEnvironmentVariable(PasswordVariable, null);
                Directory.Delete(Root, recursive: true);
            }
        }
    }
}
