using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public async Task HttpDataReaderPreservesTheBindingConfiguredByteLimit(bool declaredLength, bool inputLimitBinds) {
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
            Content = declaredLength ? new ByteArrayContent(new byte[65]) : new StreamContent(new NonSeekableMemoryStream(new byte[65]))
        }));
        var readOptions = new ExcelReaderOptions { MaxInputBytes = inputLimitBinds ? 64 : 128 };
        var httpOptions = new ExcelLoadOptions { HttpMessageHandler = handler, MaxBytes = inputLimitBinds ? 128 : 64 };
        OfficeInspectionResourceLimitException error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() =>
            ExcelDocument.OpenReaderAsync(new Uri("https://example.test/workbook.xlsx"),WithTransport(readOptions, httpOptions)));
        Assert.Equal("excel", error.Owner);
        Assert.Equal(inputLimitBinds ? nameof(ExcelReaderOptions.MaxInputBytes) : nameof(ExcelLoadOptions.MaxBytes), error.LimitName);
        Assert.Equal(64, error.Maximum);
        Assert.IsType<IOException>(error.InnerException);
    }

    [Fact]
    public async Task HttpByteLimitIdentityUsesItsAdmissionSnapshot() {
        var readOptions = new ExcelReaderOptions { MaxInputBytes = 64 };
        var httpOptions = new ExcelLoadOptions { MaxBytes = 128 };
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
            readOptions.MaxInputBytes = 256;
            httpOptions.MaxBytes = 16;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[65]) });
        });
        httpOptions.HttpMessageHandler = handler;
        OfficeInspectionResourceLimitException error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() =>
            ExcelDocument.OpenReaderAsync(new Uri("https://example.test/workbook.xlsx"),WithTransport(readOptions, httpOptions)));
        Assert.Equal("excel", error.Owner); Assert.Equal(nameof(ExcelReaderOptions.MaxInputBytes), error.LimitName); Assert.Equal(64, error.Maximum);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task HttpDocumentLoadSharesTheConfiguredByteClassification(bool inputLimitBinds) {
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[65]) }));
        OfficeInspectionResourceLimitException error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() => ExcelDocument.LoadAsync(
            new Uri("https://example.test/workbook.xlsx"),            new ExcelLoadOptions { HttpMessageHandler = handler,MaxBytes = inputLimitBinds ? 128 : 64 ,MaxInputBytes = inputLimitBinds ? 64 : 128 }));
        Assert.Equal("excel", error.Owner);
        Assert.Equal(inputLimitBinds ? nameof(ExcelLoadOptions.MaxInputBytes) : nameof(ExcelLoadOptions.MaxBytes), error.LimitName); Assert.Equal(64, error.Maximum);
    }

    [Theory]
    [InlineData(true, 128)]
    [InlineData(false, 128)]
    [InlineData(true, 64)]
    public async Task HttpDocumentLoadKeepsTheBindingPackageSecurityRole(bool declaredLength, long inputMaximum) {
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
            Content = declaredLength ? new ByteArrayContent(new byte[65]) : new StreamContent(new NonSeekableMemoryStream(new byte[65]))
        }));
        OfficePackageSecurityException error = await Assert.ThrowsAsync<OfficePackageSecurityException>(() => ExcelDocument.LoadAsync(
            new Uri("https://example.test/workbook.xlsx"),            new ExcelLoadOptions { HttpMessageHandler = handler,MaxBytes = 128 ,MaxInputBytes = inputMaximum,PackageSecurity = new OfficePackageSecurityOptions { MaxPackageBytes = 64 } }));
        Assert.Equal(OfficePackageSecurityRule.PackageSize, error.Rule); Assert.Equal(64D, error.Limit); Assert.Equal(65D, error.ObservedValue);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task HttpDocumentLoadDoesNotAttributeAStricterByteLimitToPackagePolicy(bool httpLimitBinds) {
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[65]) }));
        OfficeInspectionResourceLimitException error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() => ExcelDocument.LoadAsync(
            new Uri("https://example.test/workbook.xlsx"),            new ExcelLoadOptions { HttpMessageHandler = handler,MaxBytes = httpLimitBinds ? 32 : 128 ,MaxInputBytes = httpLimitBinds ? 128 : 32,PackageSecurity = new OfficePackageSecurityOptions { MaxPackageBytes = 64 } }));
        Assert.Equal("excel", error.Owner); Assert.Equal(httpLimitBinds ? nameof(ExcelLoadOptions.MaxBytes) : nameof(ExcelLoadOptions.MaxInputBytes), error.LimitName); Assert.Equal(32, error.Maximum);
    }

    [Fact]
    public async Task HttpPackageByteRoleUsesItsAdmissionSnapshot() {
        var packageOptions = new OfficePackageSecurityOptions { MaxPackageBytes = 64 };
        var loadOptions = new ExcelLoadOptions { MaxInputBytes = 128, PackageSecurity = packageOptions, MaxBytes = 128 };
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
            loadOptions.MaxInputBytes = 32; packageOptions.MaxPackageBytes = 256; loadOptions.MaxBytes = 16;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[65]) });
        });
        loadOptions.HttpMessageHandler = handler;
        OfficePackageSecurityException error = await Assert.ThrowsAsync<OfficePackageSecurityException>(() =>
            ExcelDocument.LoadAsync(new Uri("https://example.test/workbook.xlsx"), loadOptions));
        Assert.Equal(OfficePackageSecurityRule.PackageSize, error.Rule); Assert.Equal(64D, error.Limit); Assert.Equal(65D, error.ObservedValue);
    }

    [Fact]
    public async Task HttpTransportFailureRetainsItsOriginalIdentity() {
        var expected = new IOException("fixture transport failure");
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromException<HttpResponseMessage>(expected));
        IOException error = await Assert.ThrowsAsync<IOException>(() => ExcelDocument.OpenReaderAsync(new Uri("https://example.test/workbook.xlsx"),new ExcelReaderOptions { Transport = new ExcelLoadOptions { HttpMessageHandler = handler } }));
        Assert.Same(expected, error);
    }

    [Fact]
    public async Task HttpMalformedWorkbookRetainsItsInputClassification() {
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[] { 1, 2, 3 }) }));
        await Assert.ThrowsAsync<InvalidDataException>(() => ExcelDocument.OpenReaderAsync(new Uri("https://example.test/workbook.xlsx"),WithTransport(new ExcelReaderOptions { MaxInputBytes = 64 }, new ExcelLoadOptions { HttpMessageHandler = handler, MaxBytes = 128 })));
    }

    [Fact]
    public async Task HttpExactByteBoundaryLoadsTheWorkbook() {
        byte[] bytes = CreateRemoteWorkbookBytes();
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(CreateWorkbookResponse(bytes)));
        using ExcelWorkbookDataReader reader = await ExcelDocument.OpenReaderAsync(new Uri("https://example.test/workbook.xlsx"),WithTransport(new ExcelReaderOptions { MaxInputBytes = bytes.Length, HasHeaderRow = false }, new ExcelLoadOptions { HttpMessageHandler = handler, MaxBytes = bytes.Length }));
        Assert.True(reader.Read()); Assert.Equal("Value", reader.GetString(0));
    }

    [Fact]
    public async Task HttpCancellationRemainsTerminal() {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        using var handler = new FakeWorkbookHttpMessageHandler((_, token) => Task.FromCanceled<HttpResponseMessage>(token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelDocument.OpenReaderAsync(new Uri("https://example.test/workbook.xlsx"),new ExcelReaderOptions { Transport = new ExcelLoadOptions { HttpMessageHandler = handler } },cancellationToken: cancellation.Token));
    }
}
