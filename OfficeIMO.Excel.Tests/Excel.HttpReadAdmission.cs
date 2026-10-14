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
        var readOptions = new ExcelReadOptions { MaxInputBytes = inputLimitBinds ? 64 : 128 };
        var httpOptions = new ExcelHttpLoadOptions { HttpMessageHandler = handler, MaxBytes = inputLimitBinds ? 128 : 64 };
        OfficeInspectionResourceLimitException error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() =>
            ExcelDocument.OpenDataReaderAsync(new Uri("https://example.test/workbook.xlsx"), readOptions, httpOptions));
        Assert.Equal("excel", error.Owner);
        Assert.Equal(inputLimitBinds ? nameof(ExcelReadOptions.MaxInputBytes) : nameof(ExcelHttpLoadOptions.MaxBytes), error.LimitName);
        Assert.Equal(64, error.Maximum);
        Assert.IsType<IOException>(error.InnerException);
    }

    [Fact]
    public async Task HttpByteLimitIdentityUsesItsAdmissionSnapshot() {
        var readOptions = new ExcelReadOptions { MaxInputBytes = 64 };
        var httpOptions = new ExcelHttpLoadOptions { MaxBytes = 128 };
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
            readOptions.MaxInputBytes = 256;
            httpOptions.MaxBytes = 16;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[65]) });
        });
        httpOptions.HttpMessageHandler = handler;
        OfficeInspectionResourceLimitException error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() =>
            ExcelDocument.OpenDataReaderAsync(new Uri("https://example.test/workbook.xlsx"), readOptions, httpOptions));
        Assert.Equal("excel", error.Owner); Assert.Equal(nameof(ExcelReadOptions.MaxInputBytes), error.LimitName); Assert.Equal(64, error.Maximum);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task HttpDocumentLoadSharesTheConfiguredByteClassification(bool inputLimitBinds) {
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[65]) }));
        OfficeInspectionResourceLimitException error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() => ExcelDocument.LoadAsync(
            new Uri("https://example.test/workbook.xlsx"),
            new ExcelHttpLoadOptions { HttpMessageHandler = handler, MaxBytes = inputLimitBinds ? 128 : 64 },
            new ExcelLoadOptions { MaxInputBytes = inputLimitBinds ? 64 : 128 }));
        Assert.Equal("excel", error.Owner);
        Assert.Equal(inputLimitBinds ? nameof(ExcelLoadOptions.MaxInputBytes) : nameof(ExcelHttpLoadOptions.MaxBytes), error.LimitName); Assert.Equal(64, error.Maximum);
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
            new Uri("https://example.test/workbook.xlsx"),
            new ExcelHttpLoadOptions { HttpMessageHandler = handler, MaxBytes = 128 },
            new ExcelLoadOptions { MaxInputBytes = inputMaximum, PackageSecurity = new OfficePackageSecurityOptions { MaxPackageBytes = 64 } }));
        Assert.Equal(OfficePackageSecurityRule.PackageSize, error.Rule); Assert.Equal(64D, error.Limit); Assert.Equal(65D, error.ObservedValue);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task HttpDocumentLoadDoesNotAttributeAStricterByteLimitToPackagePolicy(bool httpLimitBinds) {
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[65]) }));
        OfficeInspectionResourceLimitException error = await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() => ExcelDocument.LoadAsync(
            new Uri("https://example.test/workbook.xlsx"),
            new ExcelHttpLoadOptions { HttpMessageHandler = handler, MaxBytes = httpLimitBinds ? 32 : 128 },
            new ExcelLoadOptions { MaxInputBytes = httpLimitBinds ? 128 : 32, PackageSecurity = new OfficePackageSecurityOptions { MaxPackageBytes = 64 } }));
        Assert.Equal("excel", error.Owner); Assert.Equal(httpLimitBinds ? nameof(ExcelHttpLoadOptions.MaxBytes) : nameof(ExcelLoadOptions.MaxInputBytes), error.LimitName); Assert.Equal(32, error.Maximum);
    }

    [Fact]
    public async Task HttpPackageByteRoleUsesItsAdmissionSnapshot() {
        var packageOptions = new OfficePackageSecurityOptions { MaxPackageBytes = 64 };
        var loadOptions = new ExcelLoadOptions { MaxInputBytes = 128, PackageSecurity = packageOptions };
        var httpOptions = new ExcelHttpLoadOptions { MaxBytes = 128 };
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
            loadOptions.MaxInputBytes = 32; packageOptions.MaxPackageBytes = 256; httpOptions.MaxBytes = 16;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[65]) });
        });
        httpOptions.HttpMessageHandler = handler;
        OfficePackageSecurityException error = await Assert.ThrowsAsync<OfficePackageSecurityException>(() =>
            ExcelDocument.LoadAsync(new Uri("https://example.test/workbook.xlsx"), httpOptions, loadOptions));
        Assert.Equal(OfficePackageSecurityRule.PackageSize, error.Rule); Assert.Equal(64D, error.Limit); Assert.Equal(65D, error.ObservedValue);
    }

    [Fact]
    public async Task HttpTransportFailureRetainsItsOriginalIdentity() {
        var expected = new IOException("fixture transport failure");
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromException<HttpResponseMessage>(expected));
        IOException error = await Assert.ThrowsAsync<IOException>(() => ExcelDocument.OpenDataReaderAsync(new Uri("https://example.test/workbook.xlsx"),
            httpOptions: new ExcelHttpLoadOptions { HttpMessageHandler = handler }));
        Assert.Same(expected, error);
    }

    [Fact]
    public async Task HttpMalformedWorkbookRetainsItsInputClassification() {
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[] { 1, 2, 3 }) }));
        await Assert.ThrowsAsync<InvalidDataException>(() => ExcelDocument.OpenDataReaderAsync(new Uri("https://example.test/workbook.xlsx"),
            new ExcelReadOptions { MaxInputBytes = 64 }, new ExcelHttpLoadOptions { HttpMessageHandler = handler, MaxBytes = 128 }));
    }

    [Fact]
    public async Task HttpExactByteBoundaryLoadsTheWorkbook() {
        byte[] bytes = CreateRemoteWorkbookBytes();
        using var handler = new FakeWorkbookHttpMessageHandler((_, _) => Task.FromResult(CreateWorkbookResponse(bytes)));
        using ExcelWorkbookDataReader reader = await ExcelDocument.OpenDataReaderAsync(new Uri("https://example.test/workbook.xlsx"),
            new ExcelReadOptions { MaxInputBytes = bytes.Length, HasHeaderRow = false }, new ExcelHttpLoadOptions { HttpMessageHandler = handler, MaxBytes = bytes.Length });
        Assert.True(reader.Read()); Assert.Equal("Value", reader.GetString(0));
    }

    [Fact]
    public async Task HttpCancellationRemainsTerminal() {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        using var handler = new FakeWorkbookHttpMessageHandler((_, token) => Task.FromCanceled<HttpResponseMessage>(token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelDocument.OpenDataReaderAsync(new Uri("https://example.test/workbook.xlsx"),
            httpOptions: new ExcelHttpLoadOptions { HttpMessageHandler = handler }, cancellationToken: cancellation.Token));
    }
}
