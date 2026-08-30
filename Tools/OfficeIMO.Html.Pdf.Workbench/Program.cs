using HtmlTinkerX;
using OfficeIMO.Html.Pdf.Workbench;
using OfficeIMO.Html.Pdf.Workbench.Components;

var builder = WebApplication.CreateBuilder(args);
string listenUrl = builder.Configuration["Workbench:Url"] ?? "http://127.0.0.1:5105";
if (!Uri.TryCreate(listenUrl, UriKind.Absolute, out Uri? listenUri)
    || !listenUri.IsLoopback
    || (listenUri.Scheme != Uri.UriSchemeHttp && listenUri.Scheme != Uri.UriSchemeHttps)
    || !string.IsNullOrEmpty(listenUri.UserInfo)
    || listenUri.AbsolutePath != "/") {
    throw new InvalidOperationException("The HTML-to-PDF workbench must bind to a loopback URL.");
}
builder.WebHost.UseUrls(listenUrl);
builder.Configuration["AllowedHosts"] = listenUri.Host;

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();
builder.Services.AddSingleton<HtmlBrowserPdfRenderer>(_ => new HtmlBrowserPdfRenderer(new HtmlBrowserPdfRendererOptions(
        maximumBrowserInstances: 1,
        maximumQueuedCaptures: 4,
        viewportWidth: 1440,
        viewportHeight: 900,
        networkPolicy: HtmlBrowserNetworkPolicy.Offline,
        setupTimeout: TimeSpan.FromSeconds(45))));
builder.Services.AddSingleton<HtmlPdfWorkbenchConversionService>();
builder.Services.AddSingleton<WorkbenchArtifactStore>();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}

app.UseHostFiltering();
app.UseWebSockets();
app.Use(async (context, next) => {
    if (!WorkbenchRequestBoundary.IsAllowedHost(context.Request.Host, listenUri)) {
        context.Response.StatusCode = StatusCodes.Status400BadRequest;
        return;
    }
    if (context.WebSockets.IsWebSocketRequest
        && !WorkbenchRequestBoundary.IsAllowedWebSocketOrigin(context.Request.Headers.Origin, listenUri)) {
        context.Response.StatusCode = StatusCodes.Status403Forbidden;
        return;
    }
    await next();
});
app.UseStaticFiles();
app.UseAntiforgery();

app.MapGet("/workbench/artifacts/{token}/pdf", (HttpContext context, string token, WorkbenchArtifactStore store) => {
    SetArtifactHeaders(context);
    return store.TryGet(token, out WorkbenchArtifact? artifact)
        ? Results.Bytes(artifact!.PdfBytes, "application/pdf")
        : Results.NotFound();
});
app.MapGet("/workbench/artifacts/{token}/evidence", (HttpContext context, string token, WorkbenchArtifactStore store) => {
    SetArtifactHeaders(context);
    return store.TryGet(token, out WorkbenchArtifact? artifact)
        ? Results.Bytes(artifact!.EvidenceBytes, "application/json; charset=utf-8")
        : Results.NotFound();
});

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();

static void SetArtifactHeaders(HttpContext context) {
    context.Response.Headers.CacheControl = "no-store, max-age=0";
    context.Response.Headers.Append("Pragma", "no-cache");
    context.Response.Headers.Append("X-Content-Type-Options", "nosniff");
    context.Response.Headers.Append("Referrer-Policy", "no-referrer");
}
