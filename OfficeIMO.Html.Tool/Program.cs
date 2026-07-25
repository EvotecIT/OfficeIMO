using OfficeIMO.Html.Tool;

using var cancellation = new CancellationTokenSource();
Console.CancelKeyPress += (_, eventArgs) => {
    eventArgs.Cancel = true;
    cancellation.Cancel();
};

return await HtmlPdfToolApp.RunAsync(
    args,
    Console.OpenStandardInput(),
    Console.OpenStandardOutput(),
    Console.Error,
    cancellation.Token).ConfigureAwait(false);
