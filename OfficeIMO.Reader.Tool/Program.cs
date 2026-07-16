using OfficeIMO.Reader.Tool;

using var cancellation = new CancellationTokenSource();
Console.CancelKeyPress += (_, eventArgs) => {
    eventArgs.Cancel = true;
    cancellation.Cancel();
};

return await ReaderToolApp.RunAsync(
    args,
    Console.OpenStandardInput(),
    Console.Out,
    Console.Error,
    cancellation.Token).ConfigureAwait(false);