using System.Text;
using System.Text.Json;
using OfficeIMO.AI;
using OfficeIMO.AI.IntelligenceX;

try {
    ExampleOptions options = ExampleOptions.Parse(args);
    if (options.Help) { Console.WriteLine(ExampleOptions.Usage); return 0; }
    if (!options.Local && !options.AllowRemote) { Console.Error.WriteLine("Hosted processing requires --allow-remote. Selected source text and, with --images, page images will be sent to the configured model."); return 2; }
    if (options.Local && options.Endpoint is null) throw new ArgumentException("Local execution requires an explicit endpoint.");
    using var lifetime = new CancellationTokenSource();
    Console.CancelKeyPress += (_, e) => { e.Cancel = true; lifetime.Cancel(); };
    if (options.Evaluate) return await EvaluationRunner.RunAsync(options, lifetime.Token);
    OfficeAiRequest request = options.RequestPath is null ? ExampleOptions.InvoiceRequest : JsonSerializer.Deserialize<OfficeAiRequest>(
        await DocumentInputs.ReadFileAsync(options.RequestPath, 64_000, lifetime.Token), ExampleOptions.RequestJsonOptions)
        ?? throw new InvalidDataException("Empty request configuration.");
    request = request with { AllowRemoteProcessing = options.AllowRemote, IncludeImages = options.Images };
    lifetime.CancelAfter(request.Limits.Timeout);
    byte[] bytes = options.SourcePath is null ? Encoding.UTF8.GetBytes(ExampleOptions.SyntheticInvoice)
        : await DocumentInputs.ReadFileAsync(options.SourcePath, request.Limits.MaxInputBytes, lifetime.Token);
    OfficeAiDocument document = await DocumentInputs.ReadAsync(bytes, options.SourcePath ?? "example.txt", options.Images,
        request.Pages, request.Limits, lifetime.Token);
    using var executor = await ExampleExecution.ConnectAsync(options, options.Images, lifetime.Token);
    var progress = new Progress<OfficeAiProgress>(value => Console.Error.WriteLine($"{value.Stage}: {value.CompletedBatches}/{value.TotalBatches}"));
    OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(document, request, progress, lifetime.Token);
    Console.WriteLine(OfficeAiArtifacts.SerializeReport(document, result));
    if (options.OutputPath is not null) ArtifactWriter.Save(options.OutputPath, document, result);
    return result.Status == OfficeAiResultStatus.Completed ? 0 : 1;
} catch (OperationCanceledException) {
    Console.Error.WriteLine("Operation cancelled or its time budget expired."); return 3;
} catch (Exception exception) {
    Console.Error.WriteLine("Operation failed: " + exception.GetType().Name + ". Check the documented input and profile requirements."); return 2;
}
