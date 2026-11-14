using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.PdfCorpusReliability;

/// <summary>
/// Child-process loop. Reads one JSON request per stdin line and answers with protocol lines on stdout,
/// so the coordinator can kill and replace this process when a file hangs, crashes or exhausts memory.
/// </summary>
internal static class ReliabilityWorker {
    public const string ProgressPrefix = "@@op ";
    public const string ResultPrefix = "@@result ";
    private static readonly byte[] PngSignature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    public static int Run() {
        var output = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
        // Keep library console noise out of the protocol stream.
        Console.SetOut(TextWriter.Null);
        string? line;
        while ((line = Console.In.ReadLine()) != null) {
            if (line.Length == 0) continue;
            WorkerRequest request = JsonSerializer.Deserialize<WorkerRequest>(line, ReliabilityJson.Line)
                ?? throw new InvalidDataException("Empty worker request.");
            bool exhausted = false;
            WorkerResponse response = Process(request, name => output.WriteLine(ProgressPrefix + name), ref exhausted);
            output.WriteLine(ResultPrefix + JsonSerializer.Serialize(response, ReliabilityJson.Line));
            if (exhausted) return 3;
            GC.Collect();
        }
        return 0;
    }

    internal static WorkerResponse Process(WorkerRequest request, Action<string> progress, ref bool exhausted) {
        var response = new WorkerResponse();
        byte[] bytes = File.ReadAllBytes(request.FullPath);
        response.Bytes = bytes.LongLength;
        response.Sha256 = Convert.ToHexString(SHA256.HashData(bytes));

        PdfDocument? document = null;
        string? password = null;
        string? sourceText = null;
        byte[]? saved = null;
        bool memoryExhausted = false;

        void Run(string name, Action action) {
            if (memoryExhausted) {
                response.Operations.Add(new OperationResult { Name = name, Outcome = Outcome.Skipped, Detail = "memory exhausted earlier" });
                return;
            }
            progress(name);
            var stopwatch = Stopwatch.StartNew();
            var result = new OperationResult { Name = name, Outcome = Outcome.Ok };
            try {
                action();
            } catch (MismatchException mismatch) {
                result.Outcome = Outcome.Mismatch;
                result.Detail = mismatch.Message;
            } catch (Exception exception) {
                Exception root = Unwrap(exception);
                result.Outcome = Classify(root);
                result.ExceptionType = root.GetType().FullName;
                result.Message = Truncate(root.Message, 300);
                result.Frame = FirstLibraryFrame(root);
                if (result.Outcome == Outcome.Memory) memoryExhausted = true;
            }
            result.Milliseconds = stopwatch.ElapsedMilliseconds;
            response.Operations.Add(result);
        }

        void Skip(string name, string reason) =>
            response.Operations.Add(new OperationResult { Name = name, Outcome = Outcome.Skipped, Detail = reason });

        Run(Operations.Load, () => {
            document = LoadWithCandidates(bytes, request.Passwords, out password);
            response.UsedPassword = password != null;
        });
        if (document is null) {
            foreach (string name in Operations.All.Skip(1)) Skip(name, "load failed");
            exhausted = memoryExhausted;
            return response;
        }

        PdfDocument loaded = document;
        Run(Operations.PageCount, () => response.PageCount = loaded.Pages.Count);
        Run(Operations.Text, () => {
            sourceText = loaded.ExtractText();
            response.TextLength = sourceText.Length;
        });
        Run(Operations.Read, () => {
            PdfDocumentReadResult read = loaded.Read();
            response.Paragraphs = read.Paragraphs.Count;
            response.Tables = read.Tables.Count;
        });
        if (response.PageCount is int pageCount && pageCount > 0) {
            Run(Operations.Render, () => RenderSample(loaded, pageCount, request, response));
        } else {
            Skip(Operations.Render, "no page count");
        }
        Run(Operations.Save, () => {
            using var stream = new MemoryStream();
            loaded.Save(stream);
            saved = stream.ToArray();
            response.SavedBytes = saved.LongLength;
        });
        if (saved is null) {
            Skip(Operations.RoundTrip, "save failed");
        } else {
            byte[] savedBytes = saved;
            Run(Operations.RoundTrip, () => CompareRoundTrip(savedBytes, password, response.PageCount, sourceText));
        }
        // Saving an unmodified document passes the source bytes through, so the writer is exercised separately by
        // extracting every page into a new document and comparing that output.
        if (response.PageCount is int total && total > 0) {
            Run(Operations.Rewrite, () => {
                PdfDocument rewritten = loaded.Pages.Extract("1-" + total.ToString(System.Globalization.CultureInfo.InvariantCulture));
                byte[] output = rewritten.ToBytes();
                response.RewrittenBytes = output.LongLength;
                CompareRoundTrip(output, password, total, sourceText);
            });
        } else {
            Skip(Operations.Rewrite, "no page count");
        }
        Run(Operations.Repair, () => {
            PdfRepairReport repair = loaded.RepairReport;
            response.RepairCount = repair.RepairCount;
            response.DetectedOnlyCount = repair.DetectionOnlyCount;
            response.RepairCodes = repair.Diagnostics
                .Select(static diagnostic => diagnostic.Code)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(static code => code, StringComparer.Ordinal)
                .Take(20)
                .ToArray();
        });
        exhausted = memoryExhausted;
        return response;
    }

    private static PdfDocument LoadWithCandidates(byte[] bytes, string[]? passwords, out string? usedPassword) {
        usedPassword = null;
        // PdfDocument.Load defers parsing; touching the page tree makes load include the first structural parse
        // and the password check.
        if (passwords is null || passwords.Length == 0) {
            PdfDocument plain = PdfDocument.Load(new MemoryStream(bytes, writable: false));
            _ = plain.Pages.Count;
            return plain;
        }
        Exception? last = null;
        foreach (string candidate in passwords) {
            try {
                PdfDocument document = PdfDocument.Load(new MemoryStream(bytes, writable: false), new PdfLoadOptions { Password = candidate });
                _ = document.Pages.Count;
                usedPassword = candidate;
                return document;
            } catch (Exception exception) when (exception is PdfInvalidPasswordException or PdfPasswordRequiredException) {
                last = exception;
            }
        }
        throw last ?? new PdfInvalidPasswordException("No published password candidate opened the document.");
    }

    private static void RenderSample(PdfDocument document, int pageCount, WorkerRequest request, WorkerResponse response) {
        int[] pages = SamplePages(pageCount, request.RenderPages);
        var options = new PdfPageRenderOptions { Dpi = 72, Format = OfficeImageExportFormat.Png, ContinueOnError = false };
        Exception? firstFailure = null;
        foreach (int page in pages) {
            response.RenderAttemptedPages++;
            try {
                PdfPageRenderResult result = document.Render.Pages(page.ToString(System.Globalization.CultureInfo.InvariantCulture), options)[0];
                byte[]? png = result.Bytes;
                if (!result.Succeeded || png is null || png.Length < PngSignature.Length || !png.AsSpan(0, PngSignature.Length).SequenceEqual(PngSignature)) {
                    throw new MismatchException($"page {page} produced no PNG ({string.Join("; ", result.Diagnostics.Take(2))})");
                }
                response.RenderedPages++;
                if (request.PngDirectory is not null) {
                    Directory.CreateDirectory(request.PngDirectory);
                    File.WriteAllBytes(System.IO.Path.Combine(request.PngDirectory, $"{response.Sha256}-p{page}.png"), png);
                }
            } catch (Exception exception) when (exception is not OutOfMemoryException) {
                firstFailure ??= exception;
            }
        }
        if (firstFailure is not null) {
            if (firstFailure is MismatchException) throw firstFailure;
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(firstFailure).Throw();
        }
    }

    internal static int[] SamplePages(int pageCount, int sample) {
        if (sample <= 1 || pageCount == 1) return new[] { 1 };
        var pages = new SortedSet<int> { 1, pageCount };
        if (sample >= 3) pages.Add((pageCount + 1) / 2);
        return pages.Take(sample).ToArray();
    }

    private static void CompareRoundTrip(byte[] saved, string? password, int? sourcePages, string? sourceText) {
        var options = password is null ? null : new PdfLoadOptions { Password = password };
        PdfDocument reloaded = PdfDocument.Load(new MemoryStream(saved, writable: false), options);
        int pages = reloaded.Pages.Count;
        if (sourcePages is int expected && pages != expected) {
            throw new MismatchException($"page count {expected} -> {pages}");
        }
        if (sourceText is null) return;
        string text = reloaded.ExtractText();
        string before = NormalizeText(sourceText);
        string after = NormalizeText(text);
        if (!string.Equals(before, after, StringComparison.Ordinal)) {
            throw new MismatchException($"text changed: {before.Length} -> {after.Length} normalized chars");
        }
    }

    internal static string NormalizeText(string text) {
        var builder = new StringBuilder(text.Length);
        bool space = false;
        foreach (char ch in text) {
            if (char.IsWhiteSpace(ch)) {
                space = builder.Length > 0;
                continue;
            }
            if (space) builder.Append(' ');
            space = false;
            builder.Append(ch);
        }
        return builder.ToString();
    }

    private static Exception Unwrap(Exception exception) {
        while (exception is AggregateException { InnerExceptions.Count: 1 } aggregate) exception = aggregate.InnerExceptions[0];
        while (exception is System.Reflection.TargetInvocationException { InnerException: { } inner }) exception = inner;
        return exception;
    }

    /// <summary>OfficeIMO's own exception types and standard bad-data signals count as reasoned rejections.</summary>
    internal static string Classify(Exception exception) {
        if (exception is OutOfMemoryException or InsufficientMemoryException) return Outcome.Memory;
        if (exception is PdfMutationBlockedException) return Outcome.Blocked;
        string? ns = exception.GetType().Namespace;
        if (ns is not null && ns.StartsWith("OfficeIMO", StringComparison.Ordinal)) return Outcome.Rejected;
        if (exception is InvalidDataException or NotSupportedException or System.Security.Cryptography.CryptographicException) return Outcome.Rejected;
        return Outcome.Unexpected;
    }

    private static string? FirstLibraryFrame(Exception exception) {
        StackFrame[] frames = new StackTrace(exception, fNeedFileInfo: false).GetFrames();
        foreach (StackFrame frame in frames) {
            System.Reflection.MethodBase? method = frame.GetMethod();
            Type? type = method?.DeclaringType;
            string? ns = type?.Namespace;
            if (ns is null || !ns.StartsWith("OfficeIMO", StringComparison.Ordinal) || ns.StartsWith("OfficeIMO.PdfCorpusReliability", StringComparison.Ordinal)) continue;
            string typeName = type!.FullName ?? type.Name;
            int generic = typeName.IndexOf('+');
            if (generic > 0 && typeName.Contains("<", StringComparison.Ordinal)) typeName = typeName[..generic];
            return typeName + "." + method!.Name;
        }
        return null;
    }

    private static string Truncate(string value, int length) => value.Length <= length ? value : value[..length] + "...";

    private sealed class MismatchException : Exception {
        public MismatchException(string message) : base(message) { }
    }
}
