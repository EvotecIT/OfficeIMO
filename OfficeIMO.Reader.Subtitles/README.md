# OfficeIMO.Reader.Subtitles

`OfficeIMO.Reader.Subtitles` adds bounded SubRip (`.srt`) and WebVTT (`.vtt`) ingestion to an isolated `OfficeDocumentReader`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Subtitles;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddSubtitleHandler(new ReaderSubtitleOptions {
        MaxCues = 10_000
    })
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("captions.vtt");
foreach (ReaderChunk chunk in result.Chunks) {
    Console.WriteLine($"{chunk.Id} lines {chunk.Location.StartLine}-{chunk.Location.EndLine}: {chunk.Text}");
}
```

Each cue becomes a source-ordered chunk with readable timestamp Markdown, source line locations, and machine-readable start/end milliseconds in metadata. WebVTT `NOTE`, `STYLE`, and `REGION` blocks are skipped; cue markup is stripped by default.

The adapter is local and BCL-only. It does not download media, call YouTube, transcribe audio, run Whisper, or add an audio codec. Applications compose those provider capabilities explicitly outside this package. `OfficeIMO.Reader.All` registers this handler.

## Registration

| Item | Value |
|---|---|
| Handler id | `officeimo.reader.subtitles` |
| Extensions | `.srt`, `.vtt` |
| Reader kind | `ReaderInputKind.Text` |
| Default input limit | 32 MiB when `ReaderOptions.MaxInputBytes` is not set |

A file is parsed as WebVTT when it does not have the `.srt` extension and its first line carries the `WEBVTT` signature; otherwise it is parsed as SubRip. Text is decoded as UTF-8, honoring a byte-order mark when present.

## Options

| `ReaderSubtitleOptions` | Default | Range | Effect |
|---|---|---|---|
| `IncludeTimestampsInMarkdown` | `true` | | Prefixes each cue's Markdown with `**hh:mm:ss.fff → hh:mm:ss.fff**`. `Text` never contains the timestamp. |
| `StripCueMarkup` | `true` | | Removes WebVTT/SRT cue tags (`b`, `i`, `u`, `c`, `v`, `lang`, `ruby`, `rt`, `span`, `font` and inline timestamps), turns `<br>` into a line break, and decodes HTML entities. Unrecognized tags are kept. |
| `MaxCues` | `50000` | 1 to 1,000,000 | Parsing stops after this many cues. |
| `MaxCueCharacters` | `32000` | 1 to 1,000,000 | Longer cue text is truncated. |

Out-of-range values throw `ArgumentOutOfRangeException` when the handler is registered. `ReaderOptions.MaxChars` (default `8000`) splits an unusually long cue into continuation chunks.

## Output

- **Chunks:** one per cue, with `Id` `subtitle-cue-000000`, `subtitle-cue-000001`, and so on. A cue split by `MaxChars` continues as `subtitle-cue-000000-part-0001` with `ContinuesPreviousChunk = true`. `Location` carries the cue's `StartLine`, `EndLine`, `SourceBlockIndex` and `SourceBlockKind = "subtitle-cue"`. Markdown special characters `&`, `<` and `>` are escaped in the Markdown projection.
- **Document metadata:** `subtitle-format` (`srt` or `webvtt`) and `subtitle-cue-count`.
- **Cue metadata:** one `subtitle.cue` entry per cue with `Value` `hh:mm:ss.fff --> hh:mm:ss.fff`, `SourceObjectId` set to the cue identifier when present, and `Attributes["startMilliseconds"]` / `Attributes["endMilliseconds"]`.

## Limits and diagnostics

| Condition | Result |
|---|---|
| Block with an invalid timing line | Skipped; recoverable `subtitle-invalid-block` diagnostic naming the line. |
| More cues than `MaxCues` | Parsing stops; recoverable `subtitle-cue-limit` diagnostic in the `Limit` category. |
| Cue longer than `MaxCueCharacters` | Cue truncated; chunk warning `Subtitle cue text was truncated at MaxCueCharacters.` |
| Cue projection longer than `MaxChars` | Cue split; chunk warning `Subtitle cue projection was split due to MaxChars.` |
| Cue that is empty after markup stripping | Omitted. |

At most 100 parser diagnostics are reported per file.

## Example output

```text
WEBVTT

1
00:00:01.000 --> 00:00:03.500
<v Anna>Welcome to the <b>quarterly</b> review.
```

produces one chunk:

| Property | Value |
|---|---|
| `Id` | `subtitle-cue-000000` |
| `Text` | `Welcome to the quarterly review.` |
| `Markdown` | `**00:00:01.000 → 00:00:03.500**` followed by a blank line and the cue text |
| `Location` | lines 3-5, `subtitle-cue` |
| Cue metadata | `00:00:01.000 --> 00:00:03.500`, `startMilliseconds = 1000`, `endMilliseconds = 3500`, `SourceObjectId = 1` |

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Reader.Core`.
