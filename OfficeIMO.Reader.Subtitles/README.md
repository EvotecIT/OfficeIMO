# OfficeIMO.Reader.Subtitles

`OfficeIMO.Reader.Subtitles` adds bounded SubRip (`.srt`) and WebVTT (`.vtt`) ingestion to an isolated `OfficeDocumentReader`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Subtitles;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddSubtitleHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("captions.vtt");
```

Each cue becomes a source-ordered chunk with readable timestamp Markdown, source line locations, and machine-readable start/end milliseconds in metadata. WebVTT notes, styles, and regions are skipped; cue markup is stripped by default.

The adapter is local and BCL-only. It does not download media, call YouTube, transcribe audio, run Whisper, or add an audio codec. Those remain explicit later/provider layers.
