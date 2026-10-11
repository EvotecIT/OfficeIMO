# Format-map videos and social exports

The format-map recorder turns OfficeIMO's generated conversion catalog into animated clips and PNG stills. It includes landscape, square, 4:5 and vertical compositions with format spotlights, conversion paths, surface badges and an OfficeIMO outro. Clips are silent: audio, narration and subtitle tracks are not generated.

This is repository tooling for producing showcase media. It is separate from document conversion and from the interactive website map. Generated videos are local artifacts, not files installed with an OfficeIMO package.

## Record a clip

Run these commands from the repository root. Use the .NET SDK selected by `global.json`, PowerShell and the shared HtmlTinkerX browser. Video encoding also needs ffmpeg on `PATH`, including the selected encoders and the PNG encoder used for posters. PNG still capture does not need ffmpeg.

```powershell
# Inspect the standard compositions without recording
./Build/Record-FormatMap.ps1 -List

# Refresh the catalog and render the standard cuts
./Build/Record-FormatMap.ps1

# Reuse the catalog while changing presentation or timing
./Build/Record-FormatMap.ps1 -Cut media-9x16 -Formats mp4 -Fps 30 -SkipCatalog

# Render the landscape composition at true 3840 x 2160
./Build/Record-FormatMap.ps1 -Cut media-16x9 -Scale 2 -Formats mp4 -SkipCatalog

# Capture exact moments, including the opening frame, as PNGs
./Build/Record-FormatMap.ps1 -Cut media-16x9 -Still 0,2000,6800 -SkipCatalog
```

The default output directory is `Artefacts/FormatMapVideos`. Set `-Output` to choose another directory. Each video cut produces a clip and a PNG poster with the same base name. Still filenames include their timestamp, for example `media-16x9-t2000.png`.

The standard `cuts.json` includes MP4 for all four compositions and GIF for the square composition. `-Formats` replaces a cut's format selection; it does not add to it.

## Export for X/Twitter, LinkedIn and YouTube

[Social presets](../Build/FormatMapMedia/social-cuts.json) provide five named, silent H.264 MP4 exports at 30 fps, including both 4K and 1080p YouTube landscape files. Each composition is rendered at its native dimensions; vertical output is not cropped from the landscape clip.

| Preset | Intended placement | Resolution | Aspect ratio |
| --- | --- | --- | --- |
| `x-twitter-feed` | X/Twitter feed | 1080 x 1080 | 1:1 |
| `linkedin-feed` | LinkedIn organic feed | 1080 x 1350 | 4:5 |
| `youtube-landscape-4k` | YouTube standard video: preferred upload | 3840 x 2160 | 16:9 |
| `youtube-landscape` | YouTube standard video: smaller 1080p alternative | 1920 x 1080 | 16:9 |
| `youtube-shorts` | YouTube Shorts | 1080 x 1920 | 9:16 |

Use the 4K landscape file for a standard YouTube upload when preserving the diagram's fine lines and small labels matters. Its recipe sets `scale: 2`, so the page draws the same landscape composition at twice the pixel dimensions; it does not enlarge an existing 1080p video. Keep the 1080p landscape file as a smaller alternative. Both play the same scenes at the same speed.

For X use the square export, for LinkedIn use the 4:5 export, and for Shorts use the 1080 x 1920 vertical export. These are editorial choices within the documented upload formats, not a promise that one size performs best for every audience. Do not apply `-Scale 2` to the entire social recipe: the 4K preset already includes its scale, and the other presets retain their platform dimensions.

```powershell
# List the platform presets
./Build/Record-FormatMap.ps1 -CutsFile Build/FormatMapMedia/social-cuts.json -List

# Refresh the catalog and render all five files and their posters
./Build/Record-FormatMap.ps1 -CutsFile Build/FormatMapMedia/social-cuts.json -Output Artefacts/FormatMapSocial

# Render only the preferred YouTube upload (already configured for native 4K)
./Build/Record-FormatMap.ps1 -CutsFile Build/FormatMapMedia/social-cuts.json -Cut youtube-landscape-4k -Output Artefacts/FormatMapSocial -SkipCatalog

# Re-render one preset after a design change
./Build/Record-FormatMap.ps1 -CutsFile Build/FormatMapMedia/social-cuts.json -Cut linkedin-feed -Output Artefacts/FormatMapSocial -SkipCatalog
```

Upload the resulting `.mp4` file, rather than a ZIP or poster. The PNG is available for a thumbnail or a companion image; platform controls determine whether a custom thumbnail is accepted. The commands prepare local files and do not upload or publish them.

Platform rules change and vary by placement or account. Check the linked sources before a campaign, especially when changing duration, resolution or frame rate:

- [X video uploads](https://help.x.com/en/using-x/x-videos): non-Premium uploads are limited to 140 seconds and 512 MB. Keep the short feed preset below both limits.
- [LinkedIn Page video guidelines](https://www.linkedin.com/help/linkedin/answer/a1311816): MP4/H.264, 10-60 fps and aspect ratios from 1:2.4 to 2.4:1 are supported. [LinkedIn advertising specifications](https://business.linkedin.com/advertise/ads/sponsored-content/video-ads/specs) also describe 1080 x 1350 at 4:5, but ads have their own rules; the supplied preset targets organic posts.
- [YouTube 4K guidance](https://support.google.com/youtube/answer/16597617?hl=en) supports recording and exporting at 3840 x 2160. [YouTube upload encoding](https://support.google.com/youtube/answer/1722171): MP4, H.264, progressive video, 4:2:0 and fast-start metadata are supported upload settings. The recorder uses quality-based variable bitrate, not a fixed platform bitrate target. Allow [high-resolution processing](https://support.google.com/youtube/answer/71674?hl=en) to finish before assessing the uploaded video's quality.
- [YouTube Shorts classification](https://support.google.com/youtube/answer/15424877): square or vertical videos up to three minutes are classified as Shorts under the current rules. Use the landscape preset for a standard watch-page video.
- [YouTube Shorts creation guidance](https://support.google.com/youtube/answer/10059070?hl=en-GB) specifies a maximum upload resolution of 1080p for that workflow; the Shorts preset stays at 1080 x 1920. LinkedIn's Page guidelines list a maximum of 4096 x 2304, so doubling the 4:5 preset to 2160 x 2700 would exceed the listed height.

## Output formats and options

| `-Formats` value | Encoding | File extension |
| --- | --- | --- |
| `mp4` | H.264, YUV 4:2:0, fast start | `.mp4` |
| `h265` | HEVC/H.265 in MP4 | `.h265.mp4` |
| `av1` | AV1 in MP4 | `.av1.mp4` |
| `webm` | VP9 | `.webm` |
| `mov` | ProRes 422 HQ | `.mov` |
| `gif` | Palette-based animated GIF | `.gif` |
| `webp` | Animated WebP | `.webp` |
| `apng` | Animated PNG | `.apng` |

```powershell
./Build/Record-FormatMap.ps1 -Cut media-1x1 -Formats mp4,gif,webp -SkipCatalog
```

`-Cut` accepts one or more cut names. `-CutsFile` selects the recipe file. `-Fps` overrides video capture at 10-120 fps; the default is 30. `-Scale` accepts 1-4 and multiplies pixel dimensions while retaining the composition. Higher resolutions and frame rates take more time and storage. GIF, animated WebP and APNG have separate `gifWidth` and `gifFps` recipe settings, shared by those three animated image formats.

Use the .NET entry point for custom media directories, an explicit ffmpeg path or fractional still timestamps:

```powershell
dotnet run --project Build/FormatMapRecorder -c Release -- --cuts=Build/FormatMapMedia/social-cuts.json --cut=x-twitter-feed --still=0,100.1,100.9
```

The CLI also accepts `--output`, `--scale`, `--fps`, `--formats`, `--media-dir`, `--ffmpeg` and `--list`. Value options use `--name=value`. Unknown, repeated or empty options fail rather than silently changing the requested render.

## Change the story as capabilities grow

[The standard cuts](../Build/FormatMapMedia/cuts.json) and [social cuts](../Build/FormatMapMedia/social-cuts.json) own scene selection and export settings. Edit a maintained preset or copy a cut into another JSON file and pass it with `-CutsFile`. Required fields are `name`, `ratio`, `width`, `height` and `formats`.

```json
{
  "cuts": [
    {
      "name": "spreadsheet-showcase",
      "ratio": "4x5",
      "width": 1080,
      "height": 1350,
      "formats": ["mp4"],
      "fps": 30,
      "speed": 1.1,
      "scenes": "intro,spot:XLSX,path:XLSX>PDF,surface:powershell,outro",
      "posterMs": 6000
    }
  ]
}
```

`scenes` is a comma-separated sequence:

| Scene | Shows |
| --- | --- |
| `intro` | The catalog's format/route network and introduction |
| `spot:DOCX` | A format and its connected routes |
| `path:DOCX>Markdown` | A conversion path |
| `surface:powershell` | Routes available on a surface |
| `outro` | The OfficeIMO closing composition |

Append a surface qualifier, such as `spot:XLSX@powershell`, to filter a spotlight or path. Surface identifiers are `studio`, `browser`, `cli`, `dotnet` and `powershell`; use the catalog's exact format names. Omitting `scenes` plays the full generated sequence. An empty scene list, unknown name or scene without matching routes is an error.

`speed` accepts 0.4-3; higher values shorten the clip. `posterMs` selects a point in the resulting playback timeline and must fit its duration. Optional defaults are `speed: 1`, `posterMs: 2500`, `fps: 30`, `gifWidth: 640`, `gifFps: 12` and `scale: 1`. Ratios are `16x9`, `1x1`, `4x5` and `9x16`; dimensions must match the ratio. JSON field names are case-insensitive and comments are allowed, but unknown or duplicate fields and missing required fields are rejected.

For a capability update:

1. Update the owning conversion catalog or the reviewed PowerShell route mapping. Do not hand-edit generated `Website/data/format_map.json`.
2. Run the wrapper without `-SkipCatalog` to regenerate the route data. Review its generated changes; the wrapper preserves working-tree edits.
3. Update the scene selections in each relevant cut. Counts and surface badges come from the regenerated data; they do not need handwritten numbers.
4. Capture opening, spotlight, path and outro stills with `-Still` before a full render. Check long format names and captions in each selected composition.
5. Render at the intended frame rate, inspect the clips and posters, and verify file size/duration against the destination's current upload rules.

To change typography, camera movement, captions or composition, edit [the media page](../Build/FormatMapMedia/index.html). Its `ALIAS` table shortens presentation labels without renaming catalog formats. Changes to the interactive website map do not automatically change this media page.

## Rendering and output boundaries

The page exposes `window.imoMedia.render(milliseconds)`: drawing is a function of time, rather than a real-time screen recording. The recorder serves the page, generated catalog, surface data and site font on a private loopback address, then captures each timestamp through HtmlTinkerX. No website build is required. Timing is repeatable; identical bytes across different operating systems or browser versions are not guaranteed.

Requested encoders and codec dimensions are checked before rendering. Stills and posters must fit the cut duration. Rendering and encoding finish in unique staging directories before replacing final artifacts. Failed replacements attempt to restore previous outputs; blocked recovery reports where the retained files live. This does not provide crash consistency or concurrent-writer isolation. Use separate output directories for concurrent runs.

Windows capture and encoding are locally qualified. macOS/Linux recorder execution, Firefox/Safari rendering and the platforms' upload/transcoding results require separate validation. Inspect the actual upload preview before publishing. The supplied silent clips do not claim spoken captions or an audio track.
