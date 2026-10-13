---
name: officeimo-studio-translations
description: Use when adding OfficeIMO Studio interface languages, completing or correcting translations, or adding interface text to the desktop Studio (Avalonia, Windows and Linux) or the native Apple apps (Mac, iPad, iPhone). Covers the shared export, translate, import and check workflow for both apps.
---

# OfficeIMO Studio translations

Studio's interface text lives in two places, one per app family, and one workflow translates both:

- **Desktop (Avalonia):** `OfficeIMO.Studio/Localization/StudioStrings.resx` holds English; `StudioStrings.<language>.resx`
  holds each language. Code and XAML read stable, feature-scoped keys through `IStudioLocalizer` (`UiText`, `UiFormat`,
  `{loc:Loc Key}`). The picker offers the languages in `StudioCultureCatalog.ShippedCultures`.
- **Apple (Mac, iPad, iPhone):** `Apple/OfficeIMOStudio/Shared/Localizable.xcstrings` holds the interface,
  `AppShortcuts.xcstrings` the Shortcuts phrases and `InfoPlist.xcstrings` the app name and permission prompts. SwiftUI
  text and `String(localized:)` are keyed by their English text. The Xcode project's `knownRegions` lists the languages.

The document engines stay culture-neutral: they don't translate document content or vary behaviour by language.

## Add languages

Run from the repository root. `Build/Studio` holds the scripts; each takes several languages at once.

1. **On a Mac, sync the Apple strings first:** `./Build/Studio/Sync-StudioAppleStrings.ps1`. It builds the iOS app,
   extracts the Swift strings and adds new ones to `Localizable.xcstrings`; `xcodebuild` alone doesn't. Without a Mac,
   skip it and say the Apple catalog may lack strings added since its last sync.
2. **Export:** `./Build/Studio/Export-StudioTranslations.ps1 -Language it,es,pt-BR,ja -OutputDirectory Artefacts/Translations`
   writes `<language>.json` with every string that language lacks in both apps (about 2,600 for a new language), its
   English `source`, a `context` and an empty `translation`. `-All` exports translated strings too, to review a language.
3. **Translate** each file: fill every `translation`, following the `rules` in the file. Work in batches; a file can be
   imported partly translated and exported again for the rest.
4. **Import:** `./Build/Studio/Import-StudioTranslations.ps1 -Path Artefacts/Translations/it.json,Artefacts/Translations/es.json`
   writes the desktop `.resx` and the Apple catalogs and registers each language in `ShippedCultures` and
   `knownRegions`. It refuses a file whose translation changes a placeholder, naming each one; nothing from that file is
   written. `-NoRegister` writes strings without offering the language.
5. **Check:** `./Build/Studio/Test-StudioTranslations.ps1` must pass (CI runs it in the Studio job). Then run the
   desktop tests (`StudioTranslationResourceTests`, `StudioInfrastructureTests`) and build the Mac and iOS apps.
6. Delete `Artefacts/Translations` when done; the translation files aren't kept in the repository.

Use culture names .NET and Apple share: `it`, `es`, `pt-BR`, `nb`, `zh-Hans`, `zh-Hant`. `ReservedCultures` lists names
already accepted in preferences.

## Translating well

- Keep every placeholder: `{0}`, `{1:N0}`, `{page}`, `%@`, `%lld`, `%1$@`, `${applicationName}`. Positional forms such as
  `%2$@ … %1$@` may reorder Apple values. Header and footer fields (`{page}`, `{pages}`, `{date}`, `{file}`, `{title}`)
  stay in English because Studio replaces them.
- Keep product and format names: OfficeIMO, Studio, PDF, PDF/A, PDF/UA, Word, Excel, PowerPoint, Markdown, JSON, Bates.
- Use the operating system's own terms for standard commands (Undo, Save, Share, Print) in that language, and stay close
  to the English length: menus, toolbars and sheets have limited room.
- Translate whole sentences naturally; counts such as "(%lld)" may move within the sentence.
- Right-to-left languages (`ar`, `he`, `fa`, `ur`) mirror both apps: desktop windows, menus and dialogs follow the
  culture (`App.ApplyFlowDirection`), and the Apple apps follow the system language. PDF pages, scans, signatures,
  images and file locations keep their own direction. After importing one, run `StudioRightToLeftLayoutTests` with
  `OFFICEIMO_STUDIO_VISUAL_OUTPUT` set and look at the screenshots; on a Mac, launch the app with
  `-AppleTextDirection YES -NSForceRightToLeftWritingDirection YES` to see the mirrored Apple layout.
- Languages are translated and reviewed by AI agents; no human translation is required. A pack ships when the check
  passes and a second agent pass has reviewed the whole language in the export file (`-All`) for meaning, consistent
  terminology, the platform's standard terms and length, with its corrections imported.

## Adding interface text

- **Desktop:** add the English string to `StudioStrings.resx` under a feature-scoped key and the same key to every shipped
  language's `.resx` in the same change. Never put user-facing text in code or XAML.
- **Apple:** write the text in SwiftUI or `String(localized:)`, run `Sync-StudioAppleStrings.ps1`, and translate the new
  catalog entries into every known region. Mark strings with no words, such as `%@ – %@`, `shouldTranslate: false`.
- For a few new strings, export with the languages concerned, translate and import, as above.
- Engine messages (exceptions from Studio.Core and the native engine) are English; the roadmap records making them
  translatable. Don't translate them in the apps by matching message text.
