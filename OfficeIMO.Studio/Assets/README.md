# OfficeIMO Studio identity

`OfficeIMO.Studio.svg` is the editable source for the Studio application icon. It adapts the current OfficeIMO website mark—a document sheet with a cyan companion spine—into a full-bleed native application tile. `OfficeIMO.Studio.png` (1024 px), `OfficeIMO.Studio-512.png`, and `OfficeIMO.Studio.ico` are generated exports used by packaging and Windows application metadata. The 512 px export is the Linux desktop icon, because the standard hicolor icon theme has no 1024 px size directory.

Keep the mark free of text so it remains legible at launcher and taskbar sizes. Platform packaging may apply its own mask; do not pre-render a circular crop or replace the OfficeIMO product identity with the Evotec corporate mark.

`OfficeIMO.Studio.AppStore.svg` uses the same artwork with a square background
that reaches every edge. Its opaque 1024 px export is
`Apple/OfficeIMOStudio/Shared/Assets.xcassets/AppIcon.appiconset/OfficeIMO.Studio-iOS.png`.
The iOS asset lets Apple apply the platform mask without transparent corners.
