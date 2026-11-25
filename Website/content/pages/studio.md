---
title: "OfficeIMO Studio"
description: "Edit supported PDF text and images, review comments, fill forms, organize pages, and convert documents in OfficeIMO Studio. See screenshots and availability."
meta.seo_title: "OfficeIMO Studio: desktop document tools and conversion"
layout: application
meta.application: "studio"
---

## Keep the document in view

Open PDFs in tabs, switch between reading layouts, and find individual text matches on the page. Use light or dark application appearance without changing the document's own colors.

Studio brings the OfficeIMO document engines into a desktop workspace on Windows and Linux. Native apps for Mac, iPad and iPhone use the same engines; see [what Studio does on each system](#features). It works with supported PDF content directly; reconstructing an original Word document from a PDF is a separate conversion task.

Read the [Studio privacy policy](/studio/privacy/) for details about local document processing, saved application data, and optional document assistant connections.

## Review before you save

- **Edit supported content:** select existing PDF text and images for object-level edits. This is not word-processor paragraph reflow; scanned text requires OCR and remains a separate workflow.
- **Comments and forms:** follow anchored comment threads, reply and resolve, fill supported fields, and inspect appearance previews before applying changes.
- **Pages:** select ranges, review a new page order, and import, extract, or split pages into separate documents.
- **Searchable scans:** review recognized words before creating a searchable PDF. On Windows and Linux, OCR uses an optional Tesseract installation and language data.
- **Protected documents:** view files according to their permissions, prepare protected copies, and inspect signed output. A drawn signature and a certificate signature are different operations.
- **Recoverable work:** restore supported editing sessions, inspect job outcomes, and control whether recent-file history and recovery snapshots are stored.

## Complete a document task

| Task | Studio workflow | Boundary |
|---|---|---|
| Assemble a delivery PDF | Combine PDFs, supported Office files, and images | Inspect conversion diagnostics for each input |
| Remove sensitive content | Mark redactions and review permanent removal | Save a separate output and inspect the verification evidence |
| Compare revisions | Compare PDFs and open the resulting report | A visual comparison does not establish semantic equivalence |
| Reduce document overhead | Run supported lossless optimization | Scan image downsampling is a separate capability |
| Print | Preview sheets and page placement, then print to an installed printer with copies, duplex and a paper source | Driver-specific color settings and physical-printer acceptance remain open |

Open any screenshot at full size to inspect the workspace. For immediate use without installing Studio, try the [browser PDF tools](/pdf/). Their focused operations and limits are described on each task page.

## Studio for Mac, iPad and iPhone

A native SwiftUI app for Mac, iPad and iPhone runs the same OfficeIMO engine, compiled with NativeAOT, and follows each platform's document conventions. It requires macOS, iPadOS or iOS 26. It is coming soon to the Mac App Store and App Store; until then, the [Apple app guide](https://github.com/EvotecIT/OfficeIMO/tree/master/Apple/OfficeIMOStudio#readme) describes its current state.

- **Read and mark up:** search, select and copy text, highlight, draw, and add notes, shapes and stamps.
- **Pages:** rotate, reorder, insert, extract and split pages, and combine PDFs, images and Office files into one PDF.
- **Forms and signatures:** fill form fields, place a handwritten signature, and sign a copy with a certificate from a .p12 or .pfx file. The certificate signature is invisible.
- **Redaction:** mark text you find or select, or areas you drag across, then remove the content and verify the result.
- **Scans and text recognition:** scan paper with the document camera on iPad and iPhone or Continuity Camera on the Mac. **Make Searchable** recognizes text on the device with Apple's Vision framework.
- **Conversion:** convert Word, Excel, PowerPoint and other documents to PDF, and export a PDF to Word, Excel, PowerPoint, HTML, Markdown, text or images.
- **View Office files:** Word, Excel, PowerPoint, RTF, Markdown and OpenDocument files open read-only as their PDF conversion, with **Save as PDF** and **Edit as PDF**. Markdown is shown as converted PDF pages, not rendered Markdown. Files larger than 100 MB are not opened.
- **Shortcuts:** actions with Siri and Spotlight phrases combine PDFs, make a PDF from images, extract pages, get a PDF's text, make a PDF searchable, and convert a document to PDF.

The start screen offers conversion, combining files, a PDF from photos or a scan, a blank PDF and a sample. Comment threads, comparison, editing existing text and images, creating form fields, visible certificate signatures and the document assistant are still desktop-only. The [feature table](#features) compares each system.

## Get Studio

<div class="imo-platforms">
  <a class="imo-platform" href="https://apps.microsoft.com/detail/xp9m6j7hvm2hz0"><b>Microsoft Store</b><span>Install and update through the Store</span><small>Windows 10 and 11</small></a>
  <a class="imo-platform" href="/downloads/#studio-downloads"><b>Windows</b><span>Signed MSI installer or portable ZIP</span><small>x64 and Arm64 · .NET runtime included</small></a>
  <a class="imo-platform" href="/downloads/#studio-linux-downloads"><b>Linux</b><span>Debian package or portable ZIP</span><small>x64 and Arm64 · .NET runtime included</small></a>
  <a class="imo-platform" href="https://github.com/EvotecIT/OfficeIMO/tree/master/Apple/OfficeIMOStudio#readme"><b>Mac, iPad and iPhone</b><span>Native app, coming soon to the App Store</span><small>macOS, iPadOS and iOS 26</small></a>
</div>

Windows and Linux x64 installation and launch were verified for the initial direct release. Arm64 packages are available, with installation and runtime testing on matching hardware still pending. Release notes and checksums are linked beside the current downloads.
