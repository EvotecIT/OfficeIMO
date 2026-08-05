---
title: "Frequently Asked Questions"
description: "Answers about OfficeIMO licensing, DOC/XLS/PPT and modern format support, conversion fidelity, .NET platforms, PowerShell, containers, and dependencies."
layout: faq
meta.faq.questions:
  - "Does OfficeIMO require Microsoft Office to be installed?|No. OfficeIMO uses managed document engines and focused package dependencies; it does not require Microsoft Office, COM automation, or Office interop assemblies at runtime."
  - "Does OfficeIMO support DOC, XLS, XLSB, and PPT as well as DOCX, XLSX, and PPTX?|OfficeIMO 3.1 classifies and reads modern and legacy Word, Excel, and PowerPoint families, writes documented native subsets, and reports conversion coverage by operation and fidelity."
  - "What .NET versions are supported?|OfficeIMO 3.1 targets .NET 8.0, .NET 10.0, .NET Standard 2.0, and .NET Framework 4.7.2."
  - "Is OfficeIMO free for commercial use?|Yes. OfficeIMO packages are MIT licensed. Commercial teams should also review the separate terms of optional third-party dependencies used by their selected package set."
  - "How does OfficeIMO compare to Aspose or GemBox?|OfficeIMO leads with MIT licensing, source access, modular packages, PowerShell, and explicit fidelity. Commercial suites can provide broader portfolio coverage, mature rendering for some workloads, formal support, and procurement SLAs."
  - "Can I use OfficeIMO in a Docker container or CI/CD pipeline?|Yes. The COM-free .NET packages fit Linux containers, GitHub Actions, Azure DevOps, and other CI/CD environments. Validate fonts and native dependencies for rendering-heavy workloads."
  - "What is PSWriteOffice?|PSWriteOffice is the first-party PowerShell surface over OfficeIMO. Its command catalog is generated from the module manifest used for this release."
  - "Does OfficeIMO support reading existing documents?|Yes, within each package's documented format and feature boundaries. OfficeIMO.Reader provides a normalized extraction API across modular format handlers."
  - "Is NativeAOT compilation supported?|The OfficeIMO 3.1 validation matrix covers 92 of 93 production projects: 90 fully rooted libraries, one bounded Google APIs workflow, one native command-line tool, and one managed-only WPF/WebView2 renderer. Check the matrix for the exact package and runtime path."
  - "What are the dependencies?|The core Office packages use DocumentFormat.OpenXml and the first-party OfficeIMO.Core foundation. Drawing primitives remain in the OfficeIMO.Drawing namespace. Optional converter and compatibility packages add focused dependencies documented on the Third-Party Dependencies page."
  - "Can I convert Word documents to PDF?|Yes. OfficeIMO.Word.Pdf provides Word-to-PDF export without Microsoft Office, and OfficeIMO.Excel.Pdf provides the corresponding Excel route."
  - "Is thread safety supported?|Separate document instances can run on separate threads. Concurrent access to the same document instance is not supported; OfficeIMO.Excel also provides parallel bulk operations such as AutoFit and bulk writes."
---

{{< faq >}}
