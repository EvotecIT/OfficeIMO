namespace OfficeIMO.Notebook.Import;

/// <summary>Input profiles for importing saved notebooks into the Jupyter document model.</summary>
public enum NotebookImportFormat {
    /// <summary>A .NET Polyglot Interactive script using the .dib format.</summary>
    PolyglotDib,
    /// <summary>A Jupytext script that uses the percent cell format.</summary>
    JupytextPercent,
    /// <summary>A Quarto .qmd source document.</summary>
    Quarto,
    /// <summary>MyST Markdown selected explicitly as notebook source.</summary>
    MySTMarkdown,
    /// <summary>R Markdown .Rmd source.</summary>
    RMarkdown,
    /// <summary>R Markdown HTML Notebook with embedded source and saved results.</summary>
    RNotebookHtml,
    /// <summary>A native Deepnote file or snapshot.</summary>
    Deepnote,
    /// <summary>An exported Apache Zeppelin note.</summary>
    Zeppelin,
    /// <summary>A PowerShell script, optionally containing OfficeIMO notebook cell boundaries.</summary>
    PowerShellScript
}
