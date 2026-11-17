namespace OfficeIMO.Reader;

/// <summary>Maps source-code file extensions and script interpreters to stable language identifiers.</summary>
/// <remarks>
/// Extensions owned by structured format readers (JSON, YAML, TOML, XML, Markdown, CSV, HTML, LaTeX and others)
/// are deliberately absent so those readers keep their richer projections.
/// </remarks>
internal static class ReaderSourceCodeLanguages {
    internal sealed class Language {
        internal Language(string id, string name) {
            Id = id;
            Name = name;
        }

        /// <summary>Lower-case identifier used as the Markdown fence info string.</summary>
        internal string Id { get; }

        /// <summary>Human-readable language name.</summary>
        internal string Name { get; }
    }

    private static readonly Language CSharp = new("csharp", "C#");
    private static readonly Language FSharp = new("fsharp", "F#");
    private static readonly Language VisualBasic = new("vbnet", "Visual Basic");
    private static readonly Language PowerShell = new("powershell", "PowerShell");
    private static readonly Language Python = new("python", "Python");
    private static readonly Language JavaScript = new("javascript", "JavaScript");
    private static readonly Language TypeScript = new("typescript", "TypeScript");
    private static readonly Language Java = new("java", "Java");
    private static readonly Language Kotlin = new("kotlin", "Kotlin");
    private static readonly Language Scala = new("scala", "Scala");
    private static readonly Language Groovy = new("groovy", "Groovy");
    private static readonly Language Go = new("go", "Go");
    private static readonly Language Rust = new("rust", "Rust");
    private static readonly Language C = new("c", "C");
    private static readonly Language Cpp = new("cpp", "C++");
    private static readonly Language ObjectiveC = new("objectivec", "Objective-C");
    private static readonly Language Swift = new("swift", "Swift");
    private static readonly Language Dart = new("dart", "Dart");
    private static readonly Language Ruby = new("ruby", "Ruby");
    private static readonly Language Php = new("php", "PHP");
    private static readonly Language Perl = new("perl", "Perl");
    private static readonly Language Lua = new("lua", "Lua");
    private static readonly Language R = new("r", "R");
    private static readonly Language Julia = new("julia", "Julia");
    private static readonly Language Haskell = new("haskell", "Haskell");
    private static readonly Language Elixir = new("elixir", "Elixir");
    private static readonly Language Erlang = new("erlang", "Erlang");
    private static readonly Language Clojure = new("clojure", "Clojure");
    private static readonly Language Sql = new("sql", "SQL");
    private static readonly Language Shell = new("bash", "Shell");
    private static readonly Language Batch = new("batch", "Batch");
    private static readonly Language Css = new("css", "CSS");
    private static readonly Language Scss = new("scss", "SCSS");
    private static readonly Language Less = new("less", "Less");
    private static readonly Language Vue = new("vue", "Vue");
    private static readonly Language Svelte = new("svelte", "Svelte");
    private static readonly Language Protobuf = new("protobuf", "Protocol Buffers");
    private static readonly Language GraphQl = new("graphql", "GraphQL");
    private static readonly Language Terraform = new("hcl", "HCL");
    private static readonly Language Ini = new("ini", "INI");
    private static readonly Language Properties = new("properties", "Properties");

    private static readonly Dictionary<string, Language> ByExtension = new(StringComparer.OrdinalIgnoreCase) {
        [".cs"] = CSharp,
        [".csx"] = CSharp,
        [".fs"] = FSharp,
        [".fsi"] = FSharp,
        [".fsx"] = FSharp,
        [".vb"] = VisualBasic,
        [".ps1"] = PowerShell,
        [".psm1"] = PowerShell,
        [".psd1"] = PowerShell,
        [".py"] = Python,
        [".pyw"] = Python,
        [".pyi"] = Python,
        [".js"] = JavaScript,
        [".mjs"] = JavaScript,
        [".cjs"] = JavaScript,
        [".jsx"] = JavaScript,
        [".ts"] = TypeScript,
        [".mts"] = TypeScript,
        [".cts"] = TypeScript,
        [".tsx"] = TypeScript,
        [".java"] = Java,
        [".kt"] = Kotlin,
        [".kts"] = Kotlin,
        [".scala"] = Scala,
        [".groovy"] = Groovy,
        [".gradle"] = Groovy,
        [".go"] = Go,
        [".rs"] = Rust,
        [".c"] = C,
        [".h"] = C,
        [".cpp"] = Cpp,
        [".cc"] = Cpp,
        [".cxx"] = Cpp,
        [".c++"] = Cpp,
        [".hpp"] = Cpp,
        [".hh"] = Cpp,
        [".hxx"] = Cpp,
        [".mm"] = ObjectiveC,
        [".swift"] = Swift,
        [".dart"] = Dart,
        [".rb"] = Ruby,
        [".php"] = Php,
        [".pl"] = Perl,
        [".pm"] = Perl,
        [".lua"] = Lua,
        [".r"] = R,
        [".jl"] = Julia,
        [".hs"] = Haskell,
        [".ex"] = Elixir,
        [".exs"] = Elixir,
        [".erl"] = Erlang,
        [".clj"] = Clojure,
        [".cljs"] = Clojure,
        [".sql"] = Sql,
        [".sh"] = Shell,
        [".bash"] = Shell,
        [".zsh"] = Shell,
        [".ksh"] = Shell,
        [".bat"] = Batch,
        [".cmd"] = Batch,
        [".css"] = Css,
        [".scss"] = Scss,
        [".less"] = Less,
        [".vue"] = Vue,
        [".svelte"] = Svelte,
        [".proto"] = Protobuf,
        [".graphql"] = GraphQl,
        [".gql"] = GraphQl,
        [".tf"] = Terraform,
        [".hcl"] = Terraform,
        [".ini"] = Ini,
        [".cfg"] = Ini,
        [".conf"] = Ini,
        [".properties"] = Properties
    };

    // Interpreter names found after "#!" (directly or through env) mapped to the extension that owns the language.
    private static readonly Dictionary<string, string> ExtensionByInterpreter = new(StringComparer.Ordinal) {
        ["sh"] = ".sh",
        ["bash"] = ".sh",
        ["dash"] = ".sh",
        ["ksh"] = ".sh",
        ["zsh"] = ".sh",
        ["ash"] = ".sh",
        ["python"] = ".py",
        ["python2"] = ".py",
        ["python3"] = ".py",
        ["pwsh"] = ".ps1",
        ["powershell"] = ".ps1",
        ["node"] = ".js",
        ["nodejs"] = ".js",
        ["deno"] = ".ts",
        ["ts-node"] = ".ts",
        ["ruby"] = ".rb",
        ["perl"] = ".pl",
        ["php"] = ".php",
        ["lua"] = ".lua",
        ["Rscript"] = ".r",
        ["julia"] = ".jl",
        ["elixir"] = ".exs",
        ["groovy"] = ".groovy",
        ["kotlin"] = ".kts",
        ["scala"] = ".scala",
        ["dotnet-script"] = ".csx"
    };

    /// <summary>Source-code extensions claimed by the Reader source-code handler, in ordinal order.</summary>
    internal static IReadOnlyList<string> Extensions { get; } = ByExtension.Keys
        .OrderBy(static extension => extension, StringComparer.Ordinal)
        .ToArray();

    internal static bool TryGetByExtension(string? extension, out Language language) {
        language = null!;
        if (string.IsNullOrEmpty(extension)) return false;
        if (!ByExtension.TryGetValue(extension!, out Language? resolved)) return false;
        language = resolved;
        return true;
    }

    /// <summary>
    /// Resolves the extension that owns the interpreter named by a leading <c>#!</c> line, such as
    /// <c>#!/usr/bin/env python3</c> or <c>#!/bin/bash -e</c>.
    /// </summary>
    internal static bool TryGetExtensionFromShebang(string text, out string extension) {
        extension = string.Empty;
        if (text == null) return false;
        int start = text.Length > 0 && text[0] == '﻿' ? 1 : 0;
        if (text.Length < start + 3 || text[start] != '#' || text[start + 1] != '!') return false;
        int end = text.IndexOfAny(new[] { '\r', '\n' }, start);
        string line = (end < 0 ? text.Substring(start + 2) : text.Substring(start + 2, end - start - 2)).Trim();
        if (line.Length == 0 || line.Length > 256) return false;
        string[] parts = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        int index = 0;
        string command = FileName(parts[index]);
        if (command == "env") {
            index++;
            while (index < parts.Length && (parts[index].StartsWith("-", StringComparison.Ordinal) || parts[index].IndexOf('=') > 0)) index++;
            if (index >= parts.Length) return false;
            command = FileName(parts[index]);
        }
        if (ExtensionByInterpreter.TryGetValue(command, out string? mapped)) {
            extension = mapped;
            return true;
        }
        // Versioned interpreters such as python3.12 or ruby3.3 share the unversioned language.
        string unversioned = command.TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.');
        if (unversioned.Length > 0 && unversioned.Length < command.Length &&
            ExtensionByInterpreter.TryGetValue(unversioned, out mapped)) {
            extension = mapped;
            return true;
        }
        return false;
    }

    private static string FileName(string value) {
        int slash = Math.Max(value.LastIndexOf('/'), value.LastIndexOf('\\'));
        return slash >= 0 ? value.Substring(slash + 1) : value;
    }
}
