# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

ox is a C# command-line tool for reading and editing Office documents (.docx, .pptx, .xlsx) from the terminal. It uses .NET 10.0 with Native AOT compilation and depends on Microsoft's Open-XML-SDK and System.CommandLine.

## Commands

```bash
make build              # dotnet build src/ox.slnx
make test               # dotnet test src/ox.slnx
make clean              # dotnet clean && rm -rf publish/

# Run a single test
dotnet test src/ox.slnx --filter "FullyQualifiedName~TextSearchTests.FindInRuns_SingleMatch"

# Run during development
dotnet run --project src/Ox -- read document.docx
dotnet run --project src/Ox -- cat presentation.pptx
dotnet run --project src/Ox -- cat spreadsheet.xlsx

# Publish native binary (platform-specific)
dotnet publish src/Ox -r linux-x64 -c Release -p:PublishAot=true -p:StripSymbols=true
```

## Architecture

**CLI layer** (`src/Ox/Commands/`): Each subcommand (cat, read, info, edit, comment) is a separate class registered in `Program.cs` via System.CommandLine. Global options (--dry-run, --force) are in `GlobalOptions.cs`. Error handling uses an extension method `SetActionWithErrorHandling()` in `ErrorHandler.cs`.

**Core logic** (`src/Ox/Core/`): The critical path for text operations is:
- `TextSearch.cs` — Cross-run text matching. DOCX splits text across multiple Run elements; this builds a concatenated text with a RunMapping array, searches it, then maps matches back to run boundaries. This is the most complex part of the codebase.
- `TextReplacer.cs` — Replaces text while preserving formatting. Handles single-run, cross-run, and newline (paragraph-creating) cases. Clones RunProperties from the start run.
- `TrackedChanges.cs` — Creates DeletedRun/InsertedRun pairs using Open-XML-SDK types with author/date/revisionId metadata.
- `CommentService.cs` — Manages comments stored in a separate XML part with CommentRangeStart/End anchors in the document body.
- `DocumentService.cs` — File I/O with read-only vs edit modes, atomic saves, paragraph extraction (including synthetic table rows). Pre-validates files (zero-byte, password-protected OLE detection).
- `PresentationService.cs` — Read-only extraction of text from .pptx slides (shapes → text bodies → paragraphs).
- `SpreadsheetService.cs` — Read-only extraction of cell values from .xlsx sheets. Resolves shared strings, outputs rows as TSV.

**Output** (`src/Ox/Output/TextFormatter.cs`): Text wrapping, track changes display (~~deleted~~ **inserted**), table formatting.

## Key Design Decisions

- **Safe-by-default**: Requires unique text matches (fails on ambiguous), lock file detection, backup support, dry-run preview, atomic writes.
- **AOT-compatible**: No reflection-based serialization. Single native binary with no runtime dependency.
- **Agent-native**: Glob patterns for batch ops, character-based offset/limit for pagination, `\n` for paragraph creation, plain text pipe-friendly output.
- **Multi-format read, docx-only edit**: cat/read/info work on all OOXML formats; edit/comment are restricted to .docx.

## Tests

xUnit tests in `tests/Ox.Tests/`. `TestHelper.cs` provides utilities for creating temporary .docx files: `CreateTestDocx()`, `CreateMultiRunDocx()`, `CreateDocxWithTable()`, and readers `ReadDocxText()`/`ReadDocxParagraphs()`. Pptx and xlsx tests have their own helpers in `PresentationServiceTests.cs` and `SpreadsheetServiceTests.cs`.

## CI/CD

- **CI**: `.github/workflows/ci.yml` — builds and tests on push to main and PRs.
- **Release**: `.github/workflows/release.yml` — on `v*` tags, builds AOT binaries for linux-x64, osx-arm64, win-x64 and creates a GitHub release.

## Binary

The published binary is named `ox` (set via `AssemblyName` in the csproj), not `Ox`.
