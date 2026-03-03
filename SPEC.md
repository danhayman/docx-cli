# ox-cli

Read and edit Office documents from the terminal.
Built for AI agents. Works for humans too.

Inspired by [gogcli](https://github.com/steipete/gogcli) — same ergonomics, but for local `.docx` files.

## Why

AI agents need to edit documents. Current options are:
- **python-docx** — library, not a CLI. Requires writing scripts.
- **WPS/Word COM automation** — needs Windows, fragile, slow.
- **XML unpack/edit/repack** — error-prone, verbose.
- **Google Docs API** — cloud only, needs auth.

There's no simple `edit --old "X" --new "Y"` for local .docx files. This fills that gap.

## Language

**C# / .NET 8+**. Microsoft's official Open XML SDK gives us first-party OOXML support — tracked changes, comments, styles, cross-run editing — all handled by the library that Microsoft built for their own format. AOT-compiled to a single native binary with no runtime dependency.

## Dependencies

- [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) (MIT) — Microsoft's official Open XML SDK. Full read/write/edit of .docx, .pptx, .xlsx. Tracked changes, comments, styles, formatting — all built in.
- [System.CommandLine](https://github.com/dotnet/command-line-api) — Microsoft's CLI framework (similar to kong in Go). Supports subcommands, flags, help generation, tab completion.
- No other external dependencies needed.

## Why .NET over Go

| | Go | .NET |
|---|---|---|
| **Best library** | fumiama/go-docx (community, partial) | Open-XML-SDK (Microsoft official, complete) |
| **OOXML coverage** | Partial — roll your own XML | Complete — first-party support |
| **Tracked changes** | Manual XML manipulation | Built-in API |
| **Comments** | Manual XML manipulation | Built-in API |
| **Cross-run editing** | Build from scratch | SDK handles run/paragraph structure |
| **Single binary** | ✅ `go build` | ✅ `dotnet publish -r linux-x64 --self-contained -p:PublishAot=true` |
| **Binary size** | ~30MB (gog is 32MB) | ~30-50MB (trimmed AOT) |
| **Startup** | Instant | Fast with AOT |
| **Cross-platform** | ✅ | ✅ (linux-x64, osx-arm64, win-x64) |

The SDK saves weeks of XML pain. Tracked changes alone would be a massive undertaking in Go.

## File Format Overview

A `.docx` file is a ZIP archive containing:
```
word/document.xml    — main content (paragraphs, runs, tables)
word/comments.xml    — comments referenced from document.xml
word/styles.xml      — style definitions
word/_rels/          — relationships
[Content_Types].xml  — MIME types
```

Text lives in **runs** (`Run`) inside **paragraphs** (`Paragraph`):
```csharp
// Open XML SDK gives you typed access
using var doc = WordprocessingDocument.Open("file.docx", true);
var body = doc.MainDocumentPart.Document.Body;

foreach (var para in body.Elements<Paragraph>())
{
    foreach (var run in para.Elements<Run>())
    {
        var text = run.InnerText;
        var props = run.RunProperties; // bold, italic, font, etc.
    }
}
```

Key challenge: a single word/phrase may span multiple runs due to formatting boundaries. The SDK gives you the structure; you still need to handle cross-run text matching.

---

## Commands

### `ox read <file>`

Display document content with character offset/limit for pagination.

```
$ ox read contract.docx
Agreement between Party A and Party B

The term is 30 days from the date of signing.
Party A shall deliver all materials within
this period.

Compensation shall not exceed £50,000.
```

**Flags:**
- `--offset` — character offset to start from (0-indexed)
- `--limit` — max characters to read (default: 100000)
- `--track-changes` — show tracked changes inline: ~~deleted~~ **inserted**

**Implementation:**
```csharp
using var doc = WordprocessingDocument.Open(path, false);
var body = doc.MainDocumentPart.Document.Body;

foreach (var para in body.Elements<Paragraph>())
{
    Console.WriteLine(para.InnerText);
}
```

### `ox cat <file>`

Plain text dump. No formatting. Pipe-friendly.

```
$ ox cat contract.docx
Agreement between Party A and Party B

The term is 30 days from the date of signing...
```

### `ox info <file>`

Document metadata.

```
$ ox info contract.docx
File:       contract.docx
Size:       45.2 KB
Pages:      3 (estimated)
Paragraphs: 47
Words:      1,234
Author:     Dan Hayman
Created:    2026-02-20T10:30:00Z
Modified:   2026-02-25T14:15:00Z
```

**Implementation:**
```csharp
var props = doc.PackageProperties;
Console.WriteLine($"Author:   {props.Creator}");
Console.WriteLine($"Created:  {props.Created}");
Console.WriteLine($"Modified: {props.Modified}");
```

### `ox edit <file> --old "X" --new "Y"`

Replace text by unique string matching. Core command.

```
$ ox edit contract.docx --old "30 days" --new "60 days"
replaced 1 occurrence

$ ox edit contract.docx --old "Party" --new "Company"
error: "Party" is not unique (found 8 occurrences). Use --replace-all to replace all.

$ ox edit contract.docx --old "Party" --new "Company" --replace-all
replaced 8 occurrences
```

**Flags:**
- `--replace-all` — replace all occurrences (required if not unique)
- `--track` — insert as tracked change instead of direct edit
- `--author "Name"` — author for tracked changes (default: "ox")
- `--dry-run` — show what would change without modifying file
- `--backup` — create `.docx.bak` before editing
- `-o, --output <file>` — write to different file instead of in-place

**Implementation — cross-run matching:**
```csharp
// 1. Build text map across runs in a paragraph
var runs = para.Elements<Run>().ToList();
var fullText = string.Concat(runs.Select(r => r.InnerText));

// 2. Find match
int idx = fullText.IndexOf(oldText);
if (idx < 0) continue;

// 3. Map character offset back to runs
var (startRun, startOffset) = MapToRun(runs, idx);
var (endRun, endOffset) = MapToRun(runs, idx + oldText.Length);

// 4. Split boundary runs, replace content
// SDK handles the XML — we just manipulate Run objects
var newRun = new Run(
    startRun.RunProperties?.CloneNode(true),  // inherit formatting
    new Text(newText)
);

// 5. Remove old runs, insert new
RemoveRunRange(startRun, startOffset, endRun, endOffset);
startRun.InsertAfterSelf(newRun);
```

### `ox edit <file> --track --old "X" --new "Y"`

Same as edit but inserts tracked changes instead of direct replacement.

**Implementation with SDK:**
```csharp
// The SDK has built-in types for tracked changes
var deleteRun = new DeletedRun(
    new RunProperties(startRun.RunProperties?.CloneNode(true)),
    new DeletedText(oldText)
);

var insertRun = new InsertedRun(
    new RunProperties(startRun.RunProperties?.CloneNode(true)),
    new Text(newText)
);

// Wrap in tracked change markers
var del = new Deleted { Author = author, Date = DateTime.UtcNow };
del.Append(deleteRun);

var ins = new Inserted { Author = author, Date = DateTime.UtcNow };
ins.Append(insertRun);
```

User opens in Word/LibreOffice → sees red strikethrough + blue insertion → accepts/rejects.

### `ox comment <file>`

Add, list, or remove comments.

```
$ ox comment add contract.docx --at "indemnify" --text "Should we cap this?"
comment added at "indemnify" (id: 1)

$ ox comment list contract.docx
ID  AUTHOR      DATE        AT                  TEXT
1   ox          2026-02-26  "indemnify"         Should we cap this?
2   Dan         2026-02-20  "30 days"           Too short?

$ ox comment delete contract.docx --id 1
comment 1 deleted
```

**Implementation with SDK:**
```csharp
// Add comment to comments part
var commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart
    ?? doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();

var comment = new Comment
{
    Id = nextId.ToString(),
    Author = author,
    Date = DateTime.UtcNow
};
comment.Append(new Paragraph(new Run(new Text(commentText))));
commentsPart.Comments.Append(comment);

// Add range markers in document body
para.InsertBefore(new CommentRangeStart { Id = nextId.ToString() }, targetRun);
para.InsertAfter(new CommentRangeEnd { Id = nextId.ToString() }, targetRun);
para.InsertAfter(new Run(new CommentReference { Id = nextId.ToString() }), targetRun);
```

### `ox diff <file1> <file2>`

Show differences between two documents. Nice-to-have for v2.

```
$ ox diff original.docx edited.docx
"30 days" → "60 days"
"£50,000" → "£75,000"
[deleted paragraph]
```

---

## Safety Features

### Lock File Detection

Before any write operation, check for:
- `~$filename.docx` (Word lock file)
- `.~lock.filename.docx#` (LibreOffice lock file)

```
$ ox edit contract.docx --old "X" --new "Y"
error: file appears to be open in another application (~$contract.docx exists)
hint: close the file first, or use --force to edit anyway
```

**Flags:**
- `--force` — ignore lock files and edit anyway
- Default: refuse to edit locked files

### Backup

```
$ ox edit contract.docx --old "X" --new "Y" --backup
backed up to contract.docx.bak
replaced 1 occurrence
```

### Dry Run

Every write command supports `--dry-run`:
```
$ ox edit contract.docx --old "30 days" --new "60 days" --dry-run
would replace 1 occurrence:
  "The term is [30 days] from the date..."
            → "The term is [60 days] from the date..."
```

### Output to Different File

```
$ ox edit contract.docx --old "X" --new "Y" -o contract-edited.docx
```

---

## Global Flags

- `-n, --dry-run` — no changes
- `-y, --force` — skip confirmations
- `--version` — print version

---

## Build & Publish

### Development
```bash
dotnet build
dotnet run -- read contract.docx
dotnet test
```

### Single binary (AOT)
```bash
# Linux
dotnet publish -r linux-x64 -c Release -p:PublishAot=true -p:StripSymbols=true

# macOS (Apple Silicon)
dotnet publish -r osx-arm64 -c Release -p:PublishAot=true -p:StripSymbols=true

# Windows
dotnet publish -r win-x64 -c Release -p:PublishAot=true -p:StripSymbols=true
```

### Homebrew
```bash
brew install danhayman/tap/ox-cli
```

---

## Build Milestones

### v0.1 — Read (week 1)
- [ ] Project setup (.NET 8, System.CommandLine, Open XML SDK)
- [ ] `read` with character offset/limit and text wrapping
- [ ] `cat` plain text
- [ ] `info` metadata (author, dates, word count)
- [ ] Handle tables (basic text extraction)
- [ ] Lock file detection

### v0.2 — Edit (week 2)
- [ ] Single-run text replacement (`edit --old --new`)
- [ ] Cross-run text replacement
- [ ] Formatting preservation (inherit from first matched run)
- [ ] Uniqueness checking and `--replace-all`
- [ ] `--backup` and `-o` output
- [ ] `--dry-run` for all write commands

### v0.3 — Track Changes (week 3)
- [ ] `--track` flag for edit
- [ ] Proper `Inserted` / `Deleted` element generation via SDK
- [ ] Author and timestamp metadata
- [ ] Read and display existing tracked changes in `read --track-changes`

### v0.4 — Comments (week 4)
- [ ] `comment add/list/delete`
- [ ] Full comment support via SDK (CommentRangeStart/End, CommentReference)

### v0.5 — Polish (week 5)
- [ ] Edge cases: graceful errors for empty docs, password-protected, corrupted files
- [ ] AOT publishing for linux-x64, osx-arm64, win-x64
- [ ] CI/CD (GitHub Actions)
- [ ] GitHub releases with binaries
- [ ] Homebrew formula
- [ ] README with examples

---

## Project Structure

```
ox-cli/
├── src/
│   └── Ox/
│       ├── Program.cs                 # Entry point, command registration
│       ├── Ox.csproj                  # Project file (AOT-enabled)
│       ├── Commands/
│       │   ├── ReadCommand.cs
│       │   ├── CatCommand.cs
│       │   ├── InfoCommand.cs
│       │   ├── EditCommand.cs
│       │   └── CommentCommand.cs
│       ├── Core/
│       │   ├── DocumentService.cs     # Open, save, backup .docx
│       │   ├── TextSearch.cs          # Cross-run text matching
│       │   ├── TextReplacer.cs        # Replace with formatting preservation
│       │   ├── TrackedChanges.cs      # Insert/delete tracked changes
│       │   ├── CommentService.cs      # Comment management
│       │   └── LockDetector.cs        # Lock file detection
│       └── Output/
│           └── TextFormatter.cs       # Text wrapping, paragraph display
├── tests/
│   └── Ox.Tests/
│       ├── TextSearchTests.cs
│       ├── TextReplacerTests.cs
│       ├── TrackedChangesTests.cs
│       ├── CommentServiceTests.cs
│       └── IntegrationTests.cs
├── testdata/
│   ├── simple.docx
│   ├── formatted.docx
│   ├── tracked_changes.docx
│   ├── cross_run.docx              # Text split across runs
│   └── comments.docx
├── ox-cli.slnx
├── Makefile
├── README.md
└── LICENSE                          # MIT
```

---

## Key Technical Challenges

### 1. Cross-Run Text Matching

The hardest problem. Text like "Hello World" might be split across runs:
```csharp
// Run 1: "Hel" (bold)
// Run 2: "lo Wor" (bold)
// Run 3: "ld" (bold)
```

**Algorithm:**
```csharp
public static List<TextMatch> FindText(Paragraph para, string searchText)
{
    var runs = para.Elements<Run>().ToList();

    // Build concatenated text with run boundary tracking
    var sb = new StringBuilder();
    var runMap = new List<(int RunIndex, int StartInConcat)>();

    foreach (var (run, i) in runs.Select((r, i) => (r, i)))
    {
        runMap.Add((i, sb.Length));
        sb.Append(run.InnerText);
    }

    var fullText = sb.ToString();
    var matches = new List<TextMatch>();

    int pos = 0;
    while ((pos = fullText.IndexOf(searchText, pos)) >= 0)
    {
        matches.Add(new TextMatch(pos, searchText.Length, runMap));
        pos += searchText.Length;
    }

    return matches;
}
```

### 2. Formatting Preservation

The Open XML SDK makes this straightforward:
```csharp
// Clone formatting from the first matched run
var sourceProps = matchedRun.RunProperties?.CloneNode(true) as RunProperties;
var newRun = new Run();
if (sourceProps != null) newRun.Append(sourceProps);
newRun.Append(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });
```

### 3. Atomic File Writes

Write to temp file, then rename for safety:
```csharp
var tempPath = path + ".tmp";
using (var doc = WordprocessingDocument.Open(path, false))
{
    // Clone to temp
    doc.Clone(tempPath);
}
using (var doc = WordprocessingDocument.Open(tempPath, true))
{
    // Make edits
    doc.Save();
}
File.Move(tempPath, path, overwrite: true);
```

### 4. AOT Compatibility

Open XML SDK works with AOT as of v3.0+. Ensure:
- No reflection-based serialization
- Trim-compatible code
- Test AOT binary on all target platforms

---

## Testing Strategy

**Unit tests:**
- Text search across various run configurations
- Cross-run matching edge cases (match at run boundary, match spanning 3+ runs)
- Tracked change generation
- Comment insertion

**Integration tests (testdata/):**
- Create test .docx files covering edge cases
- Round-trip: read → edit → read, verify content
- Verify edited files open correctly in Word, LibreOffice, Google Docs
- Verify tracked changes appear correctly in Word

**Compatibility tests:**
- Files created by Word, LibreOffice, Google Docs, WPS Office
- Files with complex formatting (nested styles, themes)
- Files with existing tracked changes and comments

---

## README Draft

```markdown
# ox

Read and edit Office documents from the terminal.
Built for AI agents. Works for humans too.

Powered by Microsoft's [Open XML SDK](https://github.com/dotnet/Open-XML-SDK).

## Install

### Homebrew
brew install danhayman/tap/ox-cli

### Binary
Download from [GitHub Releases](https://github.com/danhayman/ox-cli/releases).

### Build from source
git clone https://github.com/danhayman/ox-cli
cd ox-cli
dotnet publish -r linux-x64 -c Release -p:PublishAot=true

## Quick Start

# Read a document
ox read contract.docx

# Edit text
ox edit contract.docx --old "30 days" --new "60 days"

# Edit with tracked changes (non-destructive)
ox edit contract.docx --old "30 days" --new "60 days" --track

# Add a comment
ox comment add contract.docx --at "indemnify" --text "Cap this?"

# Works with cloud-synced files
ox edit ~/Dropbox/contract.docx --old "draft" --new "final"
```

---

## Open Questions

1. **Should `edit` default to `--track` mode?** Safer, but more friction for simple edits. Suggest: default to direct edit, recommend `--track` in docs for shared files.

2. **Headers/footers** — include in v0.1 read output? They're accessible via `doc.MainDocumentPart.HeaderParts` / `FooterParts`.

3. **Images** — display `[Image: alt text]` placeholders in read output? Extract with a separate command?

4. **PPTX support** — The Open XML SDK handles .pptx too (`PresentationDocument`). Same binary or separate `pptx-cli`? Could share the text search/replace engine.

5. **Max file size** — should we set a limit? The SDK streams efficiently but very large docs could still consume memory.

6. **NuGet package** — publish `Ox.Core` as a NuGet for other .NET projects to use the text search/replace engine?
