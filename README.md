# ox

Read and edit Office documents from the terminal.

A single native binary with no dependencies. Treats `.docx`, `.pptx`, and `.xlsx` files like text files — read, edit, search, pipe to `grep`. Designed for humans and AI agents.

## Install

Download from [GitHub Releases](https://github.com/danhayman/ox-cli/releases/latest):

| Platform | Download |
|----------|----------|
| Linux x64 | `ox-linux-x64.tar.gz` |
| macOS ARM64 | `ox-osx-arm64.tar.gz` |
| Windows x64 | `ox-win-x64.zip` |

```bash
# Linux / macOS
tar xzf ox-*.tar.gz
sudo mv ox /usr/local/bin/

# Build from source (.NET 10 SDK required)
dotnet publish src/Ox -r linux-x64 -c Release -p:PublishAot=true -p:StripSymbols=true
sudo cp src/Ox/bin/Release/net10.0/linux-x64/publish/ox /usr/local/bin/
```

The output is a native AOT-compiled binary. No .NET runtime needed.

## Commands

```
ox cat <file>...    Dump plain text (pipe to grep to search)
ox read <file>...   Read content with character offset/limit
ox info <file>      Show metadata (title, author, word count)
ox edit <file>...   Replace text (use \n for new paragraphs)
ox comment          Add, list, or remove comments
```

## Examples

Read a document:

```bash
ox read report.docx
```

Search across a directory of documents:

```bash
ox cat "**/*.docx" | grep -i "budget"
# reports/Q1.docx:Revenue exceeded budget by 15%
# reports/Q3.docx:Budget review scheduled for Monday
```

Edit text:

```bash
ox edit report.docx --old "Draft" --new "Final"
```

Insert new paragraphs (use `\n`):

```bash
ox edit report.docx --old "Introduction" --new "Introduction\nThis report covers Q1 results."
```

Delete text:

```bash
ox edit report.docx --old "Remove this sentence." --new ""
```

Find and replace across files:

```bash
ox edit "**/*.docx" --old "Acme Corp" --new "Globex Inc" --replace-all
# contracts/nda.docx: replaced 3 occurrences
# contracts/sow.docx: replaced 7 occurrences
# total: 10 replacements across 2 files
```

Preview before editing:

```bash
ox edit report.docx --old "Draft" --new "Final" --dry-run
# would replace 1 occurrence
```

Edit with tracked changes:

```bash
ox edit report.docx --old "Draft" --new "Final" --track --author "Dan"
```

Add a comment:

```bash
ox comment add report.docx --at "revenue figures" --text "Need source for this claim"
```

## Design

**Unix philosophy.** Plain text in, plain text out. Compose with `grep`, `awk`, `wc`, `diff`, or any tool that works with text. Multi-file output prefixes lines with filenames, just like `grep -r`.

**Agent native.** Built for AI coding agents like Claude Code. The `read` command has character-based `--offset` and `--limit` for deterministic pagination. The `edit` command works like an agent's code editor — match unique text and replace it. `\n` creates new paragraphs just as it creates new lines in code. Glob patterns let agents search and edit across entire directories.

**Zero dependency.** Single native binary, AOT-compiled. No .NET runtime, no Python, no Java. Copy it anywhere and it runs.

**Safe by default.** Edits require unique text matches — ambiguous matches fail with a clear error. Use `--replace-all` for bulk operations. `--dry-run` previews changes. `--backup` creates `.docx.bak` files. Lock detection warns when Word has a file open.
