# Project Agent Instructions

This repository is prepared as an LLM-maintained project wiki, based on the
`llm-wiki` pattern.

## Layers

- `Примеры таблиц/` is the raw source layer. Treat files here as immutable input
  examples. Do not edit or rename them unless the user explicitly asks.
- `wiki/` is the maintained knowledge layer. Update it when you inspect sources,
  make implementation decisions, or answer a project question worth preserving.
- This `AGENTS.md` file is the schema layer. Keep workflow rules here small and
  practical.

## Workflow

When preparing or changing the project:

1. Read `wiki/index.md` first to find relevant notes.
2. Inspect the actual source files or code before making claims.
3. Update the smallest useful wiki page when new durable knowledge is learned.
4. Append one chronological entry to `wiki/log.md` for ingests, decisions, and
   project changes.

## Wiki Conventions

- Use Markdown files with short, descriptive lowercase names.
- Link related pages with relative Markdown links.
- Keep `wiki/index.md` content-oriented: page link plus one-line summary.
- Keep `wiki/log.md` append-only, newest entries at the top.
- Prefer facts observed in the repository over guesses.

## Excel Aggregator Notes

- The target application is not specified yet. Do not invent a UI, database, or
  output format before requirements exist.
- The known input format is legacy Excel `.xls` sample files under
  `Примеры таблиц/`.
- Before implementing parsing, inspect workbook sheets, headers, merged cells,
  encodings, and repeated structures across all samples.
