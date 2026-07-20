"""LLM prompt constants for the docx -> Markdown ingestion pipeline.

Imported by script/reverse.py. Kept out of the pipeline module so the
wording can be tuned without touching the extraction logic.
"""

# Cleans a single Pandoc-exported section fragment and appends a one-line summary
# after a ``---SUMMARY---`` delimiter (see reverse._SUMMARY_DELIM).
CLEAN_SYSTEM_PROMPT = (
    "You clean up Markdown that was auto-exported from a Word document by Pandoc.\n"
    "Given ONE section fragment, return proper, clean Markdown.\n"
    "Rules:\n"
    "- Remove Pandoc artifacts: footnote refs like [^1] AND footnote definition "
    "blocks like '[^1]: text...' (Pandoc dumps them at the document end -- drop "
    "the whole block including its continuation paragraphs), attribute blocks like "
    "{#id .class}, image size attributes like {width=\"6.5in\" height=\"3.78in\"}, "
    "inline spans like [X]{.smallcaps} (keep the visible text), "
    "escaped pipes \\|, and stray *** emphasis noise.\n"
    "- Unwrap hard line breaks: the source wraps paragraphs at a fixed column width. "
    "Join those wrapped lines so each paragraph (and each list item) is a single "
    "continuous line. Keep blank lines BETWEEN paragraphs/list items. Do NOT unwrap "
    "inside fenced code blocks or tables.\n"
    "- Fix headings, lists, tables, and emphasis so they are valid Markdown.\n"
    "- Keep the heading level of the section exactly as given. If the fragment "
    "starts with a heading line, that heading MUST stay as the first line of "
    "your output — never drop or reword it.\n"
    "- Lightly polish prose for grammar and clarity, but PRESERVE the original "
    "meaning and all technical terms. Do not summarize or drop content.\n"
    "- NEVER invent, add, or expand content that is not in the input. If the "
    "fragment has little or no body text, return it as-is -- do not write an "
    "introduction, examples, or explanations for it.\n"
    "- Preserve image and link references as-is (minus the attribute blocks).\n"
    "- Return the cleaned Markdown, then a line containing only '---SUMMARY---', "
    "then 1-2 plain sentences (same language as the content) stating the content "
    "directly as a noun phrase, e.g. 'A comprehensive list of bibliographic "
    "references and instructions on footnote usage.' Never open with 'This "
    "section', 'This page', 'Bagian ini', or similar. "
    "No other commentary, no code fences around anything."
)

# Reviews the whole-document outline and returns corrected heading levels/text
# plus an optional (never auto-applied) reorder suggestion, as a JSON object.
STRUCTURE_PROMPT = (
    "You are given a JSON outline of a document's sections, in source order. Each "
    "entry has: index, level (0 = preamble with no heading, 1..6 = heading depth), "
    "heading, file.\n"
    "Return a JSON object with:\n"
    '  "title": a concise title for the whole document.\n'
    '  "changes": a list of ONLY the entries that need correction, each as '
    '{"index": i, "level": l, "heading": h}, so the document gets a clean, consistent '
    "heading hierarchy (proper H1/H2/H3 nesting) and tidy heading wording/"
    "capitalization. Strip manual numbering prefixes from headings "
    '("1. Overview" -> "Overview"). '
    "Entries that are already fine must NOT appear in the list. "
    "Never invent new index values. Keep level 0 entries at level 0.\n"
    '  "suggested_order": null if the source order is fine; otherwise a list of index '
    "values in a better reading order.\n"
    '  "reorder_reason": short string, or null.\n'
    "Return ONLY the JSON object, no commentary, no code fences."
)

# Condenses each page's per-section summaries into one page description, returned
# as a JSON object {"descriptions": [...]} with one entry per page, in order.
DESCRIBE_PROMPT = (
    "You are given a JSON list of documentation pages. Each entry has: title, and "
    "summaries (one line per section of that page, in order).\n"
    'Return a JSON object {"descriptions": [...]}: exactly one string per page, '
    "same order and count as the input. Each is 1-3 sentences (never more) that "
    "represent the whole page, written in the same language as the summaries. "
    "State the content directly as a noun phrase, e.g. 'A comprehensive list of "
    "bibliographic references and instructions on footnote usage.' Never open "
    "with 'This section', 'This page', 'This sub', 'Halaman ini', or similar. "
    "Return ONLY the JSON object, no commentary, no code fences."
)

# Locates section boundaries in a document with few or no headings. Returns only
# headings + verbatim text anchors (indices/prose are never echoed) so the caller
# can splice headings deterministically without the model rewriting the body.
SEGMENT_PROMPT = (
    "You are given the full text of a document that has few or no headings. "
    "Identify where its natural sections begin.\n"
    'Return ONLY a JSON object {"sections": [{"heading": h, "anchor": a}, ...]}:\n'
    "- heading: a short section title, in the same language as the content.\n"
    "- anchor: the opening of that section, copied EXACTLY and VERBATIM from the "
    "text (including any typos and punctuation) — never paraphrased — starting at "
    "the section's first word and long enough (roughly 8-15 words) to appear only "
    "once in the document, so it can be located again.\n"
    "Do not invent content. Do not return prose or the document body. Do not add "
    "code fences. If the document has no meaningful sections, return "
    '{"sections": []}.'
)
