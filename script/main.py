"""qlon entry point: route the input to the right pipeline.

A positional `.docx`/`.md` input runs the Markdown-extraction pipeline
(reverse); anything else (a `.yml` config, or no positional such as
`--test`) runs the Quarto book -> DOCX renderer (render). Each pipeline
owns its own argument parser and dependencies, so the imports are deferred
until a pipeline is chosen.

`qlon -h` / `qlon --help` (no input) prints the combined overview below.
Once an input is given, the chosen pipeline's own `--help` takes over
(e.g. `qlon in.docx --help`, `qlon config.yml --help`).
"""

import sys

# Options that take a separate value; their following token is consumed and must
# not be mistaken for the positional input. Covers both pipelines' value-flags so
# e.g. `--custom template.docx` never routes to the reverse path by its value.
_VALUE_FLAGS = frozenset({
    "--layout", "--model-name", "--model-endpoint", "--model-key", "--skip-toc",  # reverse
    "--preset", "--custom",                                                       # render
    "--output", "-o",                                                             # both
})

# Combined help. Hand-maintained: keep in sync with reverse.parse_args and
# render.main's argparse when their flags change.
_HELP = """Qlon - Quarto Chameleon

Usage:
  qlon <config.yml> [options]   Quarto book -> DOCX (render)
  qlon --test                   render the bundled example.yml
  qlon <input.docx> [options]   docx -> clean Markdown pages (reverse)

Routing:
  anything other than a .docx positional (a .yml config, or --test with no
  input) runs render; a .docx positional input runs the reverse (ingest) pipeline.

render options  (yml -> DOCX, output to the current directory):
  --test               use the bundled example.yml
  --preset NAME        built-in template in template/ (e.g. basic)
  --custom PATH        path to a custom .docx reference template
  --keep-work          keep the per-run render/<uuid> workspace
  --output, -o DIR     folder to write the .docx/Image output into (default: cwd)
  --no-images          skip exporting the Image/ folder; copy only the .docx

reverse options  (docx -> Markdown pages, output to ./<input name>/):
  --use-llm            run the LLM steps (segment / clean / structure / describe)
  --layout flat|fuma   output layout (default: flat; 'fuma' requires the LLM)
  --model-name ID      model id ($LLM_MODEL or the built-in default)
  --model-endpoint URL LLM API base URL ($LLM_BASE_URL or the built-in default)
  --model-key KEY      LLM API key ($LLM_API_KEY / $OPENROUTER_API_KEY)
  --no-merge           skip the final merge step
  --skip-toc HEADING   drop the page whose first heading is HEADING (the TOC page)
  --allow-reorder      apply the LLM's suggested section reorder
  --keep-work          keep the per-run render/<uuid> scratch folder
  --output, -o DIR     base folder for the <input-stem>/ output (default: cwd)

Run `qlon <input> --help` for a pipeline's own argparse help.
"""


def _input_arg(argv: list[str]) -> str | None:
    """First positional argument, skipping options and their consumed values."""
    skip = False
    for arg in argv:
        if skip:
            skip = False
            continue
        if arg.startswith("-"):
            skip = "=" not in arg and arg in _VALUE_FLAGS
            continue
        return arg
    return None


def main() -> int:
    """Dispatch on the input argument's extension: .docx -> extraction, else render."""
    argv = sys.argv[1:]
    input_arg = _input_arg(argv)

    # No input chosen yet + a help flag -> show the combined overview. Once an
    # input is present, routing forwards --help to the selected pipeline instead.
    if input_arg is None and ("-h" in argv or "--help" in argv):
        print(_HELP)
        return 0

    if input_arg and input_arg.lower().endswith(".docx"):
        from reverse import main as run  # .docx -> Markdown extraction
    else:
        from render import main as run    # .yml  -> DOCX render (default)
    return run()


if __name__ == "__main__":
    raise SystemExit(main())
