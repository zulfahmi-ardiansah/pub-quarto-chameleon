"""qlon entry point: route the input to the right pipeline.

A positional `.docx` input runs the Markdown-extraction pipeline
(reverse); anything else (a `.yml` config, or no positional such as
`--test`) runs the Quarto book -> DOCX renderer (render). Each pipeline
owns its own argument parser and dependencies, so the imports are deferred
until a pipeline is chosen.
"""

import sys

# Options that take a separate value; their following token is consumed and must
# not be mistaken for the positional input. Covers both pipelines' value-flags so
# e.g. `--custom template.docx` never routes to the docx path by its value.
_VALUE_FLAGS = frozenset({
    "-o", "--output", "--layout", "--model", "--work-dir",  # reverse
    "--preset", "--custom",                                  # render
})


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
    input_arg = _input_arg(sys.argv[1:])
    if input_arg and input_arg.lower().endswith(".docx"):
        from reverse import main as run  # .docx -> Markdown extraction
    else:
        from render import main as run    # .yml  -> DOCX render (default)
    return run()


if __name__ == "__main__":
    raise SystemExit(main())
