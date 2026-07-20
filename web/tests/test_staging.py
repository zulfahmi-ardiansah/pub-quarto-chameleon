"""Unit tests for web.staging (stage_zip + stage_content)."""

import io
import sys
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from staging import StagingError, stage_content, stage_zip  # noqa: E402


class FakeUpload:
    """Minimal stand-in for a Werkzeug FileStorage (filename + .save)."""

    def __init__(self, filename: str, data: bytes):
        self.filename = filename
        self._data = data

    def save(self, dst) -> None:
        Path(dst).write_bytes(self._data)


def make_zip(path: Path, entries: dict[str, bytes | str]) -> Path:
    """Write a zip at *path* from {arcname: content}. Bytes stay bytes; str -> utf-8."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries.items():
            data = content if isinstance(content, bytes) else content.encode("utf-8")
            zf.writestr(name, data)
    return path


PNG = b"\x89PNG\r\n\x1a\n"  # enough bytes to stand in for an image


@pytest.fixture
def content_dir(tmp_path: Path) -> Path:
    d = tmp_path / "job" / "content"
    d.mkdir(parents=True)
    return d


def test_flat_zip_stages_md_and_image(tmp_path, content_dir):
    arc = make_zip(tmp_path / "in.zip", {
        "chapter.md": "# Ch\n\n![diagram](images/pic.png)\n",
        "images/pic.png": PNG,
    })
    stage_zip(arc, content_dir)

    assert (content_dir / "chapter.md").is_file()
    assert (content_dir / "images" / "pic.png").read_bytes() == PNG


def test_single_wrapper_folder_is_stripped(tmp_path, content_dir):
    arc = make_zip(tmp_path / "in.zip", {
        "project/ch1.md": "# One\n\n![p](images/a.png)\n",
        "project/images/a.png": PNG,
    })
    stage_zip(arc, content_dir)

    assert (content_dir / "ch1.md").is_file()
    assert (content_dir / "images" / "a.png").read_bytes() == PNG


def test_nested_chapters_all_found(tmp_path, content_dir):
    arc = make_zip(tmp_path / "in.zip", {
        "a/01-intro.md": "# Intro\n",
        "b/02-body.qmd": "# Body\n",
    })
    stage_zip(arc, content_dir)

    staged = sorted(p.name for p in content_dir.glob("*.md")) + \
        sorted(p.name for p in content_dir.glob("*.qmd"))
    assert "01-intro.md" in staged
    assert "02-body.qmd" in staged


def test_basename_collision_disambiguated(tmp_path, content_dir):
    arc = make_zip(tmp_path / "in.zip", {
        "a/notes.md": "# A notes\n",
        "b/notes.md": "# B notes\n",
    })
    stage_zip(arc, content_dir)

    md = sorted(content_dir.glob("*.md"))
    assert len(md) == 2, "both colliding chapters must survive under distinct names"
    texts = sorted(p.read_text(encoding="utf-8") for p in md)
    assert texts == ["# A notes\n", "# B notes\n"]


def test_ref_climbing_out_is_rewritten_and_copied(tmp_path, content_dir):
    # md sits in a subfolder and points up-and-over to a sibling assets folder.
    arc = make_zip(tmp_path / "in.zip", {
        "docs/ch.md": "# Ch\n\n![p](../assets/x.png)\n",
        "assets/x.png": PNG,
    })
    stage_zip(arc, content_dir)

    # Image must land somewhere inside content/, and the md ref must resolve to it.
    staged_md = (content_dir / "ch.md").read_text(encoding="utf-8")
    import re
    ref = re.search(r"!\[.*?\]\(([^)\s]+)", staged_md).group(1)
    assert not ref.startswith(".."), "ref must be rewritten to stay inside content/"
    assert (content_dir / ref).is_file()
    assert (content_dir / ref).read_bytes() == PNG


def test_image_ref_escaping_staging_is_not_copied(tmp_path, content_dir):
    # A secret outside the extracted project must never be read/exfiltrated via a
    # climbing image ref like ../../../secret.txt.
    secret = tmp_path / "secret.txt"
    secret.write_bytes(b"TOP SECRET")
    arc = make_zip(tmp_path / "in.zip", {
        "ch.md": "# Ch\n\n![x](../../../secret.txt)\n",
    })
    stage_zip(arc, content_dir)

    copied = [p for p in content_dir.rglob("*") if p.is_file() and p.read_bytes() == b"TOP SECRET"]
    assert copied == [], "escaping ref must not copy files from outside the project"


def test_stage_content_pasted_written_as_single_chapter(content_dir):
    stage_content([], "# Pasted\n\nbody", content_dir)
    assert (content_dir / "document.md").read_text(encoding="utf-8") == "# Pasted\n\nbody"


def test_stage_content_loose_uploads(content_dir):
    files = [FakeUpload("01-intro.md", b"# Intro"), FakeUpload("02-body.qmd", b"# Body")]
    stage_content(files, "", content_dir)
    assert (content_dir / "01-intro.md").is_file()
    assert (content_dir / "02-body.qmd").is_file()


def test_stage_content_empty_rejected(content_dir):
    with pytest.raises(StagingError):
        stage_content([], "", content_dir)


def test_stage_content_bad_suffix_rejected(content_dir):
    with pytest.raises(StagingError):
        stage_content([FakeUpload("notes.txt", b"x")], "", content_dir)


def test_stage_content_zip_mixed_with_md_rejected(tmp_path, content_dir):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("ch.md", "# c")
    files = [FakeUpload("proj.zip", buf.getvalue()), FakeUpload("extra.md", b"# x")]
    with pytest.raises(StagingError):
        stage_content(files, "", content_dir)


def test_stage_content_single_zip_unpacked(tmp_path, content_dir):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("project/ch1.md", "# One\n\n![p](images/a.png)\n")
        zf.writestr("project/images/a.png", PNG)
    stage_content([FakeUpload("proj.zip", buf.getvalue())], "", content_dir)
    assert (content_dir / "ch1.md").is_file()
    assert (content_dir / "images" / "a.png").read_bytes() == PNG


def test_no_markdown_rejected(tmp_path, content_dir):
    arc = make_zip(tmp_path / "in.zip", {"images/only.png": PNG, "readme.txt": "hi"})
    with pytest.raises(StagingError):
        stage_zip(arc, content_dir)


def test_zip_slip_absolute_path_rejected(tmp_path, content_dir):
    arc = tmp_path / "evil.zip"
    with zipfile.ZipFile(arc, "w") as zf:
        zf.writestr("../../evil.md", "# pwn\n")
    with pytest.raises(StagingError):
        stage_zip(arc, content_dir)
    assert not (content_dir.parent.parent / "evil.md").exists()


def test_not_a_zip_rejected(tmp_path, content_dir):
    bogus = tmp_path / "in.zip"
    bogus.write_bytes(b"not a zip at all")
    with pytest.raises(StagingError):
        stage_zip(bogus, content_dir)


def test_zip_bomb_total_size_rejected(tmp_path, content_dir):
    # One highly-compressible member whose uncompressed size blows the total cap.
    arc = tmp_path / "bomb.zip"
    with zipfile.ZipFile(arc, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("doc.md", "# t\n")
        zf.writestr("big.png", b"\0" * (300 * 1024 * 1024))
    with pytest.raises(StagingError):
        stage_zip(arc, content_dir)
