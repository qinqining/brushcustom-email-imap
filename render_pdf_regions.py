import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageEnhance, ImageFilter, ImageOps


PROJECT_DIR = Path(r"D:\brushcustom-email-imap")


# Normalized crop boxes: left, top, right, bottom.
# These defaults fit landscape engineering drawings like the NEU-PMC and Motion PDFs.
DEFAULT_REGIONS = {
    "title_block_zoom": [0.55, 0.84, 0.99, 0.99],
    "notes_zoom": [0.02, 0.88, 0.46, 0.99],
    "main_dimensions_zoom": [0.03, 0.08, 0.64, 0.39],
    "section_views_zoom": [0.02, 0.36, 0.70, 0.84],
    "right_detail_zoom": [0.60, 0.10, 0.98, 0.76],
}


COMMON_POPPLER_PATHS = [
    Path.home()
    / ".cache"
    / "codex-runtimes"
    / "codex-primary-runtime"
    / "dependencies"
    / "native"
    / "poppler"
    / "Library"
    / "bin"
    / "pdftoppm.exe",
    Path(r"C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin\pdftoppm.exe"),
]


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def safe_stem(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value)
    value = re.sub(r"\s+", "_", value).strip("._ ")
    return value or "pdf"


def resolve_pdftoppm(explicit: str | None) -> str:
    candidates = []
    if explicit:
        candidates.append(Path(explicit))
    env_path = os.environ.get("PDFTOPPM_PATH")
    if env_path:
        candidates.append(Path(env_path))

    candidates.extend(COMMON_POPPLER_PATHS)

    found = shutil.which("pdftoppm.exe") or shutil.which("pdftoppm")
    if found:
        candidates.append(Path(found))

    for item in candidates:
        try:
            if item.exists() and item.suffix.lower() != ".cmd":
                return str(item)
        except OSError:
            continue

    raise SystemExit(
        "Could not find pdftoppm. Install Poppler or pass --pdftoppm "
        "C:\\path\\to\\pdftoppm.exe"
    )


def page_count(pdf_path: Path) -> int:
    try:
        from pypdf import PdfReader

        return len(PdfReader(str(pdf_path)).pages)
    except Exception:
        return 1


def parse_pages(value: str, total_pages: int) -> list[int]:
    if value.lower() == "all":
        return list(range(1, total_pages + 1))

    pages = set()
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = int(start_s)
            end = int(end_s)
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))

    result = sorted(page for page in pages if 1 <= page <= total_pages)
    if not result:
        raise SystemExit(f"No valid pages selected from {value!r}; PDF has {total_pages} page(s).")
    return result


def load_regions(path: str | None) -> dict[str, list[float]]:
    regions = dict(DEFAULT_REGIONS)
    if not path:
        return regions

    config = json.loads(Path(path).read_text(encoding="utf-8"))
    custom = config.get("regions", config)
    for name, bounds in custom.items():
        if isinstance(bounds, dict):
            box = [bounds["left"], bounds["top"], bounds["right"], bounds["bottom"]]
        else:
            box = list(bounds)
        if len(box) != 4:
            raise SystemExit(f"Region {name!r} must have four values: left, top, right, bottom.")
        regions[name] = [float(item) for item in box]
    return regions


def filter_regions(regions: dict[str, list[float]], selected: str) -> dict[str, list[float]]:
    if selected.lower() == "all":
        return regions
    names = [item.strip() for item in selected.split(",") if item.strip()]
    missing = [name for name in names if name not in regions]
    if missing:
        raise SystemExit(f"Unknown region(s): {', '.join(missing)}")
    return {name: regions[name] for name in names}


def render_page(pdftoppm: str, pdf_path: Path, page: int, dpi: int, temp_dir: Path) -> Path:
    prefix = temp_dir / f"page_{page:03d}"
    command = [
        pdftoppm,
        "-png",
        "-r",
        str(dpi),
        "-f",
        str(page),
        "-l",
        str(page),
        "-singlefile",
        str(pdf_path),
        str(prefix),
    ]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "pdftoppm failed")

    expected = prefix.with_suffix(".png")
    if expected.exists():
        return expected

    matches = sorted(temp_dir.glob(f"{prefix.name}*.png"))
    if not matches:
        raise RuntimeError(f"pdftoppm succeeded but no PNG was found for page {page}.")
    return matches[0]


def clamp_box(bounds: list[float], width: int, height: int) -> tuple[int, int, int, int]:
    left, top, right, bottom = bounds
    if not (0 <= left < right <= 1 and 0 <= top < bottom <= 1):
        raise ValueError(f"Invalid normalized bounds: {bounds}")
    return (
        max(0, min(width - 1, round(left * width))),
        max(0, min(height - 1, round(top * height))),
        max(1, min(width, round(right * width))),
        max(1, min(height, round(bottom * height))),
    )


def enhance_crop(image: Image.Image, scale: float, contrast: float, sharpen: bool) -> Image.Image:
    image = image.convert("RGB")
    image = ImageOps.autocontrast(image)
    if contrast and contrast != 1:
        image = ImageEnhance.Contrast(image).enhance(contrast)
    if sharpen:
        image = image.filter(ImageFilter.SHARPEN)
    if scale and scale != 1:
        resampling = getattr(getattr(Image, "Resampling", Image), "LANCZOS")
        image = image.resize((round(image.width * scale), round(image.height * scale)), resampling)
    return image


def process_pdf(
    pdf_path: Path,
    out_dir: Path | None,
    pdftoppm: str,
    pages_value: str,
    dpi: int,
    scale: float,
    contrast: float,
    sharpen: bool,
    regions: dict[str, list[float]],
) -> dict:
    pdf_path = pdf_path.resolve()
    if not pdf_path.exists():
        raise SystemExit(f"PDF not found: {pdf_path}")

    target_dir = out_dir or (pdf_path.parent / "pdf_regions" / safe_stem(pdf_path.stem))
    target_dir.mkdir(parents=True, exist_ok=True)

    total_pages = page_count(pdf_path)
    pages = parse_pages(pages_value, total_pages)
    manifest = {
        "source_pdf": str(pdf_path),
        "output_dir": str(target_dir.resolve()),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "renderer": pdftoppm,
        "dpi": dpi,
        "scale": scale,
        "contrast": contrast,
        "sharpen": sharpen,
        "regions": regions,
        "pages": [],
    }

    with tempfile.TemporaryDirectory(prefix="pdf_render_", dir=str(target_dir)) as temp_name:
        temp_dir = Path(temp_name)
        for page in pages:
            rendered = render_page(pdftoppm, pdf_path, page, dpi, temp_dir)
            source_image = Image.open(rendered)
            width, height = source_image.size

            page_record = {
                "page": page,
                "rendered_size": {"width": width, "height": height},
                "outputs": [],
            }

            full_name = f"page_{page:03d}_full_page.png"
            full_target = target_dir / full_name
            source_image.convert("RGB").save(full_target)
            page_record["outputs"].append(
                {
                    "name": "full_page",
                    "path": str(full_target),
                    "pixel_box": [0, 0, width, height],
                    "normalized_box": [0, 0, 1, 1],
                }
            )

            for name, bounds in regions.items():
                box = clamp_box(bounds, width, height)
                crop = source_image.crop(box)
                crop = enhance_crop(crop, scale=scale, contrast=contrast, sharpen=sharpen)
                target = target_dir / f"page_{page:03d}_{name}.png"
                crop.save(target)
                page_record["outputs"].append(
                    {
                        "name": name,
                        "path": str(target),
                        "pixel_box": list(box),
                        "normalized_box": bounds,
                    }
                )

            manifest["pages"].append(page_record)

    manifest_path = target_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    readme_path = target_dir / "README.txt"
    lines = [
        "PDF region render output",
        f"Source PDF: {pdf_path}",
        f"Created at: {manifest['created_at']}",
        f"DPI: {dpi}",
        f"Scale: {scale}",
        "",
        "Generated files:",
    ]
    for page in manifest["pages"]:
        for output in page["outputs"]:
            lines.append(f"- {output['name']}: {output['path']}")
    readme_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    manifest["manifest_path"] = str(manifest_path)
    manifest["readme_path"] = str(readme_path)
    return manifest


def discover_pdfs(pdf: str | None, customer_dir: str | None) -> list[Path]:
    pdfs = []
    if pdf:
        pdfs.append(Path(pdf))
    if customer_dir:
        root = Path(customer_dir)
        pdfs.extend(sorted(root.glob("*.pdf")))
    unique = []
    seen = set()
    for item in pdfs:
        resolved = str(item.resolve())
        if resolved not in seen:
            unique.append(item)
            seen.add(resolved)
    if not unique:
        raise SystemExit("Pass --pdf PDF_PATH or --customer-dir ATTACHMENT_DIR.")
    return unique


def main() -> None:
    configure_stdout()
    parser = argparse.ArgumentParser(
        description="Render local PDF drawings to full-page and zoomed region PNGs for inquiry review."
    )
    parser.add_argument("--pdf", help="Single PDF path to render.")
    parser.add_argument("--customer-dir", help="Attachment/customer folder; all top-level PDFs are processed.")
    parser.add_argument("--out-dir", help="Output folder. Omit to use <pdf parent>\\pdf_regions\\<pdf stem>.")
    parser.add_argument("--pages", default="all", help="Pages to render: all, 1, 1,3, or 1-2.")
    parser.add_argument("--dpi", type=int, default=350, help="Render DPI. 300-400 is usually good for drawings.")
    parser.add_argument("--scale", type=float, default=3.0, help="Zoom scale for cropped regions.")
    parser.add_argument("--contrast", type=float, default=1.25, help="Contrast multiplier for cropped regions.")
    parser.add_argument("--no-sharpen", action="store_true", help="Disable light sharpening on cropped regions.")
    parser.add_argument("--regions-json", help="Optional JSON file with custom normalized crop boxes.")
    parser.add_argument("--regions", default="all", help="Comma-separated region names, or all.")
    parser.add_argument("--pdftoppm", help="Explicit path to pdftoppm.exe.")
    args = parser.parse_args()

    pdftoppm = resolve_pdftoppm(args.pdftoppm)
    regions = filter_regions(load_regions(args.regions_json), args.regions)
    pdfs = discover_pdfs(args.pdf, args.customer_dir)

    all_results = []
    for pdf_path in pdfs:
        out_dir = Path(args.out_dir) if args.out_dir and len(pdfs) == 1 else None
        result = process_pdf(
            pdf_path=pdf_path,
            out_dir=out_dir,
            pdftoppm=pdftoppm,
            pages_value=args.pages,
            dpi=args.dpi,
            scale=args.scale,
            contrast=args.contrast,
            sharpen=not args.no_sharpen,
            regions=regions,
        )
        all_results.append(result)

    print(json.dumps(all_results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
