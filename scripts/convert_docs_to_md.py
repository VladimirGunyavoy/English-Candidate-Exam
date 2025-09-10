import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path
import re


def ensure_pandoc() -> None:
    if shutil.which("pandoc") is None:
        print("[ERROR] pandoc не найден в PATH. Установите: https://pandoc.org/installing.html", file=sys.stderr)
        sys.exit(1)


def compute_output_path(source: Path, root: Path, out_root: Path | None) -> Path:
    if out_root is None:
        return source.with_suffix(".md")
    rel = source.resolve().relative_to(root.resolve())
    return (out_root / rel).with_suffix(".md")


def needs_conversion(source: Path, target: Path, overwrite: bool) -> bool:
    if overwrite:
        return True
    if not target.exists():
        return True
    return source.stat().st_mtime > target.stat().st_mtime


def has_pdftotext() -> bool:
    return shutil.which("pdftotext") is not None


def convert_pdf_with_pdftotext(src: Path, dst: Path, dry_run: bool) -> tuple[bool, str]:
    # Конвертируем PDF -> txt, затем сохраняем как .md
    txt_dst = dst.with_suffix(".txt")
    args = [
        "pdftotext",
        "-layout",
        str(src),
        str(txt_dst),
    ]
    if dry_run:
        return True, "DRY-RUN: " + " ".join(args)
    try:
        proc = subprocess.run(args, capture_output=True, text=True)
        if proc.returncode != 0:
            return False, proc.stderr.strip() or proc.stdout.strip()
        # Переименовываем txt в md (простая текстовая разметка)
        try:
            txt_content = txt_dst.read_text(encoding="utf-8", errors="ignore")
        except UnicodeDecodeError:
            txt_content = txt_dst.read_text(encoding="latin-1", errors="ignore")
        dst.write_text(txt_content, encoding="utf-8")
        try:
            txt_dst.unlink(missing_ok=True)
        except Exception:
            pass
        return True, f"OK (pdftotext): {src} -> {dst}"
    except Exception as exc:  # noqa: BLE001
        return False, str(exc)


def convert_file(src: Path, dst: Path, dry_run: bool) -> tuple[bool, str]:
    dst.parent.mkdir(parents=True, exist_ok=True)
    media_dir = dst.parent / "media" / src.stem
    media_dir.mkdir(parents=True, exist_ok=True)

    # Специальная обработка PDF: pandoc не умеет конвертировать ИЗ PDF
    if src.suffix.lower() == ".pdf":
        if not has_pdftotext():
            return False, (
                "pdftotext не найден. Установите Poppler (https://blog.alivate.com.au/poppler-windows/) "
                "и добавьте bin в PATH, либо укажите --no-recurse и сконвертируйте вручную."
            )
        return convert_pdf_with_pdftotext(src, dst, dry_run)

    from_fmt = "docx" if src.suffix.lower() == ".docx" else ""
    args = [
        "pandoc",
        "--from",
        from_fmt or "",
        "--to",
        "gfm",
        "--wrap",
        "none",
        "--markdown-headings",
        "atx",
        "--extract-media",
        str(media_dir),
        "--output",
        str(dst),
        str(src),
    ]

    # очистка пустых значений
    args = [a for a in args if a]

    if dry_run:
        return True, "DRY-RUN: " + " ".join(args)

    try:
        proc = subprocess.run(args, capture_output=True, text=True)
        if proc.returncode != 0:
            return False, proc.stderr.strip() or proc.stdout.strip()
        return True, f"OK: {src} -> {dst}"
    except Exception as exc:  # noqa: BLE001
        return False, str(exc)


def main() -> int:
    parser = argparse.ArgumentParser(description="Рекурсивная конвертация .docx/.pdf в Markdown через pandoc")
    parser.add_argument("--root", default="src_docs", help="Корневой каталог поиска (по умолчанию: src_docs)")
    parser.add_argument("--out", default="", help="Каталог вывода. Если пусто, сохранять в корневую папку 'converted'")
    parser.add_argument("--no-recurse", action="store_true", help="Не сканировать рекурсивно")
    parser.add_argument("--overwrite", action="store_true", help="Перезаписывать существующие .md")
    parser.add_argument("--dry-run", action="store_true", help="Показывать команды, не выполняя их")
    parser.add_argument("--no-clean", action="store_true", help="Отключить авто‑очистку Markdown от хедеров/футеров")

    args = parser.parse_args()

    ensure_pandoc()

    root = Path(args.root).resolve()
    if not root.exists():
        print(f"[INFO] Каталог не найден: {root}")
        return 0

    # Если --out не указан, сохраняем в корень репозитория в подкаталог 'converted'
    if args.out:
        out_root = Path(args.out).resolve()
    else:
        out_root = Path(__file__).resolve().parents[1] / "converted"
    out_root.mkdir(parents=True, exist_ok=True)

    patterns = (".docx", ".pdf")
    if args.no_recurse:
        candidates = [p for p in root.iterdir() if p.is_file() and p.suffix.lower() in patterns]
    else:
        candidates = [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in patterns]

    if not candidates:
        print(f"[INFO] Файлы для конвертации не найдены в {root}")
        return 0

    print(f"Найдено файлов: {len(candidates)}")

    converted = 0
    skipped = 0
    failed = 0

    def clean_markdown_text(text: str) -> str:
        # Удаляем символы разрыва страницы
        text = text.replace("\f", "\n")

        lines = text.splitlines()
        cleaned: list[str] = []
        date_re = re.compile(r"^\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}.*$")
        page_counter_re = re.compile(r"\s+\d+/\d+\s*$")
        only_page_num_re = re.compile(r"^\s*\d{1,3}\s*$")
        service_phrases = (
            "место для печати",
        )

        for line in lines:
            # Удаляем строки-хедеры с датой
            if date_re.match(line):
                continue
            # Удаляем строки, состоящие только из номера страницы
            if only_page_num_re.match(line):
                continue
            # Удаляем типовые служебные строки
            if line.strip().lower() in service_phrases:
                continue
            # Удаляем счётчик страниц в конце строки (оставляя контент строки)
            line = page_counter_re.sub("", line)
            # Трим правые пробелы
            line = line.rstrip()
            cleaned.append(line)

        # Схлопываем более чем одну пустую строку подряд и нормализуем маркеры списков
        result: list[str] = []
        empty = 0
        for line in cleaned:
            # Нормализация маркеров списков в начале строки
            line = re.sub(r"^\s*[•\-–—]\s+", "- ", line)
            line = re.sub(r"^\s*\*\s+", "- ", line)
            if line.strip() == "":
                empty += 1
            else:
                empty = 0
            if empty <= 1:
                result.append(line)
        # Гарантируем H1 у первого непустого заголовка
        for idx, ln in enumerate(result):
            if ln.strip():
                if not ln.lstrip().startswith("# ") and not ln.lstrip().startswith("## "):
                    # Не делаем заголовком, если строка подозрительно длинная (>120) или выглядит как ссылка
                    if len(ln.strip()) <= 120 and not re.match(r"^https?://", ln.strip(), re.I):
                        result[idx] = "# " + ln.strip()
                break
        return "\n".join(result).strip() + "\n"

    for src in candidates:
        dst = compute_output_path(src, root, out_root)
        dst.parent.mkdir(parents=True, exist_ok=True)

        if not needs_conversion(src, dst, args.overwrite):
            skipped += 1
            continue

        ok, msg = convert_file(src, dst, args.dry_run)
        if ok:
            # Постобработка Markdown (если не отключена)
            if not args.no_clean and dst.suffix.lower() == ".md" and dst.exists():
                try:
                    original = dst.read_text(encoding="utf-8")
                    cleaned = clean_markdown_text(original)
                    if cleaned and cleaned != original:
                        dst.write_text(cleaned, encoding="utf-8")
                except Exception:
                    # Не прерываем общий процесс при ошибке очистки конкретного файла
                    pass
            converted += 1
        else:
            failed += 1
        print(msg)

    print(f"Готово. Конвертировано: {converted}, пропущено: {skipped}, ошибок: {failed}")
    if failed and not has_pdftotext():
        print(
            "Подсказка: установите 'pdftotext' (Poppler) и добавьте в PATH. "
            "Windows: скачайте архив Poppler, распакуйте и добавьте путь к \bin в PATH."
        )
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())




