#!/usr/bin/env python3
"""
Descarga materiales de Google Classroom: PDFs, Drive, Colab, videos.
No navega fuera de la página principal — cada descarga abre una pestaña aparte.
"""
import argparse
import re
import time
from pathlib import Path
from typing import List, Dict, Optional

from playwright.sync_api import sync_playwright, Page, BrowserContext
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError


def log(msg: str) -> None:
    print(msg, flush=True)


def safe_name(value: Optional[str], fallback: str = "sin_nombre") -> str:
    text = (value or "").strip()
    if not text:
        text = fallback
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:180] or fallback


# ── Clasificación de URLs ────────────────────────────────────────────────────

def get_drive_id(url: str) -> Optional[str]:
    for pat in [
        r"drive\.google\.com/file/d/([a-zA-Z0-9_-]+)",
        r"drive\.google\.com/open\?id=([a-zA-Z0-9_-]+)",
        r"drive\.google\.com/.*[?&]id=([a-zA-Z0-9_-]+)",
    ]:
        m = re.search(pat, url)
        if m:
            return m.group(1)
    return None


def get_docs_export(url: str) -> Optional[str]:
    m = re.search(r"docs\.google\.com/(document|spreadsheets|presentation)/d/([a-zA-Z0-9_-]+)", url)
    if not m:
        return None
    fmt = {"document": "pdf", "spreadsheets": "xlsx", "presentation": "pptx"}[m.group(1)]
    return f"https://docs.google.com/{m.group(1)}/d/{m.group(2)}/export?format={fmt}"


def is_colab(url: str) -> bool:
    return "colab.research.google.com" in url


def is_video(url: str) -> bool:
    return any(d in url for d in ["youtube.com", "youtu.be", "vimeo.com"])


# ── Scraping ─────────────────────────────────────────────────────────────────

def extract_links_from_page(page: Page) -> List[Dict]:
    """Extrae solo links de contenido (Drive, Docs, Colab, YouTube) ignorando UI de Classroom."""
    return page.evaluate("""
        () => {
            const CONTENT = /drive\\.google\\.com\\/file|drive\\.google\\.com\\/open|docs\\.google\\.com\\/(document|spreadsheets|presentation)|colab\\.research\\.google\\.com|youtube\\.com\\/watch|youtu\\.be\\/|vimeo\\.com\\/[0-9]/;
            const seen = new Set();
            const out  = [];
            for (const a of document.querySelectorAll('a[href]')) {
                const href = a.href || '';
                if (!CONTENT.test(href)) continue;
                if (seen.has(href)) continue;
                seen.add(href);
                const title = (a.innerText || a.title || a.getAttribute('aria-label') || '')
                    .trim().replace(/\\s+/g, ' ').substring(0, 120);
                out.push({ href, title });
            }
            return out;
        }
    """)


def scrape_links(page: Page) -> List[Dict]:
    log("[INFO] Esperando que la página se estabilice...")
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
    except Exception:
        pass
    time.sleep(3)

    # Abrir todos los acordeones que están cerrados
    collapsed = page.locator('[aria-expanded="false"]').all()
    log(f"[INFO] {len(collapsed)} acordeones cerrados, abriendo...")
    for btn in collapsed:
        try:
            btn.click(timeout=2000)
            time.sleep(0.5)
        except Exception:
            continue

    time.sleep(2)

    # Extraer links de contenido
    all_links = extract_links_from_page(page)

    log(f"\n[INFO] Total links encontrados: {len(all_links)}")
    for it in all_links:
        log(f"  {it['title'] or '—'}: {it['href'][:90]}")
    return all_links


# ── Descarga ─────────────────────────────────────────────────────────────────

def save_shortcut(url: str, target_dir: Path, prefix: str) -> None:
    dest = target_dir / (safe_name(prefix) + ".url")
    dest.write_text(f"[InternetShortcut]\nURL={url}\n", encoding="utf-8")
    log(f"[LINK] {dest.name}")


def download_url(context: BrowserContext, url: str, target_dir: Path, prefix: str) -> bool:
    """Abre una pestaña nueva, descarga el archivo y la cierra."""
    tab = context.new_page()
    try:
        tab.goto(url, wait_until="domcontentloaded", timeout=30000)
        time.sleep(2)

        # Buscar botón de descarga en el visor de Drive
        download_btn = None
        for selector in [
            '[aria-label="Descargar"]',
            '[aria-label="Download"]',
            'div[aria-label="Descargar"]',
            'div[aria-label="Download"]',
            '[data-tooltip="Descargar"]',
            '[data-tooltip="Download"]',
        ]:
            btn = tab.locator(selector).first
            try:
                if btn.is_visible(timeout=2000):
                    download_btn = btn
                    break
            except Exception:
                continue

        if download_btn:
            with tab.expect_download(timeout=40000) as dl_info:
                download_btn.click()
            dl = dl_info.value
            # Manejar confirmación de archivo grande
            try:
                confirm = tab.locator('a:has-text("Descargar de todas formas"), a:has-text("Download anyway")').first
                if confirm.is_visible(timeout=3000):
                    with tab.expect_download(timeout=60000) as dl_info2:
                        confirm.click()
                    dl = dl_info2.value
            except Exception:
                pass
        else:
            # Sin visor — intentar descarga directa
            with tab.expect_download(timeout=40000) as dl_info:
                tab.goto(url, wait_until="domcontentloaded", timeout=20000)
            dl = dl_info.value

        suffix = Path(dl.suggested_filename).suffix or ".bin"
        dest = target_dir / (safe_name(prefix) + suffix)
        dl.save_as(str(dest))
        log(f"[OK] {dest.name}")
        return True

    except Exception as e:
        log(f"[WARN] No se pudo descargar ({url[:60]}): {e}")
        return False
    finally:
        tab.close()


def process(context: BrowserContext, links: List[Dict], target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)

    for i, item in enumerate(links, 1):
        url   = item["href"]
        title = item.get("title") or f"item_{i:04d}"
        prefix = f"{i:04d} - {title}"
        log(f"\n[{i}/{len(links)}] {title[:60]}")

        # Colab → shortcut
        if is_colab(url):
            save_shortcut(url, target_dir, prefix)
            continue

        # Video → shortcut
        if is_video(url):
            save_shortcut(url, target_dir, prefix)
            continue

        # Google Docs/Sheets/Slides → exportar
        export = get_docs_export(url)
        if export:
            download_url(context, export, target_dir, prefix)
            continue

        # Google Drive file → ir al visor y hacer click en descargar
        drive_id = get_drive_id(url)
        if drive_id:
            download_url(context, f"https://drive.google.com/file/d/{drive_id}/view", target_dir, prefix)
            continue

        # Cualquier otro link → intentar descarga directa, si falla shortcut
        if not download_url(context, url, target_dir, prefix):
            save_shortcut(url, target_dir, prefix)


# ── Main ─────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("url", help="URL del classwork (.../w/XXX/t/all)")
    parser.add_argument("--output", type=Path, default=Path("./classroom_descargas"))
    args = parser.parse_args()

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            channel="msedge",
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1440, "height": 900},
            locale="es-ES",
        )
        context.set_default_timeout(20000)
        page = context.new_page()

        page.goto(args.url, wait_until="domcontentloaded", timeout=30000)

        log("\nInicia sesión y asegúrate de estar en el tab 'Trabajo en clases'.")
        input("→ Presiona Enter cuando estés listo...\n")

        links = scrape_links(page)

        if not links:
            log("[ERROR] No se encontraron links.")
            browser.close()
            return 1

        process(context, links, args.output)
        browser.close()

    log(f"\n[INFO] Listo. Archivos en: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
