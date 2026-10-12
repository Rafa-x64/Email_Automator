#!/usr/bin/env python3
"""WhatsApp semiautomático con enlaces wa.me. Ayuda: python wa.py --help

Abre cada chat con el mensaje ya escrito; tú revisas y pulsas enviar en WhatsApp.
Reutiliza la validación de teléfonos, el historial y la trazabilidad de main.py.
"""
from __future__ import annotations

import argparse
import html
import sys
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import main as core
from checks import normalize_ve_phone

WA_DEFAULTS = {
    "template": "templates/whatsapp.txt", "name_column": "Empresa", "mobile_only": True,
    "only_without_email": False, "max_per_run": 0,
    "history": "logs/whatsapp_enviados.csv", "links_file": "logs/whatsapp_links.html",
}
log = core.log


def build_url(phone: str, text: str) -> str:
    return f"https://wa.me/{phone.lstrip('+')}?text={quote(text, safe='')}"


def write_links(path: Path, items: list[tuple[int, str, str, str]]) -> None:
    h = lambda s: html.escape(s, quote=True)  # noqa: E731
    rows = "\n".join(
        f'<tr><td>{n}</td><td>{h(name)}</td><td>{h(phone)}</td>'
        f'<td><a href="{h(url)}" target="_blank" rel="noopener">Abrir chat</a></td></tr>'
        for n, name, phone, url in items)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        '<!DOCTYPE html><html lang="es"><meta charset="utf-8"><title>Enlaces WhatsApp</title>'
        '<style>body{font-family:system-ui,sans-serif;max-width:720px;margin:2rem auto;padding:0 1rem}'
        'table{border-collapse:collapse;width:100%}td,th{border-bottom:1px solid #8884;padding:.5rem;'
        'text-align:left}a{font-weight:600}</style>'
        f'<h1>Chats de WhatsApp ({len(items)})</h1>'
        '<p>Pulsa "Abrir chat", revisa el mensaje y presiona enviar.</p>'
        f'<table><tr><th>#</th><th>Contacto</th><th>Teléfono</th><th></th></tr>{rows}</table></html>',
        encoding="utf-8")


def run(args: argparse.Namespace) -> int:
    core.DEFAULTS["whatsapp"] = WA_DEFAULTS
    cfg_path = Path(args.config).resolve()
    base = cfg_path.parent
    cfg = core.load_config(cfg_path)
    core.setup_logging(base / cfg["files"]["log"])
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S") + "-wa"
    log.info("=== Ejecución WhatsApp %s ===", run_id)

    ex, val, wa = cfg["excel"], cfg["validation"], cfg["whatsapp"]
    pcol, ncol = val["phone_column"], wa["name_column"]
    columns, rows = core.load_contacts(base / ex["file"], ex["sheet"], ex["email_column"])
    if pcol not in columns:
        raise core.ConfigError(f"Falta la columna de teléfono '{pcol}'. Columnas: {columns}")
    template = core.read_text(base / wa["template"], "la plantilla de WhatsApp")
    missing = core.placeholders(template) - set(columns)
    if missing:
        raise core.ConfigError(f"La plantilla usa columnas que no existen en el Excel: {sorted(missing)}")

    hist_path = base / wa["history"]
    done = set() if args.ignore_history else core.load_sent(hist_path)
    trace = core.Tracer(base / cfg["files"]["trace"], run_id)
    queue: list[tuple[dict[str, str], str, str]] = []
    seen: set[str] = set()
    rejected = skipped = 0
    try:
        for n, row in enumerate(rows, start=2):
            row[core.K_ROW] = str(n)
            email, raw = row[ex["email_column"]], row[pcol]
            ctx = {"email": email, "fila": str(n), "tel_orig": raw}
            if wa["only_without_email"] and email:
                continue
            if not raw:
                trace.add("whatsapp", "omitido", detalle="sin teléfono", **ctx)
                continue
            try:
                phone = normalize_ve_phone(raw, allow_landline=not wa["mobile_only"])
            except ValueError as e:
                rejected += 1
                log.error("Fila %d omitida: teléfono inválido '%s': %s", n, raw, e)
                trace.add("whatsapp", "rechazado", detalle=f"teléfono inválido: {e}", **ctx)
                continue
            if phone in seen or phone.lower() in done:
                skipped += 1
                trace.add("whatsapp", "omitido", tel_norm=phone, detalle="duplicado o ya enviado", **ctx)
                continue
            seen.add(phone)
            queue.append((row, phone, row.get(ncol, "")))
        limit = args.limit if args.limit is not None else wa["max_per_run"]
        if limit:
            queue = queue[:limit]
        log.info("Filas: %d | chats a abrir: %d | ya enviados/duplicados: %d | rechazados: %d",
                 len(rows), len(queue), skipped, rejected)
        if not queue:
            log.info("Nada que enviar por WhatsApp.")
            return 2 if rejected else 0

        items = []
        for i, (row, phone, name) in enumerate(queue, 1):
            text = core.render(template, row)
            url = build_url(phone, text)
            if len(url) > 1900:
                log.warning("Fila %s: el mensaje es muy largo y el enlace puede fallar", row[core.K_ROW])
            items.append((i, name, phone, url))

        if args.dry_run:
            for i, name, phone, url in items:
                log.info("[SIMULACIÓN %d/%d] %s %s\n    %s", i, len(items), name, phone, url)
                trace.add("whatsapp", "simulacion", tel_norm=phone, fila=queue[i - 1][0][core.K_ROW])
            return 2 if rejected else 0

        if args.list:
            out = base / wa["links_file"]
            write_links(out, items)
            log.info("Lista generada: %s", out)
            trace.add("whatsapp", "listado", detalle=f"{len(items)} enlaces en {out.name}")
            webbrowser.open(out.as_uri())
            return 2 if rejected else 0

        history = core.SentLog(hist_path)
        sent = 0
        print("\nAsegúrate de tener WhatsApp Web (o la app de escritorio) con la sesión iniciada.")
        print("En cada chat: revisa el mensaje, pulsa enviar en WhatsApp y vuelve aquí.\n")
        try:
            for i, (name, phone, url) in enumerate([(x[1], x[2], x[3]) for x in items], 1):
                row = queue[i - 1][0]
                ctx = {"email": row[ex["email_column"]], "fila": row[core.K_ROW],
                       "tel_orig": row[core.K_PHONE0] if core.K_PHONE0 in row else row[pcol],
                       "tel_norm": phone}
                webbrowser.open_new_tab(url)
                trace.add("whatsapp", "abierto", detalle="chat abierto en el navegador", **ctx)
                try:
                    ans = input(f"[{i}/{len(items)}] {name} {phone}  ->  Enter = ya lo envié | s = saltar | q = salir: ")
                except EOFError:
                    ans = "q"
                ans = ans.strip().lower()
                if ans == "q":
                    log.warning("Saliste. Al volver a ejecutar continúas donde quedaste.")
                    trace.add("whatsapp", "interrumpido", detalle="el usuario salió", **ctx)
                    break
                if ans == "s":
                    log.info("%s saltado (volverá a aparecer en la próxima ejecución)", phone)
                    trace.add("whatsapp", "omitido", detalle="saltado por el usuario", **ctx)
                    continue
                sent += 1
                history.add(phone, "ok", "confirmado por el usuario")
                trace.add("whatsapp", "enviado", detalle="confirmado manualmente por el usuario", **ctx)
                log.info("%s registrado como enviado", phone)
        except KeyboardInterrupt:
            log.warning("Interrumpido (Ctrl+C). El historial conserva lo ya confirmado.")
            trace.add("whatsapp", "interrumpido", detalle="Ctrl+C")
        finally:
            history.close()
        log.info("Fin. Confirmados como enviados: %d de %d", sent, len(items))
        return 2 if rejected else 0
    finally:
        trace.close()


def main() -> None:
    p = argparse.ArgumentParser(description="WhatsApp semiautomático con enlaces wa.me")
    p.add_argument("-c", "--config", default=str(Path(__file__).parent / "config.toml"))
    p.add_argument("--dry-run", action="store_true", help="muestra los enlaces sin abrir nada")
    p.add_argument("--list", action="store_true", help="genera una página HTML con todos los enlaces")
    p.add_argument("--limit", type=int, help="máximo de chats en esta ejecución")
    p.add_argument("--ignore-history", action="store_true", help="incluye teléfonos ya enviados")
    try:
        sys.exit(run(p.parse_args()))
    except core.ConfigError as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        sys.exit(130)
    except Exception:  # noqa: BLE001
        detail = traceback.format_exc()
        if core.log.handlers:
            core.log.critical("ERROR INESPERADO:\n%s", detail)
        else:
            print(detail, file=sys.stderr)
        print("[ERROR FATAL] Fallo inesperado. Revisa logs/mailer.log y envíame el traceback.", file=sys.stderr)
        sys.exit(3)


if __name__ == "__main__":
    main()
