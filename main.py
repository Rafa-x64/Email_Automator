#!/usr/bin/env python3
"""Envío masivo de correos desde Excel/CSV. Ayuda: python main.py --help"""
from __future__ import annotations

import argparse
import csv
import html
import imaplib
import logging
import mimetypes
import os
import random
import re
import smtplib
import ssl
import sys
import time
import tomllib
import traceback
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path
from typing import Any

import pandas as pd
from dotenv import load_dotenv

from checks import check_bounces, check_email, normalize_ve_phone

log = logging.getLogger("mailer")
PLACEHOLDER = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
K_ROW, K_PHONE0 = "__fila__", "__tel_orig__"  # claves internas añadidas a cada fila

DEFAULTS: dict[str, dict[str, Any]] = {
    "smtp": {"host": "smtp.gmail.com", "port": 587, "security": "starttls", "timeout": 30},
    "excel": {"file": "data/contactos.xlsx", "sheet": 0, "email_column": "Correo"},
    "message": {
        "subject": "Mensaje para {{Empresa}}", "html_template": "templates/plantilla.html",
        "text_template": "", "from_name": "", "reply_to": "", "list_unsubscribe": "",
        "attachments": [],
    },
    "sending": {
        "delay_min": 3, "delay_max": 8, "batch_size": 20, "batch_pause": 60,
        "max_per_run": 0, "max_retries": 3, "retry_wait": 10, "max_consecutive_errors": 5,
    },
    "files": {"log": "logs/mailer.log", "sent": "logs/enviados.csv",
              "trace": "logs/trazabilidad.csv", "rejected": "logs/rechazados.csv"},
    "validation": {
        "check_email_dns": True, "dns_timeout": 10, "validate_phone": True, "phone_column": "Telefono",
        "phone_required": False, "allow_landline": True, "on_invalid_phone": "skip",
    },
    "imap": {"host": "imap.gmail.com", "port": 993, "days": 3},
    "alerts": {"mode": "problems", "to": ""},
}


class ConfigError(Exception):
    """Error de configuración o datos de entrada."""


class PermanentError(Exception):
    """Error que no se soluciona reintentando. hard=True: el destinatario fue rechazado."""

    def __init__(self, message: str, hard: bool = False) -> None:
        super().__init__(message)
        self.hard = hard


# ----------------------------- Configuración -----------------------------
def load_config(path: Path) -> dict[str, dict[str, Any]]:
    if not path.is_file():
        raise ConfigError(f"No existe el archivo de configuración: {path}")
    try:
        with path.open("rb") as f:
            user = tomllib.load(f)
    except tomllib.TOMLDecodeError as e:
        raise ConfigError(f"config.toml inválido: {e}") from e
    cfg = {k: {**v, **user.get(k, {})} for k, v in DEFAULTS.items()}
    s = cfg["sending"]
    if s["delay_min"] > s["delay_max"]:
        raise ConfigError("sending.delay_min no puede ser mayor que delay_max")
    if cfg["smtp"]["security"] not in ("starttls", "ssl", "none"):
        raise ConfigError('smtp.security debe ser "starttls", "ssl" o "none"')
    if cfg["validation"]["on_invalid_phone"] not in ("skip", "warn"):
        raise ConfigError('validation.on_invalid_phone debe ser "skip" o "warn"')
    if cfg["alerts"]["mode"] not in ("problems", "always", "never"):
        raise ConfigError('alerts.mode debe ser "problems", "always" o "never"')
    return cfg


def setup_logging(path: Path) -> None:
    if log.handlers:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    log.setLevel(logging.INFO)
    for h in (logging.StreamHandler(sys.stdout), logging.FileHandler(path, encoding="utf-8")):
        h.setFormatter(fmt)
        log.addHandler(h)


# ----------------------------- Datos y plantillas -----------------------------
def load_contacts(path: Path, sheet: int | str, email_col: str) -> tuple[list[str], list[dict[str, str]]]:
    if not path.is_file():
        raise ConfigError(f"No existe el archivo de contactos: {path}")
    try:
        if path.suffix.lower() == ".csv":
            df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        else:
            df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    except Exception as e:  # noqa: BLE001
        raise ConfigError(f"No se pudo leer {path.name}: {e}") from e
    df.columns = [str(c).strip() for c in df.columns]
    if email_col not in df.columns:
        raise ConfigError(f"Falta la columna '{email_col}'. Columnas encontradas: {list(df.columns)}")
    df = df.fillna("").astype(str).apply(lambda c: c.str.strip())
    return list(df.columns), df.to_dict("records")


def read_text(path: Path, what: str) -> str:
    if not path.is_file():
        raise ConfigError(f"No existe {what}: {path}")
    return path.read_text(encoding="utf-8")


def placeholders(*texts: str) -> set[str]:
    return {m.group(1) for t in texts for m in PLACEHOLDER.finditer(t)}


def render(template: str, row: dict[str, str], escape: bool = False) -> str:
    def repl(m: re.Match[str]) -> str:
        val = row.get(m.group(1), "")
        return html.escape(val) if escape else val
    return PLACEHOLDER.sub(repl, template)


def html_to_text(src: str) -> str:
    src = re.sub(r"(?is)<(script|style).*?</\1>", "", src)
    src = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</h\d>|</li>", "\n", src)
    text = html.unescape(re.sub(r"<[^>]+>", "", src))
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def load_attachments(paths: list[str], base: Path) -> list[tuple[str, str, str, bytes]]:
    out = []
    for p in paths:
        f = base / p
        if not f.is_file():
            raise ConfigError(f"Adjunto no encontrado: {f}")
        ctype = mimetypes.guess_type(f.name)[0] or "application/octet-stream"
        main, sub = ctype.split("/", 1)
        out.append((f.name, main, sub, f.read_bytes()))
    return out


def build_message(cfg_msg: dict[str, Any], sender: str, to: str, row: dict[str, str],
                  html_tpl: str, text_tpl: str, attachments: list[tuple[str, str, str, bytes]]) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = formataddr((cfg_msg["from_name"], sender)) if cfg_msg["from_name"] else sender
    msg["To"] = to
    msg["Subject"] = " ".join(render(cfg_msg["subject"], row).split())  # sin saltos de línea
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=sender.split("@")[-1])
    if cfg_msg["reply_to"]:
        msg["Reply-To"] = cfg_msg["reply_to"]
    if cfg_msg["list_unsubscribe"]:
        msg["List-Unsubscribe"] = f"<{cfg_msg['list_unsubscribe']}>"
    msg.set_content(render(text_tpl, row))
    msg.add_alternative(render(html_tpl, row, escape=True), subtype="html")
    for name, main, sub, data in attachments:
        msg.add_attachment(data, maintype=main, subtype=sub, filename=name)
    return msg


# ----------------------------- Historial -----------------------------
def load_sent(path: Path) -> set[str]:
    if not path.is_file():
        return set()
    with path.open(newline="", encoding="utf-8") as f:
        return {r["email"].lower() for r in csv.DictReader(f) if r.get("status") in ("ok", "bounced", "rechazado")}


class SentLog:
    def __init__(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        new = not path.is_file()
        self.f = path.open("a", newline="", encoding="utf-8")
        self.w = csv.writer(self.f)
        if new:
            self.w.writerow(["timestamp", "email", "status", "detail"])

    def add(self, email: str, status: str, detail: str = "") -> None:
        self.w.writerow([datetime.now().isoformat(timespec="seconds"), email, status, detail])
        self.f.flush()

    def close(self) -> None:
        self.f.close()


class Tracer:
    """Registro de trazabilidad: una línea por cada evento (validación, envío, rebote, sistema)."""
    COLS = ["timestamp", "run_id", "fila_excel", "correo", "telefono_original", "telefono_normalizado",
            "etapa", "estado", "detalle", "message_id"]

    def __init__(self, path: Path, run_id: str) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        new = not path.is_file()
        self.run_id = run_id
        self.f = path.open("a", newline="", encoding="utf-8")
        self.w = csv.writer(self.f)
        if new:
            self.w.writerow(self.COLS)

    def add(self, etapa: str, estado: str, email: str = "", fila: str = "", tel_orig: str = "",
            tel_norm: str = "", detalle: str = "", message_id: str = "") -> None:
        self.w.writerow([datetime.now().isoformat(timespec="seconds"), self.run_id, fila, email,
                         tel_orig, tel_norm, etapa, estado, detalle, message_id])
        self.f.flush()

    def close(self) -> None:
        self.f.close()


# ----------------------------- SMTP -----------------------------
class Mailer:
    def __init__(self, smtp: dict[str, Any], user: str, password: str) -> None:
        self.smtp, self.user, self.password = smtp, user, password
        self.server: smtplib.SMTP | None = None

    def open(self) -> None:
        s, ctx = self.smtp, ssl.create_default_context()
        if s["security"] == "ssl":
            srv: smtplib.SMTP = smtplib.SMTP_SSL(s["host"], s["port"], timeout=s["timeout"], context=ctx)
        else:
            srv = smtplib.SMTP(s["host"], s["port"], timeout=s["timeout"])
            srv.ehlo()
            if s["security"] == "starttls":
                srv.starttls(context=ctx)
                srv.ehlo()
        if self.user:
            srv.login(self.user, self.password)
        self.server = srv

    def close(self) -> None:
        if self.server:
            try:
                self.server.quit()
            except Exception:  # noqa: BLE001
                pass
            self.server = None

    def send(self, msg: EmailMessage, retries: int, wait: float) -> None:
        last: Exception | None = None
        for attempt in range(1, retries + 2):
            try:
                if self.server is None:
                    self.open()
                assert self.server is not None
                self.server.send_message(msg)
                return
            except smtplib.SMTPAuthenticationError:
                raise
            except smtplib.SMTPRecipientsRefused as e:
                raise PermanentError(f"Destinatario rechazado: {e.recipients}", hard=True) from e
            except smtplib.SMTPResponseException as e:
                if not 400 <= e.smtp_code < 500:
                    raise PermanentError(f"{e.smtp_code} {e.smtp_error!r}") from e
                last = e
            except (smtplib.SMTPException, OSError) as e:
                last = e
            self.close()
            if attempt <= retries:
                delay = wait * attempt
                log.warning("Error temporal (%s). Reintento %d/%d en %.0fs", last, attempt, retries, delay)
                time.sleep(delay)
        raise PermanentError(f"Agotados los reintentos: {last}")


# ----------------------------- Ejecución -----------------------------
def next_wait(done: int, total: int, s: dict[str, Any]) -> float:
    if done >= total:
        return 0
    if s["batch_size"] and done % s["batch_size"] == 0:
        log.info("Pausa de lote: %ss", s["batch_pause"])
        return float(s["batch_pause"])
    return random.uniform(s["delay_min"], s["delay_max"])


def send_alert(cfg: dict[str, Any], user: str, password: str, sender: str, subject: str, body: str) -> None:
    """Notifica por correo un problema. Nunca lanza excepciones (no debe tumbar el proceso)."""
    to = cfg["alerts"]["to"] or sender
    try:
        msg = EmailMessage()
        msg["From"], msg["To"], msg["Subject"] = sender, to, subject
        msg["Date"] = formatdate(localtime=True)
        msg["Message-ID"] = make_msgid(domain=sender.split("@")[-1])
        msg.set_content(body)
        m = Mailer(cfg["smtp"], user, password)
        m.open()
        assert m.server is not None
        m.server.send_message(msg)
        m.close()
        log.info("Alerta enviada a %s", to)
    except Exception as e:  # noqa: BLE001
        log.warning("No se pudo enviar la alerta por correo: %s", e)


def run_bounce_check(cfg: dict[str, Any], base: Path, user: str, password: str, sender: str,
                     trace: Tracer) -> int:
    sent_path = base / cfg["files"]["sent"]
    if not sent_path.is_file():
        log.error("No existe el historial %s: aún no se ha enviado nada.", sent_path)
        return 1
    with sent_path.open(newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    flagged = {r["email"].lower() for r in rows if r["status"] in ("bounced", "rechazado")}
    candidates = {r["email"].lower() for r in rows if r["status"] == "ok"} - flagged
    if not candidates:
        log.info("No hay correos enviados pendientes de revisar.")
        return 0
    log.info("Revisando rebotes de %d correos en %s ...", len(candidates), cfg["imap"]["host"])
    try:
        found = check_bounces(cfg["imap"], user, password, candidates)
    except (imaplib.IMAP4.error, OSError) as e:
        log.critical("No se pudo revisar el buzón por IMAP: %s", e)
        trace.add("rebote", "error", detalle=f"IMAP: {e}")
        return 1
    if not found:
        log.info("Sin rebotes detectados en los últimos %s días (revisados: %d).",
                 cfg["imap"]["days"], len(candidates))
        trace.add("rebote", "ok", detalle=f"sin rebotes; revisados {len(candidates)}")
        return 0
    history = SentLog(sent_path)
    try:
        for email, subject in sorted(found.items()):
            history.add(email, "bounced", f"Rebote: {subject}")
            trace.add("rebote", "rebotado", email=email, detalle=subject)
            log.error("REBOTE: %s no pudo entregarse (buzón inexistente o rechazado) | %s", email, subject)
    finally:
        history.close()
    if cfg["alerts"]["mode"] != "never":
        lines = "\n".join(f"- {e}: {s}" for e, s in sorted(found.items()))
        send_alert(cfg, user, password, sender, f"[Mailer] {len(found)} correos rebotaron",
                   f"Estos correos NO se entregaron. Corrígelos en el Excel:\n\n{lines}\n\n"
                   f"Detalle: {cfg['files']['trace']}")
    return 2


def run(args: argparse.Namespace) -> int:
    cfg_path = Path(args.config).resolve()
    base = cfg_path.parent
    cfg = load_config(cfg_path)
    load_dotenv(base / ".env")
    setup_logging(base / cfg["files"]["log"])
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S") + ("-sim" if args.dry_run else "")
    log.info("=== Ejecución %s ===", run_id)

    user = os.getenv("SMTP_USER", "")
    password = os.getenv("SMTP_PASSWORD", "")
    sender = os.getenv("SENDER_EMAIL") or user
    needs_login = args.check_bounces or not args.dry_run
    if needs_login and (not sender or (cfg["smtp"]["security"] != "none" and not password)):
        raise ConfigError("Faltan SMTP_USER / SMTP_PASSWORD en el archivo .env")
    if args.dry_run and not sender:
        sender = "remitente@example.com"

    trace = Tracer(base / cfg["files"]["trace"], run_id)
    try:
        if args.check_bounces:
            return run_bounce_check(cfg, base, user, password, sender, trace)
        return run_send(args, cfg, base, user, password, sender, trace, run_id)
    finally:
        trace.close()


def run_send(args: argparse.Namespace, cfg: dict[str, Any], base: Path, user: str, password: str,
             sender: str, trace: Tracer, run_id: str) -> int:
    ex, msgc, snd, val = cfg["excel"], cfg["message"], cfg["sending"], cfg["validation"]
    columns, rows = load_contacts(base / ex["file"], ex["sheet"], ex["email_column"])
    pcol = val["phone_column"] if val["validate_phone"] else ""
    if pcol and pcol not in columns:
        raise ConfigError(f"Falta la columna de teléfono '{pcol}' (o pon validation.validate_phone = false). "
                          f"Columnas encontradas: {columns}")
    html_tpl = read_text(base / msgc["html_template"], "la plantilla HTML")
    text_tpl = read_text(base / msgc["text_template"], "la plantilla de texto") if msgc["text_template"] \
        else html_to_text(html_tpl)
    missing = placeholders(msgc["subject"], html_tpl, text_tpl) - set(columns)
    if missing:
        raise ConfigError(f"Las plantillas usan columnas que no existen en el Excel: {sorted(missing)}")
    attachments = load_attachments(msgc["attachments"], base)

    def meta(row: dict[str, str]) -> dict[str, str]:
        return {"fila": row.get(K_ROW, ""), "tel_orig": row.get(K_PHONE0, ""),
                "tel_norm": row.get(pcol, "") if pcol else ""}

    # ---- Fase 1: validar contactos, deduplicar y saltar los ya procesados ----
    sent_path = base / cfg["files"]["sent"]
    done = set() if args.ignore_history else load_sent(sent_path)
    queue: list[tuple[dict[str, str], str]] = []
    first_row: dict[str, int] = {}
    rejected: list[tuple[int, str, str, str]] = []
    problems: list[str] = []
    skipped = 0
    for n, row in enumerate(rows, start=2):  # la fila 1 del Excel son los encabezados
        row[K_ROW] = str(n)
        raw_email = row[ex["email_column"]]
        raw_phone = row[pcol] if pcol else ""
        row[K_PHONE0] = raw_phone
        reasons: list[str] = []
        email, dns = raw_email, ""
        try:
            email, dns = check_email(raw_email, val["check_email_dns"], val["dns_timeout"])
        except ValueError as e:
            reasons.append(f"Correo inválido: {e}")
        if pcol:
            if raw_phone:
                try:
                    row[pcol] = normalize_ve_phone(raw_phone, val["allow_landline"])
                except ValueError as e:
                    text = f"Teléfono inválido '{raw_phone}': {e}"
                    if val["on_invalid_phone"] == "skip":
                        reasons.append(text)
                    else:
                        log.warning("Fila %d: %s (se enviará igualmente)", n, text)
                        trace.add("validacion", "advertencia", email=raw_email, detalle=text, **meta(row))
            elif val["phone_required"]:
                reasons.append("Teléfono vacío (obligatorio)")
        if reasons:
            reason = " | ".join(reasons)
            rejected.append((n, raw_email, raw_phone, reason))
            problems.append(f"Fila {n} ({raw_email or 'sin correo'}): {reason}")
            log.error("Fila %d omitida (%s): %s", n, raw_email or "(vacío)", reason)
            trace.add("validacion", "rechazado", email=raw_email, detalle=reason, **meta(row))
            continue
        key = email.lower()
        if key in first_row:
            skipped += 1
            trace.add("validacion", "omitido", email=email, detalle=f"duplicado de la fila {first_row[key]}",
                      **meta(row))
        elif key in done:
            skipped += 1
            trace.add("validacion", "omitido", email=email,
                      detalle="ya enviado/rebotado/rechazado en una ejecución previa", **meta(row))
        else:
            first_row[key] = n
            queue.append((row, email))
            trace.add("validacion", "ok", email=email, detalle=f"correo válido (DNS: {dns})", **meta(row))

    if rejected:
        rej_path = base / cfg["files"]["rejected"]
        rej_path.parent.mkdir(parents=True, exist_ok=True)
        with rej_path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["fila_excel", "correo", "telefono", "motivo"])
            w.writerows(rejected)
        log.warning("%d filas rechazadas. Corrígelas en el Excel; detalle en %s", len(rejected), rej_path)

    limit = args.limit if args.limit is not None else snd["max_per_run"]
    if args.test:
        queue, limit = queue[:1], 0
    elif limit:
        queue = queue[:limit]
    log.info("Contactos: %d | a enviar: %d | ya procesados/duplicados: %d | rechazados: %d",
             len(rows), len(queue), skipped, len(rejected))

    def alert(title: str, lines: list[str], aborted: bool = False) -> None:
        mode = cfg["alerts"]["mode"]
        if args.dry_run or args.test or mode == "never" or (mode == "problems" and not lines and not aborted):
            return
        body = f"Ejecución: {run_id}\n{title}\n\n" + "\n".join(f"- {x}" for x in lines[:50])
        if len(lines) > 50:
            body += f"\n... y {len(lines) - 50} más (ver {cfg['files']['trace']})"
        body += f"\n\nTrazabilidad completa: {cfg['files']['trace']}\nLog: {cfg['files']['log']}"
        send_alert(cfg, user, password, sender, f"[Mailer] {title}", body)

    if not queue:
        log.info("Nada que enviar.")
        if not args.dry_run:
            alert("Sin envíos: no quedó ningún contacto válido", problems)
        return 2 if rejected else 0

    # ---- Simulación ----
    if args.dry_run:
        for i, (row, email) in enumerate(queue, 1):
            msg = build_message(msgc, sender, email, row, html_tpl, text_tpl, attachments)
            log.info("[SIMULACIÓN %d/%d] %s | Asunto: %s", i, len(queue), email, msg["Subject"])
            trace.add("simulacion", "ok", email=email, detalle=msg["Subject"], **meta(row))
            if i == 1:
                (base / "logs").mkdir(exist_ok=True)
                (base / "logs" / "preview.eml").write_bytes(bytes(msg))
        log.info("Simulación OK. Vista previa del primer correo: logs/preview.eml")
        return 2 if rejected else 0

    # ---- Fase 2: envío ----
    mailer = Mailer(cfg["smtp"], user, password)
    history = None if args.test else SentLog(sent_path)
    ok = fail = consecutive = 0
    aborted = False
    try:
        mailer.open()
        log.info("Conexión SMTP establecida (%s:%s)", cfg["smtp"]["host"], cfg["smtp"]["port"])
        for n, (row, email) in enumerate(queue, 1):
            to = args.test or email
            tag = f"[{n}/{len(queue)}] fila {row[K_ROW]} {to}"
            mid = ""
            try:
                msg = build_message(msgc, sender, to, row, html_tpl, text_tpl, attachments)
                mid = str(msg["Message-ID"])
                if args.test:
                    msg.replace_header("Subject", "[PRUEBA] " + msg["Subject"])
                mailer.send(msg, snd["max_retries"], snd["retry_wait"])
                ok += 1
                consecutive = 0
                log.info("%s aceptado por el servidor SMTP", tag)
                trace.add("envio", "enviado", email=email, message_id=mid,
                          detalle="aceptado por SMTP (usa --check-bounces para confirmar entrega)", **meta(row))
                if history:
                    history.add(email, "ok", mid)
            except PermanentError as e:
                fail += 1
                consecutive += 1
                status = "rechazado" if e.hard else "error"
                problems.append(f"Fila {row[K_ROW]} ({email}): {e}")
                log.error("%s FALLÓ: %s", tag, e)
                trace.add("envio", status, email=email, message_id=mid, detalle=str(e), **meta(row))
                if history:
                    history.add(email, status, str(e))
                if consecutive >= snd["max_consecutive_errors"]:
                    aborted = True
                    log.critical("%d errores consecutivos: se aborta para proteger tu cuenta.", consecutive)
                    trace.add("sistema", "abortado", detalle=f"{consecutive} errores consecutivos")
                    break
            if n < len(queue):
                time.sleep(next_wait(n, len(queue), snd))
    except smtplib.SMTPAuthenticationError as e:
        log.critical("Autenticación rechazada (revisa usuario/contraseña de aplicación): %s", e)
        trace.add("sistema", "error", detalle=f"autenticación: {e}")
        return 1
    except (smtplib.SMTPException, OSError) as e:
        log.critical("No se pudo conectar al servidor SMTP: %s", e)
        trace.add("sistema", "error", detalle=f"conexión SMTP: {e}")
        return 1
    except KeyboardInterrupt:
        log.warning("Interrumpido por el usuario. Puedes reanudar: el historial evita duplicados.")
        trace.add("sistema", "interrumpido", detalle="Ctrl+C")
    finally:
        mailer.close()
        if history:
            history.close()
    log.info("Fin. Enviados: %d | Fallidos: %d | Rechazados en validación: %d", ok, fail, len(rejected))
    alert(f"Resumen: {ok} enviados, {fail} fallidos, {len(rejected)} rechazados"
          + (" (ABORTADO)" if aborted else ""), problems, aborted)
    return 0 if (fail == 0 and not rejected) else 2


def main() -> None:
    p = argparse.ArgumentParser(description="Envío masivo de correos desde Excel/CSV")
    p.add_argument("-c", "--config", default=str(Path(__file__).parent / "config.toml"))
    p.add_argument("--dry-run", action="store_true", help="valida y simula sin enviar")
    p.add_argument("--test", metavar="EMAIL", help="envía solo el 1er contacto a este correo de prueba")
    p.add_argument("--limit", type=int, help="máximo de correos en esta ejecución")
    p.add_argument("--ignore-history", action="store_true", help="reenvía aunque ya figure como enviado")
    p.add_argument("--check-bounces", action="store_true",
                   help="revisa el buzón (IMAP) y marca los correos que rebotaron")
    try:
        sys.exit(run(p.parse_args()))
    except ConfigError as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        sys.exit(130)
    except Exception:  # noqa: BLE001
        detail = traceback.format_exc()
        if log.handlers:
            log.critical("ERROR INESPERADO:\n%s", detail)
        else:
            print(detail, file=sys.stderr)
        print("[ERROR FATAL] Fallo inesperado. Revisa logs/mailer.log y envíame el traceback.", file=sys.stderr)
        sys.exit(3)


if __name__ == "__main__":
    main()
