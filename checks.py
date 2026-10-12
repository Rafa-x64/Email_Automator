"""Validaciones: teléfonos de Venezuela, correos (sintaxis + DNS) y rebotes (IMAP)."""
from __future__ import annotations

import imaplib
import re
from datetime import datetime, timedelta
from email import message_from_bytes, policy
from typing import Any

from email_validator import EmailNotValidError, validate_email

# Móviles: Digitel 412/422, Movistar 414/424, Movilnet 416/426
MOBILE_PREFIXES = frozenset({"412", "414", "416", "422", "424", "426"})
_ADDR = re.compile(r"[\w.+'-]+@[\w-]+(?:\.[\w-]+)+")
_SOFT_BOUNCE = ("delay", "retraso", "demora")
_domain_cache: dict[str, str] = {}  # dominio -> "" (ok) | motivo del fallo | "?" (no concluyente)


def normalize_ve_phone(raw: str, allow_landline: bool = True) -> str:
    """Devuelve el teléfono en formato +58XXXXXXXXXX o lanza ValueError con el motivo."""
    text = re.sub(r"\.0+$", "", raw.strip())  # Excel a veces guarda 4121234567.0
    if re.search(r"[^\d\s()+\-.]", text):
        raise ValueError("contiene caracteres no permitidos")
    digits = re.sub(r"\D", "", text)
    if digits.startswith("00"):
        digits = digits[2:]
    if digits.startswith("58") and len(digits) >= 12:
        digits = digits[2:]
    digits = digits.removeprefix("0")
    if len(digits) != 10:
        raise ValueError(f"debe tener 10 dígitos sin el 0 inicial ni +58 (tiene {len(digits)})")
    if digits[:3] in MOBILE_PREFIXES:
        return "+58" + digits
    if digits[0] == "2":
        if allow_landline:
            return "+58" + digits
        raise ValueError("es un número fijo y solo se aceptan móviles")
    raise ValueError(f"el prefijo {digits[:3]} no existe en Venezuela")


def check_email(address: str, check_dns: bool = True, timeout: int = 10) -> tuple[str, str]:
    """Valida sintaxis y que el dominio reciba correo (registros MX/A).
    Devuelve (correo_normalizado, estado_dns) o lanza ValueError."""
    try:
        info = validate_email(address, check_deliverability=False)
    except EmailNotValidError as e:
        raise ValueError(f"sintaxis inválida: {e}") from e
    if not check_dns:
        return info.normalized, "omitido"
    domain = info.domain
    if domain not in _domain_cache:
        try:
            validate_email(f"x@{domain}", check_deliverability=True, timeout=timeout)
            _domain_cache[domain] = ""
        except EmailNotValidError as e:
            _domain_cache[domain] = str(e)
        except Exception:  # noqa: BLE001  (red caída, DNS bloqueado, etc.)
            _domain_cache[domain] = "?"
    state = _domain_cache[domain]
    if state == "?":
        return info.normalized, "no_concluyente"
    if state:
        raise ValueError(f"el dominio '{domain}' no puede recibir correo: {state}")
    return info.normalized, "ok"


def check_bounces(imap_cfg: dict[str, Any], user: str, password: str,
                  candidates: set[str]) -> dict[str, str]:
    """Busca en el buzón avisos de rebote (mailer-daemon/postmaster) de los últimos N días
    y devuelve {correo_rebotado: asunto_del_aviso} limitado a los correos de `candidates`."""
    since = (datetime.now() - timedelta(days=int(imap_cfg["days"]))).strftime("%d-%b-%Y")
    found: dict[str, str] = {}
    conn = imaplib.IMAP4_SSL(imap_cfg["host"], int(imap_cfg["port"]))
    try:
        conn.login(user, password)
        conn.select("INBOX", readonly=True)
        for sender in ("mailer-daemon", "postmaster"):
            _, data = conn.search(None, "SINCE", since, "FROM", sender)
            for num in (data[0] or b"").split():
                _, parts = conn.fetch(num, "(BODY.PEEK[])")
                raw = next((p[1] for p in parts if isinstance(p, tuple)), b"")
                msg = message_from_bytes(raw, policy=policy.default)
                subject = str(msg.get("Subject", ""))
                if any(w in subject.lower() for w in _SOFT_BOUNCE):
                    continue  # aviso de demora, no es un rebote definitivo
                text = str(msg.get("X-Failed-Recipients", "")) + " " + raw.decode("utf-8", "ignore")
                for addr in {a.lower() for a in _ADDR.findall(text)} & candidates:
                    found.setdefault(addr, subject)
    finally:
        try:
            conn.logout()
        except Exception:  # noqa: BLE001
            pass
    return found
