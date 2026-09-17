#!/usr/bin/env python3
"""Raporu Gmail SMTP ile gönderir.

  python send_mail.py --mode onay  --to ben@ornek.com      # onay maili (üstte onay bandı)
  python send_mail.py --mode ekip  --to a@x.com,b@y.com    # ekibe nihai mail

Ortam değişkenleri: GMAIL_USER, GMAIL_APP_PASSWORD; onay modunda APPROVE_URL (isteğe bağlı).
Dosyalar output/result.json'dan okunur.
"""
import argparse
import json
import mimetypes
import os
import smtplib
import sys
from email.message import EmailMessage
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["onay", "ekip"], required=True)
    ap.add_argument("--to", required=True, help="virgülle ayrılmış adresler")
    ap.add_argument("--bcc", default="", help="virgülle ayrılmış adresler")
    args = ap.parse_args()

    res = json.loads((ROOT / "output" / "result.json").read_text())
    to = [a.strip() for a in args.to.split(",") if a.strip()]
    bcc = [a.strip() for a in args.bcc.split(",") if a.strip()]
    if not to and not bcc:
        sys.exit("Alıcı yok.")

    body_html = (ROOT / res["mail_html"]).read_text()
    body_txt = (ROOT / res["mail_txt"]).read_text()
    subject = res["mail_konu"]
    if args.mode == "onay":
        tarih = res["rapor_tarihi"]
        url = os.environ.get("APPROVE_URL", "")
        link = (f'<a href="{url}" style="color:#8B0000;font-weight:bold">Ekibe gönder workflow\'unu aç</a>'
                if url else "GitHub Actions → “Ekibe gönder” workflow’u")
        banner = (f'<div style="font-family:Arial,Helvetica,sans-serif;background:#FFF4E5;border:1px solid #E0A040;'
                  f'padding:10px 14px;margin-bottom:12px;font-size:14px"><b>ONAY BEKLİYOR — bu mail yalnızca size '
                  f'gönderildi.</b><br>Rapor uygunsa: {link} → <b>Run workflow</b> → tarih alanına '
                  f'<b>{tarih}</b> yazın. Aynı PDF/Excel ekibe gönderilir.</div>')
        body_html = banner + body_html
        body_txt = (f"ONAY BEKLİYOR — Rapor uygunsa GitHub Actions'ta 'Ekibe gönder' workflow'unu "
                    f"tarih={tarih} ile çalıştırın. {url}\n\n" + body_txt)
        subject = "[ONAY] " + subject

    user, pwd = os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"]
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"BV Portföy Hazine Raporu <{user}>"
    if to:
        msg["To"] = ", ".join(to)
    msg.set_content(body_txt)
    msg.add_alternative(body_html, subtype="html")
    for key in ("pdf", "excel"):
        if res.get(key):
            p = ROOT / res[key]
            ctype = mimetypes.guess_type(p.name)[0] or "application/octet-stream"
            maintype, subtype = ctype.split("/", 1)
            msg.add_attachment(p.read_bytes(), maintype=maintype, subtype=subtype, filename=p.name)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(user, pwd)
        s.send_message(msg, to_addrs=to + bcc)
    print(f"gönderildi ({args.mode}): {subject} → {len(to) + len(bcc)} alıcı")


if __name__ == "__main__":
    main()
