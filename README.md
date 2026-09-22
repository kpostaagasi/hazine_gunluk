# Hazine Günlük Değişim Raporu

Borsa İstanbul BAP günlük bültenlerinden Hazine kağıtlarının kesin alım satım işlemlerindeki günlük değişimi
hesaplar ve BV Portföy formatında yatay A4 PDF (+ Excel) üretir.

## Akış (GitHub Actions)

1. **Günlük rapor (onay maili)** — `.github/workflows/gunluk-rapor.yml`
   Hafta içi 08:30 (İstanbul) çalışır, bugünden önceki son bülten gününün raporunu üretir, dosyaları
   `rapor-YYYYMMDD` artifact'ı olarak 7 gün saklar ve `ONAY_ALICI` adresine **[ONAY]** mail atar.
   Tatil sonrası aynı bülten tekrar seçilirse mail atlanır. Elle de çalıştırılabilir (tarih girilerek).
2. **Ekibe gönder (onaydan sonra)** — `.github/workflows/ekibe-gonder.yml`
   Actions → *Ekibe gönder* → *Run workflow* → tarih (GG.AA.YYYY). Onay mailinde incelenen dosyaların
   aynısını `EKIP_ALICILAR` adreslerine gönderir.

## Kurulum

Repo → Settings → Secrets and variables → Actions:

| Tür | Ad | Değer |
|---|---|---|
| Secret | `GMAIL_USER` | Gönderen Gmail adresi |
| Secret | `GMAIL_APP_PASSWORD` | Gmail uygulama şifresi (Google Hesabı → Güvenlik → 2 adımlı doğrulama → Uygulama şifreleri) |
| Secret | `ONAY_ALICI` | Onay mailinin gideceği adres |
| Secret | `EKIP_ALICILAR` | Ekip adresleri, virgülle ayrılmış |

## Yerel çalıştırma

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python hazine_rapor.py --date 16.09.2026 --excel
```

Çıktılar `output/` altına yazılır; `output/_work_YYYYMMDD/page*.png` sayfa görüntüleri kontrol içindir.
