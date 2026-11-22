# Hazine Günlük Değişim Raporu

Borsa İstanbul BAP günlük bültenlerinden Hazine kağıtlarının kesin alım satım işlemlerindeki günlük değişimi
hesaplar ve BV Portföy formatında yatay A4 PDF (+ Excel) üretir.

**Teknoloji:** Python 3.12 (pandas, matplotlib, reportlab, openpyxl, pypdfium2), GitHub Actions, cron-job.org, Gmail SMTP.

## Rapor içeriği

- Kaynak: Borsa İstanbul BAP kesin alım satım günlük bülteni (`borsaistanbul.com/data/ttb/...`), normal ve
  küçük emirler pazarı. İndirilen ham veriler yerelde `data/` altında önbelleğe alınır (git dışı).
- Gruplar: TLREF endeksli, sabit kuponlu/iskontolu (DT/HB), döviz cinsi, TLREFK, TÜFE endeksli, altın
  tahvili/kira sertifikası, kamu kira sertifikaları; önceki güne göre fiyat/getiri değişimi, hacim
  dağılımı, getiri eğrisi (x ekseni Macaulay durasyon) ve 1 hafta önceki seviyeler.
- Mail gövdesi: özet rakamlar ve öne çıkanlar (`output/mail.html`, `output/mail.txt`).

## Akış (GitHub Actions)

1. **Günlük rapor (onay maili)** — `.github/workflows/gunluk-rapor.yml`
   Hafta içi 09:00 (İstanbul) cron-job.org üzerinden `workflow_dispatch` ile tetiklenir; 09:37'deki GitHub
   zamanlaması yedektir. Bugünden önceki son bülten gününün raporunu üretir, dosyaları
   `rapor-YYYYMMDD` artifact'ı olarak 7 gün saklar ve `ONAY_ALICI` adresine **[ONAY]** mail atar.
   Tatil sonrası aynı bülten tekrar seçilirse yedek koşu atlanır. Elle de çalıştırılabilir (`tarih` girilerek;
   `onay_maili` kapatılırsa yalnızca artifact üretilir).
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

Zamanlama için cron-job.org'da hafta içi 09:00 (İstanbul) `gunluk-rapor.yml` için `workflow_dispatch` çağrısı tanımlayın.

## Yerel çalıştırma

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python hazine_rapor.py --date 16.09.2026 --excel
```

Çıktılar `output/` altına yazılır; `output/_work_YYYYMMDD/page*.png` sayfa görüntüleri kontrol içindir.

Çıktılar (`output/`):

- `Hazine_Gunluk_Degisim_Raporu_YYYYMMDD.pdf` ve (`--excel` ile) `.xlsx`
- `mail.html`, `mail.txt`: mail gövdesi
- `result.json`: rapor tarihi, dosya yolları, kağıt sayıları, özet ve öne çıkanlar (workflow'lar ve
  `send_mail.py` bunu okur)

Mail göndermek için (gerçek gönderim yapar; `GMAIL_USER`, `GMAIL_APP_PASSWORD` gerekir):

```bash
.venv/bin/python send_mail.py --mode onay --to ben@ornek.com
```

Otomatik test yoktur; `hazine_rapor.py` hatasız biter ve PDF üretirse hat çalışıyordur.

## Lisans

Repoda lisans dosyası yoktur; kod varsayılan olarak tüm hakları saklı kabul edilir.
