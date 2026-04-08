# Packeta Matcher

Webová aplikace pro automatické přiřazení čísel objednávek k zásilkám **Packeta / Zásilkovna** na základě Excel exportu.

---

## Jak to funguje

1. **Nahrajte Excel soubor** – exportovaný ze Zásilkovny nebo od dopravce (obsahuje sloupec s trasovacími čísly ve formátu `Z + číslice` (např. `Z1234567890`) nebo jen číslice (např. `4170246843`), a sloupec `Customs Value (CHF)`).
2. **Aplikace automaticky detekuje** sloupec s trasovacími čísly – hledá podle záhlaví ("Tracking Number") nebo podle formátu dat. Podporuje čísla s prefixem Z i bez něj.
3. **Zásilkovna SOAP API** je dotázáno paralelně pro všechna trasovací čísla (10 vláken současně), aby zpracování bylo co nejrychlejší.
4. **Výsledný Excel** dostanete ke stažení – obsahuje dva nové sloupce hned za trasovacími čísly:
   - **Číslo objednávky** – doplněno z API; pokud není nalezeno, zobrazí se `—`.
   - **Celková cena CHF** – vypočtená cena, kterou zákazník skutečně zaplatil.

---

## Výpočet celkové ceny

Zásilkovna vždy deklaruje v poli `Customs Value (CHF)` hodnotu: **produkty (bez DPH) + 9 CHF doprava**, a to i u zásilek s dopravou zdarma.

Logika výpočtu:

```
product_value = Customs Value − 9 CHF

Pokud product_value > 99 CHF  → doprava zdarma:
    Celková cena = product_value × 1,081

Pokud product_value ≤ 99 CHF  → zákazník platí dopravu 9,99 CHF (vč. DPH):
    Celková cena = (product_value + 9,99 / 1,081) × 1,081
```

- **DPH**: 8,1 % (švýcarská MWST)
- **Doprava zdarma od**: 99 CHF hodnoty produktů (bez DPH)
- **Cena dopravy**: 9,99 CHF vč. DPH (≈ 9,24 CHF bez DPH)

---

## Technické detaily

| Komponenta | Technologie |
|---|---|
| Backend | Python / Flask |
| SOAP API klient | zeep |
| Excel zpracování | openpyxl |
| Paralelní API volání | ThreadPoolExecutor (10 vláken) |
| Hosting | Render (gunicorn) |

---

## Nasazení (Render)

1. Forkněte nebo naklonujte tento repozitář.
2. Vytvořte novou **Web Service** na [render.com](https://render.com).
3. Nastavte environment variable:
   ```
   ZASILKOVNA_API_PASSWORD = <váš API klíč ze Zásilkovny>
   ```
4. Build Command: `pip install -r requirements.txt`
5. Start Command: `gunicorn app:app`

---

## Lokální spuštění

```bash
pip install -r requirements.txt
set ZASILKOVNA_API_PASSWORD=<váš klíč>
python app.py
```

Aplikace poběží na `http://localhost:5000`.
