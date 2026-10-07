# Vabilo na Ninovo 40-letnico

Spletno vabilo v obliki spričevala. Gostje na njem potrdijo udeležbo, vsaka potrditev pa se zapiše kot nova vrstica v Google tabelo.

Povezava za goste: https://nicki89-blip.github.io/vabilo/

## Datoteke

- `index.html`: vabilo in obrazec za potrditev.
- `config.js`: edina datoteka z nastavitvami (`SCRIPT_URL`, telefonska številka za rezervni način).
- `apps-script/Code.gs`: koda za Google tabelo. Tukaj je samo za referenco, v tabelo jo prilepiš ročno.
- `og-image-v2.png`: slika, ki se pokaže v predogledu povezave (WhatsApp, Messenger).
- `robots.txt`: iskalnikom prepove indeksiranje.

## Postavitev tabele (enkrat)

1. V Google Drive ustvari prazno tabelo **Nino 40 potrditve**.
2. Razširitve > Apps Script. Pobriši, kar je v urejevalniku, prilepi celotno vsebino `apps-script/Code.gs` in shrani.
3. Zgoraj izberi funkcijo `setup` in klikni Zaženi. Odobri dovoljenja (Google opozori, da aplikacija ni preverjena: Napredno > Odpri projekt).
4. Uvedi > Nova uvedba > vrsta **Spletna aplikacija**. Izvedi kot: **jaz**. Kdo ima dostop: **Vsi**. Kopiraj URL, ki se konča z `/exec`.
5. URL vpiši v `config.js` pod `SCRIPT_URL` (ali ga pošlji Claude Code).

## Preizkus formul

V urejevalniku Apps Script zaženi `dodajTestnePodatke`. V zavihku Povzetek mora pisati: skupaj oseb 5, gospodinjstva, ki pridejo 2, odpovedala 1, s prenočiščem 1. V tabeli zadnjih odgovorov so trije testni gostje (Ana z 3 osebami, Boris ne pride, Cene z 2 osebama). Nato zaženi `pobrisiTestnePodatke`.

## Popravki skripte

**Pomembno:** Uvedi > Upravljaj uvedbe > svinčnik pri obstoječi uvedbi > Različica: **Nova različica** > Uvedi. Če narediš novo uvedbo, dobiš nov URL in vabilo preneha pošiljati potrditve.

## Zavihki v tabeli

- **Odgovori**: vsaka oddana potrditev, tudi podvojene. Tega ne urejaj ročno, razen za brisanje testnih vrstic.
- **Povzetek**: velja zadnji odgovor po imenu (velike črke in presledki na robovih se ne upoštevajo). Prikaže skupno število oseb, odpovedi, prenočišča in čas zadnjega odgovora. Osveži se sam.
- **Povabljeni**: v stolpec A vpiši imena povabljenih, stolpec B sam pokaže »pride (2)«, »ne pride« ali »ni odgovora« (rumeno).
  Stolpec C pokaže osebno povezavo za vsak par (`?za=Ime in Ime Priimek`). Pošlji jo obema partnerjema: ime je na obrazcu že vpisano in če odgovorita oba, se upošteva samo zadnji odgovor. Imena v stolpcu A po pošiljanju ne spreminjaj.

Vabilo je za pare, zato gostje vpišejo obe imeni (npr. »Ana in Marko Novak«), število oseb je 2, 1 ali 0. Ujemanje imen se zanaša na enak zapis. Če se je gost vpisal drugače (npr. »Miha« namesto »Mihael Novak«), popravi ime v zavihku Povabljeni, da se ujema z zapisom v Odgovorih.

## Rezervni način

Če pošiljanje ne uspe (ni omrežja, napaka skripte ali je `SCRIPT_URL` prazen), stran gostu pokaže sestavljeno sporočilo z gumboma Kopiraj in WhatsApp. Če v `config.js` vpišeš `PHONE`, gre WhatsApp sporočilo neposredno tebi.

## Lokalni preizkus

```bash
python3 -m http.server 8767
```

Nato odpri http://localhost:8767/.
