# ServizioAntiCopieMultiple

**Servizio Windows** in C# (.NET 10.0) per monitorare i lavori di stampa e impedire la stampa di più copie dello stesso documento.

---

## Distribuzione e release

Le release pubblicate contengono due eseguibili nella stessa cartella:
- `ServizioAntiCopieMultiple.exe` — il servizio Windows
- `gestionesacm.exe` — tool console per installare/avviare/fermare/disinstallare il servizio

Per semplificare l'uso all'utente finale i binari sono pubblicati come build "self-contained, single-file" per `win-x64` e sono posizionati in `artifacts/publish` dallo script di pubblicazione.

---

## Installazione (utente finale)

Opzioni disponibili (esegui come amministratore):

1) Usare il tool `gestionesacm.exe` (consigliato)
- Estrai la release in una cartella.
- Apri PowerShell o cmd come amministratore e avvia `gestionesacm.exe` per il menu interattivo che trova automaticamente `ServizioAntiCopieMultiple.exe` nella stessa cartella.

2) Usare direttamente `sc.exe`
- `sc create ServizioAntiCopieMultiple binPath= "C:\path\to\ServizioAntiCopieMultiple.exe" start= auto DisplayName= "Servizio Anti Copie Multiple"`
- `sc start ServizioAntiCopieMultiple`

Note:
- La creazione della sorgente del registro eventi (`EventLog` source) richiede privilegi amministrativi. Se si usa `gestionesacm.exe`, il tool tenterà di creare la sorgente durante l'installazione.

---

## Per gli sviluppatori

Script e CI:
- `publish.ps1` pubblica entrambi i progetti in `artifacts/publish` come `win-x64` self-contained single-file.
- Il workflow GitHub Actions `.github/workflows/publish.yml` esegue build/publish e allega gli artifact a una release quando tagghi `v*`.

Compilare localmente:
- `dotnet build ServizioAntiCopieMultiple.sln -c Release`
- `pwsh publish.ps1` (richiede PowerShell)

---

## Uso di `gestionesacm`

`gestionesacm.exe` è il tool console incluso nella release. Se si trova nella stessa cartella di `ServizioAntiCopieMultiple.exe` lo userà automaticamente. In alternativa, il tool chiederà il percorso dell'eseguibile del servizio.

---

## Note

- Release ufficiali sono pensate per Windows x64. Il servizio usa API Windows-specifiche.
- Testa l'installazione su una macchina Windows pulita prima di distribuire.

---

## Licenza

CC BY-NC-SA 4.0 — Nicola Cantalupo

