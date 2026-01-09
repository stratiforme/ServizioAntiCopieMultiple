# ServizioAntiCopieMultiple

**Servizio Windows** in C# (.NET 8.0) per monitorare i lavori di stampa e impedire la stampa di più copie dello stesso documento.

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

## Diagnostica e debug

Per facilitare l'analisi dei casi in cui il numero di copie non viene riconosciuto, il servizio supporta modalità più verbose e salvataggio di dump diagnostici.

- Eseguire in console (debug):
  - Avvia `ServizioAntiCopieMultiple.exe --console` per eseguire il servizio in foreground. In questa modalità il logger predefinito è `Debug` e viene abilitato il sink Console.

- Variabili d'ambiente utili:
  - `SACM_LOG_LEVEL`: fornisce un override per il livello minimo di log. Valori: `Verbose`, `Debug`, `Information`, `Warning`, `Error`, `Fatal`. Esempio PowerShell:
    ```powershell
    $env:SACM_LOG_LEVEL = 'Debug'
    ```
  - `SACM_ENABLE_CONSOLE`: se impostata a `true` forza il sink Console anche quando non si esegue con `--console`.
    ```powershell
    $env:SACM_ENABLE_CONSOLE = 'true'
    ```

- Cartelle diagnostiche (per default):
  - Logs: `C:\ProgramData\ServizioAntiCopieMultiple\logs` (file rolling Serilog `service-*.log`).
  - Diagnostics: `C:\ProgramData\ServizioAntiCopieMultiple\diagnostics` — qui vengono salvati i dump WMI (`wmi_*.json`) e, se presenti, i `PrintTicket` XML (`printticket_*.xml`).
  - Responses/simulator (utilità del servizio): `C:\ProgramData\ServizioAntiCopieMultiple\responses` e `...\\simulator`.

- Cosa viene registrato in più quando si abilita Debug:
  - Dump completo delle proprietà WMI del `Win32_PrintJob` in `diagnostics`.
  - Salvataggio separato del `PrintTicket` XML quando presente.
  - Debug dettagliato del tentativo nativo `GetJob`/`DEVMODE` incluso valore `dmCopies` quando disponibile.

---

## Note

- Il servizio tenterà di rilevare il numero di copie con più strategie: proprietà WMI (`Copies`, `TotalPages`), parsing `PrintTicket` XML e un tentativo best‑effort di leggere `DEVMODE`/`JOB_INFO_2` via Win32 `GetJob`. I tentativi nativi richiedono permessi e sono driver‑specifici.
- Se vuoi assistenza nell'interpretare i dump presenti in `diagnostics`, copia i file e incollami il contenuto — posso analizzarli.

---

## Licenza

CC BY-NC-SA 4.0 — Nicola Cantalupo

