# ServizioAntiCopieMultiple

**Servizio Windows** in C# (.NET 10.0) per monitorare i lavori di stampa e impedire la stampa di più copie dello stesso documento, invitando l’utente a utilizzare la fotocopiatrice per duplicazioni multiple.

Il progetto è open source e distribuito con **licenza CC BY-NC-SA 4.0**, ma l’uso commerciale è vietato e va sempre attribuito l’autore originale: **Nicola Cantalupo**.

---

## ⚡ Funzionalità principali
- Monitoraggio continuo delle stampanti Windows.
- Identificazione automatica dei lavori di stampa con più copie.
- Blocco automatico dei job non conformi.
- Notifica all’utente tramite popup o log di sistema.
- Logging dettagliato delle attività del servizio.
- Configurabile tramite `appsettings.json`.
- Possibilità di personalizzare e modificare il codice.

---

## 🛠 Tecnologie utilizzate
- **.NET 10.0**
- **C# 11**
- **Worker Service / Windows Service**
- **API Win32 per gestione Print Spooler**
- **WMI (opzionale)**

---

# 👨‍💻 Sezione Sviluppatore

Questa sezione è pensata per chi vuole compilare, modificare o estendere il progetto.

### 1️⃣ Requisiti
- Visual Studio 2026 con carico di lavoro **Sviluppo applicazioni desktop .NET**
- .NET 10.0 SDK installato
- Permessi amministrativi per eseguire test su servizi Windows

### 2️⃣ Struttura del progetto
```
/ServizioAntiCopieMultiple
│
├── Program.cs
├── PrintMonitorWorker.cs
├── appsettings.json
├── LICENSE
└── README.md
```

### 3️⃣ Configurazione del progetto
Il file `appsettings.json` permette di configurare:

```json
{
  "PrintMonitorSettings": {
    "IntervalSeconds": 5,
    "Printers": ["StampanteUfficio1", "StampanteUfficio2"],
    "NotifyUser": true,
    "LogLevel": "Information"
  }
}
```

- `IntervalSeconds`: intervallo di controllo della coda di stampa  
- `Printers`: stampanti da monitorare  
- `NotifyUser`: true per abilitare popup/notifiche  
- `LogLevel`: livello di log (Information, Warning, Error)

### 4️⃣ Compilazione
1. Apri il progetto in Visual Studio 2026.
2. Seleziona **Build → Build Solution**.
3. In alternativa, puoi usare **Build → Publish → Folder** per ottenere l’eseguibile.

### 5️⃣ Personalizzazione
- Modifica `PrintMonitorWorker.cs` per aggiungere regole o comportamenti personalizzati.
- Il codice è rilasciato con **CC BY-NC-SA 4.0**:
  - Obbligo di attribuzione  
  - Non commerciale  
  - Le versioni modificate devono mantenere la stessa licenza

---

# 🖥 Sezione Utente

Questa sezione è pensata per chi vuole installare il servizio su Windows.  
Puoi usare le **release già compilate** oppure compilare tu stesso.

### 1️⃣ Installazione del servizio (da eseguibile)
Apri un prompt dei comandi come amministratore ed esegui:

```powershell
sc create ServizioAntiCopieMultiple binPath= "C:\Percorso\alla\pubblicazione\ServizioAntiCopieMultiple.exe"
```

### 2️⃣ Avvio del servizio
```powershell
sc start ServizioAntiCopieMultiple
```

### 3️⃣ Stop del servizio
```powershell
sc stop ServizioAntiCopieMultiple
```

### 4️⃣ Rimozione del servizio
```powershell
sc delete ServizioAntiCopieMultiple
```

### 5️⃣ Uso tramite release GitHub
- Scarica la release dalla sezione [Releases](https://github.com/tuoaccount/ServizioAntiCopieMultiple/releases)  
- Estrai l’eseguibile in una cartella di tua scelta  
- Segui i passaggi precedenti per creare il servizio Windows

---

## 📄 Licenza

**Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International (CC BY-NC-SA 4.0)**  
Copyright © **Nicola Cantalupo, 2024**

- Modifiche permesse, ma devono mantenere la stessa licenza.  
- Obbligo di attribuzione.  
- Uso commerciale vietato.  
- Testo completo: [https://creativecommons.org/licenses/by-nc-sa/4.0/](https://creativecommons.org/licenses/by-nc-sa/4.0/)

---

## 🏗 Roadmap suggerita
- Supporto multi-stampante in parallelo  
- Notifiche avanzate (Toast Notification Windows 10/11)  
- Configurazione remota tramite JSON centralizzato o database  
- Dashboard web per monitorare i job in tempo reale  

---

## 👤 Autore
**Nicola Cantalupo**  
Progetto creato nel 2024

---

## 🙏 Ringraziamenti
Grazie a chi contribuisce al progetto o lo utilizza come base per soluzioni aziendali o educative.

