# PCF Card Order Agent

Ce dossier contient un projet Power Apps Component Framework (PCF) initialise manuellement.

## Structure

- `pcf/CardOrderAgent`: controle PCF

## Prerequis

- Node.js LTS (18 ou 20 recommande)
- npm
- Power Platform CLI (`pac`)
- .NET SDK (pour creer/packager une solution Dataverse)

## Installation

```powershell
npm install --prefix .\pcf\CardOrderAgent
```

## Build du controle

```powershell
npm run build --prefix .\pcf\CardOrderAgent
```

## Lancer le harness local

```powershell
npm start --prefix .\pcf\CardOrderAgent
```

## Installation des prerequis manquants (si besoin)

```powershell
winget install OpenJS.NodeJS.LTS
winget install Microsoft.DotNet.SDK.8
```

Pour `pac`, installez Power Platform CLI selon la documentation Microsoft (ou via VS Code Power Platform Tools).
