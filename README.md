# Bidlisiw · PSHS-CVC Batch 2007

The batch website: **https://tenrenti01.github.io/bidlisiw/**

| File | What it is |
|---|---|
| `index.html` | The website: **Teachers** (main view), **Testimonials**, **Profiles & sections** |
| `config.js` | One setting: the Apps Script API address (`…/exec`) |
| `Code.gs` | Copy of the Apps Script that reads/writes the Google Sheet (paste it into the sheet's Apps Script) |

## How it fits together
The page is static (GitHub Pages). All data stays in the Google Sheet *Bidlisiw Project Database*;
`Code.gs`, deployed as a web app, is the only thing that touches it. The page talks to it with `fetch`.

## Setup (once)
1. Google Sheet ▸ Extensions ▸ Apps Script ▸ replace `Code.gs` with the one here. The old `Index.html` there can be deleted.
2. Deploy ▸ New deployment ▸ **Web app** · Execute as: **Me** · Who has access: **Anyone** ▸ Deploy ▸ copy the URL ending in `/exec`.
3. Paste that URL into `config.js` here (between the quotes) and commit.

After changing `Code.gs` later: Deploy ▸ Manage deployments ▸ ✏️ ▸ Version: **New version** — the URL stays the same.
