# bulkatl — FBR Active Taxpayer List, refreshed weekly

This repo's only job is to keep a small, fast lookup file of every active income-tax taxpayer in Pakistan up to date. It backs the **Bulk ATL Checker** tool on [fairtaxint.org](https://fairtaxint.org).

## How it works

Every Monday at 04:00 UTC (09:00 PKT), [a GitHub Action](.github/workflows/refresh-atl.yml) downloads the official ~160 MB ATL xlsx from the FBR, extracts the NTNs and CNICs from all 20 sheets, sorts and dedupes them, and publishes the result as a ~30 MB gzipped file on a GitHub release.

The web app downloads that small file once a week and does in-memory binary-search lookups against it.

## Releases

- **`atl-latest`** — always points at the most recent build. The web app fetches from this tag.
- **`atl-YYYY-MM-DD`** — dated snapshot, kept for history.

## Manual refresh

Actions tab → "Refresh FBR ATL" → "Run workflow".

## Source

- FBR ATL landing page: <https://www.fbr.gov.pk/categ/active-taxpayer-list-income-tax/51147/30859/71169>
- Direct file: <https://download.fbr.gov.pk/IT/ATL_IT.xlsx>
