# NegosyoSheet — Inventory, Utang at Profit Tracker (PH sari-sari stores & resellers)

**Live:** https://makavelimachiavelli.github.io/negosyosheet/

## What it is
A polished Excel/Google-Sheets workbook (₱149 one-time) that runs a small PH store's whole back-office: inventory with auto stock counts and ⚠️ REORDER alerts, a sales log with VLOOKUP pricing and per-sale profit, a utang (credit) list with overdue alerts, expenses, and an auto DASHBOARD (sales / gross profit / net / inventory value / utang unpaid / reorder count / late utang count). Instructions in Taglish. Free 5-row demo as lead magnet.

## Buyer persona
- **Who:** sari-sari store owners, online resellers (Shopee/Lazada/Carousell/FB Marketplace), small carinderia/variety stores — the PH micro-retail backbone (100k+ stores).
- **Pain:** tracking stock/utang/profit in notebooks or bare spreadsheets; no cheap tool in their language/context; POS apps need hardware + subscriptions.
- **Why pay ₱149:** less than a day's profit; converts a night of manual Excel setup into instant; utang-tracking alone prevents losses that dwarf the price.
- **Where they hang out:** FB groups (Sari-Sari Store Owners PH, Online Sellers PH), TikTok negosyo content, Carousell.

## Demand evidence (per REVENUE GATES)
- Marketplace listings: Etsy/Gumroad "inventory tracker / small business bookkeeping" spreadsheets = hundreds of paid listings at $3–15, with sellers reporting $300–800/mo (Sellfy 2026 guide) and $12k from a single reseller-organizer sheet (Hustle & Slow case). Passes the ≥10-listings gate with room to spare.
- Adjacent paid: POS/inventory SaaS (Laybare/StoreHub ₱ subscriptions), QuickBooks — all overkill for a ₱50k/month store.

## Monetization
Demo = free download on the page. Full version gated by unlock code → reveals download. GCash QR + code (see `PAYMENTS.md`); Gumroad/LemonSqueezy optional for cards. Static-gate limitation honestly documented (fine at this scale; upgrade path ready).

## Tech
- `build_sheet.py` regenerates both workbooks (openpyxl): formulas, dropdowns, conditional formatting, Taglish README sheet.
- Verified 2 ways: structural check (16/16) + **real formula evaluation** via the `formulas` engine (15/15 — stock math, VLOOKUP pricing, profit, dashboard rollups all compute correctly).
- Static landing page reuses the zinvent pay-block pattern.

## Deploy
```bash
../toolkit/deploy-pages.sh . negosyosheet
```
(regular repo: git init → gh repo create negosyosheet → push → Pages)

## Owner TODO (Allen, ~5 min)
Swap GCash QR (`gcash-qr.svg`), set `PRO_CODES` in `app.js`. At volume: move full file to Gumroad auto-delivery (then set `FULL_URL` or replace gate with their link).
