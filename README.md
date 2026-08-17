# SCU · PHX Metro Space Planner

Live scenario modeling for SCU's Phoenix-metro campus footprint (Tempe / Scottsdale) across DC, DPT, OTD, SLP, PA, and GC — with the ARC-PA October 2027 site-walk deadline front and center.

Built with React + Vite. No backend — all data is computed client-side and saved to the browser's local storage, so each viewer has their own working copy.

## Run locally

```bash
npm install
npm run dev
```

Open the URL Vite prints (usually http://localhost:5173).

## Put it on GitHub

```bash
git init
git add .
git commit -m "SCU PHX Metro Space Planner"
git branch -M main
git remote add origin https://github.com/<your-org>/scu-phx-space-planner.git
git push -u origin main
```

(Create the empty repo on GitHub first — set it to **Private**.)

## Deploy to Netlify

1. In Netlify: **Add new site → Import an existing project → GitHub**, pick the repo.
2. Netlify reads `netlify.toml` automatically (build `npm run build`, publish `dist`). Just click **Deploy**.
3. Every push to `main` redeploys automatically.

## Password protection

Options, simplest first:

- **Netlify site-wide password** — Site configuration → Access & security → **Site protection** → set a password. Note: this feature generally requires a paid Netlify plan; check your plan's current inclusions since Netlify adjusts tiers periodically.
- **Netlify Identity / OAuth gate** — invite-only email logins; more setup but works on more plans.
- **Keep the GitHub repo private regardless** — the site URL is only as private as the people you share it with, so pair any option with an unguessable Netlify site name until protection is on.

The app also ships with `noindex, nofollow` meta so search engines won't index it.

## Where the numbers live

- `src/App.jsx` → `defaultData()` holds the program space model (room-level NSF), the ×1.45 GSF multiplier, site capacities (Sonoran = 10,062 RSF), lease rates (Sonoran $28.71/RSF/yr, Scottsdale $45/SF/yr), and the four preset scenarios.
- Everything is editable in the UI; edits persist per-browser. "Reset to defaults" restores this file's values.

## Modeling rules (as implemented)

- **GSF** = program NSF × multiplier (default 1.45, covering restrooms, hallways, storage, faculty offices). Anatomy is flat GSF.
- **Surge Site** is time-shared turnkey rental driven by DPT running two cohorts' labs simultaneously (one week, three times per year): programs there count flat lab SF (no multiplier), and the site's footprint is the single largest need, not the sum. Toggle it off per scenario to model absorbing that need into a larger permanent site (e.g., taking a whole floor) — displaced programs then require full-GSF permanent space.
- **Shared core** (per-scenario toggle): each program has a shareable % of its base NSF. Co-located sharers count the shared slice once (largest need wins); the multiplier remainder stays dedicated per program.
- **Anatomy Lab** can be included/excluded per scenario.
- **ARC-PA**: header counts down to Oct 1, 2027; scenarios flag when PA is unassigned or parked at an over-capacity site.
