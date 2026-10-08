# Moving Lab Shift Roster to its own domain

**labshiftroster.com** is registered (purchased on Cloudflare, 2026-10-07) and is the
permanent home of the Lab Shift Roster product -- not a subdomain, not a path under
`tools.optymumss.com`, and not GitHub Pages.

## Architecture (what actually exists now -- read this before touching DNS)

This cutover plan was originally written assuming a two-host split: GitHub Pages
serving the free app at the apex domain, with a Cloudflare Worker only on an
`app.` subdomain for the hosted tier. **That is no longer the architecture.**
As of the `lab-shift-roster-cloud` repo's Milestone 1-3 build, one Cloudflare
Worker serves the entire product from a single origin:

```
labshiftroster.com
        |
   Cloudflare Worker (lab-shift-roster-cloud)
        |
        +-- /                   public marketing pages (home, pricing, contact)
        +-- /app/*              the Free app's static files, served unmodified
        |                       via Workers Static Assets (synced from this repo
        |                       by lab-shift-roster-cloud/scripts/sync-free-app.mjs)
        +-- /login, /signup     authenticated account pages (hono/jsx, server-rendered)
        +-- /account/*          Pro/Enterprise workspace: team, configuration,
        |                       workflows, audit history, sites, reporting, billing
        +-- /auth/*, /orgs/*,   the JSON API behind all of the above
        |   /billing/*, /contact
        +-- D1                  organisations, users, subscriptions, sites, workflows
        +-- Stripe              Pro self-serve billing + Enterprise sales-led billing
```

There is no GitHub Pages involvement in serving `labshiftroster.com` at all. This
`lab-roster` repo remains the **source of truth** for the Free app's code (the
Python roster engine, the vendored Pyodide runtime, `index.html`) -- the Worker
just copies those files in at deploy time and serves them unmodified. Nothing
about how the Free app processes data changes: it is still 100% client-side,
regardless of which domain or which server serves the static files.

## Step 1 -- Domain registration: done

`labshiftroster.com` is registered and managed on Cloudflare (the same account
that will host the Worker), which is what makes attaching it as a Worker custom
domain a dashboard click rather than external DNS-provider configuration -- see
Step 3. The `.co.uk`/`.io` defensive registrations discussed earlier were not
pursued; revisit only if squatting becomes an actual problem.

## Step 2 -- Deploy the Worker (before touching DNS)

In `lab-shift-roster-cloud`:
1. `wrangler d1 create lab_shift_roster_cloud` (if not already done) and paste
   the database id into `wrangler.toml`.
2. `npm run db:migrate:remote`.
3. Set secrets: `STRIPE_SECRET_KEY`, `STRIPE_WEBHOOK_SECRET`,
   `STRIPE_PRICE_ID_PRO_MONTHLY`, `STRIPE_PRICE_ID_PRO_ANNUAL`, `SESSION_SECRET`,
   `RESEND_API_KEY` (`STRIPE_PRICE_ID_ENTERPRISE` only if a canonical Enterprise
   Price is in use -- Enterprise is sales-led and often doesn't need one, see
   that repo's README).
4. Update `wrangler.toml`'s `APP_ORIGIN` var to `https://labshiftroster.com` and
   `CONTACT_FROM_EMAIL` to a Resend-verified address once that domain is
   verified in Resend.
5. `npm run deploy` (this runs `sync:free-app` first, pulling this repo's
   current files into the Worker's static assets).
6. Confirm the Worker serves correctly on its `workers.dev` subdomain --
   including `/app/` actually running the Pyodide engine -- before attaching
   the real domain in Step 3.

## Step 3 -- Attach labshiftroster.com to the Worker

In the Cloudflare dashboard: the `lab-shift-roster-cloud` Worker -> Settings ->
Domains & Routes -> Add Custom Domain -> `labshiftroster.com` (and optionally
`www.labshiftroster.com`, redirecting to the apex). Because the domain is
already on this Cloudflare account, this single step both creates the
necessary DNS records and provisions the TLS certificate -- there is no
separate registrar-side DNS configuration to do, unlike the old two-host plan.

This is the actual go-live moment: once attached, `labshiftroster.com` serves
the Worker directly, with no window where the domain points at GitHub Pages or
anywhere else first.

**This step has not been performed.** It is a deliberate, manual, one-time
action for the account owner to take when ready -- not something any
automated process should do on its own.

## Step 4 -- Update the Optymum SS site's outbound links

`dmamphey.github.io` already has a prepared, unmerged branch
(`lab-shift-roster-new-domain`, draft PR #1, "Point Lab Shift Roster at
labshiftroster.com (blocked on DNS)") that swaps the three
`tools.optymumss.com/lab-shift-roster/` links (hero button, tool card, footer)
for `https://labshiftroster.com/`. That branch does **not** touch
`dmamphey.github.io`'s own `CNAME` or any GitHub Pages hosting configuration
for `labshiftroster.com` -- it only changes where visitors clicking through
from the Optymum SS site land. It is still correct and compatible with this
architecture and does not need rewriting; merge it once Step 3 is live, not
before, so the Optymum SS site never links to a domain that isn't serving yet.

## What does NOT change

- `tools.optymumss.com`'s own DNS and GitHub Pages custom-domain configuration
  (the Optymum SS multi-product site) -- untouched.
- The `dmamphey.github.io` repository's own `CNAME` (still `tools.optymumss.com`).
- This repo (`lab-roster`) itself is not deployed anywhere directly -- it is
  synced into the Worker's static assets at `lab-shift-roster-cloud` deploy
  time, same as before.
- Nothing about how the Free app processes data -- still 100% client-side.
