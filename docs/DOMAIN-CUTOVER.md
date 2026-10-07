# Moving Lab Shift Roster to its own domain

Recommended domain (checked available via RDAP on 2026-10-07): **labshiftroster.com**,
with `.co.uk` and `.io` registered defensively so nobody else can grab the UK or
developer-facing variant and redirect confused visitors elsewhere.

This cutover is a two-sided move: this repo (the free local-first app) gets its
own CNAME, and the separate `lab-shift-roster-cloud` repo serves the hosted team
tier from a subdomain. Nothing here should be done until the domain is actually
registered and the DNS records below resolve -- doing it earlier risks a window
where the live tool is unreachable from its existing links.

## Step 1 -- Register the domain

You do this part; I cannot purchase domains. Register:
- `labshiftroster.com` (primary)
- `labshiftroster.co.uk` and `labshiftroster.io` (defensive -- point these at the
  .com via your registrar's forwarding, or just park them; either is fine)

## Step 2 -- DNS records

At your registrar, for `labshiftroster.com`, add:

| Type  | Host | Value                              | Notes                                   |
|-------|------|-------------------------------------|------------------------------------------|
| CNAME | `@` or apex via ALIAS/ANAME | `dmamphey.github.io` | The free app (GitHub Pages). If your registrar doesn't support a CNAME at the apex, use its ALIAS/ANAME equivalent, or four `A` records to GitHub Pages' IPs (185.199.108.153, .109.153, .110.153, .111.153) plus an `AAAA` set for IPv6 -- GitHub's own custom-domain docs list the current addresses. |
| CNAME | `www` | `dmamphey.github.io` | So `www.labshiftroster.com` also resolves |
| CNAME | `app` | `<your-worker>.workers.dev` initially, then the custom domain Cloudflare issues once attached | The hosted team tier (Cloudflare Worker). Exact target comes from Cloudflare's dashboard when you add a custom domain to the Worker -- it will tell you precisely what to set. |

Wait for DNS to actually resolve (`dig labshiftroster.com` / `dig app.labshiftroster.com`)
before Step 3. This can take minutes to a day depending on the registrar.

## Step 3 -- Point this repo at the new domain

Once DNS resolves, two commits (I will do these once you confirm DNS is live):

1. Add a `CNAME` file containing `labshiftroster.com` to this repo's root.
   GitHub Pages treats a project repo with its own `CNAME` file as having an
   independent custom domain, rather than being served as a path under
   `tools.optymumss.com`'s site. This is the actual cutover moment.
2. Update every in-app reference from the `tools.optymumss.com/lab-shift-roster/`
   path to `labshiftroster.com`: `index.html` (meta tags, canonical links, any
   hardcoded absolute URLs), `user-guide.html`, `README.md`, `CHANGELOG.md`, and
   the generated-workbook footer text in `labroster/template.py` /
   `labroster/export.py` if either hardcodes the old path.

## Step 4 -- Update the Optymum SS site

Per your decision: replace the Lab Shift Roster card/links on
`tools.optymumss.com` with a link out to `labshiftroster.com`, rather than
removing the product from the site entirely. I have the exact six-line diff
prepared (see `dmamphey.github.io` repo, not committed yet) -- it swaps every
`https://tools.optymumss.com/lab-shift-roster/` URL for
`https://labshiftroster.com/` and keeps the card copy as-is.

**Do this only after Step 3 is live**, so the Optymum SS site never links to a
domain that isn't serving yet.

## Step 5 -- Cloudflare Worker custom domain

In the Cloudflare dashboard, under the `lab-shift-roster-cloud` Worker ->
Settings -> Domains & Routes, add `app.labshiftroster.com` as a custom domain.
Cloudflare provisions the TLS certificate automatically once the CNAME in Step 2
resolves. Update `wrangler.toml`'s `APP_ORIGIN` var and the Worker's CORS origin
to match (currently a placeholder of `http://localhost:5173` for local dev).

## What does NOT change

- `tools.optymumss.com`'s DNS and GitHub Pages custom-domain configuration
  (the Optymum SS multi-product site) -- untouched.
- The `dmamphey.github.io` repository's own CNAME.
- Nothing about how the free local app processes data -- it is still 100%
  client-side regardless of which domain serves the static files.
