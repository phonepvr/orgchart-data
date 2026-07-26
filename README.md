# AM/NS Org Sense

A single-page org-chart explorer for AM/NS India. Upload an Excel file with
the employee template and browse the structure, filter/sort a table view,
compare people side by side, and print A4-landscape position maps.
All parsing happens in the browser — nothing is uploaded anywhere.

## No build step

The app is **plain HTML + vanilla JavaScript + plain CSS**. The repo *is*
the site: edit a file, push to `main`, and GitHub Pages serves it directly
(the workflow in `.github/workflows/deploy.yml` just packages the checkout —
no npm, no compiler, no dependencies to install).

To run locally, serve the folder with any static server, e.g.:

```
python3 -m http.server 8000
```

then open http://localhost:8000/.

## File map

```
index.html                   page shell (CSP, fonts, overlay layers, script order)
css/app.css                  the full stylesheet (see "About the CSS" below)
vendor/xlsx.full.min.js      SheetJS 0.18.5 standalone (Excel parsing; only
                             XLSX.read + sheet_to_json are used)
orglens_sample_template.xlsx sample data template offered on the upload screen
js/
  harden.js                  prototype freeze — must load before the XLSX vendor script
  main.js                    boot, delegated event handling, actions, print lifecycle
  state.js                   central app state + derived data + region re-render dispatch
  constants.js               template schema, status styles, filter fields, print caps
  data.js                    parsing, normalization, insights graph, print pagination
  filters.js                 filter/sort/cohort/benchmark computations
  icons.js                   inline SVG icon map (extracted from lucide 0.577.0)
  util.js                    HTML escaping
  render/                    one module per UI area; each builds an HTML string
    shell.js                 lock screen, upload screen, header, sidebar, filter pills
    chart.js                 employee cards + org chart layout
    table.js                 sortable table view
    compare.js               compare view (5 color slots)
    spotlight.js             spotlight tooltip + benchmark scales
    print.js                 A4-landscape print pages
    overlays.js              tooltips + right-click context menu
    bits.js                  chips, avatar, brand marks
```

How it renders: one `state` object; every user action mutates it and
re-renders the affected region's container (`innerHTML` rebuild). At the
data sizes this app handles (hundreds of rows) that is single-digit
milliseconds — there is deliberately no framework and no DOM diffing.
Events are delegated at the document level via `data-action` attributes,
so re-renders never need listener re-wiring.

## About the CSS

`css/app.css` is the frozen output of the Tailwind build the app used
before it was converted to a no-build stack (generated at the commit that
introduced it, via `npm run build` with the 5 dynamic
`hover:ring-{color}-400` classes safelisted, then copied from
`dist/assets/index-*.css`). The markup keeps Tailwind's utility class
names, so **only class names that already appear in `css/app.css` will
have any effect**. If you need a class that is missing, append a small
hand-written rule at the end of `css/app.css`.

## Changing the access password

The lock screen compares a SHA-256 hash. Set a new one in
`js/constants.js` (`ACCESS_HASH`):

```
node -e "console.log(require('crypto').createHash('sha256').update('NEWPASS').digest('hex'))"
```

Note this is client-side obfuscation to keep casual visitors out of a
public Pages URL — not real authentication.

## Deploying

Push to `main`. The Pages workflow uploads the repo as-is and deploys it.
`workflow_dispatch` is enabled for manual re-deploys.
