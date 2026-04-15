# TravelCore RM Hub — Claude Working Notes

## Workflow
After **every set of code changes**, always:
1. `git add .`
2. `git commit -m "description"`
3. `git push`
4. `vercel --prod --yes`

Or run `./deploy.sh "commit message"` which does all four steps.

## Repo
- GitHub: https://github.com/marieduetto123/travelcore-rm-hub
- Vercel: https://travelcore-rm-hub.vercel.app
- Working dir: /Users/marie/Melia

## Files
- `travelcore-rm-hub.html` — markup / tab structure
- `travelcore-rm-hub.css`  — all styles
- `travelcore-rm-hub.js`   — all logic (AG Grid, charts, data)

## Key implementation notes
- AG Grid Community v23.2.1 (`_realAgGrid`) — uses v23 API: `new agGrid.Grid(el, opts)`, `opts.api.setRowData()`, theme `ag-theme-alpine`
- Daily H: flat-row accordion (fullWidthRow for sections, custom cellRenderer for parents)
  - `_dhCollapsed` persists state between grid rebuilds
  - `dhSetAll(collapse)` — Open All / Close All
  - `autoHeight: true` on defaultColDef — rows auto-size to content
- Vercel project ID: prj_NsOYsSZ4pDONmV2BgnB6DrJPhlGR
