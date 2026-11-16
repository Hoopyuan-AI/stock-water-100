# 市值前100大｜股價五水溫公開儀表板

This project is adapted from the user's working Windows V3 stock-water analysis program.

## What it does
- Retrieves TAIFEX top-100 market-cap ranking and Yahoo Finance daily history.
- Uses 10 trading-day, 10 weekly, and 10 monthly averages plus directional trends.
- Classifies into 滾水、熱水、溫水、冷水、冰水、過渡區; labels insufficient/failed data.
- Produces static `index.html` and downloadable `water-report.xlsx` on GitHub Pages.
- Records historical trend CSV files under `water_output/data` and commits their changes.
- Runs at **15:10 Asia/Taipei Monday–Friday** (market holidays have no new data); can also run manually.

## Publish on GitHub
1. Create a **public** GitHub repository named `stock-water-100` (or your chosen name) with default branch `main`.
2. Upload this repository's **contents** using GitHub Desktop or Git. Hidden `.github/workflows/update-and-publish.yml` must be included. Example:
   ```bash
   git init
   git branch -M main
   git add .
   git commit -m "Initial water monitor cloud version"
   git remote add origin https://github.com/YOUR_USER/stock-water-100.git
   git push -u origin main
   ```
3. Settings → Pages → Build and deployment → Source: **GitHub Actions**.
4. Settings → Actions → General → Workflow permissions: enable **Read and write permissions** for the history commit step. Repository/org policies may override this. If read/write is unavailable, skip the history commit step or use an approved persistence method.
5. Actions → **Refresh Taiwan Top 100 Water Dashboard** → **Run workflow** (main).
6. When it succeeds, find the permanent URL under Settings → Pages, usually `https://YOUR_USER.github.io/stock-water-100/`.

### Troubleshooting
- Workflow failure: Actions → failed run → check **Download rankings and prices** logs. Source websites may rate-limit cloud-hosted runners even when Windows works.
- `403` at deployment: confirm Pages is enabled, the GitHub Actions source is selected, and workflow permissions are granted.
- `git push` denied: Settings → Actions → General → Workflow permissions → Read and write.
- If latest data are missing, the existing public site keeps showing the previously successful publication; check **行情基準日** before relying on the report.
- Public Pages sites are readable by anyone; do not commit sensitive information. Historical CSV files are public in this repository.
- Scheduled workflows can be delayed; in inactive public repositories they can be disabled after 60 days without repository activity.
- Yahoo Finance and TAIFEX terms/rate limits should be reviewed for public redistribution. This is market technical analysis, not trading advice.

### Local test (without external network)
`python stock_water_100.py --self-test`

The existing Windows startup scripts are not needed on GitHub Actions.
