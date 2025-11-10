# Repository Guidelines

## Project Structure & Module Organization
`standalone_basel2_wsl.py` drives the crawl, from Playwright navigation to Excel export; keep shared helpers close to the call sites so state flow stays transparent. Use `debug/` for temporary HTML or JSON captures that document selector updates. `output/` (symlinked to `../output`) stores generated Excel files such as `dolce_*` and `_second` variants, while `log.txt` is replaced on every run by the `Tee` logger—archive it before starting another job. The reference workbook `../naver_category.xlsx` must stay versioned with script changes.

## Build, Test, and Development Commands
Activate the virtualenv with `source .venv/bin/activate`. Install runtime needs explicitly: `pip install playwright playwright-stealth beautifulsoup4 openpyxl pandas python-dotenv requests`, then provision the browser once via `playwright install chromium`. Run the job with `python standalone_basel2_wsl.py`; adjust `global_start_page` and `global_last_page` during debugging to limit the crawl window. Stream logs in a second shell using `tail -f log.txt`.

## Coding Style & Naming Conventions
Use 4-space indentation and keep code `black`-friendly even if the formatter is not auto-run. Favor short modules: extract cohesive helpers (option parsing, product transforms) into nearby functions. Follow snake_case for functions and variables, UPPER_CASE for constants, and pick descriptive names (`verify_cdp_endpoint`, `crawl_page`). Keep the defensive logging style—log new selectors, retries, and early exits.

## Testing Guidelines
No automated suite exists yet; add lightweight smoke scripts under `debug/` when you touch complex flows. Before a PR, run `python standalone_basel2_wsl.py` on a narrow range (set both globals to 51) and confirm paired Excel files arrive in `~/Desktop/excel_output`. Spot-check rows to ensure deduping, option capture, and image harvesting still succeed, and scan `log.txt` for selector warnings.

## Commit & Pull Request Guidelines
Commits follow short, imperative Korean summaries (`셀링프라이스변경`); open with the affected area and keep the message under ~40 characters, expanding in the body only when needed. PRs should list purpose, crawl scope exercised, environment changes (new `.env` keys, ports), and attach screenshots or sample Excel diffs whenever selector work alters output shape.

## Security & Configuration Tips
Never commit `.env`; it holds `PLAYWRIGHT_CONNECT_URL` and other overrides. Strip customer data from debug files, and trim saved HTML before uploading. Any new configuration knob should be guarded by environment variables with documented defaults at the top of `standalone_basel2_wsl.py`.

