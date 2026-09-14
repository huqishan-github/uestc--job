# UESTC Job Monitor maintenance instructions

This repository is a long-running recruitment monitor for the public UESTC career website.

## Non-negotiable rules

1. Never delete historical recruitment data merely to simplify a migration or rerun.
2. Never invent recruitment facts. Unknown fields remain empty.
3. Never invent or guess detail URLs. A detail URL must come from a real list/API record and the routing behavior used by the live site, and must be requested for verification.
4. Always read `data/jobs.csv` and the existing Excel workbook before scraping or merging.
5. Never append an existing unique ID as a duplicate row; update meaningful changed fields in place.
6. Save explicitly listed jobs from one announcement as separate rows.
7. Do not rewrite `data/UESTC招聘信息.xlsx` when no data changed and the workbook is healthy.
8. A site or parsing failure must never clear or replace historical data with an empty dataset.
9. Never commit or print Secrets, tokens, cookies, passwords, credentials, or API keys.
10. If push fails, preserve the local commit and clearly report that remote persistence failed.
11. Do not destabilize working scraping logic merely for stylistic refactoring. Verify live site behavior first.
12. The only allowed live-update entrypoint is `python scripts/update_jobs.py`; do not run or search for a root-level `fetch_jobs.py`.
13. Treat a non-zero exit from `python scripts/update_jobs.py` as a hard stop: do not commit or push generated data.
14. Before any data commit, verify that the post-run CSV contains every pre-run unique ID, the record count has not decreased, the Excel workbook opens successfully, and CSV/Excel unique-ID sets are identical.
15. Daily data commits must stage only `data/` and `reports/`. Never include `.github/workflows/` in the same commit as a recruitment-data update.
16. Never use force push for daily updates. Re-read the latest `main` before persisting and require a normal fast-forward.

## Safe maintenance workflow

Check Git status and preserve unrelated user changes. Read `AGENTS.md`, `config/config.json`, `data/jobs.csv`, and the existing Excel workbook. Run `python -m pytest -q` before a live update, then run only `python scripts/update_jobs.py`. The update wrapper performs fail-closed pre-run and post-run integrity checks. If it exits non-zero, preserve the last healthy repository state and do not stage generated data. If it succeeds, independently validate CSV/Excel counts and unique IDs, inspect the diff, then stage only `data/` and `reports/` and commit only when there is a real data change. Never use `git reset --hard`, force push, bypass login/verification, or increase request frequency aggressively.
