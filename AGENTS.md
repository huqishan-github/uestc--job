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

## Safe maintenance workflow

Check Git status and preserve unrelated user changes. Run `python -m pytest -q` before a live update, then run `python scripts/update_jobs.py`. Validate CSV/Excel counts and unique IDs, inspect the diff, and commit only relevant project paths when there is a real code or data change. Never use `git reset --hard`, force push, bypass login/verification, or increase request frequency aggressively.
