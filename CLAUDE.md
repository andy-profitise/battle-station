# Claude Code Notes

## Git Workflow

- **Always fetch and merge main before starting work** on a feature branch: `git fetch origin main && git merge origin/main`. Other sessions may have merged changes concurrently.
- **Always fetch and merge main again before pushing**, to catch anything that landed while you were working. This avoids merge conflicts at PR time.
- This repo has multiple concurrent Claude Code sessions working on it. Never assume main is static.
