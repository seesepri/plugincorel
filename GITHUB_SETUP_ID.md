# Corel Page Copier — GitHub-ready package

This package adds a GitHub Actions workflow to build the VB.NET Windows Forms application for .NET Framework 4.8, x64, and publish a ZIP to GitHub Releases when a version tag (`v*`) is pushed.

## Repository layout

- `CorelPageCopier_MODERN/` — source project
- `.github/workflows/build-release.yml` — automatic build/release workflow
- `.gitignore` — excludes build output and local files

## Add to your existing private repository safely

1. Back up or create a branch in your existing repository before replacing files.
2. Extract this ZIP into a temporary folder.
3. Copy `CorelPageCopier_MODERN/`, `.github/`, and `.gitignore` into the repository root. If your existing source folder has another name, keep the path expected by the workflow or update the `project` path in the YAML.
4. Commit and push to `main` (or `master`). GitHub Actions should build the project and expose a downloadable workflow artifact.
5. To publish a GitHub Release, create and push a version tag, for example:

   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```

6. Open the repository's **Actions** tab and confirm the workflow succeeds. Then open **Releases** and verify `CorelPageCopier-Windows-x64.zip` is attached.

## Required GitHub settings

- In **Settings → Actions → General → Workflow permissions**, allow the workflow to read and write repository contents. The workflow uses the repository's `GITHUB_TOKEN` to create releases; no personal access token should be committed.
- Keep the repository **Private**. Release assets are still subject to repository access permissions.
- Do not commit credentials, personal access tokens, private documents, or customer CDR files.

## Important compatibility and testing notes

- Target framework remains **.NET Framework 4.8**, x64, intended for Windows and CorelDRAW 2022.
- GitHub Actions can compile the source on a Windows runner, but it cannot verify COM connection to the user's installed CorelDRAW 2022 instance or validate real object-copy behavior.
- Test the released executable on a copy of a CDR document before using it on important files.
- This workflow creates a build artifact on every push to `main`/`master`; it creates a GitHub Release only when a `v*` tag is pushed.
- The workflow has not been run against your GitHub repository because its URL/access was not provided and no GitHub connection is available in this session.
