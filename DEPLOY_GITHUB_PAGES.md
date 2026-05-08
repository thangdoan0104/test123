# Deploy to GitHub Pages

Important: the GitHub repository page (`https://github.com/...`) shows source code and README only. The public app URL is `https://YOUR_USERNAME.github.io/REPO_NAME/` after the GitHub Actions deployment succeeds.

## 1. Push the project with Git CLI

Do not upload only the `.zip` file to GitHub. Unzip this package, open the project folder, then push the files inside it:

```bash
cd invitation-card-editor-mvp
git init
git add .
git commit -m "Deploy invitation card editor"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

## 2. Enable GitHub Pages Actions

In the GitHub repo:

```txt
Settings -> Pages -> Source -> GitHub Actions
```

Then open:

```txt
Actions -> Deploy Next.js static site to GitHub Pages
```

Run or wait for the workflow. When it becomes green, open:

```txt
https://YOUR_USERNAME.github.io/YOUR_REPO/
```

## 3. Why this workflow installs native packages

The app uses Fabric.js. Fabric's npm dependency tree can involve `canvas`, which may need Cairo/Pango native build packages on Ubuntu runners. The workflow installs these packages before `npm install`, then builds a static Next.js export into the `out` folder.

## 4. Common issues

- If GitHub shows the README, you are opening the repo URL, not the GitHub Pages URL.
- If GitHub diff says "contents could not be loaded", it is usually just the web diff viewer failing to preview a large file. It does not mean the file was not committed.
- If the app opens but CSS/JS is missing, check that `next.config.mjs` contains `basePath` and `assetPrefix` for GitHub Pages project URLs.
- If Actions fails, open the failed job and expand the red step. The important error is usually in `Install dependencies` or `Build static Next.js site`.
