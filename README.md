# Invitation Card Editor MVP

## Public deployment

For GitHub Pages, this package already includes `.github/workflows/deploy.yml` and static export config in `next.config.mjs`. See `DEPLOY_GITHUB_PAGES.md`.

Quick public URL after deploy:

```txt
https://YOUR_USERNAME.github.io/YOUR_REPOSITORY_NAME/
```

> The normal GitHub repository page only shows README.md. Use the GitHub Pages URL above to open the real app.


A client-side Next.js app for editing invitation card images and generating personalized cards in bulk from Excel or CSV data.

## Features

- Upload one or many invitation images. Each image is a template.
- Upload custom fonts: TTF, OTF, WOFF, WOFF2.
- Add editable text fields on top of the card.
- Drag, rotate, resize, and edit text directly on the canvas.
- Edit text style: font, solid color, gradient color, stroke, shadow, size, opacity, position, alignment, spacing, line height, and text box width.
- Auto layout groups: link multiple text fields such as `Ong` + `Name`, set order/gap/anchor, and prevent overlap when Excel values change.
- Auto letter spacing mode: fit text to the configured text box width per Excel row.
- Classify fields as `Ten`, `Bi danh`, `Chuc vu`, or custom.
- Configure màu / gradient / stroke / shadow presets theo từng loại field.
- Upload `.xlsx` or `.csv` guest data.
- Map each text field to an Excel column.
- Generate batch previews.
- Export high quality JPEG. If there are many outputs, the app creates a ZIP.
- Export high quality PDF with one page per generated card.
- Save and reopen project JSON, including template images, fonts, text fields, styles, and mapping.
- Save and load project from browser localStorage.

## Tech stack

- Next.js App Router
- React
- Fabric.js canvas
- SheetJS `xlsx` for Excel / CSV parsing
- jsPDF for PDF export
- JSZip for multi-image ZIP export
- FileSaver.js for client-side downloads

## Run locally

```bash
npm install
npm run dev
```

Open http://localhost:3000

## Suggested workflow

1. Upload 1 or more card images.
2. Upload custom fonts if needed.
3. Click `Add text field` and position text on the card.
4. Upload an Excel or CSV file. The first row must contain column names.
5. Select each text object and choose its `Field type` and `Excel column`.
6. For separate fields that should stay together, give them the same `Auto layout` group ID, set `Order` values, and tune `Gap`. For example: honorific `Ong` order 1, guest name order 2.
7. Use `Box width` + align buttons for real left/center/right alignment. Use `Fill type -> Gradient` for gradient text.
8. Click `Generate preview`.
9. Export JPEG, ZIP, or PDF.
10. Click `Download project` to save a `.json` project file for later.

## Example CSV

```csv
Name,Alias,Title
Nguyen Van A,Anh A,CEO
Tran Thi B,Chi B,Marketing Manager
Le Minh C,Minh C,Partner
```

## Notes and next improvements

This is an MVP. For very large batches, move rendering to a Web Worker or a backend queue to avoid blocking the browser. For production, also add authentication, cloud storage, project versioning, font licensing checks, template thumbnails, undo/redo, snap lines, and optional print bleed / crop mark settings.
