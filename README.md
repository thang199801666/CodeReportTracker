# Code Report Tracker

Code Report Tracker is a React and Tauri desktop application for managing code report links, checking PDF availability, updating report data, and organizing downloaded PDFs by tab.

## Features

- Tab-based report management with drag-and-drop reordering.
- Import `.crp`, `.xlsx`, and `.xls` files using header-based column detection.
- Check report links and identify missing PDFs.
- Update report information from remote or local PDFs.
- Download PDFs into a directory and a subfolder for each tab.
- Open a tab download folder directly in Windows Explorer when running as Tauri.
- Edit individual rows from the table context menu.
- Export all tabs to a styled Excel workbook.
- Native Tauri Save and Directory dialogs with real filesystem paths.
- Vercel PDF proxy functions for production deployments.

## Requirements

- Node.js 20 or newer.
- Rust and Cargo for Tauri development and desktop builds.
- Windows WebView2 for the Windows desktop build.

## Web Development

```bash
npm install
npm run dev
```

The Vite development server runs the PDF proxy middleware locally.

## Tauri Development

```bash
npm run tauri:dev
```

## Production Build

Build the web application:

```bash
npm run build
```

Build the Windows MSI installer:

```bash
npm run tauri:build
```

The installer is generated under `src-tauri/target/release/bundle/msi/`.

## Vercel Deployment

Deploy the repository as a Vite project. The `api/` directory contains the PDF proxy serverless functions used by production builds:

- `/api/pdf-head`
- `/api/pdf-download`
- `/api/pdf-resolve`

## License

See the project source and repository configuration for licensing details.
