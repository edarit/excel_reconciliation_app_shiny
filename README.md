# Excel Comparator Lite

A high-performance, browser-based utility for reconciling Excel datasets. Built with **Shinylive**, this application runs entirely in the browser using Pyodide, requiring no backend server for deployment.

## Features

- **Multi-Column Reconciliation**: Compare up to 5 specific column pairs between two Excel files (e.g., Finance vs. CPUS).
- **Flexible Logic**: Support for "IN" (exists in both) and "NOT IN" (missing from reference) filtering modes.
- **Dynamic Selection**: Select specific sheets and header row offsets for each source file.
- **Visual Highlighting**: Color-coded results for clear identification of matched/unmatched pairs.
- **Professional Export**: Download results as a formatted `.xlsx` file with color-coded cells preserved.
- **Privacy First**: Files are processed locally in your browser; data never leaves your machine.

## Getting Started

### Local Development

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ````

2.  **Run the App**:
    ```bash
    shiny run app.py
    ```

### Deployment (GitHub Pages)

This project is configured for **Shinylive** deployment.

1.  **Build the static site**:
    ```bash
    shinylive export . site
    ```
2.  **Push to GitHub**: The included `.github/workflows/deploy-pages.yml` will automatically deploy the `site` folder to GitHub Pages on every push to the `main` branch.

## Lite vs Pro Version

This **Lite** version provides core deterministic matching.

| Feature | Lite | Pro |
| :--- | :---: | :---: |
| 5-Column Pairing | ✅ | ✅ |
| Excel Export | ✅ | ✅ |
| Fuzzy Matching | ❌ | ✅ |
| Automated Mapping | ❌ | ✅ |
| Large Dataset Optimization (\>100k rows) | ❌ | ✅ |

*For access to the Pro Version features, contact the development team.*

## Tech Stack

  - **UI**: [Shiny for Python](https://shiny.posit.co/py/)
  - **Data**: [Pandas](https://pandas.pydata.org/)
  - **Excel Engine**: [Openpyxl](https://openpyxl.readthedocs.io/) & [XlsxWriter](https://xlsxwriter.readthedocs.io/)
  - **Runtime**: [Shinylive](https://www.google.com/search?q=https://shiny.posit.co/py/docs/shinylive.html)
