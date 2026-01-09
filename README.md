# NG-Report Generator

A comprehensive Angular application designed for generating industrial-standard Sprint Planning, Sprint Closure, and Monthly reports. This tool streamlines the reporting process by allowing data import from Excel and exporting to multiple formats including PDF, PowerPoint, and Word.

## 🚀 Features

*   **Dashboard**: Central hub for navigating to different report generators.
*   **Sprint Planning Module**: Create detailed estimation reports showing story points, resource allocation, and scope.
*   **Sprint Closure Module**: Generate closure reports analyzing committed vs. completed points, spilled points, and status breakdowns.
*   **Monthly Report Module**: specific reporting for monthly overviews.
*   **Excel Import**: Upload Excel (.xlsx) or CSV files to instantly populate report tables.
*   **Multi-Format Export**:
    *   **PDF**: Generates high-quality, print-ready PDF reports.
    *   **PowerPoint (.pptx)**: Exports editable slides for presentations.
    *   **Word (.doc)**: Exports editable documents in Print Layout mode.
    *   **Email Copy**: Formats the report as HTML compatible with Outlook and Gmail for direct pasting.
*   **Industrial Design**: Features a high-contrast Black, White, and Yellow aesthetic suitable for professional engineering and operations environments.

## 🛠️ Technology Stack

*   **Framework**: Angular 19 (Standalone Components)
*   **Language**: TypeScript
*   **Styling**: Vanilla CSS (Industrial Theme)
*   **Key Libraries**:
    *   `xlsx-js-style`: For reading Excel files with style support.
    *   `jspdf`: For PDF generation.
    *   `html2canvas` & `html-to-image`: For rendering DOM elements to images for export.
    *   `pptxgenjs`: For generating PowerPoint presentations.

## 📁 Project Structure

```
ng-report/
├── src/
│   ├── app/
│   │   ├── dashboard/          # Landing page component
│   │   ├── monthly-report/     # Monthly reporting logic and UI
│   │   ├── sprint-closure/     # Sprint closure reporting logic and UI
│   │   ├── sprint-planning/    # Sprint planning reporting logic and UI
│   │   ├── app.routes.ts       # Application routing configuration
│   │   └── ...
│   ├── index.html              # Main HTML entry point
│   └── styles.css              # Global styles
├── angular.json                # Angular CLI configuration
├── package.json                # Dependencies and scripts
└── README.md                   # Project documentation
```

## 💻 Getting Started

### Prerequisites

*   **Node.js**: Ensure you have Node.js installed (LTS version recommended).
*   **npm**: Comes with Node.js.

### Installation

1.  Clone the repository or download the source code.
2.  Navigate to the project directory:
    ```bash
    cd ng-report
    ```
3.  Install dependencies:
    ```bash
    npm install
    ```

### Running the Application

start the development server:

```bash
npm start
```

Navigate to `http://localhost:4200/` in your browser. The application will automatically reload if you change any of the source files.

### Building for Production

To build the project for production deployment:

```bash
npm run build
```

The build artifacts will be stored in the `dist/` directory.

## 📖 how to Use

1.  **Select a Report Type**: From the Dashboard, click on "Sprint Planning", "Sprint Closure", or "Monthly Report".
2.  **Import Data**:
    *   Click the "Upload Excel" button.
    *   Select your Data file. (See "Data Formats" below for structure).
3.  **Preview**: The application will parse the file and display a preview of the report.
4.  **Edit/Refine**: You can manually adjust the table data if needed (if editable fields are enabled).
5.  **Export**:
    *   Use **Download PDF** for a static document.
    *   Use **Download PPT** for a presentation slide.
    *   Use **Copy for Email** to paste the report into an email body.

## 📊 Data Formats needed for Excel Import

For the import to work correctly, your Excel file must use the following column headers.

### 1. Sprint Planning

**Required Columns:**

| Column Header | Description |
| :--- | :--- |
| `Work Item ID` | Unique ID (e.g., 12345) |
| `Work Item Type` | Type of task (e.g., User Story, Bug) |
| `Work Item` | Title/Summary of the item |
| `Scope` | Detailed description or acceptance criteria |
| `Resource` | Assigned developer/QA name |
| `Story Points` | Numeric estimation value |

### 2. Sprint Closure

**Required Columns:**

| Column Header | Description |
| :--- | :--- |
| `Work Item ID` | Unique ID |
| `Work Item Type` | Type of task |
| `Work Item` | Title/Summary |
| `Story Points` | Numeric value |
| `Status` | Current state (must use exact terms: 'Completed', 'In Progress', etc.) |
| `Dev Overview` | Comments or remarks on the item's progress |

## 🧩 Code Explanation

### Core Components

*   **`SprintPlanningComponent` (`src/app/sprint-planning`)**:
    *   Handles the logic for the Estimation Report.
    *   `onFileChange()`: Parses uploaded Excel files using `XLSX`.
    *   `calculateTotals()`: Aggregates story points per resource and total scope.
    *   `generatePDF()`: Uses `html2canvas` and `jspdf` to capture the report view and save it as a PDF.

*   **`SprintClosureComponent` (`src/app/sprint-closure`)**:
    *   Manages the Sprint Closure Report.
    *   Includes logic to identify "Spilled" points (items not completing = 'Completed').
    *   `getStatusColor()`: Dynamic styling based on item status (e.g., Green for Completed, Red for Spilled).

*   **`MonthlyReportComponent` (`src/app/monthly-report`)**:
    *   Handles broader monthly data aggregation.
    *   Similar import/export structure to Sprint components but tailored for monthly metrics.

### Key Logic Flow

1.  **Import**: The `xlsx-js-style` library reads the binary Excel file. The component maps specific column names to an internal data array (e.g., `planningData` or `closureData`).
2.  **Rendering**: Angular's `*ngFor` directive iterates over this data to build the HTML tables dynamically.
3.  **Export**:
    *   **PDF**: We capture specific HTML elements (by ID) as a canvas, convert to an image, and embed in a PDF.
    *   **Email**: We clone the DOM node, sanitize styles for email clients (inline styles), and write to the system clipboard.

## ❗ Troubleshooting

*   **Blank PDF pages**: Ensure the browser window is maximized and the report is fully visible before exporting. Background rendering sometimes requires the element to be "visible" to the DOM parser.
*   **Outlook Formatting Issues**: If pasted content looks wrong in Outlook, ensure you use the "Copy for Email" button, which specifically inlines CSS styles for compatibility.
