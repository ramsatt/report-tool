# NG-Report Generator

A comprehensive Angular application for generating Sprint Planning and Sprint Closure reports with an industrial aesthetic.

## Features

- **Sprint Planning Module**: Generate estimation reports with total story points, resource allocation, and detailed scope.
- **Sprint Closure Module**: Generate closure reports with committed vs. completed points, spilled points logic, and status breakdown.
- **Excel Data Import**: Upload Excel/CSV files to populate report data instantly.
- **Multi-Format Export**:
  - **Export to PDF**: High-quality PDF generation.
  - **Export to Word**: Editable Word document export (Print Layout compatible).
  - **Copy for Email**: Outlook-compatible HTML copy-paste for email bodies.
- **Industrial Design**: High-contrast Black, White, and Yellow theme.

## Excel Data Format

To use the Excel upload feature, your data must follow the specific column headers described below.

### 1. Sprint Planning Data format

**Required Columns:**

- `Work Item ID` (e.g., 12345)
- `Work Item Type` (e.g., User Story, Bug)
- `Work Item` (Title of the item)
- `Scope` (Description of work)
- `Resource` (Developer name)
- `Story Points` (Numeric value)

**Sample row:**

| Work Item ID | Work Item Type | Work Item       | Scope            | Resource | Story Points |
| :----------- | :------------- | :-------------- | :--------------- | :------- | :----------- |
| 56781        | User Story     | Api Integration | Create endpoints | John Doe | 5            |

### 2. Sprint Closure Data format

**Required Columns:**

- `Work Item ID`
- `Work Item Type`
- `Work Item` (Title)
- `Story Points`
- `Status` (Must be 'Completed' for done items, otherwise treated as open)
- `Dev Overview` (Remarks)

**Sample row:**

| Work Item ID | Work Item Type | Work Item       | Story Points | Status      | Dev Overview           |
| :----------- | :------------- | :-------------- | :----------- | :---------- | :--------------------- |
| 56781        | User Story     | Api Integration | 5            | Completed   | Delivered successfully |
| 56792        | Bug            | Fix Login Crash | 3            | In Progress | Spilled to next sprint |

## Getting Started

1. **Install Dependencies**

   ```bash
   npm install
   ```

2. **Run Application**

   ```bash
   npm start
   ```

   Navigate to `http://localhost:4200` (or the port shown in terminal).

## Exporting Reports

- **Copy to Email**: Click "Copy Report for Email" in the Preview tab. Paste directly into Outlook or Gmail. The formatting preserves colors and layout.
- **Word Export**: Click "Export Word" to download a `.doc` file. This file opens in Print Layout mode by default.
