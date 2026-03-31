# AGENTS.md - Shipping Tool Development Guide

## Project Overview

This is a React-based shipping documentation automation tool that parses and merges PL (Packing List) Excel files to generate SI (Shipment Instruction) and consolidated Packing List documents. Built with React 19, Vite, and Tailwind CSS v4.

**Main Application**: `shipping-tool/src/App.jsx`
**Data Configuration**: `shipping-tool/src/FactoryList.js`

---

## Commands

All commands run from the `shipping-tool/` directory:

```bash
cd shipping-tool

# Development
npm run dev          # Start Vite dev server

# Build
npm run build        # Production build
npm run preview      # Preview production build

# Linting
npm run lint         # Run ESLint on all files
npm run lint -- --fix  # Auto-fix linting issues
```

---

## Code Style Guidelines

### JavaScript/React Patterns

1. **Component Structure**
   - Use functional components with hooks
   - Default export for page components
   - Named exports for utilities/modules

2. **State Management**
   - Use `useState` for local component state
   - Use `useEffect` for side effects (file parsing, data processing)
   - Avoid unnecessary re-renders by batching state updates

3. **Imports Order**
   ```javascript
   // 1. React core
   import React, { useState, useEffect } from 'react';
   // 2. Third-party libraries
   import ExcelJS from 'exceljs';
   import { saveAs } from 'file-saver';
   // 3. Internal modules
   import { FACTORY_DB } from './FactoryList';
   ```

4. **Variable Naming**
   - camelCase for variables and functions
   - UPPER_SNAKE_CASE for constants
   - Descriptive names: `uploadedFiles`, `mergedPLData`, `fileStats`
   - Avoid single letters except in loops: `i`, `j`, `k`

5. **Function Conventions**
   - Use arrow functions for callbacks and closures
   - Use regular functions for utility/helper functions
   - Keep functions focused and small (under 50 lines)

6. **Error Handling**
   - Use try/catch for async operations (file parsing)
   - Log errors with context: `console.error("解析檔案失敗:", file.name, err);`
   - Provide user feedback via UI state

### CSS/Tailwind

1. **Tailwind Classes**
   - Use utility classes directly in JSX
   - Follow semantic class naming: `bg-white`, `text-gray-700`
   - Use responsive prefixes: `grid-cols-1 lg:grid-cols-2`

2. **Custom Styles**
   - Keep custom CSS minimal (use Tailwind where possible)
   - Use CSS variables for theme colors in `index.css`

### Excel Processing

1. **ExcelJS Usage**
   - Always handle rich text cells: `cell.value.richText`
   - Use `getCellText()` helper for consistent text extraction
   - Set number formats: `cell.numFmt = '0.00'`

2. **Sheet Operations**
   - Use `eachRow()` for iteration
   - Check for worksheet existence: `wb.getWorksheet(name)`
   - Handle merged cells properly

### Data Structures

1. **Factory Data** (`FactoryList.js`)
   - Object with keys as factory codes
   - Include: `name`, `shortName`, `address`, `contact`

2. **Form State**
   - Use single state object for related fields
   - Immutable updates: `setFormData({...formData, pi: value})`

3. **File Processing**
   - Store original workbooks for later copying
   - Track parsed data separately from raw files

---

## Project Structure

```
shipping-tool/
├── src/
│   ├── App.jsx          # Main application component
│   ├── FactoryList.js   # Factory contact database
│   ├── main.jsx         # React entry point
│   ├── index.css        # Global styles (Tailwind)
│   └── assets/          # Static assets
├── public/               # Public assets
├── vite.config.js       # Vite configuration
├── package.json         # Dependencies
└── node_modules/        # Installed packages
```

---

## Key Dependencies

- **React 19** - UI framework
- **Vite 7** - Build tool
- **ExcelJS 4.4** - Excel file manipulation
- **file-saver 2.0** - File download handling
- **lucide-react 0.577** - Icon library
- **Tailwind CSS 4.2** - Styling

---

## Common Tasks

### Running a Single Test
This project does not have tests configured. To add testing:
```bash
npm install -D vitest @testing-library/react @testing-library/jest-dom
```

### Adding New Factory
Edit `shipping-tool/src/FactoryList.js`:
```javascript
export const FACTORY_DB = {
  NEW: {
    name: "工廠名稱",
    shortName: "NEW",
    address: "地址",
    contact: "聯絡資訊"
  }
};
```

### Modifying Excel Output
Edit the `generateExcel()` function in `App.jsx`:
- SI Sheet generation: lines 146-215
- PL Sheet generation: lines 217-293
- Sheet copying: lines 295-309