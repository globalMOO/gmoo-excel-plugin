# VSME - globalMOO Excel Web Add-in

## Overview
Office Web Add-in (Office.js) that replaces the legacy VSTO Excel add-in. Uses the globalMOO REST API for VSME optimization instead of a native DLL.

## Tech Stack
- **Runtime**: Office.js (Office Web Add-in)
- **Frontend**: React 18 + TypeScript
- **UI**: Fluent UI v9 (`@fluentui/react-components`)
- **Charts**: Chart.js (task pane) + Office.js native charts (Excel)
- **Build**: Webpack 5, ts-loader
- **Test**: Jest + @testing-library/react

## Development

```bash
# Install dependencies
npm install

# Start dev server (opens Excel with sideloaded add-in)
npm start

# Build for production
npm run build

# Run tests
npm test

# Validate manifest
npm run validate
```

## Project Structure
- `manifest.xml` — Office Add-in manifest (XML format)
- `src/taskpane/` — React task pane application
  - `types/` — TypeScript interfaces (API DTOs, workbook state)
  - `services/` — API client, Excel operations, state persistence
  - `hooks/` — React hooks (API key, client, state, optimization)
  - `components/` — Wizard step components
- `src/commands/` — Ribbon button handlers

## API
- Base URL: `https://app.globalmoo.com/api/`
- Auth: Bearer token (API key)
- SDK reference: `../gmoo-sdk-csharp/`

## Key Patterns
- Wizard-style UI: Connect → Define Model → Evaluate Cases → Set Objectives → Optimize → Results
- State persisted to workbook custom XML parts with Workbook.settings fallback
- API key stored in OfficeRuntime.storage (persists across sessions)
- Formula evaluation: write inputs → trigger recalc → poll calculationState → read outputs
- Optimization loop is fully async with cancellation support
