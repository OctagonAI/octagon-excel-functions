# Octagon AI for Excel

> Access specialized AI agents for market intelligence directly in Excel spreadsheets

An Excel add-in that integrates [Octagon's AI Agents API](https://docs.octagonagents.com/) with Excel, providing custom functions for financial and market research. Use the power of specialized AI agents for analyzing SEC filings, earnings transcripts, stock data, private market intelligence, and more directly from your Excel formulas.

## Features

- **🔑 API Integration**: Securely use your Octagon AI API key
- **📊 Custom Excel Functions**: Call specialized AI agents with `=OCTAGON.OCTAGON_AGENT()`, `=OCTAGON.DEEP_RESEARCH_AGENT()`, and more
- **🔄 Smart Routing**: The main Octagon Agent automatically routes queries to the most appropriate specialized agent
- **🔍 Deep Research**: Access comprehensive research on financial topics
- **🌐 Web Scraping**: Extract structured data from websites
- **📈 Market Intelligence**: Get intelligent insights on financial data

## Why This Add-in

- **Free** - Use your own API key
- **Confidential** - API calls go directly from Excel to Octagon
- **Secure** - No data stored or processed outside of your spreadsheet
- **Professional** - Built specifically for financial professionals and analysts
- **Open Source** - Review the code, contribute improvements

## Available Functions

### 🔄 **Smart Router**

- **`OCTAGON.OCTAGON_AGENT(prompt)`** - Intelligent router that automatically selects the best specialized agent for your query

## Usage

### ⚠️ Getting Started

- This Add-In is currently undergoing Microsoft's review before being published to AppSource. Soon enough, you'll be able to use it directly from

### Examples

```
=OCTAGON.OCTAGON_AGENT("Retrieve year-over-year growth in key income-statement items for AAPL, limited to 5 records and filtered by period FY.")
```

<!-- prettier-ignore -->
> [!CAUTION]
> Be mindful of potential API usage costs. Changes to dependency cells can cause recalculation, and certain actions in Excel can trigger full recalculation. You may wish to switch the calculation mode in Excel from automatic to manual to control when API calls are made.

## Development Setup

### Prerequisites

- [Node.js](https://nodejs.org) (LTS version recommended)
- [npm](https://www.npmjs.com/)
- Microsoft Excel (desktop version for Windows/Mac or Excel on the web)
- [Visual Studio Code](https://code.visualstudio.com/) (recommended)

### Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/OctagonAI/octagon-excel-functions.git
   cd <your-directory>
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. Start the dev server and sideload the add-in in Excel:
   ```bash
   npm start:dev # this points to the manifest-local.xml (local server)
   ```

This command:

1. Builds the project
2. Starts a local HTTPS server on port 3000
3. Opens Excel and sideloads the add-in

### Development Workflow

- **Development Build**: `npm run build:dev`
- **Watch Mode**: `npm run watch`
- **Dev Server**: `npm run dev-server`
- **Linting**: `npm run lint` or `npm run lint:fix`

## Local vs Production

This repository contains two manifest files:

- `manifest-local.xml` - For local development with localhost URLs
- `manifest.xml` - For production deployment with GitHub Pages URLs

## Project Structure

```
octagon-excel-custom-functions/
├── assets/                   # Icon images for the add-in
├── src/                      # Source code
│   ├── api/                  # API integration with Octagon services
│   ├── commands/             # Excel ribbon commands
│   ├── functions/            # Excel custom functions
│   ├── taskpane/             # Task pane UI
│   └── utils/                # Utility functions
├── manifest.xml              # Production manifest
├── manifest-local.xml        # Local development manifest
├── package.json              # Project dependencies and scripts
└── webpack.config.js         # Build configuration
```

## Resources

- [Excel JavaScript API overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-overview)
- [Excel Custom Functions Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview)
- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Octagon AI Website](https://www.octagonai.co/)
- [Octagon AI API Documentation](https://docs.octagonagents.com/)

## Support

For support, please reach out to [Octagon AI Support](https://www.octagonai.co/).

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for release history.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
