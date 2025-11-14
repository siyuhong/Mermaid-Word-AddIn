# Mermaid Diagram Word Add-in

A Word add-in, which allows you to write mermaid code and insert its preview directly into a Word document. The image in the inserted Word can get mermaid code at any time and be re-edited!

Built with React 18, Fluent UI v9, and CodeMirror 6 for an enhanced editing experience.

> [!IMPORTANT]
> **This project is mainly completed by [CTO.NEW](https://cto.new/) !**

## Features

- **Interactive Diagram Editor**: Full-featured code editor with syntax highlighting for Mermaid diagram syntax
- **Real-time Preview**: See your diagrams rendered in real-time as you type
- **Multiple Diagram Types**: Support for all Mermaid diagram types including:
  - Flowcharts
  - Sequence diagrams
  - Class diagrams
  - State diagrams
  - Entity Relationship diagrams
  - Gantt charts
  - Pie charts
  - And more!
- **Syntax Highlighting**: Enhanced CodeMirror 6 editor with Mermaid-specific syntax highlighting
- **Auto-completion**: Smart auto-completion for Mermaid keywords, diagram types, and syntax
- **Line Wrapping**: Automatic line wrapping for better readability of long lines
- **Theme Selection**: Choose from multiple Mermaid themes (default, forest, dark, neutral)
- **Edit Existing Diagrams**: Select an inserted diagram in Word and click "Edit Selected" to modify it
- **High-Quality Output**: Diagrams are inserted as high-quality PNG images optimized for Word documents

## Prerequisites

- **Node.js**: Latest LTS version. Visit the [Node.js site](https://nodejs.org/) to download and install. Verify installation with `node -v` and `npm -v`.
- **Microsoft 365**: Office connected to a Microsoft 365 subscription. You might qualify for a [Microsoft 365 E5 developer subscription](https://developer.microsoft.com/microsoft-365/dev-program), or you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try?rtc=1).

## Installation & Setup

1. Clone the repository
   ```bash
   git clone git@github.com:siyuhong/Mermaid-Word-AddIn.git
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Start the development server:
   ```bash
   npm run dev-server
   ```
4. In another terminal, sideload the add-in:
   ```bash
   npm start
   ```

## Development

### Run the add-in using Office Add-ins Development Kit

1. **Install Office Add-ins Development Kit In Visual Studio Code**

   Search for Office Add-ins Development Kit in VSCode extension and install it.

2. **Preview Your Office Add-in (F5)**

   Select **Preview Your Office Add-in(F5)** to launch the add-in. In the Quick Pick menu, select **Word Desktop (Edge Chromium)**.

   The extension checks prerequisites before debugging starts. After verification, Word launches with the add-in sideloaded.

3. **Stop Previewing Your Office Add-in**

   Once finished testing, select **Stop Previewing Your Office Add-in** to close the web server and remove the add-in from the registry and cache.

### Project Scripts

- `npm run build` - Build for production
- `npm run build:dev` - Build for development
- `npm run dev-server` - Start development server with hot reload
- `npm start` - Sideload the add-in to Word
- `npm stop` - Remove the add-in from Word
- `npm run validate` - Validate the manifest file
- `npm run lint` - Check code style
- `npm run lint:fix` - Auto-fix code style issues

## Using the Add-in

1. Open Word and find the **Mermaid Editor** button in the Insert tab
2. Click to open the task pane
3. Write your Mermaid diagram code in the editor (or use the default example)
4. Preview the rendered diagram in the preview pane
5. Optionally select a theme from the dropdown
6. Click **Insert Diagram** to insert it into your Word document
7. To edit an existing diagram:
   - Select the diagram image in Word
   - Click **Edit Selected** in the task pane
   - The diagram code will load into the editor
   - Make your changes and insert the updated diagram

## Project Structure

- `/src/taskpane/`
  - `index.jsx` - React entry point with Office.js initialization
  - `taskpane.js` - Word JavaScript API integration
  - `components/MermaidEditor.jsx` - Main editor component
  - `components/enhancedMermaidLanguage.js` - Enhanced Mermaid language support for CodeMirror
  - `utils/diagramUtils.js` - Diagram detection and utility functions
- `/assets/` - Add-in icons and images
- `manifest.xml` - Add-in manifest configuration
- `webpack.config.js` - Webpack bundler configuration

## Key Technologies

- **React 18** - UI framework
- **Fluent UI v9** - Microsoft's design system
- **CodeMirror 6** - Code editor with syntax highlighting
- **Mermaid.js** - Diagram rendering engine
- **Office.js** - Word JavaScript API
- **Webpack** - Module bundler

## Troubleshooting

If you have problems running the add-in:

- Close any open instances of Word
- Stop the previous web server with **Stop Previewing Your Office Add-in**
- Clear the Office cache (if issues persist)

For more help, see [troubleshoot development errors](https://learn.microsoft.com/office/dev/add-ins/testing/troubleshoot-development-errors) or [create a GitHub issue](https://aka.ms/officedevkitnewissue).

## Additional Resources

- [Office Add-ins Documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Word JavaScript API Reference](https://learn.microsoft.com/javascript/api/word)
- [Mermaid Documentation](https://mermaid.js.org/)
- [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)

## License

MIT License

Copyright (c) 2024 Microsoft Corporation. All rights reserved.
