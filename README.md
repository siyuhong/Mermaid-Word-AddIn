# Mermaid Diagram Word Add-in

**English** | [简体中文](README.zh-CN.md)

A Microsoft Word add-in for writing Mermaid code, previewing diagrams, and inserting them into a document as PNG images. Each inserted image stores its Mermaid source in its alternative text so you can select it and edit the diagram again.

Built with React 18, Fluent UI v9, and CodeMirror 6 for an enhanced editing experience.

## Features

- **Interactive Diagram Editor**: Full-featured code editor with syntax highlighting for Mermaid diagram syntax
- **Real-time Preview**: See your diagrams rendered in real-time as you type
- **Multiple Diagram Types**: Supports diagram syntax available in the project's Mermaid 10.x dependency, including:
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
- **Preview Navigation**: Zoom in, zoom out, reset the view, and drag to pan
- **Theme Selection**: Choose from multiple Mermaid themes (default, forest, dark, neutral)
- **Edit Existing Diagrams**: Select an inserted diagram in Word and click "Edit Selected" to modify it
- **High-Quality Output**: Diagrams are inserted as high-quality PNG images optimized for Word documents

## Prerequisites

- **Node.js**: Latest LTS version. Visit the [Node.js site](https://nodejs.org/) to download and install. Verify installation with `node -v` and `npm -v`.
- **Git**: Required to clone the repository.
- **Word Desktop**: The commands below target desktop Word. For the included Windows debugging configuration, use an up-to-date Microsoft 365 Word installation with Edge Chromium/WebView2. Internet Explorer (Trident) and Edge Legacy (EdgeHTML) webviews are not supported.

The Mermaid documentation may describe newer syntax that is unavailable in Mermaid 10.x. Check the installed version when using examples from newer releases.

## Installation & Setup

1. Clone the repository and enter the project directory:

   ```bash
   git clone https://github.com/siyuhong/Mermaid-Word-AddIn.git
   cd Mermaid-Word-AddIn
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. On first setup, install the local HTTPS development certificate and accept the system trust prompt:

   ```bash
   npx office-addin-dev-certs install
   ```

4. Start the development server at `https://localhost:3000` and keep this terminal running:

   ```bash
   npm run dev-server
   ```

5. Open another terminal in the project directory and sideload the add-in into Word:

   ```bash
   npm start
   ```

When finished, run `npm stop` to remove the development add-in. Stop the manually started development server with `Ctrl+C` in its terminal.

## Development

### Run the add-in using Office Add-ins Development Kit

1. **Install Office Add-ins Development Kit In Visual Studio Code**

   Search for Office Add-ins Development Kit in VSCode extension and install it.

2. **Preview Your Office Add-in (F5)**

   Select **Preview Your Office Add-in(F5)** to launch the add-in. In the Quick Pick menu, select **Word Desktop (Edge Chromium)**.

   The extension checks prerequisites before debugging starts. After verification, Word launches with the add-in sideloaded.

3. **Stop Previewing Your Office Add-in**

   Once finished testing, select **Stop Previewing Your Office Add-in** to end the debugging session and remove the sideloaded add-in. Stop any manually started development server separately.

### Project Scripts

- `npm run build` - Build for production
- `npm run build:dev` - Build for development
- `npm run dev-server` - Start development server with hot reload
- `npm start` - Sideload the add-in to Word
- `npm stop` - Remove the add-in from Word
- `npm run validate` - Validate the manifest file
- `npm run validate:production` - Validate the generated `dist/manifest.xml` for production
- `npm run lint` - Check code style
- `npm run lint:fix` - Auto-fix code style issues
- `npm run prettier` - Format source files
- `npm run watch` - Rebuild when source files change

### Production Deployment

`npm run build` writes the production files to `dist/`; it does not publish them. Set `ADDIN_BASE_URL` to your deployed HTTPS site URL before building. A production build fails if this variable is missing or points to localhost or the Contoso placeholder.

```powershell
$env:ADDIN_BASE_URL = "https://your-project.pages.dev/"
npm run build
```

The generated `dist/manifest.xml` uses the configured site for the pages, icons, help link, and app domain. The source manifest remains configured for local development. Deploy the contents of `dist/`, run `npm run validate:production`, and share the site homepage or the generated manifest with testers. The homepage includes a manifest download and sideloading instructions.

See the [Cloudflare Pages deployment guide (Chinese)](docs/cloudflare-pages.zh-CN.md) for Git integration, direct uploads, and user installation steps. Microsoft supports sideloading for development/testing, not as a production distribution channel.

## Using the Add-in

1. Open Word and find the **Mermaid Editor** button in the Insert tab
2. Click to open the task pane
3. Write your Mermaid diagram code in the editor (or use the default example)
4. Preview the rendered diagram in the preview pane
5. Optionally select a theme from the dropdown
6. Wait for the preview to finish, place the Word cursor where you want the image, and click **Insert | Replace**. If text or an image is selected in Word, that selection is replaced.
7. To edit an existing diagram:
   - Select the diagram image in Word
   - Click **Edit Selected** in the task pane
   - The diagram code will load into the editor
   - Make your changes and wait for the updated preview
   - Keep the original image selected, or select it again, then click **Insert | Replace** to replace it

The insert button stays disabled while code or theme changes are being rendered. **Edit Selected** is available for inline images inserted by this add-in that still contain the Mermaid title and source in their alternative text. Keep that metadata intact if you want to edit the image later. Replacement always uses the current Word selection; it is not permanently linked to the image whose source was loaded.

An explicit theme in a Mermaid initialization directive or YAML configuration takes precedence over the theme dropdown.

## Project Structure

- `/src/taskpane/`
  - `index.jsx` - React entry point with Office.js initialization
  - `taskpane.js` - Word JavaScript API integration
  - `components/MermaidEditor.jsx` - Main editor component
  - `components/enhancedMermaidLanguage.js` - Enhanced Mermaid language support for CodeMirror
  - `utils/diagramUtils.js` - Diagram detection and utility functions
  - `utils/mermaidRenderer.js` - Queued Mermaid rendering and export configuration
- `/assets/` - Add-in icons and images
- `/src/commands/` - Add-in command callbacks and host page
- `/dist/` - Generated build output; do not edit manually
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
- Stop debugging with **Stop Previewing Your Office Add-in** or `npm stop`; stop a manually started development server with `Ctrl+C`
- For HTTPS trust errors, run `npx office-addin-dev-certs verify` and reinstall the certificate if needed
- Confirm that the development server is running at `https://localhost:3000` while using the locally sideloaded add-in
- Clear the Office cache (if issues persist)

For more help, see [troubleshoot development errors](https://learn.microsoft.com/office/dev/add-ins/testing/troubleshoot-development-errors) or [report a project issue](https://github.com/siyuhong/Mermaid-Word-AddIn/issues).

## Additional Resources

- [Office Add-ins Documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Word JavaScript API Reference](https://learn.microsoft.com/javascript/api/word)
- [Mermaid Documentation](https://mermaid.js.org/)
- [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)

## License

MIT License. See [LICENSE](LICENSE).

Copyright (c) 2024 Microsoft Corporation. All rights reserved.
