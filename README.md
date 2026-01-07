# Copilot Word Add-In

A **Microsoft 365 Copilot Declarative Agent** that analyzes Word documents using custom add-in functions. This is a working reference implementation demonstrating how to build Copilot agents that read document content and return intelligent analysis.

## What This Does

When a user asks the Copilot agent a question like "Analyze this document for compliance issues", the agent:

1. Invokes the add-in's `analyzeDocument` action
2. The action reads the document content via Office.js Word API
3. Sends the content to Azure OpenAI (or uses demo fallback)
4. Returns the analysis to Copilot, which displays it conversationally

## Prerequisites

- **Node.js 18+**
- **Microsoft 365 Copilot license** (for your tenant)
- **Microsoft 365 Developer tenant** or work account with admin access
- **Office Add-in development certificates** installed

## Quick Start

### 1. Install Dependencies

```bash
npm install
```

### 2. Install Office Add-in Dev Certs

```bash
npx office-addin-dev-certs install
```

This creates trusted SSL certificates at `C:\Users\<username>\.office-addin-dev-certs\` (Windows) or `~/.office-addin-dev-certs/` (Mac/Linux) that webpack will use for HTTPS.

### 3. Configure Azure OpenAI (Optional)

For AI-powered analysis, edit `src/commands/commands.ts` and update:

```typescript
const AZURE_OPENAI_ENDPOINT = "https://YOUR-RESOURCE.openai.azure.com";
const AZURE_OPENAI_KEY = "YOUR-API-KEY";
const DEPLOYMENT_NAME = "gpt-4o";
```

If not configured, the add-in returns demo responses to show the flow works.

### 4. Start Development Server

```bash
npm run dev-server
```

Verify https://localhost:3000/commands.html loads without SSL errors.

### 5. Create App Package & Sideload via Teams

> **Note:** The `teamsfx provision` CLI approach often fails. Manual sideloading works reliably.

Run this PowerShell script from the project root:
```powershell
cd appPackage
New-Item -ItemType Directory -Path "build" -Force
Copy-Item -Path "..\assets" -Destination "assets" -Recurse -Force
Compress-Archive -Path "manifest.json", "declarativeAgent.json", "document-plugin.json", "assets" -DestinationPath "build\appPackage.zip" -Force
Remove-Item -Path "assets" -Recurse -Force
cd ..
```

Or manually:
1. Copy `manifest.json`, `declarativeAgent.json`, `document-plugin.json`, and the `assets/` folder into a temporary folder
2. Zip those files (at root level, not in a subfolder)

Then sideload via Teams:
1. Go to **Microsoft Teams** → **Apps** → **Manage your apps**
2. Click **Upload an app** → **Upload a custom app**
3. Select `appPackage/build/appPackage.zip`

### 6. Test in Word

1. Open **Word** (desktop or web)
2. The add-in should appear in the ribbon (if TaskPane is configured)
3. Open the **Copilot pane** (right side)
4. Click the agent picker and select **"Copilot Word Add-In"**
5. Ask questions like:
   - "Analyze this document for compliance issues"
   - "Check if this document is complete"
   - "Find inconsistencies in this document"
   - "Identify sensitive information"

## Project Structure

```
CopilotWordAddIn/
├── appPackage/
│   ├── manifest.json           # M365 unified manifest (devPreview)
│   ├── declarativeAgent.json   # Agent persona and capabilities
│   └── document-plugin.json    # Plugin with analyzeDocument action
├── src/
│   ├── commands/
│   │   ├── commands.ts         # Office.actions.associate() handler
│   │   └── commands.html       # Commands loader (webpack injects script)
│   └── taskpane/
│       ├── taskpane.ts         # Optional task pane logic
│       └── taskpane.html       # Optional task pane UI
├── assets/                     # Icons (icon-16.png, icon-32.png, icon-80.png, icon-128.png)
├── env/                        # Environment files (gitignored)
├── webpack.config.js           # Dev server with Office SSL certs
├── teamsapp.yaml              # Teams Toolkit config (optional)
├── package.json
└── tsconfig.json
```

## How It Works

### Key Architecture

1. **manifest.json** declares two runtimes:
   - `CopilotAgentRuntime` - Headless, runs `commands.html`, hosts `Office.actions.associate()`
   - `TaskPaneRuntime` - Optional UI runtime

2. **declarativeAgent.json** defines the agent's persona and references the plugin

3. **document-plugin.json** declares the `analyzeDocument` function with OpenAPI-style schema

4. **commands.ts** implements the action handler:
   ```typescript
   Office.actions.associate("analyzeDocument", async (message: string) => {
     const { analysisType } = JSON.parse(message);
     const content = await getDocumentContent();
     const analysis = await callAzureOpenAI(content, analysisType);
     return analysis;
   });
   ```

### SSL Certificates

The webpack dev server uses Office Add-in dev certs (not self-signed). This is **critical** - Office will block loading if certs aren't trusted.

```javascript
// webpack.config.js
const devCertsPath = path.join(process.env.USERPROFILE, ".office-addin-dev-certs");
const httpsOptions = {
  key: fs.readFileSync(path.join(devCertsPath, "localhost.key")),
  cert: fs.readFileSync(path.join(devCertsPath, "localhost.crt")),
  ca: fs.readFileSync(path.join(devCertsPath, "ca.crt")),
};
```

## Features

### Copilot Agent Analysis Types
- **Compliance**: Regulatory issues, policy violations, missing disclaimers
- **Completeness**: Missing sections, placeholder text, incomplete content
- **Consistency**: Contradictions, terminology inconsistencies, formatting issues
- **Sensitivity**: PII, financial data, confidential information

### Task Pane (Optional)
A traditional add-in task pane providing document statistics and analysis type selection.

## Development

### Build for Production
```bash
npm run build
```

### Watch Mode
```bash
npm run dev-server
```

### Debugging
- Check browser DevTools console for `commands.html` errors
- Look for `[DocumentAnalyzer]` log prefix in commands.ts
- Verify https://localhost:3000/commands.html loads without SSL warnings

## Troubleshooting

| Issue | Solution |
|-------|----------|
| SSL certificate errors | Run `npx office-addin-dev-certs install` |
| Agent not appearing in Copilot | Verify sideload succeeded in Teams |
| "Agent encountered an error" | Check DevTools console in Office |
| Actions return nothing | Ensure `Office.actions.associate()` is in `Office.onReady()` |

## Documentation

See [docs/Quick Start.md](docs/Quick%20Start.md) for a comprehensive guide covering:
- Step-by-step project creation
- Manifest configuration deep-dive
- Debugging techniques
- Common issues and solutions
- Lessons learned

## Platform Support

- ✅ Word on Windows (desktop)
- ✅ Word on the Web
- ❌ Word on Mac (not yet supported for Copilot agents)

## License

MIT
