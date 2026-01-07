# Quick Start: Build a Copilot Agent with Office Add-in

This guide walks you through building an Office Add-in that integrates with Microsoft 365 Copilot as a declarative agent. Users can interact with your add-in via natural language in the Copilot chat pane.

## Table of Contents

1. [Prerequisites](#prerequisites) - What you need before starting
2. [Expected Project Structure](#expected-project-structure) - Final folder layout
3. [Part 1: Create the Base Office Add-in](#part-1-create-the-base-office-add-in) - Steps 1-2
4. [Part 2: Add the Copilot Declarative Agent](#part-2-add-the-copilot-declarative-agent) - Steps 3-5 (JSON configs)
5. [Part 3: Implement the JavaScript Functions](#part-3-implement-the-javascript-functions) - Steps 6-7 (TypeScript code)
6. [Part 4: Configure Project Files](#part-4-configure-project-files) - Steps 8-9 (yaml/env)
7. [Part 5: Configure SSL Certificates](#part-5-configure-ssl-certificates-critical) - Steps 10-11 ⚠️ CRITICAL
8. [Part 6: Create and Upload the App Package](#part-6-create-and-upload-the-app-package) - Steps 12-14
9. [Part 7: Test the Agent](#part-7-test-the-agent) - Steps 15-17
10. [Part 8: Making Changes](#part-8-making-changes) - Hot reload limitations
11. [Next Steps](#next-steps) - Azure OpenAI integration
12. [Troubleshooting](#troubleshooting) - Common issues and fixes
13. [Lessons Learned](#lessons-learned-from-building-this-poc) - Hard-won insights

---

## Prerequisites

### Knowledge Prerequisites
- Basic understanding of declarative agents in Microsoft 365 Copilot
- Recommended reading: [Declarative agents for Microsoft 365 Copilot overview](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-declarative-agent)

### Software Prerequisites
- [Node.js](https://nodejs.org/) (LTS version recommended)
- [Visual Studio Code](https://code.visualstudio.com/)
- [Microsoft 365 Agent Toolkit](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/teams-toolkit-overview) extension for VS Code
- Microsoft 365 tenant with Copilot license
- **Word for Windows** (Microsoft 365 subscription, version 2404 or later)
  - Mac is not supported yet
  - Word on the web has limited support
- **Office Add-in Dev Certificates** (installed in Step 10)

### Licensing Requirements
- **Microsoft 365 Copilot license** - Required! Without this, the agent won't appear in Word
- Microsoft 365 Developer Program account (free) for development tenant

### Verify Copilot is Available

Before starting, confirm Copilot works in your Word:

1. Open Word and create a new document
2. Look for a **Copilot button** on the Home ribbon
3. Click it - you should see the Copilot pane open on the right
4. If no Copilot button exists, you don't have a Copilot license and this guide won't work

---

## Expected Project Structure

After completing this guide, your project should look like this:

```
DocumentAnalyzerAgent/
├── appPackage/
│   ├── manifest.json              # Main app manifest
│   ├── declarativeAgent.json      # Copilot agent config
│   ├── document-plugin.json       # Plugin functions config
│   └── build/                     # Generated package
│       └── appPackage.zip
├── assets/
│   ├── icon-16.png               # Created by Agent Toolkit
│   ├── icon-32.png
│   ├── icon-80.png
│   └── icon-128.png
├── env/
│   └── .env.dev                  # Environment variables
├── src/
│   └── commands/
│       ├── commands.html         # Runtime page (loads Office.js)
│       └── commands.ts           # Agent action handlers
├── package.json
├── teamsapp.yaml                 # Deployment config (optional)
├── tsconfig.json                 # TypeScript config (created by toolkit)
└── webpack.config.js             # Build config (MUST MODIFY for SSL)
```

> **Note:** The Agent Toolkit may also create a `taskpane/` folder for UI. You can delete it if you only need the Copilot agent (no task pane). If you keep it, ensure manifest.json doesn't reference `taskpane.html` if it doesn't exist.

---

## Part 1: Create the Base Office Add-in

### Step 1: Create a new add-in project

1. Open Visual Studio Code
2. Open the **Microsoft 365 Agent Toolkit** extension
3. Create a new Office Add-in project:
   - Select **Create a New App**
   - Choose **Office Add-in** as the project type
   - Select **Word** (this guide uses Word; Excel/PowerPoint require different code)
   - Name it something like "Document Analyzer Agent"

4. When the project opens in a new VS Code window, close the original window

5. Install dependencies:
   ```powershell
   npm install
   ```

> **Note:** The Agent Toolkit creates an `assets/` folder with placeholder icons. You can customize these later.

### Verify package.json scripts

Your `package.json` should have these scripts (Agent Toolkit creates them):

```json
"scripts": {
    "build": "webpack --mode production",
    "dev-server": "webpack serve --mode development",
    "start": "npm run dev-server",
    "stop": "taskkill /im node.exe /f"
}
```

If these are missing, add them.

### Step 2: Test the base add-in (optional but recommended)

1. Select **View > Run** in VS Code
2. In the RUN AND DEBUG dropdown, select **Word Desktop (Edge Chromium)**
3. Press **F5** to build and launch
4. If prompted about certificates, accept both prompts
5. In Word, select **Add-ins** button on Home ribbon, then select your add-in
6. Test that the task pane opens and basic functionality works
7. Stop debugging:
   ```powershell
   npm run stop
   ```

---

## Part 2: Add the Copilot Declarative Agent

### Step 3: Update the manifest

Open your `manifest.json` file (in the `appPackage` folder) and make these changes:

#### 3a. Add the copilotAgents declaration

Add this object to the root of the manifest (conventionally after `"validDomains"`):

```json
"copilotAgents": {
  "declarativeAgents": [
    {
      "id": "DocumentAnalyzerAgent",
      "file": "declarativeAgent.json"
    }
  ]
},
```

#### 3b. Add a runtime for the agent actions

In the `"extensions.runtimes"` array, add a new runtime object (or modify an existing CommandsRuntime):

```json
{
    "id": "CopilotAgentRuntime",
    "type": "general",
    "code": {
        "page": "https://localhost:3000/commands.html"
    },
    "lifetime": "short",
    "actions": [
        {
            "id": "analyzeDocument",
            "type": "executeDataFunction"
        }
    ]
}
```

> **Important:** The `"actions.id"` must match what you'll use in `Office.actions.associate()` later.

### Step 4: Create the declarative agent configuration

Create a new file `appPackage/declarativeAgent.json`:

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.5/schema.json",
    "version": "v1.5",
    "name": "Document Analyzer Agent",
    "description": "Agent for analyzing documents for completeness, compliance, consistency, and sensitivity.",
    "instructions": "You are an agent that analyzes documents. You can check for completeness, compliance issues, consistency problems, and sensitive content.",
    "conversation_starters": [
        {
            "title": "Analyze for compliance",
            "text": "Analyze this document for compliance issues"
        },
        {
            "title": "Check completeness",
            "text": "Check if this document is complete"
        },
        {
            "title": "Find sensitive content",
            "text": "Identify any sensitive information in this document"
        }
    ],
    "actions": [
        {
            "id": "localDocumentPlugin",
            "file": "document-plugin.json"
        }
    ]
}
```

### Step 5: Create the API plug-in configuration

Create a new file `appPackage/document-plugin.json`:

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.3/schema.json",
    "schema_version": "v2.3",
    "name_for_human": "Document Analyzer",
    "description_for_human": "Analyzes documents for various criteria",
    "namespace": "addinfunction",
    "functions": [
        {
            "name": "analyzeDocument",
            "description": "Analyzes the current document for specified criteria like compliance, completeness, consistency, or sensitivity.",
            "parameters": {
                "type": "object",
                "properties": {
                    "analysisType": {
                        "type": "string",
                        "description": "The type of analysis to perform: 'compliance', 'completeness', 'consistency', or 'sensitivity'",
                        "default": "compliance"
                    }
                },
                "required": ["analysisType"]
            },
            "returns": {
                "type": "string",
                "description": "A detailed analysis report"
            },
            "states": {
                "reasoning": {
                    "description": "`analyzeDocument` reads the document content and analyzes it based on the specified criteria.",
                    "instructions": "Determine what type of analysis the user wants from their prompt."
                },
                "responding": {
                    "description": "`analyzeDocument` returns analysis findings.",
                    "instructions": "Present the analysis findings in a clear, organized format with specific recommendations."
                }
            }
        }
    ],
    "runtimes": [
        {
            "type": "LocalPlugin",
            "spec": {
                "local_endpoint": "Microsoft.Office.Addin",
                "allowed_host": ["document"]
            },
            "run_for_functions": ["analyzeDocument"],
            "auth": {
                "type": "None"
            }
        }
    ]
}
```

> **Note:** For Word, use `"allowed_host": ["document"]`. For Excel, use `["workbook"]`. For PowerPoint, use `["presentation"]`.

### Complete Manifest Reference

For reference, here's what a complete `manifest.json` should look like after all modifications:

<details>
<summary>Click to expand full manifest.json example</summary>

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json",
    "manifestVersion": "devPreview",
    "version": "1.0.0",
    "id": "YOUR-GUID-HERE",
    "packageName": "com.contoso.documentanalyzer",
    "developer": {
        "name": "Your Company",
        "websiteUrl": "https://www.example.com",
        "privacyUrl": "https://www.example.com/privacy",
        "termsOfUseUrl": "https://www.example.com/terms"
    },
    "name": {
        "short": "Document Analyzer",
        "full": "Document Analyzer with Copilot Agent"
    },
    "description": {
        "short": "Analyze documents for compliance, completeness, consistency, and sensitivity",
        "full": "A Word add-in with a Copilot agent that analyzes documents."
    },
    "icons": {
        "color": "assets/icon-128.png",
        "outline": "assets/icon-32.png"
    },
    "accentColor": "#4F6BED",
    "localizationInfo": {
        "defaultLanguageTag": "en-us"
    },
    "validDomains": [
        "localhost:3000"
    ],
    "copilotAgents": {
        "declarativeAgents": [
            {
                "id": "DocumentAnalyzerAgent",
                "file": "declarativeAgent.json"
            }
        ]
    },
    "authorization": {
        "permissions": {
            "resourceSpecific": [
                {
                    "name": "Document.ReadWrite.User",
                    "type": "Delegated"
                }
            ]
        }
    },
    "extensions": [
        {
            "requirements": {
                "scopes": ["document"],
                "capabilities": [
                    {
                        "name": "WordApi",
                        "minVersion": "1.3"
                    }
                ]
            },
            "runtimes": [
                {
                    "id": "CopilotAgentRuntime",
                    "type": "general",
                    "code": {
                        "page": "https://localhost:3000/commands.html"
                    },
                    "lifetime": "short",
                    "actions": [
                        {
                            "id": "analyzeDocument",
                            "type": "executeDataFunction"
                        }
                    ]
                }
            ]
        }
    ]
}
```

> **Note:** This minimal manifest only includes the Copilot agent runtime. If you want a task pane with a ribbon button (like the MS quickstart example), you'd add a second runtime for TaskPane and a `"ribbons"` section. The Copilot agent works without these.

</details>

---

## Part 3: Implement the JavaScript Functions

### Step 6: Create the agent action function

Open `src/commands/commands.ts` (or `commands.js`) and replace its contents with:

```typescript
/*
 * Document Analyzer Agent - Command Handlers
 * These functions are called by the Copilot agent
 */

/**
 * Reads the entire Word document content
 */
async function getDocumentContent(): Promise<string> {
    return await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        return body.text;
    });
}

/**
 * Main analysis function - reads document and returns analysis
 * For demo: Returns document statistics
 * For production: Call Azure OpenAI (see "Next Steps" section)
 */
async function analyzeDocument(analysisType: string): Promise<string> {
    try {
        const content = await getDocumentContent();
        
        // Count basic statistics
        const wordCount = content.trim().split(/\s+/).filter(w => w.length > 0).length;
        const charCount = content.length;
        
        // Demo response - proves your code is running
        // Replace with Azure OpenAI call for real analysis
        return `🔐 DOCUMENT ANALYZER AGENT

📊 Document Statistics:
- Words: ${wordCount}
- Characters: ${charCount}
- Analysis type requested: ${analysisType}

⚠️ This is a demo response. Configure Azure OpenAI for real ${analysisType} analysis.`;
        
    } catch (err) {
        const error = err as Error;
        console.error("Error in analyzeDocument:", error);
        return `Error analyzing document: ${error.message}`;
    }
}

/**
 * Register the action handler with Office
 * The "analyzeDocument" string MUST match:
 * - manifest.json → extensions.runtimes[].actions[].id
 * - document-plugin.json → functions[].name
 */
Office.onReady(() => {
    Office.actions.associate("analyzeDocument", async (message: string) => {
        console.log("analyzeDocument called with:", message);
        
        // Parse the message from Copilot
        let analysisType = "compliance"; // default
        try {
            const parsed = JSON.parse(message);
            analysisType = parsed.analysisType || "compliance";
        } catch {
            // If message isn't JSON, use it directly or default
            analysisType = message || "compliance";
        }
        
        const result = await analyzeDocument(analysisType);
        console.log("Returning result:", result);
        return result;
    });
});
```

> **Key Points:**
> - The `🔐 DOCUMENT ANALYZER AGENT` text proves YOUR code is running (not regular Copilot)
> - The `Office.actions.associate` name MUST match the action ID in manifest.json and function name in document-plugin.json
> - Error handling prevents crashes if something goes wrong

### Step 7: Create the commands.html file (if it doesn't exist)

Create or verify `src/commands/commands.html`:

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Document Analyzer Commands</title>
    
    <!-- Office JavaScript Library -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head>
<body>
    <!-- This page has no UI - it only loads the commands JavaScript -->
</body>
</html>
```

> **Note:** If you're using Webpack with `HtmlWebpackPlugin` (typical setup from Agent Toolkit), the `commands.js` script tag is injected automatically during the build. The template only needs the Office.js script. If you're NOT using Webpack, add `<script src="commands.js"></script>` before `</body>`.

---

## Part 4: Configure Project Files

### Step 8: Update teamsapp.yaml (or m365agents.yaml)

Replace the contents of `teamsapp.yaml` in the project root:

```yaml
# yaml-language-server: $schema=https://aka.ms/teams-toolkit/v1.7/yaml.schema.json
version: v1.7

environmentFolderPath: ./env

provision:
  - uses: teamsApp/create
    with:
      name: Document Analyzer ${{APP_NAME_SUFFIX}}
    writeToEnvironmentFile:
      teamsAppId: TEAMS_APP_ID

  - uses: teamsApp/zipAppPackage
    with:
      manifestPath: ./appPackage/manifest.json
      outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
      outputFolder: ./appPackage/build

  - uses: teamsApp/validateAppPackage
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

  - uses: teamsApp/extendToM365
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
    writeToEnvironmentFile:
      titleId: M365_TITLE_ID
      appId: M365_APP_ID

publish:
  - uses: teamsApp/zipAppPackage
    with:
      manifestPath: ./appPackage/manifest.json
      outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
      outputFolder: ./appPackage/build

  - uses: teamsApp/validateAppPackage
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

  - uses: teamsApp/update
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

  - uses: teamsApp/publishAppPackage
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
    writeToEnvironmentFile:
      publishedAppId: TEAMS_APP_PUBLISHED_APP_ID

projectId: YOUR-RANDOM-GUID-HERE
```

> **Important:** Replace `YOUR-RANDOM-GUID-HERE` with a new randomly generated GUID (e.g., from [guidgenerator.com](https://www.guidgenerator.com/)).

> **Note about provisioning approaches:**
> - **Official MS approach:** Run `teamsapp provision` which uses `teamsApp/extendToM365` to register the app
> - **What actually works (as of Jan 2026):** The CLI often fails with cryptic errors. Skip provisioning and manually sideload via Teams (Step 13-14)
> - The teamsapp.yaml is included for completeness but you may not need it

### Step 9: Update environment file

Open `env/.env.dev` and add these lines at the end:

```
TEAMS_APP_ID=
TEAMS_APP_TENANT_ID=
M365_TITLE_ID=
M365_APP_ID=
```

---

## Part 5: Configure SSL Certificates (CRITICAL)

Office Add-ins require trusted SSL certificates. Webpack's default self-signed certificates will be **blocked by Office**, causing "Sorry, I wasn't able to respond" errors.

### Step 10: Install Office Add-in Dev Certificates

1. In your project folder, run:
   ```powershell
   npx office-addin-dev-certs install
   ```

2. Verify certificates are installed:
   ```powershell
   npx office-addin-dev-certs verify
   ```
   You should see certificates at `C:\Users\<username>\.office-addin-dev-certs\`

### Step 11: Configure Webpack to Use Trusted Certificates

Open `webpack.config.js` and make these changes:

1. Add `fs` require at the top (with other requires):
   ```javascript
   const fs = require("fs");
   ```

2. Add certificate configuration after the requires:
   ```javascript
   // Use Office Add-in dev certs for trusted HTTPS
   const certPath = require('os').homedir() + '/.office-addin-dev-certs';
   const httpsOptions = fs.existsSync(certPath + '/localhost.crt') ? {
     key: fs.readFileSync(certPath + '/localhost.key'),
     cert: fs.readFileSync(certPath + '/localhost.crt'),
     ca: fs.readFileSync(certPath + '/ca.crt')
   } : undefined;
   ```

3. In the `devServer` section, change `https: true` to:
   ```javascript
   server: {
     type: "https",
     options: httpsOptions
   },
   ```

<details>
<summary>Click to see a complete webpack.config.js example</summary>

```javascript
const path = require("path");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

// Use Office Add-in dev certs for trusted HTTPS
const certPath = path.join(process.env.USERPROFILE || process.env.HOME, ".office-addin-dev-certs");
const httpsOptions = fs.existsSync(path.join(certPath, "localhost.crt")) ? {
  key: fs.readFileSync(path.join(certPath, "localhost.key")),
  cert: fs.readFileSync(path.join(certPath, "localhost.crt")),
  ca: fs.readFileSync(path.join(certPath, "ca.crt")),
} : true;  // Fallback to self-signed if certs don't exist

module.exports = {
  entry: {
    commands: "./src/commands/commands.ts",
    // Add taskpane entry if you have a task pane UI:
    // taskpane: "./src/taskpane/taskpane.ts",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/commands/commands.html",
      filename: "commands.html",
      chunks: ["commands"],
    }),
    // Add taskpane HTML if you have a task pane UI:
    // new HtmlWebpackPlugin({
    //   template: "./src/taskpane/taskpane.html",
    //   filename: "taskpane.html",
    //   chunks: ["taskpane"],
    // }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets", to: "assets" },
      ],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    port: 3000,
    server: {
      type: "https",
      options: httpsOptions,
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    hot: true,
  },
};
```

> **Important:** This webpack config uses `path.join()` which works better on Windows. The `httpsOptions` fallback to `true` means it will use self-signed certs if Office dev certs aren't found - but this may cause SSL errors in Office.

</details>

---

## Part 6: Create and Upload the App Package

### Step 12: Build the project

```powershell
npm run build
```

Verify it completes without errors.

> **✅ Checkpoint:** You should see `dist/` folder created with `commands.html` and `commands.js` inside.

### Step 13: Create the app package zip

The `teamsapp provision` command may fail. If so, create the package manually.

**From your project root folder**, run:

```powershell
cd appPackage

# Create build folder
New-Item -ItemType Directory -Path "build" -Force

# Copy assets folder into appPackage temporarily
Copy-Item -Path "..\assets" -Destination "assets" -Recurse -Force

# Create the zip with all required files
Compress-Archive -Path "manifest.json", "declarativeAgent.json", "document-plugin.json", "assets" -DestinationPath "build\appPackage.zip" -Force

# Clean up
Remove-Item -Path "assets" -Recurse -Force

Write-Host "Package created at appPackage/build/appPackage.zip"

# Return to project root
cd ..
```

> **Important:** The zip must contain:
> - `manifest.json` (at root level, NOT in a subfolder)
> - `declarativeAgent.json`
> - `document-plugin.json`
> - `assets/` folder with icon files

### Step 14: Sideload via Microsoft Teams

1. Open **Microsoft Teams** (web or desktop)
2. Click **Apps** in the left sidebar
3. Click **Manage your apps** at the bottom
4. Click **Upload an app** → **Upload a custom app**
5. Select your `appPackage/build/appPackage.zip` file
6. Click **Add** when prompted

If you get a manifest parsing error, verify the zip structure (see Troubleshooting).

> **✅ Checkpoint:** You should see "Document Analyzer" appear in your "Manage your apps" list with a blue square icon.

---

## Part 7: Test the Agent

### Step 15: Start the dev server

```powershell
npm run dev-server
```

Wait until you see "compiled successfully". **Keep this running!**

### Step 16: Verify HTTPS is working

1. Open a browser and go to `https://localhost:3000/commands.html`
2. You should see a blank page with no certificate warnings
3. If you see certificate errors, see Troubleshooting section

> **✅ Checkpoint:** The page loads with NO security warnings. If you see "Your connection is not private" or a certificate error, STOP - fix this before continuing (see Part 5).

### Step 17: Open Word and test

1. Close Word if it's open, then reopen it
2. Create or open a document with some text
3. Open the **Copilot pane** (Copilot button on ribbon)
4. Click the **hamburger menu (☰)** in the Copilot pane
5. Look for your agent ("Document Analyzer Agent")
   - May need to click "See more agents"
   - If not visible, wait 1-2 minutes and press **Ctrl+R** to refresh
6. Select your agent
7. Try a conversation starter or type: "Analyze this document"

### Expected Result

You should see a response that includes your custom text (e.g., document statistics, character count). If you see regular Copilot responding instead of your agent, your JavaScript function didn't execute - see Troubleshooting.

---

## Part 8: Making Changes

Hot reload is **not supported** during preview. To make changes:

1. Stop the server: `Ctrl+C` in terminal or `npm run stop`
2. Make your code changes
3. Rebuild: `npm run build`
4. Restart server: `npm run dev-server`
5. In Word: Close and reopen the document, or press **Ctrl+R** with Copilot pane focused

### If changes don't appear:

1. Clear Office cache:
   - Windows: Delete contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
2. Remove the app from Teams:
   - Teams → Apps → Manage your apps → find your app → trash icon
3. Re-upload the app package (Step 14)
4. Restart Word

---

## Next Steps

### Add Azure OpenAI Integration

To replace the demo response with real AI analysis, modify the `analyzeDocument` function in `src/commands/commands.ts`:

```typescript
// Add this configuration at the top of the file
const AZURE_OPENAI_CONFIG = {
    endpoint: "https://YOUR-RESOURCE.openai.azure.com",
    deployment: "gpt-4o",
    apiKey: "YOUR-API-KEY" // Move to secure storage in production!
};

const ANALYSIS_PROMPTS: Record<string, string> = {
    compliance: "Analyze this document for regulatory compliance issues, policy violations, and legal concerns.",
    completeness: "Check if this document has all required sections and information. Identify any gaps.",
    consistency: "Look for contradictions, inconsistencies, or conflicting statements in this document.",
    sensitivity: "Identify any sensitive information like PII, financial data, or confidential content."
};

// Replace the analyzeDocument function with this version
async function analyzeDocument(analysisType: string): Promise<string> {
    try {
        const content = await getDocumentContent();
        
        const response = await fetch(
            `${AZURE_OPENAI_CONFIG.endpoint}/openai/deployments/${AZURE_OPENAI_CONFIG.deployment}/chat/completions?api-version=2024-02-15-preview`,
            {
                method: 'POST',
                headers: {
                    'api-key': AZURE_OPENAI_CONFIG.apiKey,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    messages: [
                        { role: 'system', content: ANALYSIS_PROMPTS[analysisType] || ANALYSIS_PROMPTS.compliance },
                        { role: 'user', content: `Analyze this document:\n\n${content}` }
                    ],
                    max_tokens: 2000
                })
            }
        );
        
        if (!response.ok) {
            return `Error: Azure OpenAI returned ${response.status}. Check your endpoint and API key.`;
        }
        
        const data = await response.json();
        return data.choices[0].message.content;
        
    } catch (err) {
        const error = err as Error;
        return `Error analyzing document: ${error.message}`;
    }
}
```

> **Note:** Keep the `getDocumentContent()` function from Step 6 - only replace `analyzeDocument()`.

### Production Considerations

- [ ] Move API keys to Azure Key Vault or environment variables
- [ ] Add error handling and retry logic
- [ ] Implement token-based authentication for the add-in
- [ ] Handle large documents (chunking for long content)
- [ ] Add loading indicators in the UI
- [ ] Test across different Office versions and platforms

---

## Troubleshooting

### ❌ "Sorry, I wasn't able to respond" or "Content blocked because not signed by valid security certificate"

**This is the most common issue.** Office is blocking your dev server because it's using an untrusted SSL certificate.

**Solution:**
1. Install Office Add-in dev certificates:
   ```powershell
   npx office-addin-dev-certs install
   ```
2. Verify they exist:
   ```powershell
   npx office-addin-dev-certs verify
   ```
3. Update `webpack.config.js` to use these certificates (see Step 11)
4. Restart the dev server
5. Test by opening `https://localhost:3000/commands.html` in a browser - should show no certificate warnings

### ❌ "Something went wrong" - Copilot crashes

This usually happens when the dev server is not running when Copilot tries to invoke your agent.

**Solution:**
1. Ensure `npm run dev-server` is running and shows "compiled successfully"
2. Remove the app from Teams (Manage your apps → trash icon)
3. Clear Office cache: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
4. Wait 30 seconds
5. Re-upload the app package via Teams
6. Restart Word

### ❌ Manifest parsing failed / Upload error in Teams

The zip file structure is incorrect.

**Solution:**
1. Verify the zip contains files at the **root level**, not in a subfolder
2. Required files: `manifest.json`, `declarativeAgent.json`, `document-plugin.json`, `assets/` folder
3. To verify, extract the zip and check - `manifest.json` should be directly inside, not in a subfolder like `appPackage/manifest.json`

**Correct structure:**
```
appPackage.zip
├── manifest.json
├── declarativeAgent.json
├── document-plugin.json
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-80.png
    └── icon-128.png
```

### ❌ Agent not appearing in Copilot hamburger menu

1. Wait 2-3 minutes after uploading - propagation can be slow
2. With Copilot pane open, press **Ctrl+R** to refresh
3. Click "See more agents" if available
4. Verify the app appears in Teams → Manage your apps
5. Try closing and reopening Word

### ❌ Regular Copilot responds instead of my agent

Your agent is selected, but the JavaScript function isn't executing. This happens if:
- The dev server isn't running
- The SSL certificate isn't trusted
- The function name doesn't match

**How to verify your code is running:**
1. Add a unique string to your response, like:
   ```typescript
   return `🔐 CUSTOM AGENT: Analysis complete for ${analysisType}`;
   ```
2. Rebuild and restart the dev server
3. Test again - you should see your unique string
4. If you see generic AI response without your string, check SSL certificates

### ❌ Agent action fails with "action not found"

The action ID doesn't match between files.

**Check these match exactly:**
- `manifest.json` → `extensions.runtimes[].actions[].id` 
- `document-plugin.json` → `functions[].name`
- `commands.ts` → `Office.actions.associate("actionName", ...)`

All three must use the **exact same string** (case-sensitive).

### ❌ Agent action fails with "handler registration not found"

The JavaScript function isn't registering properly.

**Solution:**
1. Ensure `Office.actions.associate()` is inside `Office.onReady()` callback
2. Ensure the first parameter matches `functions.name` exactly
3. Check browser console for errors at `https://localhost:3000/commands.html`

### ❌ teamsapp provision command fails

The Teams Toolkit CLI can be unreliable. Use manual sideloading instead:

1. Build the project: `npm run build`
2. Create the zip manually (see Step 13)
3. Upload via Teams → Manage your apps → Upload custom app

### ❌ "teamsApp/extendToM365" action failed

This happens with CLI provisioning. Skip it and sideload manually via Teams.

---

## Debugging Tips

### Check the browser console
1. While dev server is running, open `https://localhost:3000/commands.html` in Edge
2. Press F12 to open DevTools
3. Check Console tab for JavaScript errors
4. Check Network tab - all requests should show 200 status

### Add logging to your function
```typescript
Office.actions.associate("analyzeDocument", async (message) => {
    console.log("analyzeDocument called with:", message);
    try {
        const { analysisType } = JSON.parse(message);
        console.log("Parsed analysisType:", analysisType);
        const result = await analyzeDocument(analysisType);
        console.log("Returning result:", result);
        return result;
    } catch (err) {
        const error = err as Error;
        console.error("Error in analyzeDocument:", error);
        return `Error: ${error.message}`;
    }
});
```

### Verify certificate files exist
```powershell
$certPath = "$env:USERPROFILE\.office-addin-dev-certs"
Get-ChildItem $certPath
```

You should see:
- `ca.crt`
- `localhost.crt`  
- `localhost.key`

---

## Lessons Learned (From Building This POC)

These are hard-won lessons from actually building and debugging this integration:

### 1. SSL Certificates Are the #1 Blocker
The Microsoft docs don't emphasize this enough. If you use webpack's default self-signed certificate, Office will **silently block** your add-in. You'll see "Sorry, I wasn't able to respond" with no helpful error message. The fix is configuring webpack to use the Office Add-in trusted certificates.

### 2. teamsapp CLI Is Unreliable
The `teamsapp provision` command frequently fails with cryptic errors. Don't waste time debugging it - just create the zip manually and sideload via Teams → Manage your apps → Upload custom app.

### 3. Sideloading via Teams Is Required
You can't just F5 debug like a normal add-in. Copilot agents must be sideloaded through Teams to register with M365.

### 4. Regular Copilot May Respond Instead of Your Agent
Just because you get a response doesn't mean your code ran. Regular Copilot can intercept and respond to prompts. Always add a unique identifier string to your response (like `🔐 CUSTOM AGENT:`) to prove your JavaScript function executed.

### 5. Dev Server Must Be Running
If the dev server stops while Word is open with your agent, Copilot will crash with "Something went wrong." You'll need to remove the app, clear caches, and re-upload.

### 6. Zip Structure Is Picky
The manifest files must be at the **root** of the zip, not in a subfolder. If you create the zip wrong, Teams will give a manifest parsing error.

### 7. Propagation Takes Time
After uploading to Teams, wait 1-2 minutes before expecting the agent to appear in Word's Copilot. Press Ctrl+R in the Copilot pane to force refresh.

### 8. Office Cache Can Cause Stale Behavior
When things aren't updating, clear: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`

---

## Reference Documentation

- [Combine Copilot Agents with Office Add-ins (Overview)](https://learn.microsoft.com/en-us/office/dev/add-ins/design/agent-and-add-in-overview)
- [Add a Copilot agent to an add-in (How-to)](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/agent-and-add-in)
- [Build your first add-in as a Copilot skill (Quickstart)](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/agent-and-add-in-quickstart)
- [Declarative agent manifest schema](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/declarative-agent-manifest-1.5)
- [API plugin manifest schema](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api-plugin-manifest-2.3)

---

*Last Updated: January 6, 2026 - Based on hands-on POC experience*
