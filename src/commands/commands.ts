/*
 * Document Analyzer - Copilot Agent Commands
 * These functions are called by the Copilot agent to analyze Word documents
 */

// Analysis prompts for different types
const ANALYSIS_PROMPTS: Record<string, string> = {
    compliance: `You are a compliance analyst. Analyze this document for:
- Regulatory compliance issues
- Policy violations
- Legal concerns
- Missing required disclaimers or disclosures
- Improper language or terminology

Provide specific findings with severity (High/Medium/Low) and recommendations.`,

    completeness: `You are a document reviewer. Check if this document is complete by looking for:
- Missing required sections
- Incomplete paragraphs or sentences
- Placeholder text (like "[TBD]", "XXX", "TODO")
- Missing references or citations
- Gaps in logical flow

Provide specific findings and recommendations for what needs to be added.`,

    consistency: `You are an editor. Analyze this document for consistency issues:
- Contradictory statements
- Inconsistent terminology (same concept called different names)
- Conflicting dates or numbers
- Tone or style inconsistencies
- Formatting inconsistencies

Provide specific findings with locations and recommendations.`,

    sensitivity: `You are a data protection specialist. Identify sensitive information in this document:
- Personal Identifiable Information (PII): names, SSNs, addresses, phone numbers, emails
- Financial data: account numbers, credit card numbers, salaries
- Health information (PHI)
- Confidential business information
- Passwords or credentials

Provide specific findings with recommendations for redaction or protection.`
};

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
 * Calls Azure OpenAI to analyze the document
 * TODO: Replace with your actual Azure OpenAI endpoint and key
 */
async function callAzureOpenAI(content: string, analysisType: string): Promise<string> {
    // Configuration - UPDATE THESE VALUES
    const AZURE_OPENAI_ENDPOINT = "https://YOUR-RESOURCE.openai.azure.com";
    const AZURE_OPENAI_KEY = "YOUR-API-KEY"; // Move to secure storage in production!
    const DEPLOYMENT_NAME = "gpt-4o";
    const API_VERSION = "2024-02-15-preview";

    const systemPrompt = ANALYSIS_PROMPTS[analysisType] || ANALYSIS_PROMPTS.compliance;

    try {
        const response = await fetch(
            `${AZURE_OPENAI_ENDPOINT}/openai/deployments/${DEPLOYMENT_NAME}/chat/completions?api-version=${API_VERSION}`,
            {
                method: "POST",
                headers: {
                    "api-key": AZURE_OPENAI_KEY,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    messages: [
                        { role: "system", content: systemPrompt },
                        { role: "user", content: `Please analyze the following document:\n\n${content}` }
                    ],
                    max_tokens: 2000,
                    temperature: 0.3
                })
            }
        );

        if (!response.ok) {
            const errorText = await response.text();
            console.error("Azure OpenAI error:", errorText);
            return `Error calling Azure OpenAI: ${response.status} ${response.statusText}. Please check your API configuration.`;
        }

        const data = await response.json();
        return data.choices[0].message.content;

    } catch (error) {
        console.error("Error calling Azure OpenAI:", error);
        
        // Return a demo response if Azure OpenAI is not configured
        return generateDemoResponse(content, analysisType);
    }
}

/**
 * Generates a demo response when Azure OpenAI is not configured
 */
function generateDemoResponse(content: string, analysisType: string): string {
    const wordCount = content.split(/\s+/).length;
    const charCount = content.length;
    
    const demoResponses: Record<string, string> = {
        compliance: `## Compliance Analysis Report

**Document Statistics:** ${wordCount} words, ${charCount} characters

### Findings

⚠️ **Note:** This is a demo response. Configure Azure OpenAI for actual analysis.

**Sample findings that would be detected:**
1. **[Medium]** Missing privacy disclaimer in footer
2. **[Low]** Informal language detected in section 3
3. **[High]** No signature block present

### Recommendations
- Add standard privacy disclaimer
- Review and formalize language
- Include signature and date fields`,

        completeness: `## Completeness Analysis Report

**Document Statistics:** ${wordCount} words, ${charCount} characters

### Findings

⚠️ **Note:** This is a demo response. Configure Azure OpenAI for actual analysis.

**Sample findings that would be detected:**
1. Missing Executive Summary section
2. Placeholder text "[TBD]" found in 2 locations
3. References section is empty

### Recommendations
- Add executive summary
- Replace all placeholder text
- Complete references section`,

        consistency: `## Consistency Analysis Report

**Document Statistics:** ${wordCount} words, ${charCount} characters

### Findings

⚠️ **Note:** This is a demo response. Configure Azure OpenAI for actual analysis.

**Sample findings that would be detected:**
1. "Client" and "Customer" used interchangeably
2. Date format inconsistent (MM/DD/YYYY vs DD-MM-YYYY)
3. Heading styles vary between sections

### Recommendations
- Standardize terminology
- Use consistent date format
- Apply uniform heading styles`,

        sensitivity: `## 🔐 CUSTOM AGENT RESPONSE - Sensitivity Analysis

**🤖 Agent ID: DocumentAnalyzerAgent v1.0**
**Document Statistics:** ${wordCount} words, ${charCount} characters

### Findings

⚠️ **This is YOUR custom agent running, not regular Copilot!**

If you see this message, the Document Analyzer Agent add-in successfully:
1. Received your request via Copilot
2. Read ${charCount} characters from your Word document
3. Returned this custom response

**Sample sensitivity findings that would be detected:**
1. **[High]** Email addresses: Look for @domain patterns
2. **[Medium]** Phone numbers: Look for XXX-XXX-XXXX patterns  
3. **[Low]** Names that may be PII

### To Enable Real AI Analysis
Configure Azure OpenAI in commands.ts with your endpoint and key.`
    };

    return demoResponses[analysisType] || demoResponses.compliance;
}

/**
 * Main analysis function - called by Copilot agent
 */
async function analyzeDocument(analysisType: string): Promise<string> {
    try {
        console.log(`Starting ${analysisType} analysis...`);
        
        // Get document content
        const content = await getDocumentContent();
        
        if (!content || content.trim().length === 0) {
            return "The document appears to be empty. Please add content to analyze.";
        }
        
        console.log(`Document loaded: ${content.length} characters`);
        
        // Call Azure OpenAI for analysis
        const analysis = await callAzureOpenAI(content, analysisType);
        
        return analysis;
        
    } catch (error) {
        console.error("Error in analyzeDocument:", error);
        return `Error analyzing document: ${error instanceof Error ? error.message : String(error)}`;
    }
}

// Initialize Office and register the agent action
Office.onReady((info) => {
    console.log("Office.js initialized. Host:", info.host);
    
    // Register the analyzeDocument function with Copilot
    Office.actions.associate("analyzeDocument", async (message: string) => {
        console.log("Copilot called analyzeDocument with:", message);
        
        try {
            const params = JSON.parse(message);
            const analysisType = params.analysisType || "compliance";
            
            const result = await analyzeDocument(analysisType);
            return result;
            
        } catch (error) {
            console.error("Error parsing message or running analysis:", error);
            return `Error: ${error instanceof Error ? error.message : String(error)}`;
        }
    });
    
    console.log("analyzeDocument action registered with Copilot");
});
