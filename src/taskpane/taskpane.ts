/*
 * Document Analyzer - Task Pane
 * This provides a traditional task pane UI in addition to the Copilot agent
 */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("app-body")!.style.display = "flex";
        document.getElementById("analyze-btn")!.onclick = runAnalysis;
    }
});

async function runAnalysis() {
    const resultDiv = document.getElementById("result")!;
    const analysisType = (document.getElementById("analysis-type") as HTMLSelectElement).value;
    
    resultDiv.innerHTML = "<p>Analyzing document...</p>";
    
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.load("text");
            await context.sync();
            
            const content = body.text;
            const wordCount = content.split(/\s+/).length;
            
            resultDiv.innerHTML = `
                <h3>Analysis Complete</h3>
                <p><strong>Type:</strong> ${analysisType}</p>
                <p><strong>Word Count:</strong> ${wordCount}</p>
                <p><strong>Character Count:</strong> ${content.length}</p>
                <hr>
                <p><em>For full AI-powered analysis, use the Copilot agent!</em></p>
                <p>Open Copilot and select "Document Analyzer Agent" to get detailed analysis.</p>
            `;
        });
    } catch (error) {
        resultDiv.innerHTML = `<p style="color: red;">Error: ${error}</p>`;
    }
}
