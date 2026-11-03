// Minimal `vscode` API stub so out/*.js can run under plain node.
// Only what pythonLocator.ts touches: workspace configuration reads.
module.exports = {
    workspace: {
        getConfiguration: () => ({
            get: () => undefined,   // simulate "markdown-hub.pythonPath" not set
            update: () => Promise.resolve()
        })
    }
};
