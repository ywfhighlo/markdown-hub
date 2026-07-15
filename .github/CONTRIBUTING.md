# Contributing to Markdown Hub

First off — **thank you** for taking the time to contribute! 🎉

This document describes how to set up the project and submit contributions.

## 📦 Project Structure

```
markdown-hub/
├── src/                      # TypeScript — VS Code extension (UI, commands, config)
│   ├── extension.ts          # Entry point, command registration
│   ├── commandHandler.ts     # Command dispatch + progress UI
│   ├── pythonService.ts      # Spawns Python backend
│   └── dependencyChecker.ts  # Dependency detection
├── backend/                  # Python — actual conversion logic
│   ├── cli.py                # Entry point called by TypeScript
│   └── converters/
│       ├── md_to_office.py   # Markdown → DOCX/PDF/HTML/PPTX
│       ├── office_to_md.py   # Office → Markdown (PDF/DOCX/XLSX/PPTX)
│       ├── diagram_to_png.py # SVG/Mermaid/PlantUML → PNG
│       └── base_converter.py # Abstract base + registry
├── tools/                    # Bundled binaries (PlantUML.jar, Poppler, Batik)
├── media/                    # Icon
└── package.json              # VS Code manifest
```

**Architecture**: TypeScript (frontend) spawns Python (backend) via `cli.py`.
Each converter is independent and lazy-loads its own dependencies —
a missing Python library only breaks its own feature.

## 🚀 Development Setup

### Prerequisites
- [Node.js](https://nodejs.org/) 18+
- [Python](https://www.python.org/) 3.8+
- [Pandoc](https://pandoc.org/installing.html) (for MD → DOCX/PDF/HTML testing)
- VS Code

### Install & build
```bash
git clone https://github.com/ywfhighlo/markdown-hub.git
cd markdown-hub
npm install
npm run compile       # TypeScript → out/
npm run watch         # or watch mode during development
```

### Install Python dependencies
Each conversion feature is independent — install only what you need.
See [README.md](../README.md#prerequisites) for the per-feature list.

### Debug in VS Code
1. Open the project folder in VS Code.
2. Press `F5` to launch an Extension Development Host.
3. Test your changes against real files.

### Package & install locally
```bash
npx @vscode/vsce package
code --install-extension markdown-hub-0.3.6.vsix
```

## 🧪 Testing

We use lightweight Python tests for the core logic (table parsing, highlight, etc.):

```bash
python -m pytest backend/tests/ -v
```

When adding a feature or fixing a bug, please add a focused test under `backend/tests/`.

## 📝 Code Style

### TypeScript
- Follow the existing style in `src/`.
- Run `npm run lint` before submitting.
- Use `const` / `async-await` consistently.

### Python
- Follow [PEP 8](https://peps.python.org/pep-0008/).
- Use type hints (`from typing import List, Optional, ...`) as the existing code does.
- Lazy-load optional dependencies (`lib_available` + `importlib`) so a missing
  library never breaks unrelated features — see `_resolve_pptx()` in
  `md_to_office.py` for the pattern.

## 🔄 Pull Request Workflow

1. **Fork** the repo and create a branch from `main`:
   ```bash
   git checkout -b fix/my-bug-fix
   ```
2. **Commit** your changes with a clear message (Chinese or English are both fine).
3. **Test** locally: `npm run compile && python -m pytest backend/tests/`.
4. **Update CHANGELOG.md** under the `[Unreleased]` section.
5. **Open a Pull Request** and fill in the template.

### Commit message conventions
We don't enforce strict conventions, but here are good prefixes:
- `fix:` bug fix
- `feat:` new feature
- `docs:` documentation only
- `refactor:` code restructuring without behavior change
- `chore:` tooling, dependencies, CI

Example: `fix: table column alignment ignored when separator has colons`

## 🐛 Reporting Bugs

Use the [Bug Report template](https://github.com/ywfhighlo/markdown-hub/issues/new?template=bug_report.md).
The more detail you provide (OS, versions, input file, dependency check output),
the faster we can reproduce and fix it.

## 💡 Requesting Features

Use the [Feature Request template](https://github.com/ywfhighlo/markdown-hub/issues/new?template=feature_request.md).

## 🌍 Internationalization

The project's primary audience is bilingual (Chinese / English).
- Code comments: English or Chinese are both acceptable.
- User-facing strings (command titles, error messages): English preferred,
  with Chinese as a secondary option if you can provide both.
- Documentation: English in `README.md` (primary), Chinese in `README_zh.md`.

## ❓ Questions?

Open a [Discussion](https://github.com/ywfhighlo/markdown-hub/discussions) or an issue.
We're happy to help.

Thanks again for contributing! 💚
