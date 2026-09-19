# Office AI Assistant


<div align="center">

[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![Office](https://img.shields.io/badge/office-Excel%20Word%20PowerPoint-green.svg)](https://www.microsoft.com/office)

**🌐 Language / 语言 / Язык**

[Русский](README.md) | [English](README_EN.md) | [中文](README_ZH.md)

</div>

> **Note**: This project is developed with 90%+ code using Cursor + Copilot + Qoder + Claude programming tools
>
> Source code local running tutorial：https://www.bilibili.com/cheese/play/ep2098181
>
> Excel/Word/PPT Plugin Tri-Pack Deployment Packaging Tutorial：https://www.bilibili.com/cheese/play/ep2098188
>
> 


![ExcelView](./AiHelper.assets/excelai_display.png)

![WordView](./AiHelper.assets/wordai_display.png)

![PPTView](./AiHelper.assets/pptai_display.png)



## 📖 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [AI Native Architecture](#ai-native-architecture)
- [Supported Products](#supported-products)
- [Screenshots](#screenshots)
- [Installation](#installation)
- [Usage](#usage)
- [Development](#development)
- [Contributing](#contributing)
- [License](#license)

---

## Overview

Office AI Assistant is an AI Native Office add-in built with **Visual Studio Community 2022/2026 + Visual Basic.NET + VSTO + WebView2**. It brings chat-driven task panes, document context reading, smart formatting and proofreading, data analysis, content generation, MCP/Skills integration, and agentic execution to Excel, Word, and PowerPoint.

### 🎯 Project Goals

- **Intelligent Office Automation**: Provide AI-powered assistance for daily office tasks
- **Multi-Platform Support**: Compatible with Microsoft Office and WPS
- **User-Friendly Interface**: Simple and intuitive operation
- **Continuous Improvement**: Regular updates and feature enhancements

---

## Features

### ✨ Core Features

- **AI Native Chat**: Ask the current document, selected range, table, or slide to do real work from the Office task pane.
- **Word Smart Formatting**: Natural-language formatting, heading/numbering restructuring, automatic numbering repair, style unification, preview, and execution feedback.
- **Word Proofreading**: Review typos, punctuation, wording, and formal-document expression with focused suggestions.
- **Excel Data Analysis**: Read selected ranges and worksheets, answer data questions, organize analysis, assist formulas, and generate chart ideas.
- **PowerPoint Assistance**: Generate, continue, translate, format, and review presentation content.
- **AI Translation**: Translate selected paragraphs, pages, or documents with configurable models.
- **Memory and Context**: Combine short-term conversation state, long-term memory retrieval, and current Office file context.
- **MCP / Skills Integration**: Built-in MCP Client support for MCP Servers and skills.
- **Agent Loop**: Complex tasks follow plan -> act -> observe -> repair -> explain instead of stopping at a chat response.
- **Multi-Model Configuration**: Supports DeepSeek, Doubao, and extensible model configuration.

### 🔧 Technical Features

- **VSTO Integration**: Seamless integration with Microsoft Office
- **Cross-Platform Compatibility**: Works with both Microsoft Office and WPS
- **Modern UI**: Clean and responsive user interface
- **Extensible Architecture**: Easy to extend and customize
- **MCP Protocol Support**: Native MCP-Client implementation for server communication
- **DeepSeek Optimization**: Enhanced DeepSeek API integration with improved performance

---

## AI Native Architecture

The project is moving from a chat-style add-in to an executable Office Agent. The goal is for AI to observe the current Office file, plan an operation, call tools, verify the result, and repair failures when needed.

#### 🎯 Core Capabilities

| Capability | Status | Description |
|------------|--------|-------------|
| **Harness** | In progress | Wrap Office operations as AI-callable, observable, and verifiable capabilities |
| **Agent Loop** | In progress | Route complex tasks through planning, execution, observation, repair, and explanation |
| **Context Awareness** | Available | Read selections, whole documents, paragraphs, styles, data tables, and conversation context |
| **Word Action Harness** | Available | Natural-language entry point for Word formatting, proofreading, and numbering operations |
| **MCP Client** | Available | Connect MCP Servers to extend external tool capabilities |
| **Memory** | Available | Retrieve recent sessions and long-term memory |
| **Preview / Undo Direction** | In progress | Prefer preview, explanation, and recoverable execution for important write operations |

#### ✨ Experience Goals

**1. Word actions should execute, not just chat**

```text
User: "Change the front numbering to 1 2 3 4 5"
  ↓
Harness reads Word paragraphs, numbering, and range context
  ↓
Numbering Agent plans and applies automatic numbering repair
  ↓
Observer verifies the numbering and reports the result
```

**2. The add-in should read context by itself**

```text
User selects several Word paragraphs
User: "Restructure the numbering and headings"

AI should not repeatedly ask for range, numbering style, or selected text when the add-in can observe them. It should read the context first, then provide an executable plan or preview.
```

**3. Excel data understanding**

```text
User selects A1:B100 sales data
AI automatically knows:
- Range: A1:B100
- Data Type: Table with headers
- Column Names: Salesperson, Sales Amount
- Preview: First 3 rows

User: "Calculate commission"
AI knows where to work without asking again.
```

#### 🔍 Technical Highlights

- **Harness atomic tools**: Expose Office actions such as `ListParagraphs`, `GetParagraphInfo`, and `SetParagraphFormat` as composable capabilities.
- **Planner / Agent**: Produce structured plans from user goals instead of turning every request into a chat reply.
- **Observer / Repair**: Read the real Office state after execution and repair the plan if the result is not correct.
- **Capability / Skill**: Add new features as discoverable, explainable, and testable units.

#### 📖 Documentation

- [Visual Studio 2026 Installer Projects troubleshooting](docs/VS2026-InstallerProjects.md) (Chinese)

---

## Supported Products

| Product | Status | Features |
|---------|--------|----------|
| **Microsoft Excel** | ✅ Supported | Data analysis, chart generation, formula assistance, ALLM/CLLM functions, selected-range Q&A |
| **Microsoft Word** | ✅ Supported | Document processing, content generation/completion, proofreading, continuation, formatting, heading/numbering restructuring, translation |
| **Microsoft PowerPoint** | ✅ Supported | Presentation creation, slide design, review, continuation, formatting, translation |
| **WPS Office** | ✅ Compatible | Full compatibility with WPS suite |

---

## Screenshots

### Excel Examples

Analyze selected sheets, cells, and external Excel files for intelligent Q&A and data organization.

![Excel GDP Analysis 1](./AiHelper.assets/excelgdp1.png)

![Excel GDP Analysis 2](./AiHelper.assets/excelgdp2.png)

### Word Examples

Import and analyze data from external websites, then insert results into Word documents.

![Word Data Integration 1](./AiHelper.assets/word1.png)

![Word Data Integration 2](./AiHelper.assets/word2.png)

---

## Installation

### Prerequisites

- **Operating System**: Windows 10/11
- **Office Suite**: Microsoft Office 2016+ or WPS Office
- **Runtime**: .NET Framework 4.7.2+, Microsoft Edge WebView2 Runtime, VSTO Runtime
- **Development Environment**: Visual Studio Community 2022/2026 with the VSTO workload
- **MSI Packaging**: `OfficeAgent/OfficeAgent.vdproj` requires the `Microsoft Visual Studio Installer Projects` extension in the current Visual Studio instance

### User Installation

1. Close Word, Excel, PowerPoint, and WPS.
2. Download `OfficeAgent.msi` from the official site or a release page.
3. Run the installer. If enterprise policy blocks Office add-in registry writes, run it as administrator.
4. Open Word, Excel, or PowerPoint.
5. Find the AI Assistant entry in the ribbon and configure models, API keys, MCP, and skills on first use.
6. If Office disables the add-in, re-enable it from `File -> Options -> Add-ins`.

### 📦 Download Installer

- **Official Download**: [https://www.officeso.cn/](https://www.officeso.cn/)
- **Latest Version**: Get the most recent stable release
- **Easy Installation**: One-click installer for Windows

### Build and Debug from Source

```bash
# Restore dependencies
msbuild AiHelper.sln -t:Restore

# Build all code projects
msbuild AiHelper.sln

# Or open AiHelper.sln in Visual Studio and run Rebuild Solution
```

For VSTO debugging, launch the corresponding Office host from Visual Studio. Debug builds are for development; Release builds plus MSI installation are recommended for distribution.

### Visual Studio 2026 Cannot Load OfficeAgent.vdproj

If VS 2026 reports that `OfficeAgent.vdproj` is incompatible and shows `54435603-dbb4-11d2-8724-00a0c9a8b90c`, the usual cause is missing `.vdproj` project type support in that Visual Studio instance, not an outdated VB.NET project.

Fix:

1. Install and enable [Microsoft Visual Studio Installer Projects](https://marketplace.visualstudio.com/items?itemName=VisualStudioClient.MicrosoftVisualStudio2022InstallerProjects) in VS 2026.
2. Restart Visual Studio and reopen `AiHelper.sln`.
3. If the extension is temporarily unavailable, build and debug `ShareRibbon`, `WordAi`, `ExcelAi`, and `PowerPointAi` first, then build the MSI with a Visual Studio instance that supports the extension.
4. See [Visual Studio 2026 Installer Projects troubleshooting](docs/VS2026-InstallerProjects.md) for details.

### Excel XLL Digital Signature Prompt

In Debug builds, `ExcelAi-AddIn64.xll` may show a missing digital signature warning. For local development, use Office trusted locations, enable the add-in, or install from an internal trusted path. Paid code signing is only needed when you decide to distribute a signed production build.

---

## Usage

### Getting Started

1. **Launch Office Application**: Open Excel, Word, or PowerPoint
2. **Access AI Assistant**: Find the AI Assistant tab in the ribbon
3. **Configure Models**: Set API keys, model endpoints, and preferences
4. **Reference Context**: Select document text, a table range, or a slide, then type what you want
5. **Execute Tasks**: Let AI analyze, generate, format, proofread, or explain based on the current file

### Word Examples

- `Restructure the numbering and headings`
- `Change the front numbering to 1 2 3 4 5`
- `Increase the whole document font size by 2 points`
- `Proofread the whole document and only fix obvious typos and punctuation`
- `Rewrite the selected content in a formal report style`

### Excel Examples

- `Analyze anomalies in the selected range`
- `Generate a business analysis based on this table`
- `Suggest a chart for the current data`
- `Fill this column with a formula and explain it`

### PowerPoint Examples

- `Generate a summary slide for the current topic`
- `Improve the wording on the selected slide`
- `Make this slide look like a business report`
- `Check the whole deck for typos and wording issues`

### Advanced Features

- **MCP/Skills**: Connect external tools, knowledge bases, or enterprise systems.
- **Memory**: Use recent conversations and long-term memory to reduce repeated instructions.
- **Execution Explanation**: Show what will be changed, what was changed, and what still needs confirmation.
- **Preview First**: Prefer preview and execution feedback for formatting and proofreading operations.

---

## Development

### Development Environment

- **IDE**: Visual Studio Community 2022/2026
- **Language**: Visual Basic.NET
- **Framework**: VSTO (Visual Studio Tools for Office)
- **Version Control**: Git

### Project Structure

```
AiHelper/
├── ExcelAi/          # Excel Add-in
├── WordAi/           # Word Add-in
├── PowerPointAi/     # PowerPoint Add-in
├── ShareRibbon/      # Shared Components
├── OfficeAgent/      # Installer
├── docs/             # Debugging and installation docs
└── openspec/         # Requirements and design specs
```

### Building from Source

```bash
# Clone the repository
git clone https://github.com/it235/office-ai-agent.git

# Restore dependencies
msbuild AiHelper.sln -t:Restore

# Build the solution
msbuild AiHelper.sln
```

Before building the installer, make sure Visual Studio has the `Microsoft Visual Studio Installer Projects` extension installed; otherwise `OfficeAgent.vdproj` cannot load.

---

## Contributing

We welcome contributions from the community! Here's how you can help:

### 🤝 How to Contribute

1. **Fork the Repository**: Create your own fork of the project
2. **Create a Branch**: Make your changes in a new branch
3. **Make Changes**: Implement your features or fixes
4. **Test Thoroughly**: Ensure your changes work correctly
5. **Submit Pull Request**: Create a PR with detailed description

### 📋 Contribution Guidelines

- **Code Style**: Follow existing code conventions
- **Documentation**: Update documentation for new features
- **Testing**: Include tests for new functionality
- **Communication**: Discuss major changes in issues first

---

## License

This project is licensed under the Apache 2.0 License - see the [LICENSE](LICENSE) file for details.

---

## 📞 Contact & Support

### 🌐 Official Website
- **Website**: https://www.officeso.cn/
- **Course**: [OfficeAI办公智能体开发(基于vb.net)](https://www.bilibili.com/cheese/play/ep1540657)

### 📅 Update Schedule
- **Release Cycle**: Every 2 weeks
- **Follow Us**: Stay updated with latest releases

### 💬 Community
- **Issues**: Report bugs and request features
- **Discussions**: Share ideas and ask questions
- **Contributions**: Help improve the project

---

## 🙏 Acknowledgments

Special thanks to:

- **DeepSeek**: For providing excellent AI models
- **Microsoft**: For VSTO framework and Office APIs
- **Open Source Community**: For inspiration and support
- **Cursor & Copilot**: For AI-powered development assistance

---

<div align="center">

**Made with ❤️ for the Office community**

[![GitHub stars](https://img.shields.io/github/stars/it235/office-ai-agent?style=social)](https://github.com/it235/office-ai-agent)
[![GitHub forks](https://img.shields.io/github/forks/it235/office-ai-agent?style=social)](https://github.com/it235/office-ai-agent)

</div>

