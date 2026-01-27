# llm-dom-to-pptx

**Turn AI-generated HTML into native, editable PowerPoint slides.**

`llm-dom-to-pptx` is a lightweight JavaScript library designed to bridge the gap between LLM-generated web designs (HTML/CSS) and Office productivity tools. Unlike screenshot-based tools, this library parses the DOM to create **fully editable** shapes, text blocks, and tables in PowerPoint.

https://github.com/user-attachments/assets/527efb28-b0ae-450a-9710-60cac2924acc


## 🚀 Features

- **Semantic Parsing**: Intelligently maps HTML structure (Flexbox, Grid, Tables) to PPTX layouts.
- **Style Preservation**: captures background colors, rounded corners, borders, shadows, and fonts.
- **Text Precision**: Handles complex text runs, bolding, colors, and alignments.
- **Zero Backend**: Runs entirely in the browser using [PptxGenJS](https://gitbrent.github.io/PptxGenJS/).

## 📦 Installation & Usage

You can use the library directly via a script tag (e.g., via CDN once hosted, or locally).

### 1. Include Dependencies
This library depends on `PptxGenJS`.

```html
<!-- PptxGenJS (Required) -->
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.min.js"></script>

<!-- llm-dom-to-pptx -->
<script src="https://cdn.jsdelivr.net/npm/llm-dom-to-pptx@1.0.0/dist/llm-dom-to-pptx.js"></script>
```

### 2. Prepare Your HTML
Create a container for your slide. A fixed width (e.g., 960px) works best for 16:9 mapping.

```html
<div id="slide-canvas" style="width: 960px; height: 540px; background: white;">
    <!-- Your AI-generated content here -->
    <h1>Quarterly Report</h1>
    <p>Success driven by innovation.</p>
</div>
```

### 3. Export to PPTX
Call the export function when ready.

```javascript
// Export using the element ID
window.LLMDomToPptx.export('slide-canvas', { fileName: 'My_AI_Presentation.pptx' });
```

## 🧠 The Secret Sauce: System Prompt & Spec

**Crucial:** This library is not a generic "convert any website to PPTX" tool. It is designed to work with **LLM-generated HTML** that follows specific constraints (the "Laws of Physics").

To get good results, you must instruct your LLM (GPT-4, Claude 3.5, etc.) to generate HTML that this parser understands.

### 1. The System Prompt
We have provided a robust `System_Prompt.md` file in this repo. You **MUST** use this (or a variation of it) when asking an LLM to generate slides.

**Key constraints enforced by the prompt:**
*   **Root Container:** `<div id="slide-canvas" style="width: 960px; height: 540px;">`
*   **Layouts:** Uses absolute positioning for major sections, Flexbox for internals.
*   **Shapes:** Defines how CSS `border-radius` maps to PPTX Shapes (Rectangles vs Ellipses).
*   **Typography:** Enforces standard fonts that map to PPTX-safe fonts.

### 2. The Spec
We will include a more complete version of the specification in future updates to fully explain all supported DOM and CSS elements.

## 🙏 Acknowledgements

*   **PptxGenJS**: The core engine that powers the PPTX generation.
*   **Open Source Community**: For the continuous inspiration and tools.

---



