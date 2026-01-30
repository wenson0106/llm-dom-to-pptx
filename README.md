# llm-dom-to-pptx

**Turn AI-generated HTML into native, editable PowerPoint slides.**

`llm-dom-to-pptx` is a lightweight JavaScript library designed to bridge the gap between LLM-generated web designs (HTML/CSS) and Office productivity tools. Unlike screenshot-based tools, this library parses the DOM to create **fully editable** shapes, text blocks, and tables in PowerPoint.

https://github.com/user-attachments/assets/527efb28-b0ae-450a-9710-60cac2924acc


## 🚀 Features

| Feature | Status |
|---------|--------|
| Background Colors & Transparency | ✅ |
| Borders & Rounded Corners | ✅ |
| **Transparent Borders (New)** | ✅ |
| Box Shadows | ✅ |
| Gradient Backgrounds (Linear/Radial) | ✅ |
| SVG Icons | ✅ |
| Text Styling (Bold/Italic/Underline/Letter Spacing) | ✅ |
| Multi-Slide Export | ✅ |
| Modular API (Manual Control) | ✅ |

## 📦 Installation

### CDN (Recommended)
```html
<!-- Dependencies -->
<script src="https://cdn.jsdelivr.net/gh/gitbrent/PptxGenJS@3.12.0/libs/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/gh/gitbrent/PptxGenJS@3.12.0/dist/pptxgen.bundle.js"></script>

<!-- llm-dom-to-pptx v1.2.6 -->
<script src="https://cdn.jsdelivr.net/npm/llm-dom-to-pptx@1.2.6/dist/llm-dom-to-pptx.js"></script>
```

## 🎯 Usage

### Basic Export (Single Slide)
```javascript
// Export by element ID
await LLMDomToPptx.export('slide-canvas', { fileName: 'presentation.pptx' });
```

### Multi-Slide Export
```javascript
// Export all matching elements as separate slides
await LLMDomToPptx.export('.slide-page', { fileName: 'multi_slides.pptx' });
```

### Advanced Usage (Manual Control)
```javascript
const pres = new PptxGenJS();
pres.layout = 'LAYOUT_16x9';

const slide = pres.addSlide();
slide.background = { color: 'FFFFFF' };

// Add DOM element to slide with offset positioning
await LLMDomToPptx.addSlide(pres, slide, document.getElementById('my-card'), {
    x: 0.5,  // Horizontal offset (inches)
    y: 0.3   // Vertical offset (inches)
});

await pres.writeFile({ fileName: 'custom.pptx' });
```

## 📐 HTML Structure

For best results, use a fixed-size container (960×540px for 16:9):

```html
<!-- Added class for multi-slide selection -->
<div id="slide-canvas" class="slide-page" style="width: 960px; height: 540px; position: relative;">
    <h1>Quarterly Report</h1>
    <p>Success driven by innovation.</p>
</div>
```

## 🧠 System Prompt

This library works best with **LLM-generated HTML** that follows specific constraints. Use the included `System_Prompt.md` when instructing your LLM (GPT-4, Claude, etc.) to generate slides.

**Key constraints:**
- Root container: `width: 960px; height: 540px`
- Use absolute positioning for major sections
- Use Flexbox for internal layouts
- Stick to standard web fonts

## 📋 API Reference

### `LLMDomToPptx.export(selector, options)`
| Parameter | Type | Description |
|-----------|------|-------------|
| `selector` | `string \| HTMLElement` | Element ID, CSS selector, or DOM element |
| `options.fileName` | `string` | Output filename (default: `presentation.pptx`) |
| `options.x` | `number` | Horizontal offset in inches |
| `options.y` | `number` | Vertical offset in inches |

### `LLMDomToPptx.addSlide(pres, slide, element, options)`
| Parameter | Type | Description |
|-----------|------|-------------|
| `pres` | `PptxGenJS` | PptxGenJS presentation instance |
| `slide` | `Slide` | Slide object from `pres.addSlide()` |
| `element` | `HTMLElement` | DOM element to convert |
| `options.x` | `number` | Horizontal offset in inches |
| `options.y` | `number` | Vertical offset in inches |

## 🙏 Acknowledgements

- **PptxGenJS**: The core engine powering PPTX generation.
- **Open Source Community**: For continuous inspiration and tools.

---

**Version:** 1.2.6
