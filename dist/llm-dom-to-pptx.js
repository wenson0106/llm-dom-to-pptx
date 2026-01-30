/**
 * LLM DOM to PPTX - v1.2.3-FONT-SIM
 * Implements "Pre-Export Font Simulation" to fix text wrapping.
 * Function: Swaps browser fonts to PPTX-safe fonts (e.g. Arial) before measuring to ensure metrics match.
 * Use: LLMDomToPptx.export('selector', { fileName: '...' })
 */
(function (root) {
    'use strict';

    // ==========================================
    // 1. CONFIGURATION & CONSTANTS
    // ==========================================
    const CONFIG = {
        PPI: 96,
        SLIDE_WIDTH_PX: 960,
        PPT_WIDTH_IN: 10,
        PPT_HEIGHT_IN: 5.625
    };
    const SCALE = CONFIG.PPT_WIDTH_IN / CONFIG.SLIDE_WIDTH_PX;

    function safeNum(val, min = 0) {
        let n = parseFloat(val);
        if (isNaN(n) || !isFinite(n)) return min;
        return n < min ? min : n;
    }

    const FONT_MAP = {
        "Inter": "Arial", "Roboto": "Arial", "Open Sans": "Calibri", "Lato": "Calibri",
        "Montserrat": "Arial", "Source Sans Pro": "Arial", "Noto Sans": "Arial",
        "Helvetica": "Arial", "San Francisco": "Arial", "Segoe UI": "Segoe UI",
        "Times New Roman": "Times New Roman", "Georgia": "Georgia",
        "Courier New": "Courier New", "Fira Code": "Courier New", "monospace": "Courier New"
    };

    function getSafeFont(fontFamily) {
        if (!fontFamily) return "Arial";
        const clean = fontFamily.replace(/['"]/g, '');
        const fonts = clean.split(',').map(f => f.trim());
        for (let f of fonts) {
            if (FONT_MAP[f]) return FONT_MAP[f];
            const key = Object.keys(FONT_MAP).find(k => k.toLowerCase() === f.toLowerCase());
            if (key) return FONT_MAP[key];
        }
        return "Arial";
    }

    function pxToInch(px) {
        return safeNum(parseFloat(px) * SCALE);
    }

    // ==========================================
    // 1.1 FONT SIMULATION HELPERS
    // ==========================================

    /**
     * Temporarily swaps the font-family of the element and all its children
     * to the target PPTX-safe font.
     */
    function simulatePPTXFonts(container) {
        const elements = [container, ...container.querySelectorAll('*')];
        elements.forEach(el => {
            // Only process safe-to-style elements
            if (el.style) {
                const computed = window.getComputedStyle(el);
                const originalFont = computed.fontFamily;
                const originalWeight = computed.fontWeight; // string "400", "700", "bold"

                const safe = getSafeFont(originalFont);

                // Normalize Weight: PPTX usually treats >=600 as Bold
                let safeWeight = originalWeight;
                const wNum = parseInt(originalWeight);
                if (!isNaN(wNum)) {
                    safeWeight = wNum >= 600 ? '700' : '400';
                } else if (originalWeight === 'bold') {
                    safeWeight = '700';
                } else {
                    safeWeight = '400';
                }

                // If the safe font is different, swap it
                // We simply set inline style. Dataset preserves old if needed, 
                // but we can just clear inline style if it wasn't there?
                // Safer: Store ORIGINAL inline style in dataset
                if (!el.dataset.pptxOrigFont) {
                    el.dataset.pptxOrigFont = el.style.fontFamily || 'null';
                }
                if (!el.dataset.pptxOrigWeight) el.dataset.pptxOrigWeight = el.style.fontWeight || 'null';

                // Apply safe font
                el.style.fontFamily = safe;
                el.style.fontWeight = safeWeight;
            }
        });
        // Force Reflow (?) - Usually reading offsetWidth forces it, but browsers might optimize.
        // We'll trust the engine to re-layout before next read.
        return elements;
    }

    /**
     * Restores original fonts.
     */
    function restoreFonts(container) {
        const elements = [container, ...container.querySelectorAll('*')];
        elements.forEach(el => {
            if (el.dataset.pptxOrigFont !== undefined) {
                const orig = el.dataset.pptxOrigFont;
                if (orig === 'null') {
                    el.style.removeProperty('font-family');
                } else {
                    el.style.fontFamily = orig;
                }
                delete el.dataset.pptxOrigFont;
            }
            if (el.dataset.pptxOrigWeight !== undefined) {
                const orig = el.dataset.pptxOrigWeight;
                if (orig === 'null') el.style.removeProperty('font-weight');
                else el.style.fontWeight = orig;
                delete el.dataset.pptxOrigWeight;
            }
        });
    }

    // ==========================================
    // 2. COLOR & GRADIENT UTILITIES
    // ==========================================
    // REPLACEMENT: Canvas-based color parsing (Robus support for Named, HSL, RGB, etc.)
    const _colorCanvas = document.createElement('canvas');
    _colorCanvas.width = 1;
    _colorCanvas.height = 1;
    const _colorCtx = _colorCanvas.getContext('2d');

    function parseColor(colorStr, opacity = 1) {
        if (!colorStr || colorStr === 'none') return null;
        if (colorStr === 'transparent' || colorStr === 'rgba(0, 0, 0, 0)') {
            return { color: 'FFFFFF', transparency: 100 };
        }

        _colorCtx.clearRect(0, 0, 1, 1);
        _colorCtx.fillStyle = colorStr;
        _colorCtx.fillRect(0, 0, 1, 1);

        const data = _colorCtx.getImageData(0, 0, 1, 1).data;
        const r = data[0];
        const g = data[1];
        const b = data[2];
        const a = data[3] / 255;

        // If completely transparent according to canvas (and not explicitly handled above),
        // it usually means invalid color or fully transparent.
        if (a === 0 && colorStr !== 'transparent' && !colorStr.includes('rgba(0')) {
            return null;
        }

        const finalAlpha = a * safeNum(opacity, 1);
        const toHex = (n) => {
            const h = Math.round(safeNum(n)).toString(16).toUpperCase();
            return h.length === 1 ? '0' + h : h;
        };

        return {
            color: toHex(r) + toHex(g) + toHex(b),
            transparency: Math.round((1 - finalAlpha) * 100)
        };
    }

    function gradientToImage(cssGradient, width, height) {
        try {
            const w = safeNum(width, 10);
            const h = safeNum(height, 10);
            const canvas = document.createElement('canvas');
            canvas.width = w;
            canvas.height = h;
            const ctx = canvas.getContext('2d');

            const gradient = ctx.createLinearGradient(0, 0, 0, h);

            const firstParen = cssGradient.indexOf('(');
            const lastParen = cssGradient.lastIndexOf(')');
            if (firstParen === -1 || lastParen === -1) throw new Error('Invalid Gradient');

            const content = cssGradient.substring(firstParen + 1, lastParen);
            const parts = content.split(/,(?![^(]*\))/).map(s => s.trim());

            let stops = parts;
            if (parts[0].startsWith('to ') || parts[0].includes('deg')) {
                stops = parts.slice(1);
            }

            stops.forEach((stop, i) => {
                const colorMatch = stop.match(/(#[0-9a-f]{3,6}|rgb[^)]+\)|[a-z]+)/i);
                if (colorMatch) {
                    try {
                        gradient.addColorStop(i / (stops.length - 1), colorMatch[0]);
                    } catch (e) { }
                }
            });

            ctx.fillStyle = gradient;
            ctx.fillRect(0, 0, w, h);
            return canvas.toDataURL('image/png');
        } catch (e) {
            return null;
        }
    }

    async function svgToImage(svgElement) {
        try {
            const clonedKey = svgElement.cloneNode(true);
            const rect = svgElement.getBoundingClientRect();
            const w = safeNum(rect.width, 24);
            const h = safeNum(rect.height, 24);
            clonedKey.setAttribute('width', w);
            clonedKey.setAttribute('height', h);

            const xml = new XMLSerializer().serializeToString(clonedKey);
            const blob = new Blob([xml], { type: 'image/svg+xml;charset=utf-8' });
            const url = URL.createObjectURL(blob);

            return new Promise(resolve => {
                const img = new Image();
                img.onload = () => {
                    const canvas = document.createElement('canvas');
                    canvas.width = w * 2;
                    canvas.height = h * 2;
                    const ctx = canvas.getContext('2d');
                    ctx.scale(2, 2);
                    ctx.drawImage(img, 0, 0, w, h);
                    URL.revokeObjectURL(url);
                    resolve(canvas.toDataURL('image/png'));
                };
                img.onerror = () => resolve(null);
                img.src = url;
            });
        } catch (e) { return null; }
    }

    // ==========================================
    // 3. CORE EXPORT CLASS
    // ==========================================

    class SlideExporter {
        constructor(pres, slide, container, options = {}) {
            this.pres = pres;
            this.slide = slide;
            this.container = container;
            this.containerRect = container.getBoundingClientRect();
            this.options = options;
            this.processedNodes = new Set();
            this.xOffset = safeNum(options.x, 0);
            this.yOffset = safeNum(options.y, 0);
        }

        async process() {
            await this.processRoot();
            const sortedChildren = this.getSortedChildren(this.container);
            for (const item of sortedChildren) {
                await this.processNode(item.node);
            }
        }

        getSortedChildren(el) {
            return Array.from(el.children)
                .map(c => ({ node: c, z: parseInt(window.getComputedStyle(c).zIndex) || 0 }))
                .sort((a, b) => a.z - b.z);
        }

        async processRoot() {
            const notes = this.container.getAttribute('data-speaker-notes');
            if (notes) this.slide.addNotes(notes);
            await this.renderVisuals(this.container, true);
        }

        async processNode(node) {
            if (node.nodeType !== Node.ELEMENT_NODE) return;
            if (this.processedNodes.has(node)) return;
            this.processedNodes.add(node);

            const style = window.getComputedStyle(node);
            const rect = node.getBoundingClientRect();
            if (style.display === 'none' || style.visibility === 'hidden' || parseFloat(style.opacity) === 0) return;
            if (rect.width < 1 || rect.height < 1) return;

            if (node.tagName === 'svg' || node.tagName === 'SVG') {
                const dataUri = await svgToImage(node);
                if (dataUri) {
                    const pos = this.getPos(rect);
                    this.slide.addImage({ data: dataUri, x: pos.x, y: pos.y, w: pos.w, h: pos.h });
                }
                const mark = (n) => { this.processedNodes.add(n); Array.from(n.children).forEach(mark); };
                mark(node);
                return;
            }

            await this.renderVisuals(node, false);

            if (this.isTextBlock(node)) {
                this.renderText(node, style, rect);
            } else {
                const children = this.getSortedChildren(node);
                for (const item of children) await this.processNode(item.node);
            }
        }

        getPos(rect) {
            return {
                x: safeNum(pxToInch(rect.left - this.containerRect.left) + this.xOffset),
                y: safeNum(pxToInch(rect.top - this.containerRect.top) + this.yOffset),
                w: safeNum(pxToInch(rect.width), 0.01),
                h: safeNum(pxToInch(rect.height), 0.01)
            };
        }

        // ==========================================
        // 4. SHARED RENDER LOGIC
        // ==========================================
        async renderVisuals(node, isRoot) {
            const style = window.getComputedStyle(node);
            const rect = node.getBoundingClientRect();
            const pos = this.getPos(rect);
            const { x, y, w, h } = pos;

            const opacity = parseFloat(style.opacity) || 1;
            let bgParsed = parseColor(style.backgroundColor, opacity);
            let hasFill = bgParsed && bgParsed.transparency < 100;

            const bTop = safeNum(parseFloat(style.borderTopWidth));
            const bRight = safeNum(parseFloat(style.borderRightWidth));
            const bBot = safeNum(parseFloat(style.borderBottomWidth));
            const bLeft = safeNum(parseFloat(style.borderLeftWidth));
            const hasBorder = (bTop > 0 || bRight > 0 || bBot > 0 || bLeft > 0);

            const borderRadius = parseFloat(style.borderRadius) || 0;
            const rTL = parseFloat(style.borderTopLeftRadius) || 0;
            const rTR = parseFloat(style.borderTopRightRadius) || 0;
            const rBL = parseFloat(style.borderBottomLeftRadius) || 0;
            const rBR = parseFloat(style.borderBottomRightRadius) || 0;
            const isRounded = borderRadius > 0;

            const getR = (wIn, hIn) => {
                if (!isRounded || Math.min(wIn, hIn) <= 0) return null;
                const minD = Math.min(wIn, hIn);
                const avgR = (rTL + rTR + rBL + rBR) / 4;
                const rInch = pxToInch(avgR);
                const ratio = rInch / (safeNum(pxToInch(minD)) * 0.5);
                return safeNum(Math.min(ratio, 1.0), 0);
            };

            let stdRadius = getR(rect.width, rect.height);
            let shapeName = isRounded ? 'roundRect' : 'rect';
            if ((Math.abs(rect.width - rect.height) < 2) && borderRadius >= rect.width / 2 - 1) {
                shapeName = 'ellipse';
                stdRadius = undefined;
            }

            const isUniformWidth = (bTop === bRight && bRight === bBot && bBot === bLeft && bTop > 0);
            const isUniformColor = (style.borderTopColor === style.borderRightColor && style.borderRightColor === style.borderBottomColor && style.borderBottomColor === style.borderLeftColor);
            const isUniformStyle = (style.borderTopStyle === style.borderRightStyle && style.borderRightStyle === style.borderBottomStyle && style.borderBottomStyle === style.borderLeftStyle);
            const isUniform = isUniformWidth && isUniformColor && isUniformStyle;

            if (isUniform) {
                // Native Uniform - handled in main layer below
            } else if (hasBorder && hasFill) {
                // Layered approach for NON-TRANSPARENT background
                // Draw border "underlay" shapes, then the main fill covers the interior
                if (bTop > 0 && style.borderTopStyle !== 'none') {
                    const c = parseColor(style.borderTopColor);
                    let bLi = safeNum(pxToInch(parseFloat(style.borderLeftWidth) || 0));
                    let bRi = safeNum(pxToInch(parseFloat(style.borderRightWidth) || 0));
                    let bTi = safeNum(pxToInch(parseFloat(style.borderTopWidth) || 0));

                    let lX = safeNum(x - bLi);
                    let lY = safeNum(y - bTi);
                    let lW = safeNum(w + bLi + bRi, 0.01);
                    let lH = safeNum(h + bTi, 0.01);

                    let lR = getR(safeNum(rect.width + parseFloat(style.borderLeftWidth) + parseFloat(style.borderRightWidth)), safeNum(rect.height + parseFloat(style.borderTopWidth)));

                    this.slide.addShape(shapeName, {
                        x: lX, y: lY, w: lW, h: lH,
                        fill: { color: c.color, transparency: c.transparency },
                        rectRadius: lR
                    });
                }
                if (bBot > 0 && style.borderBottomStyle !== 'none') {
                    const c = parseColor(style.borderBottomColor);
                    let bLi = safeNum(pxToInch(parseFloat(style.borderLeftWidth) || 0));
                    let bRi = safeNum(pxToInch(parseFloat(style.borderRightWidth) || 0));
                    let bBi = safeNum(pxToInch(parseFloat(style.borderBottomWidth) || 0));

                    let lX = safeNum(x - bLi);
                    let lY = safeNum(y);
                    let lW = safeNum(w + bLi + bRi, 0.01);
                    let lH = safeNum(h + bBi, 0.01);

                    let lR = getR(safeNum(rect.width + parseFloat(style.borderLeftWidth) + parseFloat(style.borderRightWidth)), safeNum(rect.height + parseFloat(style.borderBottomWidth)));

                    this.slide.addShape(shapeName, {
                        x: lX, y: lY, w: lW, h: lH,
                        fill: { color: c.color, transparency: c.transparency },
                        rectRadius: lR
                    });
                }
                if (bLeft > 0 && style.borderLeftStyle !== 'none') {
                    const c = parseColor(style.borderLeftColor);
                    let bLi = safeNum(pxToInch(parseFloat(style.borderLeftWidth) || 0));
                    let lX = safeNum(x - bLi);
                    let lY = safeNum(y);
                    let lW = safeNum(w + bLi, 0.01);
                    let lH = safeNum(h, 0.01);
                    let lR = stdRadius; // Matches main box

                    this.slide.addShape(shapeName, {
                        x: lX, y: lY, w: lW, h: lH,
                        fill: { color: c.color, transparency: c.transparency },
                        rectRadius: lR
                    });
                }
                if (bRight > 0 && style.borderRightStyle !== 'none') {
                    const c = parseColor(style.borderRightColor);
                    let bRi = safeNum(pxToInch(parseFloat(style.borderRightWidth) || 0));
                    let lX = safeNum(x);
                    let lY = safeNum(y);
                    let lW = safeNum(w + bRi, 0.01);
                    let lH = safeNum(h, 0.01);
                    let lR = stdRadius;

                    this.slide.addShape(shapeName, {
                        x: lX, y: lY, w: lW, h: lH,
                        fill: { color: c.color, transparency: c.transparency },
                        rectRadius: lR
                    });
                }
            } else if (hasBorder && !hasFill && !isRoot) {
                // TRANSPARENT background: Use lines instead of shapes
                // This prevents "exposing" a white fill where none was intended
                if (bTop > 0 && style.borderTopStyle !== 'none') {
                    const c = parseColor(style.borderTopColor);
                    const lineY = safeNum(y - pxToInch(bTop) / 2);
                    this.slide.addShape('line', {
                        x: x, y: lineY,
                        w: w, h: 0,
                        line: { color: c.color, width: bTop * 0.75, transparency: c.transparency }
                    });
                }
                if (bBot > 0 && style.borderBottomStyle !== 'none') {
                    const c = parseColor(style.borderBottomColor);
                    const lineY = safeNum(y + h + pxToInch(bBot) / 2);
                    this.slide.addShape('line', {
                        x: x, y: lineY,
                        w: w, h: 0,
                        line: { color: c.color, width: bBot * 0.75, transparency: c.transparency }
                    });
                }
                if (bLeft > 0 && style.borderLeftStyle !== 'none') {
                    const c = parseColor(style.borderLeftColor);
                    const lineX = safeNum(x - pxToInch(bLeft) / 2);
                    this.slide.addShape('line', {
                        x: lineX, y: y,
                        w: 0, h: h,
                        line: { color: c.color, width: bLeft * 0.75, transparency: c.transparency }
                    });
                }
                if (bRight > 0 && style.borderRightStyle !== 'none') {
                    const c = parseColor(style.borderRightColor);
                    const lineX = safeNum(x + w + pxToInch(bRight) / 2);
                    this.slide.addShape('line', {
                        x: lineX, y: y,
                        w: 0, h: h,
                        line: { color: c.color, width: bRight * 0.75, transparency: c.transparency }
                    });
                }
            }

            // MAIN LAYER - No longer force white fill for transparent+border cases

            if (isRoot) {
                if (hasBorder && !isUniform) {
                    let rootFill = bgParsed;
                    if (rootFill.transparency === 100) rootFill = { color: 'FFFFFF', transparency: 0 };
                    this.slide.addShape(shapeName, {
                        x, y, w, h,
                        fill: { color: rootFill.color, transparency: rootFill.transparency },
                        rectRadius: stdRadius
                    });
                }
            } else {
                let mainOpts = { x, y, w, h };
                if (hasFill) mainOpts.fill = { color: bgParsed.color, transparency: bgParsed.transparency };
                if (isUniform && hasBorder) {
                    const c = parseColor(style.borderTopColor);
                    if (c) {
                        mainOpts.line = { color: c.color, width: parseFloat(style.borderTopWidth) * 0.75 };
                        // Add Dash Support
                        const dashMap = {
                            'dashed': 'dash',
                            'dotted': 'dot',
                        };
                        const dashType = dashMap[style.borderTopStyle];
                        if (dashType) mainOpts.line.dashType = dashType;
                    }
                }
                if (style.boxShadow && style.boxShadow !== 'none') {
                    mainOpts.shadow = { type: 'outer', angle: 45, blur: 6, offset: 2, opacity: 0.2 };
                }
                if (isRounded && stdRadius !== undefined) mainOpts.rectRadius = stdRadius;

                if (hasFill || (isUniform && hasBorder) || mainOpts.shadow) {
                    this.slide.addShape(shapeName, mainOpts);
                }
            }
            if (style.backgroundImage && style.backgroundImage.includes('gradient')) {
                const dataUri = gradientToImage(style.backgroundImage, rect.width, rect.height);
                if (dataUri) this.slide.addImage({ data: dataUri, x, y, w, h });
            }
        }

        renderText(node, style, rect) {
            const runs = this.collectTextRuns(node, style);
            if (runs.length === 0) return;

            const pos = this.getPos(rect);
            const pt = safeNum(pxToInch(parseFloat(style.paddingTop) || 0));
            const pb = safeNum(pxToInch(parseFloat(style.paddingBottom) || 0));
            const pl = safeNum(pxToInch(parseFloat(style.paddingLeft) || 0));
            const pr = safeNum(pxToInch(parseFloat(style.paddingRight) || 0));

            const bt = safeNum(pxToInch(parseFloat(style.borderTopWidth) || 0));
            const bb = safeNum(pxToInch(parseFloat(style.borderBottomWidth) || 0));
            const bl = safeNum(pxToInch(parseFloat(style.borderLeftWidth) || 0));
            const br = safeNum(pxToInch(parseFloat(style.borderRightWidth) || 0));

            let align = style.textAlign || 'left';
            if (align === 'start') align = 'left';
            if (align === 'end') align = 'right';

            // Flexbox Alignment Mapping
            if (style.display.includes('flex')) {
                const justify = style.justifyContent || 'flex-start';
                const alignItems = style.alignItems || 'stretch';
                const direct = style.flexDirection || 'row';

                if (direct.includes('row') && justify === 'center') align = 'center';
                if (direct.includes('column') && alignItems === 'center') align = 'center';
                if (direct.includes('row') && (justify === 'flex-end' || justify === 'right')) align = 'right';
                if (direct.includes('column') && (alignItems === 'flex-end' || alignItems === 'right')) align = 'right';
            }

            let valign = 'top';
            if (style.display.includes('flex') && style.alignItems === 'center') valign = 'middle';

            let tX = safeNum(pos.x + bl + pl);
            let tY = safeNum(pos.y + bt + pt);
            let tW = safeNum(pos.w - bl - br - pl - pr, 0.1);
            let tH = safeNum(pos.h - bt - bb - pt - pb, 0.1);

            // Get fontSize for lineSpacing calculation
            const fontSize = parseFloat(style.fontSize);

            // 2. Text Width Buffer Strategy (Only for BOLD)
            // Check if any run is bold or if parent is bold
            let isBold = (parseInt(style.fontWeight) || 400) >= 600;
            if (!isBold) {
                // Check runs
                for (let r of runs) {
                    if (r.options && r.options.bold) {
                        isBold = true;
                        break;
                    }
                }
            }

            // Exclude Headings from Buffer and Spacing
            // Bare text in div, p, span, etc. are NOT headings, so they get adjustments
            const isHeading = /^H[1-6]$/.test(node.tagName);

            let widthBuffer = 0;
            if (isBold && !isHeading) {
                // Add 7% buffer for bold text
                widthBuffer = tW * 0.07;
                tW += widthBuffer;

                // 3. Alignment-based X-Offset
                if (align === 'center') {
                    // Center: Move left by half the buffer to keep visual center
                    tX -= widthBuffer / 2;
                } else if (align === 'right') {
                    // Right: Move left by full buffer so right edge stays pinned
                    tX -= widthBuffer;
                }
                // Left: No offset (expands to the right)
            }

            if (tW > 0 && tH > 0) {
                // PptxGenJS lineSpacing is in Points (space after each line, not top+bottom)
                const defaultLineSpacing = safeNum(fontSize * 1.15);

                let textOpts = {
                    x: tX, y: tY, w: tW, h: tH,
                    align: align, valign: valign, wrap: true,
                    inset: 0
                };

                // Only apply default line spacing to non-headings (p, span, div text, etc.)
                if (!isHeading) {
                    textOpts.lineSpacing = defaultLineSpacing;
                }

                this.slide.addText(runs, textOpts);
            }
        }

        collectTextRuns(node, parentStyle) {
            let runs = [];
            node.childNodes.forEach(child => {
                if (child.nodeType === Node.TEXT_NODE) {
                    const txt = child.textContent;
                    if (!txt || !txt.trim()) return;
                    let clean = txt.replace(/\s+/g, ' ');
                    if (clean === ' ' && runs.length > 0 && runs[runs.length - 1].text.endsWith(' ')) return;

                    const s = (node.nodeType === Node.ELEMENT_NODE) ? window.getComputedStyle(node) : parentStyle;
                    const c = parseColor(s.color, parseFloat(s.opacity) || 1);

                    let run = {
                        text: clean, options: {
                            color: c ? c.color : '000000',
                            fontSize: safeNum(parseFloat(s.fontSize) * 0.75, 10), // Removed 0.99 scaling
                            bold: (parseInt(s.fontWeight) || 400) >= 600,
                            italic: (s.fontStyle === 'italic' || s.fontStyle === 'oblique'),
                            fontFace: getSafeFont(s.fontFamily)
                        }
                    };
                    if (c && c.transparency > 0) run.options.transparency = c.transparency;
                    if (s.textDecoration && s.textDecoration.includes('underline')) run.options.underline = { style: 'sng' };

                    runs.push(run);
                } else if (child.nodeType === Node.ELEMENT_NODE) {
                    if (child.tagName === 'BR') runs.push({ text: '', options: { breakLine: true } });
                    else runs.push(...this.collectTextRuns(child, window.getComputedStyle(child)));
                }
            });
            return runs;
        }

        isTextBlock(node) {
            if (node.tagName === 'TABLE') return false;
            let hasText = false;
            node.childNodes.forEach(c => {
                if (c.nodeType === Node.TEXT_NODE && c.textContent.trim().length > 0) hasText = true;
            });
            return hasText;
        }
    }

    async function addSlide(pres, slide, container, options) {
        const exporter = new SlideExporter(pres, slide, container, options);
        await exporter.process();
    }

    // UPDATED EXPORT FUNCTION WITH FONT SIMULATION
    async function exportToPPTX(elementOrId, options = {}) {
        let PptxClass = window.PptxGenJS;
        if (!PptxClass && typeof PptxGenJS !== 'undefined') PptxClass = PptxGenJS;

        if (!PptxClass) return console.error('PptxGenJS not found. Ensure script is loaded.');

        const fileName = options.fileName || 'presentation.pptx';
        const pres = new PptxClass();
        pres.layout = 'LAYOUT_16x9';

        let targets = [];
        if (typeof elementOrId === 'string') {
            const el = document.getElementById(elementOrId);
            if (el) targets = [el];
            else targets = Array.from(document.querySelectorAll(elementOrId));
        } else if (elementOrId instanceof HTMLElement) targets = [elementOrId];

        // START FONT SIMULATION
        let allSimulatedElements = [];
        try {
            targets.forEach(t => {
                const sims = simulatePPTXFonts(t);
                allSimulatedElements.push(...sims);
            });

            // Force reflow/repaint to ensure metrics update?
            // Reading offsetHeight acts as a reflow trigger.
            if (targets.length > 0) { const _ = targets[0].offsetHeight; }

            for (const t of targets) {
                const s = pres.addSlide();
                const bg = window.getComputedStyle(t).backgroundColor;
                const bgP = parseColor(bg);
                if (bgP && bgP.transparency < 100) {
                    s.background = { color: bgP.color, transparency: bgP.transparency };
                }
                await addSlide(pres, s, t, options);
            }
            await pres.writeFile({ fileName });

        } catch (err) {
            console.error("Export Failed:", err);
        } finally {
            // RESTORE FONTS
            targets.forEach(t => restoreFonts(t));
        }
    }

    root.LLMDomToPptx = { export: exportToPPTX, addSlide: addSlide };

})(window);
