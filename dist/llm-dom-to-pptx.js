/**
 * LLM DOM to PPTX - v1.0.3
 * Converts Semantic HTML/CSS (e.g. from LLMs) into editable PPTX.
 * 
 * Dependencies:
 * - PptxGenJS (https://gitbrent.github.io/PptxGenJS/)
 */

(function (root) {
    'use strict';

    // --- 1. Font Mapping Utilities ---

    const FONT_MAP = {
        // Sans-Serif
        "Inter": "Arial",
        "Roboto": "Arial",
        "Open Sans": "Calibri",
        "Lato": "Calibri",
        "Montserrat": "Arial",
        "Source Sans Pro": "Arial",
        "Noto Sans": "Arial",
        "Helvetica": "Arial",
        "San Francisco": "Arial",
        "Segoe UI": "Segoe UI",
        "System-UI": "Segoe UI",

        // Serif
        "Times New Roman": "Times New Roman",
        "Georgia": "Georgia",
        "Merriweather": "Times New Roman",
        "Playfair Display": "Georgia",

        // Monospace
        "Courier New": "Courier New",
        "Fira Code": "Courier New",
        "Roboto Mono": "Courier New",

        // Fallbacks
        "sans-serif": "Arial",
        "serif": "Times New Roman",
        "monospace": "Courier New"
    };

    /**
     * Returns a PPTX-safe font name based on the input web font family.
     */
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

    // --- 2. Color & Unit Utilities ---

    /**
     * Parses a color string into Hex and Transparency (%).
     * Handles #Hex, rgb(), and rgba().
     * @returns {Object|null} { color: "RRGGBB", transparency: 0-100 } or null
     */
    function parseColor(colorStr, opacity = 1) {
        if (!colorStr || colorStr === 'rgba(0, 0, 0, 0)' || colorStr === 'transparent') return null;

        let r, g, b, a = 1;

        if (colorStr.startsWith('#')) {
            const hex = colorStr.substring(1);
            if (hex.length === 3) {
                r = parseInt(hex[0] + hex[0], 16);
                g = parseInt(hex[1] + hex[1], 16);
                b = parseInt(hex[2] + hex[2], 16);
            } else {
                r = parseInt(hex.substring(0, 2), 16);
                g = parseInt(hex.substring(2, 4), 16);
                b = parseInt(hex.substring(4, 6), 16);
            }
        } else if (colorStr.startsWith('rgb')) {
            const values = colorStr.match(/(\d+(\.\d+)?)/g);
            if (!values || values.length < 3) return null;
            r = parseFloat(values[0]);
            g = parseFloat(values[1]);
            b = parseFloat(values[2]);
            if (values.length > 3) {
                a = parseFloat(values[3]);
            }
        } else {
            return null;
        }

        // Combine CSS opacity with Color Alpha
        const finalAlpha = a * opacity;

        // Convert to Hex string "RRGGBB"
        const toHex = (n) => {
            const h = Math.round(n).toString(16);
            return h.length === 1 ? '0' + h : h;
        };
        const hex = toHex(r) + toHex(g) + toHex(b);

        // Calculate Transparency Percent (0 = Opaque, 100 = Transparent)
        // PPTX gen uses percent transparency
        const transparency = Math.round((1 - finalAlpha) * 100);

        return {
            color: hex,
            transparency: transparency
        };
    }

    // Deprecated but kept for compatibility references if any
    function rgbToHex(rgb) {
        const p = parseColor(rgb);
        return p ? '#' + p.color : null;
    }

    const PPI = 96;
    const SLIDE_WIDTH_PX = 960; // Default reference width
    const PPT_WIDTH_IN = 10;

    // We calculate generic scale based on 960px. 
    // If the user's container is different, we might want to adjust, 
    // but usually 960px is a good "logical" width for a slide.
    const SCALE = PPT_WIDTH_IN / SLIDE_WIDTH_PX;

    function pxToInch(px) {
        return parseFloat(px) * SCALE;
    }

    // --- 3. Main Export Function ---

    /**
     * Exports a DOM element to PPTX.
     * @param {string|HTMLElement} elementOrId - The DOM element or ID to export.
     * @param {object} options - Options { fileName: "presentation.pptx" }
     */
    async function exportToPPTX(elementOrId = 'slide-canvas', options = {}) {
        const fileName = options.fileName || "presentation.pptx";

        console.log(`Starting PPTX Export for ${elementOrId}...`);

        if (typeof PptxGenJS === 'undefined') {
            console.error("PptxGenJS is not loaded. Please include it via CDN.");
            alert("Error: PptxGenJS library missing.");
            return;
        }

        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';

        const PPT_HEIGHT_IN = 5.625; // 16:9 aspect ratio of 10 inch width

        let slide = pres.addSlide();
        let currentSlideYOffset = 0; // Tracks Y offset for multi-page support

        // Resolve Container
        let container;
        if (typeof elementOrId === 'string') {
            container = document.getElementById(elementOrId);
        } else {
            container = elementOrId;
        }

        if (!container) {
            console.error(`Slide container '${elementOrId}' not found!`);
            return;
        }

        const containerRect = container.getBoundingClientRect();

        // --- 0. Slide Background ---
        const containerStyle = window.getComputedStyle(container);
        const bgParsed = parseColor(containerStyle.backgroundColor);
        if (bgParsed) {
            slide.background = { color: bgParsed.color, transparency: bgParsed.transparency };
        }

        // Visited nodes set to prevent ghosting/duplication
        const processedNodes = new Set();

        // --- Helper: recurse gather text runs ---
        function collectTextRuns(node, parentStyle) {
            let runs = [];
            if (!node) return runs;

            // Skip non-visible if element
            if (node.nodeType === Node.ELEMENT_NODE) {
                const s = window.getComputedStyle(node);
                if (s.display === 'none' || s.visibility === 'hidden' || parseFloat(s.opacity) === 0) return runs;
            }

            node.childNodes.forEach(child => {
                if (child.nodeType === Node.TEXT_NODE) {
                    const text = child.textContent;
                    if (!text) return;

                    // Inherit style from parent element if current is text node
                    const style = (node.nodeType === Node.ELEMENT_NODE) ? window.getComputedStyle(node) : parentStyle;

                    const colorParsed = parseColor(style.color, parseFloat(style.opacity) || 1);
                    const fontSize = parseFloat(style.fontSize);
                    const fontWeight = (style.fontWeight === '700' || style.fontWeight === 'bold' || parseInt(style.fontWeight) >= 600);

                    // Normalize Whitespace:
                    let runText = text.replace(/\s+/g, ' ');

                    if (style.textTransform === 'uppercase') runText = runText.toUpperCase();

                    if (!runText) return;

                    const runOpts = {
                        color: colorParsed ? colorParsed.color : '000000',
                        fontSize: fontSize * 0.75, // px to pt
                        bold: fontWeight,
                        fontFace: getSafeFont(style.fontFamily),
                        breakLine: false
                    };

                    // Letter Spacing (Tracking) Support
                    // Converts CSS letter-spacing (px/em) to PPTX charSpacing (points)
                    const letterSpacingStr = style.letterSpacing;
                    if (letterSpacingStr && letterSpacingStr !== 'normal') {
                        let spacingPx = 0;
                        if (letterSpacingStr.includes('em')) {
                            // e.g. "0.2em" -> 0.2 * fontSize
                            spacingPx = parseFloat(letterSpacingStr) * (parseFloat(style.fontSize) || 16);
                        } else if (letterSpacingStr.includes('px')) {
                            spacingPx = parseFloat(letterSpacingStr);
                        } else if (!isNaN(parseFloat(letterSpacingStr))) {
                            // Fallback if just number (unlikely in CSS computed style but safe)
                            spacingPx = parseFloat(letterSpacingStr);
                        }

                        // Convert px to pt (1px = 0.75pt)
                        if (spacingPx !== 0) {
                            runOpts.charSpacing = spacingPx * 0.75;
                        }
                    }

                    if (colorParsed && colorParsed.transparency > 0) {
                        runOpts.transparency = colorParsed.transparency;
                    }

                    // NOTE: We intentionally do NOT apply 'highlight' here because:
                    // 1. If this text is in a parent that has a background, that background
                    //    is already drawn as a shape by processNode.
                    // 2. Adding highlight would create duplicate/overlapping backgrounds.
                    // 3. PPTX highlight is meant for inline text highlighting, not block backgrounds.

                    // Border-bottom as text underline support (only for inline/paragraph elements)
                    // Common pattern: <span class="border-b-2 border-b-indigo-500">text</span>
                    // Exclude headings (h1-h6) as they use border-b as section separators
                    const parentTag = node.tagName ? node.tagName.toUpperCase() : '';
                    const isInlineOrParagraph = ['SPAN', 'P', 'A', 'LABEL', 'STRONG', 'EM', 'B', 'I'].includes(parentTag);

                    const borderBottomWidth = parseFloat(style.borderBottomWidth) || 0;
                    const borderBottomStyle = style.borderBottomStyle;
                    if (isInlineOrParagraph && borderBottomWidth > 0 && borderBottomStyle !== 'none') {
                        runOpts.underline = { style: 'sng' }; // Single underline
                        // Try to get underline color from border color
                        const borderColorParsed = parseColor(style.borderBottomColor);
                        if (borderColorParsed) {
                            runOpts.underline.color = borderColorParsed.color;
                        }
                    }

                    // CSS text-decoration: underline support
                    if (style.textDecoration && style.textDecoration.includes('underline')) {
                        runOpts.underline = { style: 'sng' };
                    }

                    runs.push({
                        text: runText,
                        options: runOpts
                    });
                } else if (child.nodeType === Node.ELEMENT_NODE) {
                    if (child.tagName === 'BR') {
                        runs.push({ text: '', options: { breakLine: true } });
                    } else {
                        runs.push(...collectTextRuns(child, window.getComputedStyle(child)));
                    }
                }
            });
            return runs;
        }

        // --- Helper: Identify logical Text Block ---
        function isTextBlock(node) {
            if (node.nodeType !== Node.ELEMENT_NODE) return false;
            if (node.childNodes.length === 0) return false;

            let hasText = false;
            node.childNodes.forEach(c => {
                if (c.nodeType === Node.TEXT_NODE && c.textContent.trim().length > 0) hasText = true;
            });
            return hasText;
        }

        function processNode(node) {
            if (node.nodeType !== Node.ELEMENT_NODE) return;
            if (processedNodes.has(node)) return;
            processedNodes.add(node);

            const style = window.getComputedStyle(node);
            const rect = node.getBoundingClientRect();
            const opacity = parseFloat(style.opacity) || 1;

            // Skip invisible
            if (style.display === 'none' || style.visibility === 'hidden' || opacity === 0) return;

            // Skip zero size (unless overflow)
            if (rect.width < 1 || rect.height < 1) return;

            // Relative Coordinates
            const x = pxToInch(rect.left - containerRect.left);
            const y = pxToInch(rect.top - containerRect.top) - currentSlideYOffset;
            let w = pxToInch(rect.width);
            let h = pxToInch(rect.height);

            // Enforce minimum dimensions for very thin elements (like h-px lines)
            const MIN_DIM = 0.02; // ~2px in PowerPoint
            if (h < MIN_DIM && h > 0) h = MIN_DIM;
            if (w < MIN_DIM && w > 0) w = MIN_DIM;

            // --- TABLE HANDLING ---
            if (node.tagName === 'TABLE') {
                // Shadow Handler
                if (style.boxShadow && style.boxShadow !== 'none') {
                    let tableBg = style.backgroundColor;
                    let tableOp = parseFloat(style.opacity) || 1;
                    let shadowFill = parseColor(tableBg, tableOp);

                    if (!shadowFill || shadowFill.transparency === 100) {
                        shadowFill = { color: 'FFFFFF', transparency: 99 };
                    }

                    slide.addShape(pres.ShapeType.rect, {
                        x: x, y: y, w: w, h: h,
                        fill: { color: shadowFill.color, transparency: shadowFill.transparency },
                        shadow: { type: 'outer', angle: 45, blur: 10, offset: 4, opacity: 0.3 },
                        rectRadius: 0
                    });
                }

                const tableRows = [];
                let colWidths = [];
                let rowHeights = [];

                if (node.rows.length > 0) {
                    colWidths = Array.from(node.rows[0].cells).map(c => pxToInch(c.getBoundingClientRect().width));
                    rowHeights = Array.from(node.rows).map(r => pxToInch(r.getBoundingClientRect().height));
                }

                Array.from(node.rows).forEach(row => {
                    const rowData = [];
                    Array.from(row.cells).forEach(cell => {
                        const cStyle = window.getComputedStyle(cell);
                        const cRuns = collectTextRuns(cell, cStyle);

                        let effectiveBg = cStyle.backgroundColor;
                        let effectiveOpacity = parseFloat(cStyle.opacity) || 1;

                        if ((!effectiveBg || effectiveBg === 'rgba(0, 0, 0, 0)' || effectiveBg === 'transparent') && row) {
                            const rStyle = window.getComputedStyle(row);
                            effectiveBg = rStyle.backgroundColor;
                            effectiveOpacity = parseFloat(rStyle.opacity) || 1;

                            if ((!effectiveBg || effectiveBg === 'rgba(0, 0, 0, 0)' || effectiveBg === 'transparent') && row.parentElement) {
                                const pStyle = window.getComputedStyle(row.parentElement);
                                effectiveBg = pStyle.backgroundColor;
                                effectiveOpacity = parseFloat(pStyle.opacity) || 1;
                            }
                        }

                        const bgP = parseColor(effectiveBg, effectiveOpacity);

                        let vAlign = 'top';
                        if (cStyle.verticalAlign === 'middle') vAlign = 'middle';
                        if (cStyle.verticalAlign === 'bottom') vAlign = 'bottom';

                        const pt = pxToInch(parseFloat(cStyle.paddingTop) || 0);
                        const pr = pxToInch(parseFloat(cStyle.paddingRight) || 0);
                        const pb = pxToInch(parseFloat(cStyle.paddingBottom) || 0);
                        const pl = pxToInch(parseFloat(cStyle.paddingLeft) || 0);
                        const margin = [pt, pr, pb, pl];

                        const getBdr = (w, c, s, fallbackW, fallbackC, fallbackS) => {
                            if (w && parseFloat(w) > 0 && s !== 'none') {
                                const co = parseColor(c) || { color: '000000' };
                                return { pt: parseFloat(w) * 0.75, color: co.color };
                            }
                            if (fallbackW && parseFloat(fallbackW) > 0 && fallbackS !== 'none') {
                                const co = parseColor(fallbackC) || { color: '000000' };
                                return { pt: parseFloat(fallbackW) * 0.75, color: co.color };
                            }
                            return null;
                        };

                        const rStyle = row ? window.getComputedStyle(row) : null;
                        let rbTopW = 0, rbTopC = null, rbTopS = 'none';
                        let rbBotW = 0, rbBotC = null, rbBotS = 'none';

                        if (rStyle) {
                            rbTopW = rStyle.borderTopWidth; rbTopC = rStyle.borderTopColor; rbTopS = rStyle.borderTopStyle;
                            rbBotW = rStyle.borderBottomWidth; rbBotC = rStyle.borderBottomColor; rbBotS = rStyle.borderBottomStyle;
                        }

                        const bTop = getBdr(cStyle.borderTopWidth, cStyle.borderTopColor, cStyle.borderTopStyle, rbTopW, rbTopC, rbTopS);
                        const bRight = getBdr(cStyle.borderRightWidth, cStyle.borderRightColor, cStyle.borderRightStyle);
                        const bBot = getBdr(cStyle.borderBottomWidth, cStyle.borderBottomColor, cStyle.borderBottomStyle, rbBotW, rbBotC, rbBotS);
                        const bLeft = getBdr(cStyle.borderLeftWidth, cStyle.borderLeftColor, cStyle.borderLeftStyle);

                        const cellOpts = {
                            valign: vAlign,
                            align: cStyle.textAlign === 'center' ? 'center' : (cStyle.textAlign === 'right' ? 'right' : 'left'),
                            margin: margin,
                            border: [bTop, bRight, bBot, bLeft]
                        };

                        if (bgP && bgP.transparency < 100) {
                            cellOpts.fill = { color: bgP.color };
                            if (bgP.transparency > 0) cellOpts.fill.transparency = bgP.transparency;
                        }

                        rowData.push({
                            text: cRuns,
                            options: cellOpts
                        });

                        const markChildren = (n) => {
                            processedNodes.add(n);
                            Array.from(n.children).forEach(markChildren);
                        }
                        markChildren(cell);
                        processedNodes.add(cell);
                    });
                    tableRows.push(rowData);
                    processedNodes.add(row);
                });

                if (tableRows.length > 0) {
                    const availableH = PPT_HEIGHT_IN - y - 0.5;
                    if (h <= availableH) {
                        slide.addTable(tableRows, {
                            x: x, y: y, w: w, colW: colWidths, rowH: rowHeights, autoPage: false
                        });
                    } else {
                        console.warn('Table does not fit on current slide, splitting manually (disabled autoPage)...');
                        if (y > 1.0) {
                            slide = pres.addSlide();
                            currentSlideYOffset += PPT_HEIGHT_IN;
                            slide.addTable(tableRows, {
                                x: x, y: y, w: w, colW: colWidths, rowH: rowHeights, autoPage: false
                            });
                        } else {
                            slide.addTable(tableRows, {
                                x: x, y: y, w: w, colW: colWidths, rowH: rowHeights, autoPage: false
                            });
                        }
                    }
                }
                processedNodes.add(node);
                return;
            }

            // --- A. BACKGROUNDS & BORDERS ---
            const bgParsed = parseColor(style.backgroundColor, opacity);
            const borderW = parseFloat(style.borderWidth) || 0;
            const borderParsed = parseColor(style.borderColor);

            let hasFill = bgParsed && bgParsed.transparency < 100;
            let hasBorder = borderW > 0 && borderParsed;

            let shapeOpts = { x, y, w, h };

            // -- NEW: Precise Border Radius Logic --
            const rtl = parseFloat(style.borderTopLeftRadius) || 0;
            const rtr = parseFloat(style.borderTopRightRadius) || 0;
            const rbr = parseFloat(style.borderBottomRightRadius) || 0;
            const rbl = parseFloat(style.borderBottomLeftRadius) || 0;

            const borderRadius = parseFloat(style.borderRadius) || 0;

            // Strict circle check
            const isCircle = (Math.abs(rect.width - rect.height) < 2) && (borderRadius >= rect.width / 2 - 1);

            // Shadow Support
            if (style.boxShadow && style.boxShadow !== 'none') {
                shapeOpts.shadow = { type: 'outer', angle: 45, blur: 6, offset: 2, opacity: 0.2 };
            }

            let shapeType = pres.ShapeType.rect;
            let rotation = 0;

            if (isCircle) {
                shapeType = pres.ShapeType.ellipse;
            } else {
                // Determine Shape based on corners
                // Case 1: All Uniform
                if (rtl === rtr && rtr === rbr && rbr === rbl && rtl > 0) {
                    const minDim = Math.min(rect.width, rect.height);
                    let ratio = rtl / (minDim / 2);
                    shapeOpts.rectRadius = Math.min(ratio, 1.0);
                    shapeType = pres.ShapeType.roundRect;
                }
                // Case 2: Vertical Rounding (Left or Right) -> PptxGenJS doesn't inherently support 'Right Rounded' well without rotation?
                // Actually PptxGenJS has 'round2SameRect' which is Top Corners Rounded by default.

                // Case 3: Top Rounded (Common in cards)
                else if (rtl > 0 && rtr > 0 && rbr === 0 && rbl === 0) {
                    shapeType = pres.ShapeType.round2SameRect;
                    // Rotate? Default is Top.
                    rotation = 0;
                }
                // Case 4: Bottom Rounded
                else if (rtl === 0 && rtr === 0 && rbr > 0 && rbl > 0) {
                    shapeType = pres.ShapeType.round2SameRect;
                    rotation = 180;
                }
                // Case 5: Single corners or diagonal? Fallback to rect (Square) to avoid "All Rounded" bug.
                else if (borderRadius > 0) {
                    // It has some radius, but not uniform and not simple top/bottom pair.
                    // Fallback: If we use 'roundRect' it rounds all. 
                    // Better to use 'rect' (Sharp) than incorrect 'roundRect' for things like "only top-left".
                    shapeType = pres.ShapeType.rect;
                }
            }

            if (rotation !== 0) {
                shapeOpts.rotate = rotation;
            }

            // --- Border Logic ---
            // Check all 4 sides for true uniformity
            const bTop = parseFloat(style.borderTopWidth) || 0;
            const bRight = parseFloat(style.borderRightWidth) || 0;
            const bBot = parseFloat(style.borderBottomWidth) || 0;
            const bLeft = parseFloat(style.borderLeftWidth) || 0;

            const isUniformBorder = (bTop === bRight && bRight === bBot && bBot === bLeft && bTop > 0);

            if (isUniformBorder && borderParsed) {
                const lineOpts = { color: borderParsed.color, width: bTop * 0.75 };
                if (borderParsed.transparency > 0) {
                    lineOpts.transparency = borderParsed.transparency;
                }
                shapeOpts.line = lineOpts;
            } else {
                shapeOpts.line = null;
            }

            // --- B. LEFT ACCENT BORDER (Custom Strategy) ---
            const lW = parseFloat(style.borderLeftWidth) || 0;
            const leftBorderParsed = parseColor(style.borderLeftColor);
            const hasLeftBorder = lW > 0 && leftBorderParsed && style.borderStyle !== 'none';

            if (hasLeftBorder && !isUniformBorder) {
                if (hasFill) {
                    const underlayOpts = { ...shapeOpts };
                    underlayOpts.fill = { color: leftBorderParsed.color };
                    underlayOpts.line = null;

                    // If rotated, the underlay needs careful handling. 
                    // Simpler: Just draw a side strip if rotation is involved, or complex underlay.
                    // For now, keep original logic but verify rotation impact.
                    slide.addShape(shapeType, underlayOpts);

                    const borderInch = pxToInch(lW);
                    shapeOpts.x += borderInch;
                    shapeOpts.w -= borderInch;
                    delete shapeOpts.shadow;
                } else {
                    slide.addShape(pres.ShapeType.rect, {
                        x: x, y: y, w: pxToInch(lW), h: h,
                        fill: { color: leftBorderParsed.color },
                        rectRadius: isCircle ? 0 : (shapeOpts.rectRadius || 0)
                    });
                }
            }

            // --- B2. RIGHT ACCENT BORDER (Custom Strategy) ---
            const rW = parseFloat(style.borderRightWidth) || 0;
            const rightBorderParsed = parseColor(style.borderRightColor);
            const hasRightBorder = rW > 0 && rightBorderParsed && style.borderRightStyle !== 'none';

            if (hasRightBorder && !isUniformBorder) {
                if (hasFill) {
                    const underlayOpts = { ...shapeOpts };
                    underlayOpts.fill = { color: rightBorderParsed.color };
                    if (rightBorderParsed.transparency > 0) {
                        underlayOpts.fill.transparency = rightBorderParsed.transparency;
                    }
                    underlayOpts.line = null;
                    slide.addShape(shapeType, underlayOpts);

                    // Shrink main shape from right to reveal right border
                    const borderInch = pxToInch(rW);
                    shapeOpts.w -= borderInch;
                    delete shapeOpts.shadow;
                } else {
                    // No fill: Draw simple strip at right edge
                    const stripOpts = {
                        x: x + w - pxToInch(rW), y: y, w: pxToInch(rW), h: h,
                        fill: { color: rightBorderParsed.color }
                    };
                    if (rightBorderParsed.transparency > 0) {
                        stripOpts.fill.transparency = rightBorderParsed.transparency;
                    }
                    slide.addShape(pres.ShapeType.rect, stripOpts);
                }
            }

            // --- C. TOP ACCENT BORDER (Underlay Strategy - BEFORE main shape) ---
            const tW = parseFloat(style.borderTopWidth) || 0;
            const topBorderParsed = parseColor(style.borderTopColor);
            const hasTopBorder = tW > 0 && topBorderParsed && style.borderTopStyle !== 'none';

            if (hasTopBorder && !isUniformBorder) {
                if (hasFill) {
                    // Draw full shape in border color as underlay
                    const underlayOpts = { ...shapeOpts };
                    underlayOpts.fill = { color: topBorderParsed.color };
                    if (topBorderParsed.transparency > 0) {
                        underlayOpts.fill.transparency = topBorderParsed.transparency;
                    }
                    underlayOpts.line = null;
                    slide.addShape(shapeType, underlayOpts);

                    // Offset main shape to reveal top border
                    const borderInch = pxToInch(tW);
                    shapeOpts.y += borderInch;
                    shapeOpts.h -= borderInch;
                    delete shapeOpts.shadow;
                } else {
                    // No fill: Draw simple strip
                    const stripOpts = {
                        x: x, y: y, w: w, h: pxToInch(tW),
                        fill: { color: topBorderParsed.color }
                    };
                    if (topBorderParsed.transparency > 0) {
                        stripOpts.fill.transparency = topBorderParsed.transparency;
                    }
                    slide.addShape(pres.ShapeType.rect, stripOpts);
                }
            }

            // --- D. BOTTOM ACCENT BORDER (Underlay Strategy - BEFORE main shape) ---
            const bW = parseFloat(style.borderBottomWidth) || 0;
            const bottomBorderParsed = parseColor(style.borderBottomColor);
            const hasBottomBorder = bW > 0 && bottomBorderParsed && style.borderBottomStyle !== 'none';

            if (hasBottomBorder && !isUniformBorder) {
                if (hasFill && !hasTopBorder) {
                    // Only do underlay if we didn't already do it for top border
                    const underlayOpts = { ...shapeOpts };
                    underlayOpts.fill = { color: bottomBorderParsed.color };
                    if (bottomBorderParsed.transparency > 0) {
                        underlayOpts.fill.transparency = bottomBorderParsed.transparency;
                    }
                    underlayOpts.line = null;
                    slide.addShape(shapeType, underlayOpts);

                    // Shrink main shape from bottom to reveal bottom border
                    const borderInch = pxToInch(bW);
                    shapeOpts.h -= borderInch;
                    delete shapeOpts.shadow;
                } else if (hasFill && hasTopBorder) {
                    // Both top and bottom: already have underlay, just shrink from bottom too
                    const borderInch = pxToInch(bW);
                    shapeOpts.h -= borderInch;
                } else {
                    // No fill: Draw simple strip at bottom
                    const bH = pxToInch(bW);
                    const stripOpts = {
                        x: x, y: y + h - bH, w: w, h: bH,
                        fill: { color: bottomBorderParsed.color }
                    };
                    if (bottomBorderParsed.transparency > 0) {
                        stripOpts.fill.transparency = bottomBorderParsed.transparency;
                    }
                    slide.addShape(pres.ShapeType.rect, stripOpts);
                }
            }

            // Draw Main Shape
            if (hasFill) {
                shapeOpts.fill = { color: bgParsed.color };
                if (bgParsed.transparency > 0) {
                    shapeOpts.fill.transparency = bgParsed.transparency;
                }
                slide.addShape(shapeType, shapeOpts);
            } else if (hasBorder && shapeOpts.line) {
                slide.addShape(shapeType, shapeOpts);
            }

            // --- Gradient Fallback ---
            if (style.backgroundImage && style.backgroundImage.includes('gradient')) {
                if (rect.height < 15 && rect.width > 100) {
                    slide.addShape(pres.ShapeType.rect, {
                        x: x, y: y, w: w, h: h,
                        fill: { color: '4F46E5' }
                    });
                }
            }

            // --- C. TEXT CONTENT ---
            if (isTextBlock(node)) {
                const runs = collectTextRuns(node, style);
                if (runs.length > 0) {
                    if (runs.length > 0) {
                        runs[0].text = runs[0].text.replace(/^\s+/, '');
                        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
                    }
                    const validRuns = runs.filter(r => r.text !== '' || r.options.breakLine);

                    if (validRuns.length > 0) {
                        let align = 'left';
                        if (style.textAlign === 'center') align = 'center';
                        if (style.textAlign === 'right') align = 'right';
                        if (style.textAlign === 'justify') align = 'justify';

                        let valign = 'top';
                        const ptPx = parseFloat(style.paddingTop) || 0;
                        const pbPx = parseFloat(style.paddingBottom) || 0;
                        const plPx = parseFloat(style.paddingLeft) || 0;
                        const prPx = parseFloat(style.paddingRight) || 0;

                        const boxH = rect.height;
                        const textH = parseFloat(style.fontSize) * 1.2;

                        if (style.display.includes('flex') && style.alignItems === 'center') valign = 'middle';
                        else if (Math.abs(ptPx - pbPx) < 5 && ptPx > 5) valign = 'middle';
                        else if (boxH < 40 && boxH > textH) valign = 'middle';

                        if (style.display.includes('flex')) {
                            if (style.justifyContent === 'center') align = 'center';
                            else if (style.justifyContent === 'flex-end' || style.justifyContent === 'right') align = 'right';
                        }

                        // Convert to Inches for Geometry
                        const pt = pxToInch(ptPx);
                        const pr = pxToInch(prPx);
                        const pb = pxToInch(pbPx);
                        const pl = pxToInch(plPx);

                        // Geometry Shift Strategy for Padding
                        let tx = x + pl;
                        let ty = y + pt;
                        let tw = w - pl - pr;
                        let th = h - pt - pb;

                        if (tw < 0) tw = 0;
                        if (th < 0) th = 0;

                        // Line Height (Leading) Support
                        // Convert CSS line-height to PPTX lineSpacing (Points)
                        let lineSpacingPoints = null;
                        const lhStr = style.lineHeight;
                        if (lhStr && lhStr !== 'normal') {
                            const lhPx = parseFloat(lhStr);
                            if (!isNaN(lhPx)) {
                                lineSpacingPoints = lhPx * 0.75; // px to pt
                            }
                        }

                        // Add small buffer to text box width to prevent premature wrapping
                        // due to minor font rendering differences
                        const widthBuffer = pxToInch(2);
                        let finalTx = tx;

                        // Adjust x position for center alignment to keep it visually centered
                        if (align === 'center') finalTx -= widthBuffer / 2;

                        // We use inset:0 because we already applied padding via x/y/w/h
                        const textOpts = {
                            x: finalTx, y: ty, w: tw + widthBuffer, h: th,
                            align: align, valign: valign, margin: 0, inset: 0,
                            autoFit: false, wrap: true
                        };

                        if (lineSpacingPoints) {
                            textOpts.lineSpacing = lineSpacingPoints;
                        }

                        slide.addText(runs, textOpts);
                    }

                    const markSeen = (n) => {
                        n.childNodes.forEach(c => {
                            if (c.nodeType === Node.ELEMENT_NODE) {
                                processedNodes.add(c);
                                markSeen(c);
                            }
                        });
                    };
                    markSeen(node);
                }
            } else {
                // -- NEW: Z-INDEX SORTING & RECURSION --

                // Get all element children
                const children = Array.from(node.children);

                // Map to object with z-index
                const sortedChildren = children.map(c => {
                    const zStr = window.getComputedStyle(c).zIndex;
                    return {
                        node: c,
                        zIndex: (zStr === 'auto') ? 0 : parseInt(zStr)
                    };
                });

                // Sort: ascending z-index
                sortedChildren.sort((a, b) => a.zIndex - b.zIndex);

                // Recurse
                sortedChildren.forEach(item => processNode(item.node));
            }
        }

        // Start Processing (Sorted Top-Level Children)
        const rootChildren = Array.from(container.children).map(c => {
            const zStr = window.getComputedStyle(c).zIndex;
            return {
                node: c,
                zIndex: (zStr === 'auto') ? 0 : parseInt(zStr)
            };
        });
        rootChildren.sort((a, b) => a.zIndex - b.zIndex);

        rootChildren.forEach(item => processNode(item.node));

        // Save
        pres.writeFile({ fileName: fileName });
    }

    // --- Exports ---
    root.LLMDomToPptx = {
        export: exportToPPTX
    };

    // Keep old global for compatibility if needed
    root.exportToPPTX = exportToPPTX;

})(window);
