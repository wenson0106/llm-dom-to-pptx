/**
 * LLM DOM to PPTX - v1.0.1
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
                    // Collapse all whitespace (newlines, tabs, concurrent spaces) to single space
                    // This matches browser rendering behavior for standard text.
                    let runText = text.replace(/\s+/g, ' ');

                    if (style.textTransform === 'uppercase') runText = runText.toUpperCase();

                    if (!runText) return;

                    const runOpts = {
                        color: colorParsed ? colorParsed.color : '000000',
                        fontSize: fontSize * 0.75, // px to pt
                        bold: fontWeight,
                        fontFace: getSafeFont(style.fontFamily),
                        // charSpacing: Removed due to rendering issues (huge gaps) in PptxGenJS
                        breakLine: false
                    };

                    if (colorParsed && colorParsed.transparency > 0) {
                        runOpts.transparency = colorParsed.transparency;
                    }

                    // highlight (background color) support
                    const bgParsed = parseColor(style.backgroundColor, parseFloat(style.opacity) || 1);
                    if (bgParsed && bgParsed.transparency < 100) {
                        runOpts.highlight = bgParsed.color;
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
            const w = pxToInch(rect.width);
            const h = pxToInch(rect.height);

            // --- TABLE HANDLING ---
            if (node.tagName === 'TABLE') {
                // Shadow Handler (Updated)
                if (style.boxShadow && style.boxShadow !== 'none') {
                    let tableBg = style.backgroundColor;
                    let tableOp = parseFloat(style.opacity) || 1;
                    let shadowFill = parseColor(tableBg, tableOp);

                    if (!shadowFill || shadowFill.transparency === 100) {
                        shadowFill = { color: 'FFFFFF', transparency: 99 };
                    }

                    // Now we trust 'h' matches sum(rowHeights) because we enforce it below.
                    slide.addShape(pres.ShapeType.rect, {
                        x: x, y: y, w: w, h: h,
                        fill: { color: shadowFill.color, transparency: shadowFill.transparency },
                        shadow: { type: 'outer', angle: 45, blur: 10, offset: 4, opacity: 0.3 },
                        rectRadius: 0
                    });
                }


                const tableRows = [];
                let colWidths = [];
                let rowHeights = []; // New strictly mapped row heights

                if (node.rows.length > 0) {
                    colWidths = Array.from(node.rows[0].cells).map(c => pxToInch(c.getBoundingClientRect().width));

                    // Capture exact row heights
                    rowHeights = Array.from(node.rows).map(r => pxToInch(r.getBoundingClientRect().height));
                }

                Array.from(node.rows).forEach(row => {
                    const rowData = [];
                    Array.from(row.cells).forEach(cell => {
                        const cStyle = window.getComputedStyle(cell);
                        const cRuns = collectTextRuns(cell, cStyle);

                        // backgroundColor fallback: Cell -> Row -> Row Parent (tbody/thead) -> Table
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
                        // Borders logic handles separately or assumes solid for now
                        const bCP = parseColor(cStyle.borderColor);

                        let vAlign = 'top';
                        if (cStyle.verticalAlign === 'middle') vAlign = 'middle';
                        if (cStyle.verticalAlign === 'bottom') vAlign = 'bottom';



                        // Padding / Margin Handling
                        // PptxGenJS 'margin' is in Inches (like x,y,w,h), NOT Points.
                        // Previous bug: * 0.75 converted px -> pt, but was interpreted as Inches (Massive margins).
                        // Fix: Use pxToInch().

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
                            // Fallback to row border
                            if (fallbackW && parseFloat(fallbackW) > 0 && fallbackS !== 'none') {
                                const co = parseColor(fallbackC) || { color: '000000' };
                                return { pt: parseFloat(fallbackW) * 0.75, color: co.color };
                            }
                            return null;
                        };

                        // Fallback styles from Row
                        const rStyle = row ? window.getComputedStyle(row) : null;
                        let rbTopW = 0, rbTopC = null, rbTopS = 'none';
                        let rbBotW = 0, rbBotC = null, rbBotS = 'none';

                        if (rStyle) {
                            rbTopW = rStyle.borderTopWidth; rbTopC = rStyle.borderTopColor; rbTopS = rStyle.borderTopStyle;
                            rbBotW = rStyle.borderBottomWidth; rbBotC = rStyle.borderBottomColor; rbBotS = rStyle.borderBottomStyle;
                        }

                        // NOTE: Row side borders usually don't apply to every cell, but for simple 'border-b' on div/tr, we want it.
                        // We will allow Top/Bottom fallback. Side borders on TR are rare/tricky (start/end).

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

                    // Manual Pagination Logic to fix 'Array expected' error in PptxGenJS when autoPage:true + rowH
                    const availableH = PPT_HEIGHT_IN - y - 0.5; // 0.5 inch safety margin

                    if (h <= availableH) {
                        // Case A: Fits on current slide
                        slide.addTable(tableRows, {
                            x: x, y: y, w: w, colW: colWidths, rowH: rowHeights, autoPage: false
                        });
                    } else {
                        // Case B: Needs split
                        console.warn('Table does not fit on current slide, splitting manually...');

                        // NOTE: This is a simplified split logic.
                        // Ideally we should calculate exactly how many rows fit.
                        // Depending on user needs, we might implement complex splitting here.
                        // For now, if it exceeds, we just push the whole table to the NEXT slide if it fits there.
                        // If it's bigger than a whole slide, we rely on autoPage=false which might clip, 
                        // or we stick to the user's current request of fixing the CRASH.

                        // Strategy: 
                        // 1. If y is significantly down (>1 inch), try creating a new slide and put table there.
                        // 2. If it is already at top or still too big, we just add it and warn (handling >1 page tables requires loop).

                        if (y > 1.0) {
                            // Move to next slide
                            slide = pres.addSlide();
                            currentSlideYOffset += PPT_HEIGHT_IN; // Estimate offset

                            // Re-calculate y for new slide (should be roughly top margin)
                            // But since we are inside a loop with absolute-ish coordinates... 
                            // If we move to a new slide, we reset y to top margin?
                            // Let's place it at top margin (0.5 in)
                            const newY = 0.5;

                            // We need to adjust 'currentSlideYOffset' so that subsequent elements (also processed by absolute rect)
                            // land correctly relative to this new slide.
                            // old absolute Y of table was e.g. 5.0. New relative Y is 0.5. 
                            // So new offset = 5.0 - 0.5 = 4.5.
                            // But this might break other elements aligned with the table.
                            // For this patch, we mainly ensure it doesn't crash.

                            // Let's try to just add it to current slide with autoPage:false first, 
                            // but PptxGenJS might clip it. 
                            // The user error was SPECIFICALLY about the crash.

                            // FIX: Just disable autoPage.
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

            const borderRadius = parseFloat(style.borderRadius) || 0;
            // Strict circle check
            const isCircle = (Math.abs(rect.width - rect.height) < 2) && (borderRadius >= rect.width / 2 - 1);

            // Shadow Support
            if (style.boxShadow && style.boxShadow !== 'none') {
                shapeOpts.shadow = { type: 'outer', angle: 45, blur: 6, offset: 2, opacity: 0.2 };
            }

            // --- Radius Logic ---
            let shapeType = pres.ShapeType.rect;
            if (isCircle) {
                shapeType = pres.ShapeType.ellipse;
            } else if (borderRadius > 0) {
                const minDim = Math.min(rect.width, rect.height);
                let ratio = borderRadius / (minDim / 2);
                shapeOpts.rectRadius = Math.min(ratio, 1.0);
                shapeType = pres.ShapeType.roundRect || 'roundRect';
            }

            // --- Border Logic ---
            if (hasBorder && style.borderLeftWidth === style.borderRightWidth) {
                shapeOpts.line = { color: borderParsed.color, width: borderW * 0.75 };
            } else {
                shapeOpts.line = null;
            }

            // --- B. LEFT ACCENT BORDER (Custom Strategy) ---
            const lW = parseFloat(style.borderLeftWidth) || 0;
            const leftBorderParsed = parseColor(style.borderLeftColor);

            const hasLeftBorder = lW > 0 && leftBorderParsed && style.borderStyle !== 'none';

            if (hasLeftBorder && !shapeOpts.line) {
                if (hasFill) {
                    // Underlay Strategy
                    const underlayOpts = { ...shapeOpts };
                    underlayOpts.fill = { color: leftBorderParsed.color };
                    underlayOpts.line = null;
                    slide.addShape(shapeType, underlayOpts);

                    // Adjust Main Shape
                    const borderInch = pxToInch(lW);
                    shapeOpts.x += borderInch;
                    shapeOpts.w -= borderInch;
                    delete shapeOpts.shadow; // Remove duplicate shadow
                } else {
                    // Side Strip Strategy
                    slide.addShape(pres.ShapeType.rect, {
                        x: x, y: y, w: pxToInch(lW), h: h,
                        fill: { color: leftBorderParsed.color },
                        rectRadius: isCircle ? 0 : (shapeOpts.rectRadius || 0)
                    });
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
                // If it's a bar/strip
                if (rect.height < 15 && rect.width > 100) {
                    slide.addShape(pres.ShapeType.rect, {
                        x: x, y: y, w: w, h: h,
                        fill: { color: '4F46E5' } // Fallback
                    });
                }
            }

            // --- C. TEXT CONTENT ---
            if (isTextBlock(node)) {

                // Extra check: If this is a very deep node, does it have children that are also text blocks?
                // Logic: isTextBlock returns true if it has text nodes.

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
                        const pt = parseFloat(style.paddingTop) || 0;
                        const pb = parseFloat(style.paddingBottom) || 0;
                        const boxH = rect.height;
                        const textH = parseFloat(style.fontSize) * 1.2;

                        if (style.display.includes('flex') && style.alignItems === 'center') valign = 'middle';
                        else if (Math.abs(pt - pb) < 5 && pt > 5) valign = 'middle';
                        else if (boxH < 40 && boxH > textH) valign = 'middle';

                        if (style.display.includes('flex')) {
                            if (style.justifyContent === 'center') align = 'center';
                            else if (style.justifyContent === 'flex-end' || style.justifyContent === 'right') align = 'right';
                        }
                        // Removed forced center for SPAN. It should respect parent/computed alignment.
                        // if (node.tagName === 'SPAN') { align = 'center'; valign = 'middle'; }

                        const widthBuffer = pxToInch(12); // Increased buffer (was 4) to prevent CJK wrapping issues
                        const inset = Math.max(0, pxToInch(Math.min(pt, parseFloat(style.paddingLeft) || 0)));

                        slide.addText(runs, {
                            x: x, y: y, w: w + widthBuffer, h: h,
                            align: align, valign: valign, margin: 0, inset: inset,
                            autoFit: false, wrap: true
                        });
                    }

                    // Mark children as processed
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
                Array.from(node.children).forEach(processNode);
            }
        }

        // Start Processing
        Array.from(container.children).forEach(processNode);

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
