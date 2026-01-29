/**
 * LLM DOM to PPTX - v1.1.3
 * Converts Semantic HTML/CSS (e.g. from LLMs) into editable PPTX.
 * 
 * New in v1.1.0:
 * - SVG icon support (converted to images)
 * - CSS gradient background support (linear and radial)
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
        if (!colorStr) return null;
        if (colorStr === 'transparent' || colorStr === 'rgba(0, 0, 0, 0)') {
            return { color: 'FFFFFF', transparency: 100 };
        }

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

        // Convert to Hex string "RRGGBB" (Uppercase)
        const toHex = (n) => {
            const h = Math.round(n).toString(16).toUpperCase();
            return h.length === 1 ? '0' + h : h;
        };
        const hex = toHex(r) + toHex(g) + toHex(b);

        // Calculate Transparency Percent (0 = Opaque, 100 = Transparent)
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

    // --- 2.1 SVG to Image Conversion ---

    /**
     * Converts an SVG element to a base64 PNG data URI.
     * Handles currentColor by replacing it with the computed color.
     * @param {SVGElement} svgElement - The SVG element to convert.
     * @returns {Promise<string>} - A promise that resolves to the base64 data URI.
     */
    async function svgToImage(svgElement) {
        return new Promise((resolve, reject) => {
            try {
                // Clone the SVG to avoid mutating the original
                const clonedSvg = svgElement.cloneNode(true);

                // Get computed color for currentColor replacement
                const computedStyle = window.getComputedStyle(svgElement);
                const currentColor = computedStyle.color || '#000000';

                // Replace currentColor in stroke and fill attributes
                const replaceCurrentColor = (el) => {
                    if (el.getAttribute) {
                        const stroke = el.getAttribute('stroke');
                        const fill = el.getAttribute('fill');
                        if (stroke === 'currentColor') {
                            el.setAttribute('stroke', currentColor);
                        }
                        if (fill === 'currentColor') {
                            el.setAttribute('fill', currentColor);
                        }
                    }
                    // Recurse into children
                    if (el.children) {
                        Array.from(el.children).forEach(replaceCurrentColor);
                    }
                };
                replaceCurrentColor(clonedSvg);

                // Ensure the SVG has explicit width and height for canvas rendering
                const rect = svgElement.getBoundingClientRect();
                const width = rect.width || 24;
                const height = rect.height || 24;

                clonedSvg.setAttribute('width', width);
                clonedSvg.setAttribute('height', height);

                // Serialize SVG to string
                const serializer = new XMLSerializer();
                let svgString = serializer.serializeToString(clonedSvg);

                // Add xmlns if missing
                if (!svgString.includes('xmlns')) {
                    svgString = svgString.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
                }

                // Create a blob and image
                const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
                const url = URL.createObjectURL(blob);

                const img = new Image();
                img.onload = () => {
                    // Create canvas with appropriate scaling for quality
                    const scale = 2; // 2x for better quality
                    const canvas = document.createElement('canvas');
                    canvas.width = width * scale;
                    canvas.height = height * scale;

                    const ctx = canvas.getContext('2d');
                    ctx.scale(scale, scale);
                    ctx.drawImage(img, 0, 0, width, height);

                    URL.revokeObjectURL(url);

                    // Return as base64 PNG
                    const dataUri = canvas.toDataURL('image/png');
                    resolve(dataUri);
                };

                img.onerror = (err) => {
                    URL.revokeObjectURL(url);
                    console.warn('SVG to image conversion failed:', err);
                    resolve(null); // Resolve with null instead of rejecting
                };

                img.src = url;
            } catch (err) {
                console.warn('SVG to image conversion error:', err);
                resolve(null);
            }
        });
    }

    // --- 2.1.1 Gradient to Image Conversion ---

    /**
     * Renders a CSS gradient to a Canvas and returns a base64 PNG data URI.
     * This is used as a workaround since PptxGenJS doesn't support native gradient fills.
     * @param {string} cssGradient - The CSS background-image value (e.g., "linear-gradient(...)").
     * @param {number} width - The width of the gradient image in pixels.
     * @param {number} height - The height of the gradient image in pixels.
     * @returns {string} - A base64 data URI of the rendered gradient.
     */
    function gradientToImage(cssGradient, width, height) {
        console.log('gradientToImage input:', { cssGradient, width, height });
        try {
            // Robust extraction: find content between first ( and last )
            const firstParen = cssGradient.indexOf('(');
            const lastParen = cssGradient.lastIndexOf(')');
            if (firstParen === -1 || lastParen === -1) {
                console.warn('Invalid gradient string: missing parentheses');
                return null;
            }

            const content = cssGradient.substring(firstParen + 1, lastParen).trim();

            const canvas = document.createElement('canvas');
            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d');

            if (cssGradient.includes('linear-gradient')) {
                // Determine angle
                let angle = 180; // Default: to bottom
                const dirMatch = content.match(/^to\s+(top|bottom|left|right|top\s+left|top\s+right|bottom\s+left|bottom\s+right|right\s+bottom|right\s+top|left\s+bottom|left\s+top)/i);
                const degMatch = content.match(/^(-?\d+(\.\d+)?)deg/i);

                if (dirMatch) {
                    // Normalize direction (e.g., "right bottom" -> "bottom right")
                    let dirParts = dirMatch[1].toLowerCase().split(/\s+/);
                    if (dirParts.length === 2) {
                        // Standardize order: vertical first, then horizontal (e.g., "bottom right")
                        if (dirParts[0] === 'left' || dirParts[0] === 'right') {
                            dirParts = [dirParts[1], dirParts[0]];
                        }
                    }
                    const dir = dirParts.join(' ');
                    const dirMap = {
                        'top': 0, 'top right': 45, 'right': 90, 'bottom right': 135,
                        'bottom': 180, 'bottom left': 225, 'left': 270, 'top left': 315
                    };
                    angle = dirMap[dir] !== undefined ? dirMap[dir] : 180;
                } else if (degMatch) {
                    angle = parseFloat(degMatch[1]);
                }

                // Convert angle to canvas gradient coordinates
                const radians = (angle - 90) * Math.PI / 180;
                const cx = width / 2;
                const cy = height / 2;
                const length = Math.sqrt(width * width + height * height) / 2;

                const x1 = cx - Math.cos(radians) * length;
                const y1 = cy - Math.sin(radians) * length;
                const x2 = cx + Math.cos(radians) * length;
                const y2 = cy + Math.sin(radians) * length;

                const gradient = ctx.createLinearGradient(x1, y1, x2, y2);

                // Extract color stops string
                let colorStopsStr = content;
                const commaIndex = content.indexOf(',');
                if (commaIndex !== -1 && (dirMatch || degMatch)) {
                    colorStopsStr = content.substring(commaIndex + 1).trim();
                }

                // Parse color stops correctly handling nested parentheses
                const colorStops = [];
                let currentToken = '';
                let parenDepth = 0;
                for (let i = 0; i < colorStopsStr.length; i++) {
                    const char = colorStopsStr[i];
                    if (char === '(') parenDepth++;
                    if (char === ')') parenDepth--;
                    if (char === ',' && parenDepth === 0) {
                        if (currentToken.trim()) colorStops.push(currentToken.trim());
                        currentToken = '';
                    } else {
                        currentToken += char;
                    }
                }
                if (currentToken.trim()) colorStops.push(currentToken.trim());

                if (colorStops.length === 0) return null;

                // Add color stops to gradient
                colorStops.forEach((stop, idx) => {
                    const percentMatch = stop.match(/\s+(\d+)%$/);
                    let colorStr = stop;
                    // Fix: Avoid division by zero if there's only one stop
                    let position = colorStops.length > 1 ? idx / (colorStops.length - 1) : 0;

                    if (percentMatch) {
                        position = parseInt(percentMatch[1]) / 100;
                        colorStr = stop.substring(0, percentMatch.index).trim();
                    }

                    position = Math.max(0, Math.min(1, position));

                    try {
                        gradient.addColorStop(position, colorStr);
                    } catch (e) {
                        console.error('Invalid color stop:', { stop, colorStr, position, error: e.message });
                    }
                });

                ctx.fillStyle = gradient;
                ctx.fillRect(0, 0, width, height);

            } else if (cssGradient.includes('radial-gradient')) {
                // Simplified radial gradient (center to edge)
                const gradient = ctx.createRadialGradient(
                    width / 2, height / 2, 0,
                    width / 2, height / 2, Math.max(width, height) / 2
                );

                let colorStopsStr = content;
                const shapeMatch = content.match(/^(circle|ellipse)?\s*(at\s+[^,]+)?(?:,\s*)?/i);
                if (shapeMatch && (shapeMatch[1] || shapeMatch[2])) {
                    colorStopsStr = content.substring(shapeMatch[0].length);
                }

                const colorStops = [];
                let currentToken = '';
                let parenDepth = 0;
                for (let i = 0; i < colorStopsStr.length; i++) {
                    const char = colorStopsStr[i];
                    if (char === '(') parenDepth++;
                    if (char === ')') parenDepth--;
                    if (char === ',' && parenDepth === 0) {
                        if (currentToken.trim()) colorStops.push(currentToken.trim());
                        currentToken = '';
                    } else {
                        currentToken += char;
                    }
                }
                if (currentToken.trim()) colorStops.push(currentToken.trim());

                colorStops.forEach((stop, idx) => {
                    const percentMatch = stop.match(/\s+(\d+)%$/);
                    let colorStr = stop;
                    let position = colorStops.length > 1 ? idx / (colorStops.length - 1) : 0;

                    if (percentMatch) {
                        position = parseInt(percentMatch[1]) / 100;
                        colorStr = stop.substring(0, percentMatch.index).trim();
                    }

                    position = Math.max(0, Math.min(1, position));

                    try {
                        gradient.addColorStop(position, colorStr);
                    } catch (e) {
                        console.error('Invalid radial color stop:', { stop, colorStr, position, error: e.message });
                    }
                });

                ctx.fillStyle = gradient;
                ctx.fillRect(0, 0, width, height);
            }

            const dataUrl = canvas.toDataURL('image/png');
            console.log('gradientToImage success, dataUrl length:', dataUrl.length);
            return dataUrl;
        } catch (err) {
            console.error('Gradient to image conversion failed:', err);
            return null;
        }
    }

    // --- 2.2 Gradient Parsing Utilities ---

    /**
     * Parses a CSS linear-gradient value into PptxGenJS gradient format.
     * @param {string} cssValue - The CSS background-image value containing linear-gradient.
     * @returns {Object|null} - PptxGenJS gradient fill object or null if parsing fails.
     * 
     * Example input: "linear-gradient(to bottom right, rgb(15, 23, 42), rgb(30, 27, 75), rgb(30, 58, 138))"
     * Example output: { path: 'linear', stops: [{color:'0F172A',position:0},...], rotate: 135 }
     */
    function parseLinearGradient(cssValue) {
        if (!cssValue || !cssValue.includes('linear-gradient')) return null;

        try {
            // Robust extraction: find content between first ( and last )
            const firstParen = cssValue.indexOf('(');
            const lastParen = cssValue.lastIndexOf(')');
            if (firstParen === -1 || lastParen === -1) return null;

            const content = cssValue.substring(firstParen + 1, lastParen).trim();

            // Parse direction and color stops
            let rotate = 180; // Default: to bottom
            let colorStopsStr = content;

            // Check for "to <direction>" syntax
            const dirMatch = content.match(/^to\s+(top|bottom|left|right|top\s+left|top\s+right|bottom\s+left|bottom\s+right)/i);

            // Check for degree syntax
            const degMatch = content.match(/^(-?\d+(\.\d+)?)deg/i);

            if (dirMatch) {
                const dir = dirMatch[1].toLowerCase().replace(/\s+/g, ' ');
                const dirMap = {
                    'top': 0,
                    'top right': 45,
                    'right': 90,
                    'bottom right': 135,
                    'bottom': 180,
                    'bottom left': 225,
                    'left': 270,
                    'top left': 315
                };
                rotate = dirMap[dir] !== undefined ? dirMap[dir] : 180;
                // Slice after the comma following the direction
                const commaIndex = content.indexOf(',');
                if (commaIndex !== -1) {
                    colorStopsStr = content.substring(commaIndex + 1).trim();
                }
            }
            else if (degMatch) {
                rotate = parseFloat(degMatch[1]);
                // Slice after the deg unit and comma
                const commaIndex = content.indexOf(',');
                if (commaIndex !== -1) {
                    colorStopsStr = content.substring(commaIndex + 1).trim();
                }
            }

            // Normalize angle to 0-360 positive
            rotate = rotate % 360;
            if (rotate < 0) rotate += 360;

            // Parse color stops splitting by comma (handling nested parentheses for rgb/rgba)
            const colorStops = [];
            let currentToken = '';
            let parenDepth = 0;

            for (let i = 0; i < colorStopsStr.length; i++) {
                const char = colorStopsStr[i];
                if (char === '(') parenDepth++;
                if (char === ')') parenDepth--;

                if (char === ',' && parenDepth === 0) {
                    if (currentToken.trim()) {
                        colorStops.push(currentToken.trim());
                    }
                    currentToken = '';
                } else {
                    currentToken += char;
                }
            }
            if (currentToken.trim()) {
                colorStops.push(currentToken.trim());
            }

            if (colorStops.length < 2) return null;

            // Convert to PptxGenJS format
            const stops = colorStops.map((stop, idx) => {
                // Each stop 'color position%' or just 'color'
                // Robust regex to separate color and percentage
                // Match last occurring space followed by digits and %
                const percentMatch = stop.match(/\s+(\d+)%$/);

                let colorStr = stop;
                let position = -1;

                if (percentMatch) {
                    position = parseInt(percentMatch[1]);
                    colorStr = stop.substring(0, percentMatch.index).trim();
                } else {
                    // Auto-distribute
                    position = Math.round((idx / (colorStops.length - 1)) * 100);
                }

                const parsed = parseColor(colorStr.trim());
                if (!parsed) return null;

                return {
                    color: parsed.color,
                    transparency: parsed.transparency,
                    position: position
                };
            }).filter(s => s !== null);

            if (stops.length < 2) return null;

            return {
                type: 'linear',
                rotate: rotate,
                stops: stops
            };
        } catch (err) {
            console.warn('Failed to parse linear gradient:', err);
            return null;
        }
    }

    /**
     * Parses a CSS radial-gradient value into PptxGenJS gradient format.
     * @param {string} cssValue - The CSS background-image value containing radial-gradient.
     * @returns {Object|null} - PptxGenJS gradient fill object or null if parsing fails.
     */
    function parseRadialGradient(cssValue) {
        if (!cssValue || !cssValue.includes('radial-gradient')) return null;

        try {
            const match = cssValue.match(/radial-gradient\(([^)]+(?:\([^)]*\)[^)]*)*)\)/);
            if (!match) return null;

            const content = match[1].trim();

            // For radial gradients, we'll simplify by using the colors only
            // PptxGenJS supports path: 'circle' for radial

            // Parse color stops (same logic as linear)
            const colorStops = [];
            let currentToken = '';
            let parenDepth = 0;

            // Skip shape definition (circle, ellipse, at center, etc.)
            let colorStopsStr = content;
            const shapeMatch = content.match(/^(circle|ellipse)?\s*(at\s+[^,]+)?(?:,\s*)?/i);
            if (shapeMatch) {
                colorStopsStr = content.substring(shapeMatch[0].length);
            }

            for (let i = 0; i < colorStopsStr.length; i++) {
                const char = colorStopsStr[i];
                if (char === '(') parenDepth++;
                if (char === ')') parenDepth--;

                if (char === ',' && parenDepth === 0) {
                    if (currentToken.trim()) {
                        colorStops.push(currentToken.trim());
                    }
                    currentToken = '';
                } else {
                    currentToken += char;
                }
            }
            if (currentToken.trim()) {
                colorStops.push(currentToken.trim());
            }

            if (colorStops.length < 2) return null;

            const stops = colorStops.map((stop, idx) => {
                const parts = stop.match(/^(.+?)\s+(\d+)%$/);
                let colorStr, position;

                if (parts) {
                    colorStr = parts[1];
                    position = parseInt(parts[2]);
                } else {
                    colorStr = stop;
                    position = Math.round((idx / (colorStops.length - 1)) * 100);
                }

                const parsed = parseColor(colorStr.trim());
                if (!parsed) return null;

                return {
                    color: parsed.color,
                    transparency: parsed.transparency,
                    position: position
                };
            }).filter(s => s !== null);

            if (stops.length < 2) return null;

            return {
                type: 'radial',
                stops: stops
            };
        } catch (err) {
            console.warn('Failed to parse radial gradient:', err);
            return null;
        }
    }

    /**
     * Creates a PptxGenJS fill object from a gradient definition.
     * @param {Object} gradient - Parsed gradient object from parseLinearGradient or parseRadialGradient.
     * @returns {Object} - PptxGenJS fill object.
     */
    function createGradientFill(gradient) {
        if (!gradient || !gradient.stops || gradient.stops.length < 2) return null;

        // Normalize positions to 0-1 range
        const normalizedStops = gradient.stops.map(s => ({
            color: s.color,
            transparency: s.transparency, // Pass transparency (0-100)
            position: s.position > 1 ? s.position / 100 : s.position
        }));

        if (gradient.type === 'linear') {
            return {
                type: 'linear',
                angle: gradient.rotate,
                stops: normalizedStops
            };
        } else if (gradient.type === 'radial') {
            return {
                type: 'linear', // Fallback to linear for radial
                angle: 45,
                stops: normalizedStops
            };
        }

        return null;
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
    /**
     * Internal function to add a slide from a DOM element.
     */
    async function addSlideToPres(pres, slide, container, options = {}) {
        if (!pres || !slide || !container) {
            console.error('Missing required arguments for addSlide: [pres, slide, container]');
            return;
        }

        const containerRect = container.getBoundingClientRect();
        const containerStyle = window.getComputedStyle(container);

        // --- 0. Slide Background ---
        const bgImage = containerStyle.backgroundImage;
        let slideBackgroundSet = false;

        if (bgImage && bgImage !== 'none') {
            const isGradient = bgImage.includes('linear-gradient') || bgImage.includes('radial-gradient');
            if (isGradient) {
                const imgWidth = 1920;
                const imgHeight = 1080;
                const gradientDataUri = gradientToImage(bgImage, imgWidth, imgHeight);

                if (gradientDataUri) {
                    slide.addImage({
                        data: gradientDataUri,
                        x: 0, y: 0, w: 10, h: 5.625
                    });
                    slideBackgroundSet = true;
                    console.log('Slide background gradient rendered as image');
                }
            }
        }

        if (!slideBackgroundSet) {
            const bgParsed = parseColor(containerStyle.backgroundColor);
            if (bgParsed) {
                slide.background = { color: bgParsed.color, transparency: bgParsed.transparency };
            }
        }

        const processedNodes = new Set();
        const svgPromises = [];

        // Options for positioning the container content on the slide
        const xOffset = options.x !== undefined ? options.x : 0;
        const yOffset = options.y !== undefined ? options.y : 0;
        const currentSlideYOffset = 0;

        // --- 0.1 Container Border & Shadow (NEW) ---
        // Since the container itself is normally skipped in processNode loop,
        // we manually draw its border, shadow and radius here if they exist.
        const borderW = parseFloat(containerStyle.borderWidth) || 0;
        const borderParsed = parseColor(containerStyle.borderColor);
        const borderRadius = parseFloat(containerStyle.borderRadius) || 0;
        const hasBorder = borderW > 0 && borderParsed;
        const hasShadow = containerStyle.boxShadow && containerStyle.boxShadow !== 'none';

        if (hasBorder || (borderRadius > 0 && !slideBackgroundSet) || hasShadow) {
            let shapeOpts = { x: xOffset, y: yOffset, w: pxToInch(containerRect.width), h: pxToInch(containerRect.height) };
            let shapeName = 'rect';
            if (borderRadius > 0) {
                const minDim = Math.min(containerRect.width, containerRect.height);
                shapeOpts.rectRadius = Math.min(borderRadius / (minDim / 2), 1.0);
                shapeName = 'roundRect';
            }
            if (hasShadow) {
                shapeOpts.shadow = { type: 'outer', angle: 45, blur: 6, offset: 2, opacity: 0.2 };
            }
            if (hasBorder) {
                shapeOpts.line = { color: borderParsed.color, width: borderW * 0.75 };
            }
            // If we have a slide background, we don't want to fill this shape unless it has its own bg
            // But since addSlideToPres already sets slide.background, we only add a shape for borders/shadows.
            slide.addShape(shapeName, shapeOpts);
        }

        // Helper: recurse gather text runs (nested for closure access if needed, though they are currently sibling level)
        // I will move identifies and collectors to be sibling level so they remain efficient.

        async function processNode(node) {
            if (node.nodeType !== Node.ELEMENT_NODE) return;
            if (processedNodes.has(node)) return;
            processedNodes.add(node);

            const style = window.getComputedStyle(node);
            const rect = node.getBoundingClientRect();
            const opacity = parseFloat(style.opacity) || 1;

            if (style.display === 'none' || style.visibility === 'hidden' || opacity === 0) return;
            if (rect.width < 1 || rect.height < 1) return;

            const x = pxToInch(rect.left - containerRect.left) + xOffset;
            const y = pxToInch(rect.top - containerRect.top) - currentSlideYOffset + yOffset;
            let w = pxToInch(rect.width);
            let h = pxToInch(rect.height);

            const MIN_DIM = 0.02;
            if (h < MIN_DIM && h > 0) h = MIN_DIM;
            if (w < MIN_DIM && w > 0) w = MIN_DIM;

            if (node.tagName === 'svg' || node.tagName === 'SVG') {
                const imgPromise = svgToImage(node).then(dataUri => {
                    if (dataUri) {
                        slide.addImage({ data: dataUri, x, y, w, h });
                    }
                });
                svgPromises.push(imgPromise);
                const markSvgChildren = (n) => {
                    processedNodes.add(n);
                    if (n.children) Array.from(n.children).forEach(markSvgChildren);
                };
                markSvgChildren(node);
                return;
            }

            const bgParsed = parseColor(style.backgroundColor, opacity);
            const borderW = parseFloat(style.borderWidth) || 0;
            const borderParsed = parseColor(style.borderColor);
            const borderRadius = parseFloat(style.borderRadius) || 0;
            const isCircle = (Math.abs(rect.width - rect.height) < 2) && (borderRadius >= rect.width / 2 - 1);

            let hasFill = bgParsed && bgParsed.transparency < 100;
            let hasBorder = borderW > 0 && borderParsed;

            let shapeOpts = { x, y, w, h };
            let shapeName = isCircle ? 'ellipse' : 'rect';

            if (!isCircle && borderRadius > 0) {
                const minDim = Math.min(rect.width, rect.height);
                shapeOpts.rectRadius = Math.min(borderRadius / (minDim / 2), 1.0);
                shapeName = 'roundRect';
            }

            if (style.boxShadow && style.boxShadow !== 'none') {
                shapeOpts.shadow = { type: 'outer', angle: 45, blur: 6, offset: 2, opacity: 0.2 };
            }

            if (hasFill || hasBorder) {
                if (hasFill) {
                    shapeOpts.fill = { color: bgParsed.color, transparency: bgParsed.transparency };
                }
                if (hasBorder) {
                    shapeOpts.line = { color: borderParsed.color, width: borderW * 0.75 };
                }
                slide.addShape(shapeName, shapeOpts);
            }

            if (!hasFill && style.backgroundImage && style.backgroundImage.includes('gradient') && style.backgroundClip !== 'text' && style.webkitBackgroundClip !== 'text') {
                const imgW = Math.max(100, Math.round(rect.width * 2));
                const imgH = Math.max(100, Math.round(rect.height * 2));
                const dataUri = gradientToImage(style.backgroundImage, imgW, imgH);
                if (dataUri) slide.addImage({ data: dataUri, x, y, w, h });
            }

            if (isTextBlock(node)) {
                const runs = collectTextRuns(node, style);
                if (runs.length > 0) {
                    let align = style.textAlign || 'left';
                    if (align === 'start') align = 'left';
                    if (align === 'end') align = 'right';

                    let valign = 'top';
                    const pt = parseFloat(style.paddingTop) || 0;
                    const pb = parseFloat(style.paddingBottom) || 0;
                    const boxH = rect.height;
                    const textH = parseFloat(style.fontSize) * 1.2;

                    if (style.display.includes('flex') && style.alignItems === 'center') valign = 'middle';
                    else if (Math.abs(pt - pb) < 5 && pt > 5) valign = 'middle';
                    else if (boxH < 40 && boxH > textH && Math.abs(pt - pb) < 2) valign = 'middle';

                    slide.addText(runs, {
                        x: x + pxToInch(parseFloat(style.paddingLeft) || 0),
                        y: y + pxToInch(pt),
                        w: w - pxToInch((parseFloat(style.paddingLeft) || 0) + (parseFloat(style.paddingRight) || 0)) + pxToInch(5),
                        h: h - pxToInch(pt + pb),
                        align, valign, margin: 0, inset: 0, wrap: true
                    });

                    const mark = (n) => { n.childNodes.forEach(c => { if (c.nodeType === Node.ELEMENT_NODE) { processedNodes.add(c); mark(c); } }); };
                    mark(node);
                }
            } else {
                const children = Array.from(node.children);
                const sorted = children.map(c => ({ node: c, z: parseInt(window.getComputedStyle(c).zIndex) || 0 })).sort((a, b) => a.z - b.z);
                for (const item of sorted) await processNode(item.node);
            }
        }

        const rootSorted = Array.from(container.children).map(c => ({ node: c, z: parseInt(window.getComputedStyle(c).zIndex) || 0 })).sort((a, b) => a.z - b.z);
        for (const item of rootSorted) await processNode(item.node);

        if (svgPromises.length > 0) await Promise.all(svgPromises);
    }

    /**
     * Exports DOM elements to PPTX. Supports single ID, element, or selector.
     */
    async function exportToPPTX(elementOrId = 'slide-canvas', options = {}) {
        const fileName = options.fileName || "presentation.pptx";
        if (typeof PptxGenJS === 'undefined') { alert("Error: PptxGenJS missing."); return; }

        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';

        let targets = [];
        if (typeof elementOrId === 'string') {
            const el = document.getElementById(elementOrId);
            if (el) targets = [el];
            else targets = Array.from(document.querySelectorAll(elementOrId));
        } else if (elementOrId instanceof HTMLElement) {
            targets = [elementOrId];
        }

        if (targets.length === 0) { console.error("No export targets found."); return; }

        for (const container of targets) {
            const slide = pres.addSlide();
            await addSlideToPres(pres, slide, container, options);
        }

        await pres.writeFile({ fileName });
    }

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

                // Gradient Text Override: If it's a clipped text (usually transparent), force to black
                if ((style.backgroundClip === 'text' || style.webkitBackgroundClip === 'text') && (!colorParsed || colorParsed.transparency === 100)) {
                    runOpts.color = '000000';
                    delete runOpts.transparency;
                }

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
                    // Only apply transparency if it's NOT a gradient text override
                    if (!(runOpts.color === '000000' && (style.backgroundClip === 'text' || style.webkitBackgroundClip === 'text'))) {
                        runOpts.transparency = colorParsed.transparency;
                    }
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

    // Store SVG conversion promises for async handling
    const svgPromises = [];

    async function processNode(node) {
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

        // --- SVG HANDLING ---
        if (node.tagName === 'svg' || node.tagName === 'SVG') {
            // Convert SVG to image
            const imgPromise = svgToImage(node).then(dataUri => {
                if (dataUri) {
                    slide.addImage({
                        data: dataUri,
                        x: x,
                        y: y,
                        w: w,
                        h: h
                    });
                    console.log(`SVG converted to image at (${x.toFixed(2)}", ${y.toFixed(2)}") size ${w.toFixed(2)}" x ${h.toFixed(2)}"`);
                }
            });
            svgPromises.push(imgPromise);

            // Mark all SVG children as processed to avoid duplicates
            const markSvgChildren = (n) => {
                processedNodes.add(n);
                if (n.children) {
                    Array.from(n.children).forEach(markSvgChildren);
                }
            };
            markSvgChildren(node);
            return;
        }

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

        let shapeName = 'rect';
        let rotation = 0;

        if (isCircle) {
            shapeName = 'ellipse';
        } else {
            // Determine Shape based on corners
            // Case 1: All Uniform
            if (rtl === rtr && rtr === rbr && rbr === rbl && rtl > 0) {
                const minDim = Math.min(rect.width, rect.height);
                let ratio = rtl / (minDim / 2);
                shapeOpts.rectRadius = Math.min(ratio, 1.0);
                shapeName = 'roundRect';
            }
            // Case 2: Vertical Rounding (Left or Right) -> PptxGenJS doesn't inherently support 'Right Rounded' well without rotation?
            // Actually PptxGenJS has 'round2SameRect' which is Top Corners Rounded by default.

            // Case 3: Top Rounded (Common in cards)
            else if (rtl > 0 && rtr > 0 && rbr === 0 && rbl === 0) {
                shapeName = 'round2SameRect';
                // Rotate? Default is Top.
                rotation = 0;
            }
            // Case 4: Bottom Rounded
            else if (rtl === 0 && rtr === 0 && rbr > 0 && rbl > 0) {
                shapeName = 'round2SameRect';
                rotation = 180;
            }
            // Case 5: Single corners or diagonal? Fallback to rect (Square) to avoid "All Rounded" bug.
            else if (borderRadius > 0) {
                // It has some radius, but not uniform and not simple top/bottom pair.
                // Fallback: If we use 'roundRect' it rounds all. 
                // Better to use 'rect' (Sharp) than incorrect 'roundRect' for things like "only top-left".
                shapeName = 'rect';
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
                slide.addShape('rect', stripOpts);
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
                slide.addShape(shapeName, underlayOpts);

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
                slide.addShape('rect', stripOpts);
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
                slide.addShape(shapeName, underlayOpts);

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
                slide.addShape('rect', stripOpts);
            }
        }

        // Draw Main Shape
        if (hasFill) {
            shapeOpts.fill = { color: bgParsed.color };
            if (bgParsed.transparency > 0) {
                shapeOpts.fill.transparency = bgParsed.transparency;
            }
            slide.addShape(shapeName, shapeOpts);
        } else if (hasBorder && shapeOpts.line) {
            slide.addShape(shapeName, shapeOpts);
        }

        // Removal of old processNode loop from here...

    }

    // --- Exports ---
    root.LLMDomToPptx = {
        export: exportToPPTX,
        addSlide: addSlideToPres,
        version: "1.1.3"
    };

    root.exportToPPTX = exportToPPTX;

})(window);
