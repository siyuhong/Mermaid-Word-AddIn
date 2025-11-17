/* global Word console, DOMParser, document, Image, btoa */

import { detectDiagramType } from "./utils/diagramUtils";

// Constants for Word page dimensions
// Default fallback values: US Letter size (8.5" x 11" = 612 x 792 points) with 1" margins (72 points)
// 1 point = 1/72 inch (Word uses points for sizing)
const DEFAULT_PAGE_WIDTH_POINTS = 612; // 8.5 inches
const DEFAULT_PAGE_HEIGHT_POINTS = 792; // 11 inches
const DEFAULT_MARGIN_POINTS = 72; // 1 inch margins
const POINTS_PER_INCH = 72;
const PNG_DPI = 300; // Use 300 DPI for high-quality PNG export
const PIXELS_PER_POINT = PNG_DPI / POINTS_PER_INCH; // 4.167 pixels per point at 300 DPI
const MIN_DIAGRAM_SIZE_POINTS = 144; // 2 inches minimum

// Get actual page dimensions and margins from the document
export async function getDocumentPageDimensions() {
  try {
    return await Word.run(async (context) => {
      // Get the current selection to determine which section we're in
      const selection = context.document.getSelection();
      selection.load("parentSectionOrNullObject");
      await context.sync();

      const parentSection = selection.parentSectionOrNullObject;
      if (!parentSection || parentSection.isNullObject) {
        throw new Error("Selection is not associated with a section");
      }

      // Load the pageSetup properties
      parentSection.load("pageSetup");
      await context.sync();

      const pageSetup = parentSection.pageSetup;
      pageSetup.load([
        "pageWidth",
        "pageHeight",
        "leftMargin",
        "rightMargin",
        "topMargin",
        "bottomMargin",
      ]);
      await context.sync();

      console.log("Retrieved page dimensions from document:", {
        pageWidth: pageSetup.pageWidth,
        pageHeight: pageSetup.pageHeight,
        leftMargin: pageSetup.leftMargin,
        rightMargin: pageSetup.rightMargin,
        topMargin: pageSetup.topMargin,
        bottomMargin: pageSetup.bottomMargin,
      });

      return {
        pageWidth: pageSetup.pageWidth,
        pageHeight: pageSetup.pageHeight,
        marginLeft: pageSetup.leftMargin,
        marginRight: pageSetup.rightMargin,
        marginTop: pageSetup.topMargin,
        marginBottom: pageSetup.bottomMargin,
      };
    });
  } catch (error) {
    console.log("Could not retrieve page dimensions from document, using defaults:", error);
    // Fallback to default values if retrieval fails
    return {
      pageWidth: DEFAULT_PAGE_WIDTH_POINTS,
      pageHeight: DEFAULT_PAGE_HEIGHT_POINTS,
      marginLeft: DEFAULT_MARGIN_POINTS,
      marginRight: DEFAULT_MARGIN_POINTS,
      marginTop: DEFAULT_MARGIN_POINTS,
      marginBottom: DEFAULT_MARGIN_POINTS,
    };
  }
}

export async function insertText(text) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

// Parse SVG viewBox to get accurate dimensions
function parseSvgViewBox(svgString) {
  try {
    // Try to extract viewBox from SVG tag (most accurate method)
    const viewBoxMatch = svgString.match(/viewBox\s*=\s*["']([^"']+)["']/);
    if (viewBoxMatch) {
      const viewBoxValues = viewBoxMatch[1].split(/\s+/);
      if (viewBoxValues.length === 4) {
        const width = parseFloat(viewBoxValues[2]);
        const height = parseFloat(viewBoxValues[3]);

        if (!isNaN(width) && !isNaN(height) && width > 0 && height > 0) {
          console.log("Parsed viewBox successfully:", width, "x", height);
          return { width, height };
        }
      }
    }

    // Fallback: try to extract width/height attributes
    const widthMatch = svgString.match(/width\s*=\s*["']?([^"'\s>]+)/);
    const heightMatch = svgString.match(/height\s*=\s*["']?([^"'\s>]+)/);

    if (widthMatch && heightMatch) {
      const width = parseFloat(widthMatch[1]);
      const height = parseFloat(heightMatch[1]);

      if (!isNaN(width) && !isNaN(height) && width > 0 && height > 0) {
        console.log("Parsed width/height attributes:", width, "x", height);
        return { width, height };
      }
    }

    console.log("Could not parse SVG dimensions, using fallback");
    return { width: 0, height: 0 }; // Signal to use natural dimensions
  } catch (error) {
    console.error("Error parsing SVG dimensions:", error);
    return { width: 0, height: 0 };
  }
}

function getSvgDimensions(svgString) {
  try {
    // First, try to parse viewBox from SVG string (most reliable)
    const viewBoxDimensions = parseSvgViewBox(svgString);
    if (viewBoxDimensions.width > 0 && viewBoxDimensions.height > 0) {
      return viewBoxDimensions;
    }

    // Fallback: Use DOM parsing for complex cases
    const parser = new DOMParser();
    const svgDoc = parser.parseFromString(svgString, "image/svg+xml");
    const svgElement = svgDoc.documentElement;

    // Try to get dimensions from width/height attributes
    let width = parseFloat(svgElement.getAttribute("width"));
    let height = parseFloat(svgElement.getAttribute("height"));

    // Try viewBox
    if (!width || !height) {
      const viewBox = svgElement.getAttribute("viewBox");
      if (viewBox) {
        const viewBoxValues = viewBox.split(" ").map(parseFloat);
        if (viewBoxValues.length >= 4) {
          width = width || viewBoxValues[2];
          height = height || viewBoxValues[3];
        }
      }
    }

    // Try to get from the first child element (common for Mermaid)
    if (!width || !height) {
      const firstChild = svgElement.firstElementChild;
      if (firstChild) {
        width = width || parseFloat(firstChild.getAttribute("width"));
        height = height || parseFloat(firstChild.getAttribute("height"));

        // Try viewBox from first child
        if (!width || !height) {
          const childViewBox = firstChild.getAttribute("viewBox");
          if (childViewBox) {
            const childViewBoxValues = childViewBox.split(" ").map(parseFloat);
            if (childViewBoxValues.length >= 4) {
              width = width || childViewBoxValues[2];
              height = height || childViewBoxValues[3];
            }
          }
        }
      }
    }

    // Try DOM measurement as last resort
    if (typeof document !== "undefined" && document.body && (!width || !height)) {
      let tempContainer = null;
      try {
        tempContainer = document.createElement("div");
        tempContainer.style.position = "fixed";
        tempContainer.style.pointerEvents = "none";
        tempContainer.style.opacity = "0";
        tempContainer.style.top = "-10000px";
        tempContainer.style.left = "-10000px";
        tempContainer.innerHTML = svgString;
        document.body.appendChild(tempContainer);

        const tempSvg = tempContainer.querySelector("svg");
        if (tempSvg) {
          let measuredWidth = null;
          let measuredHeight = null;

          try {
            const bbox = tempSvg.getBBox();
            if (bbox && bbox.width && bbox.height) {
              measuredWidth = bbox.width;
              measuredHeight = bbox.height;
            }
          } catch {
            // Ignore measurement errors from getBBox
          }

          if (!measuredWidth || !measuredHeight) {
            const rect = tempSvg.getBoundingClientRect();
            if (rect && rect.width && rect.height) {
              measuredWidth = measuredWidth || rect.width;
              measuredHeight = measuredHeight || rect.height;
            }
          }

          if (tempSvg.viewBox && tempSvg.viewBox.baseVal) {
            measuredWidth = measuredWidth || tempSvg.viewBox.baseVal.width;
            measuredHeight = measuredHeight || tempSvg.viewBox.baseVal.height;
          }

          if (measuredWidth && measuredWidth > 0) {
            width = measuredWidth;
          }
          if (measuredHeight && measuredHeight > 0) {
            height = measuredHeight;
          }
        }
      } catch (measurementError) {
        console.log("Error measuring SVG dimensions via DOM:", measurementError);
      } finally {
        if (tempContainer && tempContainer.parentNode) {
          tempContainer.parentNode.removeChild(tempContainer);
        }
      }
    }

    const fallbackWidth = width && !Number.isNaN(width) ? width : 800;
    const fallbackHeight = height && !Number.isNaN(height) ? height : 600;

    return { width: fallbackWidth, height: fallbackHeight };
  } catch (error) {
    console.log("Error getting SVG dimensions:", error);
    return { width: 800, height: 600 };
  }
}

function normalizePageDimensions(pageDimensions) {
  const layout = pageDimensions || {};

  const ensure = (value, fallback) =>
    typeof value === "number" && Number.isFinite(value) && value > 0 ? value : fallback;

  return {
    pageWidth: ensure(layout.pageWidth, DEFAULT_PAGE_WIDTH_POINTS),
    pageHeight: ensure(layout.pageHeight, DEFAULT_PAGE_HEIGHT_POINTS),
    marginLeft: ensure(layout.marginLeft, DEFAULT_MARGIN_POINTS),
    marginRight: ensure(layout.marginRight, DEFAULT_MARGIN_POINTS),
    marginTop: ensure(layout.marginTop, DEFAULT_MARGIN_POINTS),
    marginBottom: ensure(layout.marginBottom, DEFAULT_MARGIN_POINTS),
  };
}

function calculateDiagramDisplaySize(
  originalWidth,
  originalHeight,
  diagramType = "",
  pageDimensions
) {
  const hasValidWidth = Number.isFinite(originalWidth) && originalWidth > 0;
  const hasValidHeight = Number.isFinite(originalHeight) && originalHeight > 0;

  const DEFAULT_ASPECT_RATIO = 4 / 3;

  const svgAspectRatio =
    hasValidWidth && hasValidHeight && originalHeight !== 0
      ? originalWidth / originalHeight
      : DEFAULT_ASPECT_RATIO;

  const normalizedDimensions = normalizePageDimensions(pageDimensions);
  const availableWidth = Math.max(
    normalizedDimensions.pageWidth -
      normalizedDimensions.marginLeft -
      normalizedDimensions.marginRight,
    MIN_DIAGRAM_SIZE_POINTS
  );
  const availableHeight = Math.max(
    normalizedDimensions.pageHeight -
      normalizedDimensions.marginTop -
      normalizedDimensions.marginBottom,
    MIN_DIAGRAM_SIZE_POINTS
  );

  console.log("Available space (points):", availableWidth, "x", availableHeight);
  console.log("SVG aspect ratio:", svgAspectRatio);

  // Use 90% of available space for better layout and text readability
  const maxUsableWidth = availableWidth * 0.9;
  const maxUsableHeight = availableHeight * 0.9;

  const availableAspectRatio = maxUsableWidth / maxUsableHeight;

  let targetWidth, targetHeight;

  if (svgAspectRatio > availableAspectRatio) {
    // SVG is wider relative to available space - fit to width
    targetWidth = maxUsableWidth;
    targetHeight = targetWidth / svgAspectRatio;
    console.log("Fitting to width (85% of available)");
  } else {
    // SVG is taller relative to available space - fit to height
    targetHeight = maxUsableHeight;
    targetWidth = targetHeight * svgAspectRatio;
    console.log("Fitting to height (85% of available)");
  }

  // Ensure minimum size for better text readability (at least 2.5 inches wide or tall)
  const MIN_READABLE_SIZE_POINTS = 180; // 2.5 inches
  if (targetWidth < MIN_READABLE_SIZE_POINTS && targetHeight < MIN_READABLE_SIZE_POINTS) {
    if (svgAspectRatio >= 1) {
      targetWidth = MIN_READABLE_SIZE_POINTS;
      targetHeight = targetWidth / svgAspectRatio;
    } else {
      targetHeight = MIN_READABLE_SIZE_POINTS;
      targetWidth = targetHeight * svgAspectRatio;
    }
    console.log("Applied minimum readable size constraint");
  }

  // Ensure we don't exceed the available space even with minimum size
  if (targetWidth > maxUsableWidth) {
    targetWidth = maxUsableWidth;
    targetHeight = targetWidth / svgAspectRatio;
    console.log("Capped to maximum width");
  }

  if (targetHeight > maxUsableHeight) {
    targetHeight = maxUsableHeight;
    targetWidth = targetHeight * svgAspectRatio;
    console.log("Capped to maximum height");
  }

  console.log("Calculated diagram display size (points):", {
    diagramType,
    targetWidth,
    targetHeight,
    svgAspectRatio,
    availableWidth,
    availableHeight,
    maxUsableWidth,
    maxUsableHeight,
  });

  return { width: Math.round(targetWidth), height: Math.round(targetHeight) };
}

async function svgToBase64Png(
  svgContent,
  diagramType,
  baseDimensions,
  targetDisplayDimensions,
  pageDimensions
) {
  return new Promise((resolve, reject) => {
    try {
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");

      if (!ctx) {
        reject(new Error("Failed to get canvas context"));
        return;
      }

      const img = new Image();
      img.crossOrigin = "anonymous";

      img.onload = () => {
        try {
          // Parse SVG viewBox for accurate dimensions
          const viewBoxDimensions = parseSvgViewBox(svgContent);
          let svgWidth, svgHeight;

          // ALWAYS prefer viewBox dimensions over natural dimensions for accuracy
          if (viewBoxDimensions.width > 0 && viewBoxDimensions.height > 0) {
            svgWidth = viewBoxDimensions.width;
            svgHeight = viewBoxDimensions.height;
            console.log("Using viewBox dimensions:", svgWidth, "x", svgHeight);
          } else {
            // Fallback to natural/provided dimensions
            svgWidth = baseDimensions?.width || img.naturalWidth || 800;
            svgHeight = baseDimensions?.height || img.naturalHeight || 600;
            console.log("Using natural/fallback dimensions:", svgWidth, "x", svgHeight);
          }

          const hasValidTargetSize =
            targetDisplayDimensions &&
            Number.isFinite(targetDisplayDimensions.width) &&
            Number.isFinite(targetDisplayDimensions.height) &&
            targetDisplayDimensions.width > 0 &&
            targetDisplayDimensions.height > 0;

          // Calculate optimal display size in points (Word's unit)
          const { width: displayWidth, height: displayHeight } = hasValidTargetSize
            ? targetDisplayDimensions
            : calculateDiagramDisplaySize(svgWidth, svgHeight, diagramType, pageDimensions);

          // Calculate 300 DPI pixel dimensions for the target display size
          const canvasWidth = Math.max(Math.round(displayWidth * PIXELS_PER_POINT), 1);
          const canvasHeight = Math.max(Math.round(displayHeight * PIXELS_PER_POINT), 1);

          console.log("Creating 300 DPI PNG:", canvasWidth, "x", canvasHeight, "pixels");
          console.log("For display size:", displayWidth, "x", displayHeight, "points");

          // Set canvas to 300 DPI resolution for target display size
          canvas.width = canvasWidth;
          canvas.height = canvasHeight;

          // Enable high-quality rendering
          ctx.imageSmoothingEnabled = true;
          ctx.imageSmoothingQuality = "high";

          // Fill with white background (Word expects non-transparent images)
          ctx.fillStyle = "white";
          ctx.fillRect(0, 0, canvasWidth, canvasHeight);

          // Draw the SVG scaled to fill the canvas exactly
          ctx.drawImage(img, 0, 0, canvasWidth, canvasHeight);

          // Convert to PNG with maximum quality
          const pngBase64 = canvas.toDataURL("image/png", 1.0).split(",")[1];

          console.log("300 DPI PNG created successfully");

          // Return display size in points (Word's unit), not pixel size
          resolve({ base64: pngBase64, width: displayWidth, height: displayHeight });
        } catch (err) {
          reject(err);
        }
      };

      img.onerror = () => {
        reject(new Error("Failed to load SVG image"));
      };

      // Fix SVG content to ensure proper namespace and dimensions
      let fixedSvgContent = svgContent;
      if (!svgContent.includes('xmlns="http://www.w3.org/2000/svg"')) {
        fixedSvgContent = svgContent.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
      }

      // Remove any potential external references that could cause CORS issues
      fixedSvgContent = fixedSvgContent.replace(/href="http[^"]*"/g, "");
      fixedSvgContent = fixedSvgContent.replace(/xlink:href="http[^"]*"/g, "");
      fixedSvgContent = fixedSvgContent.replace(/@import[^;]*;/g, "");

      // Create a data URL directly from the SVG to avoid CORS issues
      const svgBase64 = btoa(unescape(encodeURIComponent(fixedSvgContent)));
      const svgDataUrl = `data:image/svg+xml;base64,${svgBase64}`;

      img.src = svgDataUrl;
    } catch (err) {
      reject(err);
    }
  });
}

export async function getSelectedImageAltText() {
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("inlinePictures");

      await context.sync();

      const inlinePictures = selection.inlinePictures;
      if (!inlinePictures || inlinePictures.items.length === 0) {
        return null;
      }

      const selectedPicture = inlinePictures.items[0];
      selectedPicture.load("altTextTitle, altTextDescription");

      await context.sync();

      if (
        selectedPicture.altTextTitle === "Mermaid Diagram" &&
        selectedPicture.altTextDescription
      ) {
        return selectedPicture.altTextDescription;
      }

      return null;
    });
  } catch (error) {
    console.log("Error getting selected image alt text:", error);
    return null;
  }
}

export async function insertDiagram(svgContent, mermaidCode, format = "png") {
  // Insert diagram as PNG or SVG with proper scaling for Word document
  try {
    console.log("Starting diagram insertion with format:", format);
    console.log("SVG content length:", svgContent.length);

    // Validate SVG content
    if (!svgContent || typeof svgContent !== "string") {
      throw new Error("Invalid SVG content provided");
    }

    if (!svgContent.includes("<svg")) {
      throw new Error("SVG content does not contain valid SVG markup");
    }

    const diagramType = detectDiagramType(mermaidCode);
    if (diagramType) {
      console.log("Detected Mermaid diagram type:", diagramType);
    }

    const { width: svgWidth, height: svgHeight } = getSvgDimensions(svgContent);
    console.log("SVG dimensions:", { svgWidth, svgHeight });

    // Validate dimensions
    if (
      !Number.isFinite(svgWidth) ||
      !Number.isFinite(svgHeight) ||
      svgWidth <= 0 ||
      svgHeight <= 0
    ) {
      throw new Error(`Invalid SVG dimensions: ${svgWidth}x${svgHeight}`);
    }

    // Get actual page dimensions from the document
    const pageDimensions = await getDocumentPageDimensions();
    console.log("Page dimensions:", pageDimensions);

    const targetDisplaySize = calculateDiagramDisplaySize(
      svgWidth,
      svgHeight,
      diagramType,
      pageDimensions
    );
    console.log("Calculated target display size (points):", targetDisplaySize);

    if (format === "svg") {
      // Insert SVG directly
      await insertSvgDiagram(svgContent, mermaidCode, targetDisplaySize);
    } else {
      // Insert PNG (default behavior)
      await insertPngDiagram(
        svgContent,
        mermaidCode,
        diagramType,
        { width: svgWidth, height: svgHeight },
        targetDisplaySize,
        pageDimensions
      );
    }

    console.log("Diagram successfully inserted into Word as", format.toUpperCase());
  } catch (error) {
    console.error("Error inserting diagram:", error);

    // Provide more specific error messages for common issues
    let errorMessage = `Failed to insert diagram: ${error.message || error}`;

    if (error.message && error.message.includes("InvalidArgument")) {
      errorMessage =
        "Failed to insert diagram: Invalid argument. This may be caused by invalid SVG content or dimensions. Please check your Mermaid syntax.";
    } else if (error.message && error.message.includes("base64")) {
      errorMessage =
        "Failed to insert diagram: Unable to convert SVG to PNG. Please try a simpler diagram.";
    }

    throw new Error(errorMessage);
  }
}

async function insertSvgDiagram(svgContent, mermaidCode, targetDisplaySize) {
  // Prepare SVG content with proper dimensions
  let fixedSvgContent = svgContent;

  // Ensure SVG has proper namespace
  if (!fixedSvgContent.includes('xmlns="http://www.w3.org/2000/svg"')) {
    fixedSvgContent = fixedSvgContent.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
  }

  // Set width and height attributes based on calculated display size (in points)
  const width = Math.round(targetDisplaySize.width);
  const height = Math.round(targetDisplaySize.height);

  // Remove existing width/height attributes and add new ones
  fixedSvgContent = fixedSvgContent.replace(/width\s*=\s*["'][^"']*["']/g, "");
  fixedSvgContent = fixedSvgContent.replace(/height\s*=\s*["'][^"']*["']/g, "");

  // Add width and height attributes after the svg tag
  const svgTagMatch = fixedSvgContent.match(/<svg[^>]*>/);
  if (svgTagMatch) {
    const svgTag = svgTagMatch[0];
    const newSvgTag = svgTag.replace(">", ` width="${width}" height="${height}">`);
    fixedSvgContent = fixedSvgContent.replace(svgTag, newSvgTag);
  }

  // Convert SVG to base64
  const svgBase64 = btoa(unescape(encodeURIComponent(fixedSvgContent)));

  await Word.run(async (context) => {
    try {
      // Get the current selection to insert at cursor position
      let selection = context.document.getSelection();

      // Insert SVG as base64 encoded image
      let image = selection.insertInlinePictureFromBase64(svgBase64, Word.InsertLocation.replace);

      image.altTextTitle = "Mermaid Diagram";
      image.altTextDescription = mermaidCode;

      // Set image size to calculated display dimensions (in points)
      image.width = width;
      image.height = height;

      await context.sync();
      console.log("SVG diagram successfully inserted into Word");
    } catch (error) {
      console.error("Error inserting SVG diagram:", error);
      // If SVG insertion fails, fall back to PNG
      console.log("Falling back to PNG format due to SVG insertion error");
      throw new Error(
        "SVG insertion not supported in this version of Word. Please use PNG format instead."
      );
    }
  });
}

async function insertPngDiagram(
  svgContent,
  mermaidCode,
  diagramType,
  svgDimensions,
  targetDisplaySize,
  pageDimensions
) {
  // Convert SVG to PNG with optimal sizing
  const pngResult = await svgToBase64Png(
    svgContent,
    diagramType,
    svgDimensions,
    targetDisplaySize,
    pageDimensions
  );

  const pngBase64 = pngResult?.base64;
  const displayWidth = pngResult?.width;
  const displayHeight = pngResult?.height;

  console.log("PNG conversion successful, base64 length:", pngBase64?.length ?? 0);
  console.log("Display dimensions (points):", { displayWidth, displayHeight });

  if (!pngBase64 || typeof pngBase64 !== "string" || pngBase64.length === 0) {
    throw new Error("Failed to convert SVG to PNG - invalid base64 output");
  }

  if (
    !Number.isFinite(displayWidth) ||
    !Number.isFinite(displayHeight) ||
    displayWidth <= 0 ||
    displayHeight <= 0
  ) {
    throw new Error(`Invalid display dimensions: ${displayWidth}x${displayHeight}`);
  }

  await Word.run(async (context) => {
    // Get the current selection to insert at cursor position
    let selection = context.document.getSelection();
    let image = selection.insertInlinePictureFromBase64(pngBase64, Word.InsertLocation.replace);

    image.altTextTitle = "Mermaid Diagram";
    image.altTextDescription = mermaidCode;

    // Set image size to calculated display dimensions (in points)
    image.width = Math.round(displayWidth);
    image.height = Math.round(displayHeight);

    await context.sync();
    console.log("PNG diagram successfully inserted into Word");
  });
}
