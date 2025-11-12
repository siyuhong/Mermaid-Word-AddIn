/* global Word console, DOMParser, document, Image, window, btoa */

import { detectDiagramType, isGanttDiagram, isStateDiagram } from "./utils/diagramUtils";

const WORD_PAGE_WIDTH_INCHES = 6;
const DPI = 96;
const WORD_MAX_WIDTH_PX = WORD_PAGE_WIDTH_INCHES * DPI;
const WORD_TARGET_WIDTH_PX = WORD_MAX_WIDTH_PX * 0.8; // 80% of max width for better fit
const PNG_EXPORT_SCALE = 2;

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

function getSvgDimensions(svgString) {
  try {
    const parser = new DOMParser();
    const svgDoc = parser.parseFromString(svgString, "image/svg+xml");
    const svgElement = svgDoc.documentElement;

    // Try to get dimensions from width/height attributes first
    let width = parseFloat(svgElement.getAttribute("width"));
    let height = parseFloat(svgElement.getAttribute("height"));

    // If not found, try viewBox
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

    // If still not found, try to get from the first child element (common for Mermaid)
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

    if (typeof document !== "undefined" && document.body) {
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
            // Ignore measurement errors from getBBox and fall back to other strategies
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

function calculateScaledDimensions(originalWidth, originalHeight, diagramType = "") {
  const hasValidWidth = Number.isFinite(originalWidth) && originalWidth > 0;
  const hasValidHeight = Number.isFinite(originalHeight) && originalHeight > 0;

  const fallbackWidth = WORD_TARGET_WIDTH_PX;
  const fallbackHeight = WORD_TARGET_WIDTH_PX * 0.75;

  const width = hasValidWidth ? originalWidth : fallbackWidth;

  let height;
  if (hasValidHeight) {
    height = originalHeight;
  } else if (hasValidWidth) {
    height = (originalWidth * fallbackHeight) / fallbackWidth;
  } else {
    height = fallbackHeight;
  }

  const safeAspectRatio = width > 0 ? height / width : fallbackHeight / fallbackWidth;

  const preferredTargetWidth = isGanttDiagram(diagramType)
    ? WORD_MAX_WIDTH_PX
    : isStateDiagram(diagramType)
      ? WORD_TARGET_WIDTH_PX * 0.75
      : WORD_TARGET_WIDTH_PX;

  const minimumAcceptableWidth = isGanttDiagram(diagramType)
    ? WORD_TARGET_WIDTH_PX
    : WORD_TARGET_WIDTH_PX * 0.6;

  const targetWidth = Math.min(
    Math.max(preferredTargetWidth, minimumAcceptableWidth),
    WORD_MAX_WIDTH_PX
  );

  let finalWidth = width;
  let finalHeight = height;

  if (width > targetWidth) {
    finalWidth = targetWidth;
    finalHeight = targetWidth * safeAspectRatio;
  } else if (width < minimumAcceptableWidth) {
    finalWidth = minimumAcceptableWidth;
    finalHeight = minimumAcceptableWidth * safeAspectRatio;
  }

  return { width: finalWidth, height: finalHeight };
}

async function svgToBase64Png(svgContent, diagramType, baseDimensions) {
  return new Promise((resolve, reject) => {
    try {
      const fallbackDimensions = baseDimensions || getSvgDimensions(svgContent);
      const fallbackWidth =
        Number.isFinite(fallbackDimensions?.width) && fallbackDimensions.width > 0
          ? fallbackDimensions.width
          : WORD_TARGET_WIDTH_PX;
      const fallbackHeight =
        Number.isFinite(fallbackDimensions?.height) && fallbackDimensions.height > 0
          ? fallbackDimensions.height
          : WORD_TARGET_WIDTH_PX * 0.75;

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
          const intrinsicWidth =
            Number.isFinite(img.naturalWidth) && img.naturalWidth > 0
              ? img.naturalWidth
              : fallbackWidth;
          const intrinsicHeight =
            Number.isFinite(img.naturalHeight) && img.naturalHeight > 0
              ? img.naturalHeight
              : fallbackHeight;

          const { width: scaledWidth, height: scaledHeight } = calculateScaledDimensions(
            intrinsicWidth,
            intrinsicHeight,
            diagramType
          );

          const devicePixelRatio = window.devicePixelRatio || 1;
          const exportScale = Math.max(devicePixelRatio, PNG_EXPORT_SCALE);
          const scaledCanvasWidth = Math.max(Math.round(scaledWidth * exportScale), 1);
          const scaledCanvasHeight = Math.max(Math.round(scaledHeight * exportScale), 1);

          canvas.width = scaledCanvasWidth;
          canvas.height = scaledCanvasHeight;
          ctx.setTransform(exportScale, 0, 0, exportScale, 0, 0);

          canvas.style.width = `${scaledWidth}px`;
          canvas.style.height = `${scaledHeight}px`;

          ctx.fillStyle = "white";
          ctx.fillRect(0, 0, scaledWidth, scaledHeight);
          ctx.imageSmoothingEnabled = true;
          ctx.imageSmoothingQuality = "high";
          ctx.drawImage(img, 0, 0, scaledWidth, scaledHeight);

          const pngBase64 = canvas.toDataURL("image/png", 1.0).split(",")[1];

          resolve({ base64: pngBase64, width: scaledWidth, height: scaledHeight });
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

export async function insertDiagram(svgContent, mermaidCode) {
  // Insert diagram as PNG with proper scaling for Word document
  try {
    console.log("Starting diagram insertion...");
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

    const pngResult = await svgToBase64Png(svgContent, diagramType, {
      width: svgWidth,
      height: svgHeight,
    });
    const pngBase64 = pngResult?.base64;
    const scaledWidth = pngResult?.width;
    const scaledHeight = pngResult?.height;
    console.log("PNG conversion successful, base64 length:", pngBase64?.length ?? 0);

    if (!pngBase64 || typeof pngBase64 !== "string" || pngBase64.length === 0) {
      throw new Error("Failed to convert SVG to PNG - invalid base64 output");
    }

    console.log("Scaled dimensions for Word:", { scaledWidth, scaledHeight });

    if (
      !Number.isFinite(scaledWidth) ||
      !Number.isFinite(scaledHeight) ||
      scaledWidth <= 0 ||
      scaledHeight <= 0
    ) {
      throw new Error(`Invalid scaled dimensions: ${scaledWidth}x${scaledHeight}`);
    }

    await Word.run(async (context) => {
      let body = context.document.body;
      let image = body.insertInlinePictureFromBase64(pngBase64, Word.InsertLocation.end);

      image.altTextTitle = "Mermaid Diagram";
      image.altTextDescription = mermaidCode;

      image.width = Math.round(scaledWidth);
      image.height = Math.round(scaledHeight);

      await context.sync();
      console.log("Diagram successfully inserted into Word");
    });
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
