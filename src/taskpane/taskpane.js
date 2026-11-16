/* global Word console, document, Image, btoa, URL, Blob */

import { readSvgGeometry, fitIntoBox } from "./utils/diagramUtils";

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
      const parentSection = context.document.getSelection().parentSectionOrNullObject;
      parentSection.load("$none");
      await context.sync();

      if (!parentSection || parentSection.isNullObject) {
        throw new Error("Selection is not associated with a section");
      }

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

async function svgToBase64Png(svgContent, geometry, targetDisplaySize) {
  return new Promise((resolve, reject) => {
    let objectUrl = null;
    let hasRetriedWithDataUrl = false;

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
          if (
            objectUrl &&
            typeof URL !== "undefined" &&
            typeof URL.revokeObjectURL === "function"
          ) {
            URL.revokeObjectURL(objectUrl);
            objectUrl = null;
          }

          // The SVG was normalised at render time and preview uses the same
          // geometry, so the canvas keeps exactly the same aspect ratio.
          console.log("Rasterising SVG geometry:", geometry.width, "x", geometry.height);
          const { width: displayWidth, height: displayHeight } = targetDisplaySize;

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
        if (objectUrl && typeof URL !== "undefined" && typeof URL.revokeObjectURL === "function") {
          URL.revokeObjectURL(objectUrl);
          objectUrl = null;
        }

        if (!hasRetriedWithDataUrl) {
          hasRetriedWithDataUrl = true;
          img.src = svgDataUrl;
          return;
        }

        reject(new Error("Failed to load SVG image"));
      };

      // Fix SVG content to ensure proper namespace and dimensions
      let fixedSvgContent = svgContent;
      if (!svgContent.includes('xmlns="http://www.w3.org/2000/svg"')) {
        fixedSvgContent = svgContent.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
      }
      if (!fixedSvgContent.includes('xmlns:xlink="http://www.w3.org/1999/xlink"')) {
        fixedSvgContent = fixedSvgContent.replace(
          "<svg",
          '<svg xmlns:xlink="http://www.w3.org/1999/xlink"'
        );
      }

      // Remove any potential external references that could cause CORS issues
      fixedSvgContent = fixedSvgContent.replace(/href="http[^"]*"/g, "");
      fixedSvgContent = fixedSvgContent.replace(/xlink:href="http[^"]*"/g, "");
      fixedSvgContent = fixedSvgContent.replace(/@import[^;]*;/g, "");

      // Create a data URL directly from the SVG to avoid CORS issues
      const svgBase64 = btoa(unescape(encodeURIComponent(fixedSvgContent)));
      const svgDataUrl = `data:image/svg+xml;base64,${svgBase64}`;

      if (typeof Blob !== "undefined" && typeof URL !== "undefined" && URL.createObjectURL) {
        const svgBlob = new Blob([fixedSvgContent], { type: "image/svg+xml;charset=utf-8" });
        objectUrl = URL.createObjectURL(svgBlob);
        img.src = objectUrl;
      } else {
        img.src = svgDataUrl;
      }
    } catch (err) {
      if (objectUrl && typeof URL !== "undefined" && typeof URL.revokeObjectURL === "function") {
        URL.revokeObjectURL(objectUrl);
      }
      reject(err);
    }
  });
}

export async function getSelectedImageAltText() {
  try {
    return await Word.run(async (context) => {
      const inlinePictures = context.document.getSelection().inlinePictures;
      inlinePictures.load("items/altTextTitle,items/altTextDescription");

      await context.sync();

      if (!inlinePictures || inlinePictures.items.length === 0) {
        return null;
      }

      const selectedPicture = inlinePictures.items[0];
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
    console.log("Starting diagram insertion with optimized sizing...");
    console.log("SVG content length:", svgContent.length);

    // Validate SVG content
    if (!svgContent || typeof svgContent !== "string") {
      throw new Error("Invalid SVG content provided");
    }

    if (!svgContent.includes("<svg")) {
      throw new Error("SVG content does not contain valid SVG markup");
    }

    // One geometry reader and one aspect-preserving fit, identical to the
    // task pane preview. No diagram-type specific sizes anywhere.
    const geometry = readSvgGeometry(svgContent);
    console.log("SVG intrinsic geometry:", geometry);

    // Validate dimensions
    if (
      !Number.isFinite(geometry.width) ||
      !Number.isFinite(geometry.height) ||
      geometry.width <= 0 ||
      geometry.height <= 0
    ) {
      throw new Error(`Invalid SVG dimensions: ${geometry.width}x${geometry.height}`);
    }

    // Get actual page dimensions from the document
    const pageDimensions = normalizePageDimensions(await getDocumentPageDimensions());
    console.log("Page dimensions:", pageDimensions);

    const maxWidth =
      Math.max(
        pageDimensions.pageWidth - pageDimensions.marginLeft - pageDimensions.marginRight,
        MIN_DIAGRAM_SIZE_POINTS
      ) * 0.9;
    const maxHeight =
      Math.max(
        pageDimensions.pageHeight - pageDimensions.marginTop - pageDimensions.marginBottom,
        MIN_DIAGRAM_SIZE_POINTS
      ) * 0.9;
    const targetDisplaySize = fitIntoBox(geometry, { width: maxWidth, height: maxHeight });
    console.log("Calculated target display size (points):", targetDisplaySize);

    // Convert SVG to PNG with optimal sizing
    const pngResult = await svgToBase64Png(svgContent, geometry, targetDisplaySize);

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
      console.log("Diagram successfully inserted into Word with optimized sizing");
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
