/* global Word console, document, Image, btoa, URL, Blob, FileReader */

import { readSvgGeometry } from "./utils/diagramUtils";
import { DEFAULT_PAGE_DIMENSIONS, getDiagramInsertLayout } from "./utils/diagramSizing";
import { prepareSvgForRasterization } from "./utils/svgSanitizer";
import { createSvgPictureOoxml } from "./utils/wordSvg";
import {
  decodeMermaidSource,
  encodeMermaidSource,
  MERMAID_IMAGE_TITLE,
} from "./utils/diagramMetadata";

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
    return { ...DEFAULT_PAGE_DIMENSIONS };
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

// toBlob keeps the PNG in the binary domain; toDataURL would build one huge
// string per attempt and copy it again when base64-encoded for Office.js.
function canvasToBase64Png(canvas) {
  return new Promise((resolve, reject) => {
    try {
      canvas.toBlob((blob) => {
        if (!blob) {
          reject(new Error("Canvas export produced no image"));
          return;
        }
        const reader = new FileReader();
        reader.onload = () => {
          const result = typeof reader.result === "string" ? reader.result : "";
          const comma = result.indexOf(",");
          resolve(comma >= 0 ? result.slice(comma + 1) : "");
        };
        reader.onerror = () => reject(new Error("Could not read the rendered PNG"));
        reader.readAsDataURL(blob);
      }, "image/png");
    } catch (error) {
      // SecurityError: the SVG still contained HTML content. Let the caller
      // classify it instead of reporting a generic canvas failure.
      reject(error);
    }
  });
}

async function svgToBase64Png(svgContent, layout) {
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

      img.onload = async () => {
        try {
          if (
            objectUrl &&
            typeof URL !== "undefined" &&
            typeof URL.revokeObjectURL === "function"
          ) {
            URL.revokeObjectURL(objectUrl);
            objectUrl = null;
          }

          const { width: displayWidth, height: displayHeight } = layout.displaySize;
          const { width: canvasWidth, height: canvasHeight } = layout.rasterSize;
          const { effectiveDpi } = layout;

          console.log(
            `Creating PNG ${canvasWidth}x${canvasHeight}px (${effectiveDpi} DPI) for`,
            displayWidth,
            "x",
            displayHeight,
            "pt"
          );

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

          const pngBase64 = await canvasToBase64Png(canvas);

          console.log("PNG created successfully at", effectiveDpi, "DPI");

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

      // Decode the SVG at the final pixel dimensions as well as drawing it
      // onto a large canvas, avoiding an intermediate low-resolution image.
      const fixedSvgContent = prepareSvgForRasterization(svgContent, layout.rasterSize);

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
        selectedPicture.altTextTitle === MERMAID_IMAGE_TITLE &&
        selectedPicture.altTextDescription
      ) {
        return decodeMermaidSource(selectedPicture.altTextDescription);
      }

      return null;
    });
  } catch (error) {
    console.log("Error getting selected image alt text:", error);
    return null;
  }
}

export async function insertDiagram(svgContent, mermaidCode, options = {}) {
  const format = options.format || "png";
  try {
    if (format !== "png" && format !== "svg") {
      throw new Error(`Unsupported image format: ${format}`);
    }
    console.log("Starting diagram insertion with optimized sizing...");

    // Validate SVG content
    if (!svgContent || typeof svgContent !== "string") {
      throw new Error("Invalid SVG content provided");
    }

    if (!svgContent.includes("<svg")) {
      throw new Error("SVG content does not contain valid SVG markup");
    }

    // Keep the export aspect ratio independent of the preview viewport.
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
    const pageDimensions = await getDocumentPageDimensions();
    console.log("Page dimensions:", pageDimensions);
    const layout = getDiagramInsertLayout(geometry, pageDimensions, options);
    console.log("Calculated target display size (points):", layout.displaySize);

    // SVG also embeds a PNG fallback for Word versions without SVG rendering.
    const pngResult = await svgToBase64Png(svgContent, layout);

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

    const ooxml =
      format === "svg"
        ? createSvgPictureOoxml(svgContent, pngBase64, layout.displaySize, mermaidCode)
        : null;
    await Word.run(async (context) => {
      // Get the current selection to insert at cursor position
      const selection = context.document.getSelection();
      if (format === "svg") {
        // Metadata and size are embedded in the same drawing as the SVG.
        // No second selection lookup or post-insertion metadata write is needed.
        selection.insertOoxml(ooxml, Word.InsertLocation.replace);
      } else {
        const image = selection.insertInlinePictureFromBase64(
          pngBase64,
          Word.InsertLocation.replace
        );
        image.altTextTitle = MERMAID_IMAGE_TITLE;
        image.altTextDescription = encodeMermaidSource(mermaidCode);
        image.width = displayWidth;
        image.height = displayHeight;
      }

      await context.sync();
      console.log("Diagram successfully inserted into Word with optimized sizing");
    });
  } catch (error) {
    console.error("Error inserting diagram:", error);

    // Provide more specific error messages for common issues
    const rawMessage = `${error?.message || ""} ${error?.name || ""}`;
    let errorMessage = `Failed to insert diagram: ${error.message || error}`;

    if (/tainted|securityerror|todataurl/i.test(rawMessage)) {
      errorMessage =
        "Failed to insert diagram: this diagram still renders HTML labels, which Word cannot store as an image. Please report this diagram type.";
    } else if (/canvas export|could not read the rendered png|invalid base64/i.test(rawMessage)) {
      errorMessage =
        "Failed to insert diagram: the rendered image could not be encoded. Try splitting the diagram into smaller sections.";
    } else if (
      format === "svg" &&
      /InvalidArgument|NotImplemented|NotSupported|insertOoxml/i.test(rawMessage)
    ) {
      errorMessage =
        "Failed to insert SVG: this Word version could not import the vector image. Select PNG in the format menu and try again.";
    } else if (error.message && error.message.includes("InvalidArgument")) {
      errorMessage =
        "Failed to insert diagram: Invalid argument. This may be caused by invalid SVG content or dimensions. Please check your Mermaid syntax.";
    } else if (error.message && error.message.includes("base64")) {
      errorMessage =
        "Failed to insert diagram: Unable to convert SVG to PNG. Please try a simpler diagram.";
    }

    throw new Error(errorMessage);
  }
}
