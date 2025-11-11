/* global Word console, DOMParser, document, Image, Blob, URL, window */

const WORD_PAGE_WIDTH_INCHES = 6;
const DPI = 96;
const WORD_MAX_WIDTH_PX = WORD_PAGE_WIDTH_INCHES * DPI;

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

    // Fallback dimensions with better aspect ratios for different diagram types
    const fallbackWidth = width || 800;
    const fallbackHeight = height || 600;

    return { width: fallbackWidth, height: fallbackHeight };
  } catch (error) {
    console.log("Error getting SVG dimensions:", error);
    return { width: 800, height: 600 };
  }
}

function calculateScaledDimensions(originalWidth, originalHeight) {
  // Maintain aspect ratio
  const aspectRatio = originalHeight / originalWidth;

  // Only scale down if the image is wider than the max width
  if (originalWidth > WORD_MAX_WIDTH_PX) {
    const scaledWidth = WORD_MAX_WIDTH_PX;
    const scaledHeight = scaledWidth * aspectRatio;
    return { width: scaledWidth, height: scaledHeight };
  }

  // If the image is smaller than max width, keep original dimensions
  return { width: originalWidth, height: originalHeight };
}

async function svgToBase64Png(svgContent) {
  return new Promise((resolve, reject) => {
    let objectUrl = null;
    try {
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");

      if (!ctx) {
        reject(new Error("Failed to get canvas context"));
        return;
      }

      const img = new Image();
      img.onload = () => {
        try {
          const { width: svgWidth, height: svgHeight } = getSvgDimensions(svgContent);
          const { width: scaledWidth, height: scaledHeight } = calculateScaledDimensions(
            svgWidth,
            svgHeight
          );

          // Set canvas dimensions with device pixel ratio for better quality
          const dpr = window.devicePixelRatio || 1;
          canvas.width = scaledWidth * dpr;
          canvas.height = scaledHeight * dpr;

          // Scale the context to match device pixel ratio
          ctx.scale(dpr, dpr);

          // Set canvas CSS dimensions
          canvas.style.width = scaledWidth + "px";
          canvas.style.height = scaledHeight + "px";

          // Fill white background
          ctx.fillStyle = "white";
          ctx.fillRect(0, 0, scaledWidth, scaledHeight);

          // Enable image smoothing for better quality
          ctx.imageSmoothingEnabled = true;
          ctx.imageSmoothingQuality = "high";

          // Draw the image
          ctx.drawImage(img, 0, 0, scaledWidth, scaledHeight);

          // Use higher quality PNG
          const pngBase64 = canvas.toDataURL("image/png", 1.0).split(",")[1];

          // Clean up the object URL
          if (objectUrl) {
            URL.revokeObjectURL(objectUrl);
          }

          resolve(pngBase64);
        } catch (err) {
          if (objectUrl) {
            URL.revokeObjectURL(objectUrl);
          }
          reject(err);
        }
      };

      img.onerror = () => {
        if (objectUrl) {
          URL.revokeObjectURL(objectUrl);
        }
        reject(new Error("Failed to load SVG image"));
      };

      // Fix SVG content to ensure proper namespace and dimensions
      let fixedSvgContent = svgContent;
      if (!svgContent.includes('xmlns="http://www.w3.org/2000/svg"')) {
        fixedSvgContent = svgContent.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
      }

      const svgBlob = new Blob([fixedSvgContent], { type: "image/svg+xml;charset=utf-8" });
      objectUrl = URL.createObjectURL(svgBlob);
      img.src = objectUrl;
    } catch (err) {
      if (objectUrl) {
        URL.revokeObjectURL(objectUrl);
      }
      reject(err);
    }
  });
}

export async function insertDiagram(svgContent, mermaidCode) {
  // Insert diagram as PNG with proper scaling for Word document
  try {
    console.log("Starting diagram insertion...");
    console.log("SVG content length:", svgContent.length);

    const { width: svgWidth, height: svgHeight } = getSvgDimensions(svgContent);
    console.log("SVG dimensions:", { svgWidth, svgHeight });

    const pngBase64 = await svgToBase64Png(svgContent);
    console.log("PNG conversion successful, base64 length:", pngBase64.length);

    await Word.run(async (context) => {
      let body = context.document.body;
      let image = body.insertInlinePictureFromBase64(pngBase64, Word.InsertLocation.end);

      // Set the alt-text to the Mermaid code for re-editing
      image.altTextTitle = "Mermaid Diagram";
      image.altTextDescription = mermaidCode;

      // Set width to max Word page width while maintaining aspect ratio
      const { width: scaledWidth, height: scaledHeight } = calculateScaledDimensions(
        svgWidth,
        svgHeight
      );

      console.log("Scaled dimensions for Word:", { scaledWidth, scaledHeight });

      image.width = scaledWidth;
      image.height = scaledHeight;

      await context.sync();
      console.log("Diagram successfully inserted into Word");
    });
  } catch (error) {
    console.error("Error inserting diagram:", error);
    throw new Error(`Failed to insert diagram: ${error.message || error}`);
  }
}
