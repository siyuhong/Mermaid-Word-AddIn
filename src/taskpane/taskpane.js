/* global Word console, DOMParser, document, Image, Blob, URL */

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

    const width =
      parseFloat(svgElement.getAttribute("width")) ||
      parseFloat(svgElement.getAttribute("viewBox")?.split(" ")[2]) ||
      800;
    const height =
      parseFloat(svgElement.getAttribute("height")) ||
      parseFloat(svgElement.getAttribute("viewBox")?.split(" ")[3]) ||
      600;

    return { width, height };
  } catch {
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

          canvas.width = scaledWidth;
          canvas.height = scaledHeight;

          ctx.fillStyle = "white";
          ctx.fillRect(0, 0, scaledWidth, scaledHeight);

          ctx.drawImage(img, 0, 0, scaledWidth, scaledHeight);

          const pngBase64 = canvas.toDataURL("image/png").split(",")[1];

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

      const svgBlob = new Blob([svgContent], { type: "image/svg+xml;charset=utf-8" });
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
    const pngBase64 = await svgToBase64Png(svgContent);

    await Word.run(async (context) => {
      let body = context.document.body;
      let image = body.insertInlinePictureFromBase64(pngBase64, Word.InsertLocation.end);

      // Set the alt-text to the Mermaid code for re-editing
      image.altTextTitle = "Mermaid Diagram";
      image.altTextDescription = mermaidCode;

      // Set width to max Word page width while maintaining aspect ratio
      const { width: svgWidth, height: svgHeight } = getSvgDimensions(svgContent);
      const { width: scaledWidth, height: scaledHeight } = calculateScaledDimensions(
        svgWidth,
        svgHeight
      );

      image.width = scaledWidth;
      image.height = scaledHeight;

      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
    throw error;
  }
}
