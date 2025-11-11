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
  const ratio = originalHeight / originalWidth;
  const scaledWidth = WORD_MAX_WIDTH_PX;
  const scaledHeight = scaledWidth * ratio;

  return { width: scaledWidth, height: scaledHeight };
}

async function svgToBase64Png(svgContent) {
  return new Promise((resolve, reject) => {
    try {
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");

      if (!ctx) {
        reject(new Error("Failed to get canvas context"));
        return;
      }

      const img = new Image();
      img.onload = () => {
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
        resolve(pngBase64);
      };

      img.onerror = () => {
        reject(new Error("Failed to load SVG image"));
      };

      const svgBlob = new Blob([svgContent], { type: "image/svg+xml" });
      const url = URL.createObjectURL(svgBlob);
      img.src = url;
    } catch (err) {
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
