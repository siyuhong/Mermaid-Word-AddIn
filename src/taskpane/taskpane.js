/* global Word console, DOMParser, document, Image, window, btoa */

const WORD_PAGE_WIDTH_INCHES = 6;
const DPI = 96;
const WORD_MAX_WIDTH_PX = WORD_PAGE_WIDTH_INCHES * DPI;
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

function calculateScaledDimensions(originalWidth, originalHeight) {
  const hasValidWidth = Number.isFinite(originalWidth) && originalWidth > 0;
  const hasValidHeight = Number.isFinite(originalHeight) && originalHeight > 0;

  const fallbackWidth = WORD_MAX_WIDTH_PX;
  const fallbackHeight = WORD_MAX_WIDTH_PX * 0.75;

  const width = hasValidWidth ? originalWidth : fallbackWidth;

  let height;
  if (hasValidHeight) {
    height = originalHeight;
  } else if (hasValidWidth) {
    height = (originalWidth * fallbackHeight) / fallbackWidth;
  } else {
    height = fallbackHeight;
  }

  // Maintain aspect ratio with safe values
  const aspectRatio = height / width || 1;

  // Only scale down if the image is wider than the max width
  if (width > WORD_MAX_WIDTH_PX) {
    const scaledWidth = WORD_MAX_WIDTH_PX;
    const scaledHeight = scaledWidth * aspectRatio;
    return { width: scaledWidth, height: scaledHeight };
  }

  // If the image is smaller than max width, keep original dimensions
  return { width, height };
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

      // Set crossOrigin to anonymous to prevent CORS issues
      img.crossOrigin = "anonymous";

      img.onload = () => {
        try {
          const { width: svgWidth, height: svgHeight } = getSvgDimensions(svgContent);
          const { width: scaledWidth, height: scaledHeight } = calculateScaledDimensions(
            svgWidth,
            svgHeight
          );

          // Set canvas dimensions with additional scale for better quality exports
          const devicePixelRatio = window.devicePixelRatio || 1;
          const exportScale = Math.max(devicePixelRatio, PNG_EXPORT_SCALE);
          canvas.width = scaledWidth * exportScale;
          canvas.height = scaledHeight * exportScale;

          // Scale the context to match the export scale
          ctx.scale(exportScale, exportScale);

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

          resolve(pngBase64);
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
      context.load(selection, "inlinePictures");

      await context.sync();

      const inlinePictures = selection.inlinePictures;
      context.load(inlinePictures, "items");

      await context.sync();

      if (!inlinePictures || inlinePictures.items.length === 0) {
        return null;
      }

      const selectedPicture = inlinePictures.items[0];
      context.load(selectedPicture, "altTextTitle, altTextDescription");

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
