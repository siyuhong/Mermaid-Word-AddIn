/* global Word console, btoa */

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

export async function insertDiagram(svgContent, mermaidCode) {
  // Insert SVG diagram with alt-text containing the Mermaid code
  try {
    await Word.run(async (context) => {
      // Create a base64 encoded SVG
      const svgBase64 = btoa(unescape(encodeURIComponent(svgContent)));
      const dataUrl = `data:image/svg+xml;base64,${svgBase64}`;

      // Insert the image at the end of the document
      let body = context.document.body;
      let image = body.insertInlinePictureFromBase64(dataUrl, Word.InsertLocation.end);

      // Set the alt-text to the Mermaid code for re-editing
      image.altTextTitle = "Mermaid Diagram";
      image.altTextDescription = mermaidCode;

      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
    throw error;
  }
}
