// Office is loaded via CDN in taskpane.html

/* global Word, console, setTimeout */

// Constants for Mermaid diagram identification
const MERMAID_SIGNATURE_PREFIX = "mermaid-diagram:";

/**
 * Insert a Mermaid diagram into the Word document
 * @param {string} code - The Mermaid diagram code
 * @returns {Promise<void>}
 */
export const insertMermaidDiagram = async (code) => {
  try {
    // Generate diagram metadata
    const metadata = {
      code,
      timestamp: new Date().toISOString(),
      version: "1.0",
    };

    // Create alt text description with metadata
    const altTextDescription = `${MERMAID_SIGNATURE_PREFIX}${JSON.stringify(metadata)}`;

    await Word.run(async (context) => {
      // Get the current selection
      const selection = context.document.getSelection();

      // Insert a placeholder for the diagram
      // In a real implementation, this would generate the actual PNG
      // For now, we'll insert a placeholder image or use a service
      const base64Image = await generateMermaidImage(code);

      // Insert the inline picture
      const inlinePicture = selection.insertInlinePictureFromBase64(
        base64Image,
        Word.InsertLocation.end
      );

      // Set the alt text with metadata
      inlinePicture.altTextDescription = altTextDescription;

      await context.sync();
    });
  } catch (error) {
    console.error("Failed to insert Mermaid diagram:", error);
    throw new Error(`Failed to insert diagram: ${error.message}`);
  }
};

/**
 * Replace an existing Mermaid diagram in the Word document
 * @param {Word.InlinePicture} picture - The existing inline picture to replace
 * @param {string} code - The new Mermaid diagram code
 * @returns {Promise<void>}
 */
export const replaceMermaidDiagram = async (picture, code) => {
  try {
    // Generate new diagram metadata
    const metadata = {
      code,
      timestamp: new Date().toISOString(),
      version: "1.0",
      originalTimestamp: picture.altTextDescription
        ? JSON.parse(picture.altTextDescription.replace(MERMAID_SIGNATURE_PREFIX, "")).timestamp
        : null,
    };

    // Create alt text description with new metadata
    const altTextDescription = `${MERMAID_SIGNATURE_PREFIX}${JSON.stringify(metadata)}`;

    await Word.run(async (context) => {
      // Load the picture to ensure we have the latest state
      picture.load("altTextDescription");
      await context.sync();

      // Generate new image
      const base64Image = await generateMermaidImage(code);

      // Replace the existing image
      picture.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.replace);

      // Get the new picture (the one that replaced the old one)
      const range = picture.getRange();
      const newPicture = range.inlinePictures.getFirst();
      newPicture.altTextDescription = altTextDescription;

      await context.sync();
    });
  } catch (error) {
    console.error("Failed to replace Mermaid diagram:", error);
    throw new Error(`Failed to replace diagram: ${error.message}`);
  }
};

/**
 * Generate a PNG image from Mermaid code
 * @param {string} code - The Mermaid diagram code
 * @returns {Promise<string>} - Base64 encoded PNG image
 */
const generateMermaidImage = async (code) => {
  try {
    // In a real implementation, this would:
    // 1. Use mermaid-cli or an online service to render the diagram
    // 2. Convert the SVG output to PNG
    // 3. Return as base64

    // For now, we'll return a placeholder image
    // This is a 1x1 transparent PNG in base64
    const placeholderImage =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

    console.log("Generating Mermaid diagram for code:", code);

    // Simulate async generation
    await new Promise((resolve) => setTimeout(resolve, 1000));

    return placeholderImage;
  } catch (error) {
    console.error("Failed to generate Mermaid image:", error);
    throw new Error(`Failed to generate diagram image: ${error.message}`);
  }
};

/**
 * Check if a picture contains Mermaid diagram metadata
 * @param {Word.InlinePicture} picture - The inline picture to check
 * @returns {boolean} - True if the picture contains Mermaid metadata
 */
export const isMermaidDiagram = (picture) => {
  try {
    const { altTextDescription } = picture;
    return altTextDescription && altTextDescription.startsWith(MERMAID_SIGNATURE_PREFIX);
  } catch {
    return false;
  }
};

/**
 * Extract Mermaid metadata from a picture
 * @param {Word.InlinePicture} picture - The inline picture containing Mermaid metadata
 * @returns {Object|null} - Parsed metadata or null if invalid
 */
export const extractMermaidMetadata = (picture) => {
  try {
    const { altTextDescription } = picture;
    if (!altTextDescription || !altTextDescription.startsWith(MERMAID_SIGNATURE_PREFIX)) {
      return null;
    }

    const jsonStr = altTextDescription.substring(MERMAID_SIGNATURE_PREFIX.length);
    return JSON.parse(jsonStr);
  } catch (error) {
    console.warn("Failed to extract Mermaid metadata:", error);
    return null;
  }
};
