import { useState, useEffect, useCallback, useRef } from "react";

/* global Office, Word, console, setTimeout */

// Constants for Mermaid diagram identification
const MERMAID_SIGNATURE_PREFIX = "mermaid-diagram:";

/**
 * Custom hook to handle Word document selection changes and detect Mermaid diagrams
 * @param {Function} onDiagramDetected - Callback when a Mermaid diagram is detected
 * @param {Function} onDiagramCleared - Callback when selection is cleared or non-Mermaid
 * @returns {Object} - Selection state and loading status
 */
export const useMermaidSelection = (onDiagramDetected, onDiagramCleared) => {
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const handlerRef = useRef(null);
  const isProcessingRef = useRef(false);

  /**
   * Parse alt text description to extract Mermaid diagram metadata
   * @param {string} altTextDescription - The alt text description from the inline picture
   * @returns {Object|null} - Parsed metadata or null if not a Mermaid diagram
   */
  const parseMermaidMetadata = useCallback((altTextDescription) => {
    try {
      if (!altTextDescription || !altTextDescription.startsWith(MERMAID_SIGNATURE_PREFIX)) {
        return null;
      }

      const jsonStr = altTextDescription.substring(MERMAID_SIGNATURE_PREFIX.length);
      const metadata = JSON.parse(jsonStr);

      // Validate required fields
      if (!metadata.code || !metadata.timestamp) {
        return null;
      }

      return metadata;
    } catch (parseError) {
      console.warn("Failed to parse Mermaid metadata:", parseError);
      return null;
    }
  }, []);

  /**
   * Process the current selection to detect Mermaid diagrams
   */
  const processSelection = useCallback(async () => {
    // Guard against concurrent processing
    if (isProcessingRef.current) {
      return;
    }

    try {
      isProcessingRef.current = true;
      setIsLoading(true);
      setError(null);

      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("inlinePictures");
        await context.sync();

        const inlinePictures = selection.inlinePictures;
        inlinePictures.load("items/altTextDescription");
        await context.sync();

        const pictures = inlinePictures.items;

        if (pictures.length === 0) {
          onDiagramCleared?.();
          return;
        }

        // Process the first inline picture (Word typically has one picture in selection)
        const picture = pictures[0];
        const metadata = parseMermaidMetadata(picture.altTextDescription);

        if (metadata) {
          onDiagramDetected?.(metadata, picture);
        } else {
          onDiagramCleared?.();
        }
      });
    } catch (err) {
      console.error("Error processing selection:", err);
      setError(err.message || "Failed to process selection");
      onDiagramCleared?.();
    } finally {
      setIsLoading(false);
      isProcessingRef.current = false;
    }
  }, [onDiagramDetected, onDiagramCleared, parseMermaidMetadata]);

  /**
   * Handle document selection change events
   */
  const handleSelectionChanged = useCallback(() => {
    // Debounce selection changes to avoid excessive processing
    setTimeout(() => {
      processSelection();
    }, 100);
  }, [processSelection]);

  /**
   * Set up selection change handler when Office is ready
   */
  useEffect(() => {
    const setupSelectionHandler = async () => {
      try {
        // Office is already loaded via CDN, check if it's ready
        if (typeof Office !== "undefined" && Office.onReady) {
          await Office.onReady();
        }

        // Remove existing handler if any
        if (handlerRef.current) {
          Office.context.document.addHandlerRemovedAsync(
            Office.EventType.DocumentSelectionChanged,
            handlerRef.current
          );
        }

        // Add new selection change handler
        handlerRef.current = handleSelectionChanged;
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          handlerRef.current
        );

        // Process initial selection
        processSelection();
      } catch (err) {
        console.error("Failed to set up selection handler:", err);
        setError(err.message || "Failed to initialize selection monitoring");
      }
    };

    setupSelectionHandler();

    // Cleanup function
    return () => {
      if (handlerRef.current) {
        try {
          Office.context.document.removeHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            handlerRef.current
          );
        } catch (err) {
          console.warn("Failed to remove selection handler:", err);
        }
        handlerRef.current = null;
      }
    };
  }, [handleSelectionChanged, processSelection]);

  return {
    isLoading,
    error,
    processSelection, // Expose for manual refresh if needed
  };
};
