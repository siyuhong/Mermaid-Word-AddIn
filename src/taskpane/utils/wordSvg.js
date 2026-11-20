/* global btoa, TextEncoder */

import { prepareSvgForRasterization } from "./svgSanitizer";
import { normalizeWordSvgStyles } from "./wordSvgStyles";
import { encodeMermaidSource, MERMAID_IMAGE_TITLE } from "./diagramMetadata";

const EMUS_PER_POINT = 12700;
let nextDrawingId = Date.now() % 2000000000;

function escapeXmlAttribute(value) {
  // Escape metadata for XML attributes, including whitespace that XML parsers
  // otherwise normalize. The source itself uses the shared versioned encoding.
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    .replace(/\r/g, "&#13;")
    .replace(/\n/g, "&#10;")
    .replace(/\t/g, "&#9;");
}

function toUtf8Base64(text) {
  const bytes = new TextEncoder().encode(text);
  const chunks = [];
  for (let offset = 0; offset < bytes.length; offset += 8192) {
    chunks.push(String.fromCharCode(...bytes.subarray(offset, offset + 8192)));
  }
  return btoa(chunks.join(""));
}

export function prepareSvgForWord(svg) {
  // Share the self-contained SVG cleanup with PNG export, retaining vector
  // geometry and text. Word does not support Mermaid's HTML foreignObjects.
  const cleaned = prepareSvgForRasterization(svg);
  if (/<foreignObject\b/i.test(cleaned)) {
    throw new Error("SVG export still contains HTML labels. Render the diagram for export first.");
  }
  // Word can hide solid edges when Mermaid expresses their dash array as 0.
  // Preserve real dashed lines, changing only arrays made entirely of zeroes.
  return normalizeWordSvgStyles(cleaned).replace(
    /stroke-dasharray\s*:\s*((?:0(?:\.0+)?(?:px)?[\s,]*)+)(?=[;}"!])/gi,
    "stroke-dasharray:none"
  );
}

export function createSvgPictureOoxml(svg, pngBase64, displaySize, mermaidCode) {
  const width = Math.max(1, Math.round(displaySize.width * EMUS_PER_POINT));
  const height = Math.max(1, Math.round(displaySize.height * EMUS_PER_POINT));
  const title = escapeXmlAttribute(MERMAID_IMAGE_TITLE);
  const source = escapeXmlAttribute(encodeMermaidSource(mermaidCode));
  const svgBase64 = toUtf8Base64(prepareSvgForWord(svg));
  const drawingId = ++nextDrawingId;

  // Flat OPC lets Word import the vector, fallback and source metadata in one
  // replace operation. The PNG is the main blip; SVG-capable Word uses svgBlip.
  // https://learn.microsoft.com/openspecs/office_standards/ms-odrawxml/7e1f1524-1569-4aa2-a6c9-aab2d855bd48
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <w:body><w:p><w:r><w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="${width}" cy="${height}"/>
            <wp:docPr id="${drawingId}" name="Mermaid Diagram" title="${title}" descr="${source}"/>
            <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:nvPicPr>
                  <pic:cNvPr id="0" name="mermaid.svg" title="${title}" descr="${source}"/>
                  <pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="rIdPng">
                    <a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
                      <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/>
                    </a:ext></a:extLst>
                  </a:blip>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing></w:r></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/mermaid.png"/>
        <Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/mermaid.svg"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/media/mermaid.png" pkg:contentType="image/png" pkg:compression="store">
    <pkg:binaryData>${pngBase64}</pkg:binaryData>
  </pkg:part>
  <pkg:part pkg:name="/word/media/mermaid.svg" pkg:contentType="image/svg+xml" pkg:compression="store">
    <pkg:binaryData>${svgBase64}</pkg:binaryData>
  </pkg:part>
</pkg:package>`;
}
