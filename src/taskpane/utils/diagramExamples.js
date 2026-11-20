import { diagramData } from "@mermaid-js/examples";

// Use the same default examples and ordering as Mermaid Live Editor.
// The package is bundled locally; loading a sample never needs a network call.
// ZenUML source: mermaid-js/mermaid-live-editor, src/lib/components/Preset.svelte.
// The upstream MIT copyright notice is preserved in LICENSE.
const PRIMARY_DIAGRAM_IDS = [
  "flowchart-v2",
  "classDiagram",
  "sequence",
  "er",
  "stateDiagram",
  "mindmap",
];

const officialExamples = diagramData
  .filter((diagram) => diagram.name && diagram.examples?.length)
  .map((diagram) => {
    const example = diagram.examples.find((item) => item.isDefault) || diagram.examples[0];
    return {
      id: diagram.id,
      label: diagram.name.replace(/ (Diagram|Chart|Graph)/, ""),
      title: example.title,
      code: example.code,
    };
  });

// Mermaid Live provides ZenUML separately from @mermaid-js/examples. This is
// also the diagram encoded in the user's mermaid-live.nodejs.cn share link.
const zenumlExample = {
  id: "zenuml",
  label: "ZenUML",
  title: "Order Service",
  code: `zenuml
    title Order Service
    @Actor Client #FFEBE6
    @Boundary OrderController #0747A6
    @EC2 <<BFF>> OrderService #E3FCEF
    group BusinessService {
      @Lambda PurchaseService
      @AzureFunction InvoiceService
    }

    @Starter(Client)
    // \`POST /orders\`
    OrderController.post(payload) {
      OrderService.create(payload) {
        order = new Order(payload)
        if(order != null) {
          par {
            PurchaseService.createPO(order)
            InvoiceService.createInvoice(order)
          }
        }
      }
    }
    `,
};

export const DIAGRAM_EXAMPLES = [...officialExamples, zenumlExample].sort((a, b) => {
  const priority = (id) => {
    const index = PRIMARY_DIAGRAM_IDS.indexOf(id);
    return index < 0 ? PRIMARY_DIAGRAM_IDS.length : index;
  };
  return priority(a.id) - priority(b.id) || a.label.localeCompare(b.label, "en");
});
