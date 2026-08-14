/**
 * VB6 event bindings.
 *
 * VB6 wires an event to its handler by NAME: a procedure called
 * `<producer>_<Event>` in the same module is what runs when `<producer>`
 * raises `<Event>`. Nothing in the source references the handler, so static
 * extraction leaves the flow broken exactly where the interesting part is —
 * a click, a timer tick, a barcode read.
 *
 * This pass reconnects it, producer → handler, so a call-graph traversal walks
 * the flow the way it happens at run time:
 *
 *     ReadBarcode --references(raises_event)--> ItemRead
 *                 --calls(withevents)--------> mReader_ItemRead
 *
 * The binding is a naming convention, not a parsed fact, so every edge it
 * creates is marked `provenance: 'heuristic'` with the wiring detail in
 * metadata (prompt §14) — it is never presented as a static call.
 *
 * A handler only binds when its producer really exists in the same file, which
 * is what keeps `IWorker_Run` (an `Implements` method, not an event handler)
 * and two controls sharing an event name from being wired to each other.
 */
import { Node, Edge } from '../types';
import { QueryBuilder } from '../db/queries';

export type Vb6EventBinding = 'control_event' | 'form_event' | 'withevents' | 'ocx_event';

/** The producer half of a `<producer>_<Event>` handler name. */
interface Producer {
  node: Node;
  binding: Vb6EventBinding;
}

/**
 * Create the event-binding edges for every indexed VB6 file.
 * Returns the number of edges added.
 */
export function synthesizeVb6EventBindings(queries: QueryBuilder): number {
  const methods = queries.getNodesByKind('method').filter((n) => n.language === 'vb6' && n.name.includes('_'));
  if (methods.length === 0) return 0;

  const nodesByFile = new Map<string, Node[]>();
  const fileNodes = (filePath: string): Node[] => {
    let hit = nodesByFile.get(filePath);
    if (!hit) {
      hit = queries.getNodesByFile(filePath);
      nodesByFile.set(filePath, hit);
    }
    return hit;
  };

  const edges: Edge[] = [];
  const seen = new Set<string>();

  for (const handler of methods) {
    const siblings = fileNodes(handler.filePath);
    const found = findProducer(handler, siblings);
    if (!found) continue;
    const { producer, eventName } = found;

    addEdge(edges, seen, producer.node.id, handler.id, producer.binding, eventName, producer.node);

    // For WithEvents the producing class is indexed too, so the Event
    // declaration itself can reach the handler. That is the hop that makes the
    // whole RaiseEvent → handler path traversable.
    if (producer.binding === 'withevents') {
      const typeName = producer.node.returnType;
      if (!typeName) continue;
      for (const type of queries.getNodesByLowerName(typeName.toLowerCase())) {
        if (type.language !== 'vb6' || type.kind !== 'class') continue;
        const declaration = fileNodes(type.filePath).find(
          (n) => n.name.toLowerCase() === eventName.toLowerCase() && n.decorators?.includes('vb6:event')
        );
        if (declaration) addEdge(edges, seen, declaration.id, handler.id, 'withevents', eventName, declaration);
      }
    }
  }

  if (edges.length > 0) queries.insertEdges(edges);
  return edges.length;
}

/**
 * Split `<producer>_<Event>` and find what raises it.
 *
 * A name may contain several underscores (`my_grid_Click`), so every split is
 * tried and the one whose producer actually exists wins. No producer, no edge.
 */
function findProducer(handler: Node, siblings: Node[]): { producer: Producer; eventName: string } | null {
  const name = handler.name;
  for (let i = name.indexOf('_'); i !== -1; i = name.indexOf('_', i + 1)) {
    const prefix = name.slice(0, i);
    const eventName = name.slice(i + 1);
    if (prefix === '' || eventName === '') continue;

    // The form or UserControl itself: `Form_Load`, `UserControl_Initialize`.
    const lower = prefix.toLowerCase();
    if (lower === 'form' || lower === 'usercontrol' || lower === 'mdiform') {
      const container = siblings.find(
        (n) =>
          n.kind === 'class' &&
          (n.decorators?.includes('vb6:form') || n.decorators?.includes('vb6:usercontrol'))
      );
      if (container) return { producer: { node: container, binding: 'form_event' }, eventName };
      continue;
    }

    const producer = siblings.find(
      (n) => n.name.toLowerCase() === lower && (n.kind === 'field' || n.kind === 'variable')
    );
    if (!producer) continue;

    if (producer.decorators?.includes('vb6:withevents')) {
      return { producer: { node: producer, binding: 'withevents' }, eventName };
    }
    if (producer.decorators?.includes('vb6:control')) {
      const binding: Vb6EventBinding = producer.decorators.includes('vb6:ocx') ? 'ocx_event' : 'control_event';
      return { producer: { node: producer, binding }, eventName };
    }
  }
  return null;
}

function addEdge(
  edges: Edge[],
  seen: Set<string>,
  source: string,
  target: string,
  binding: Vb6EventBinding,
  event: string,
  wiringSite: Node
): void {
  const key = `${source}\0${target}\0${binding}`;
  if (seen.has(key)) return;
  seen.add(key);
  edges.push({
    source,
    target,
    kind: 'calls',
    line: wiringSite.startLine,
    provenance: 'heuristic',
    metadata: {
      synthesizedBy: 'vb6-event-binding',
      binding,
      event,
      // Where the wiring is declared — the designer line for a control, the
      // `WithEvents` declaration for a field. Tooling shows this as the
      // registration site, which is the one place a reader can check a
      // synthesized binding for themselves.
      registeredAt: `${wiringSite.filePath}:${wiringSite.startLine}`,
    },
  });
}
