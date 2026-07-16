// Pure layout math shared by the renderers (which draw the boxes this
// computes) and the validator (which checks them before anything is drawn).
// Keeping this in one place means the renderer and the validator can never
// silently drift apart.

export const SLIDE_W = 13.33, SLIDE_H = 7.5;

export const WORKFLOW_LAYOUT = {
  LANE_X: 0.4, LANE_W: 12.5,
  LANE_START_Y: 1.15, LANE_GAP: 0.15, DEFAULT_LANE_H: 1.75,
  COL_X0: 0.6, COL_WIDTH: 2.0,
  NODE_W: 1.8, NODE_H: 0.95, NODE_Y_OFFSET: 0.35,
};

export function computeWorkflowLayout(raw) {
  const L = WORKFLOW_LAYOUT;
  let y = L.LANE_START_Y;
  const laneBoxes = {};
  for (const lane of raw.lanes) {
    const h = lane.height || L.DEFAULT_LANE_H;
    laneBoxes[lane.id] = { x: L.LANE_X, y, w: L.LANE_W, h };
    y += h + L.LANE_GAP;
  }
  const cardsY = y + 0.1;

  const usedCols = {};
  for (const n of raw.nodes) {
    if (n.col !== undefined) (usedCols[n.lane] ||= new Set()).add(n.col);
  }
  const nodeBoxes = {};
  const nodeMeta = {};
  for (const n of raw.nodes) {
    const laneBox = laneBoxes[n.lane];
    if (!laneBox) throw new Error(`Node "${n.id}" references unknown lane "${n.lane}"`);
    let col = n.col;
    if (col === undefined) {
      const used = (usedCols[n.lane] ||= new Set());
      col = 0;
      while (used.has(col)) col++;
      used.add(col);
    }
    const w = n.width || L.NODE_W;
    const h = n.height || L.NODE_H;
    nodeBoxes[n.id] = { x: L.COL_X0 + col * L.COL_WIDTH, y: laneBox.y + L.NODE_Y_OFFSET, w, h };
    nodeMeta[n.id] = { lane: n.lane, col };
  }
  return { laneBoxes, nodeBoxes, nodeMeta, cardsY };
}

export const SEQUENCE_LAYOUT = {
  X0: 0.7, GAP: 2.2, BOX_W: 1.5, BOX_H: 0.7,
  TOP_Y: 1.1, LIFELINE_BOTTOM: 7.0,
  MSG_START_Y: 2.3, MSG_STEP: 0.55,
};

export function computeSequenceLayout(raw) {
  const L = SEQUENCE_LAYOUT;
  const participantBoxes = {};
  const centerX = {};
  raw.participants.forEach((p, i) => {
    const x = L.X0 + i * L.GAP;
    participantBoxes[p.id] = { x, y: L.TOP_Y, w: L.BOX_W, h: L.BOX_H };
    centerX[p.id] = x + L.BOX_W / 2;
  });
  const messageYs = raw.messages.map((_, i) => L.MSG_START_Y + i * L.MSG_STEP);
  // cards need ~1.75in of vertical room, so diagrams that use them get a
  // shorter lifeline (fewer messages fit) instead of overflowing the slide
  const lifelineBottom = raw.cards && raw.cards.length ? 5.3 : L.LIFELINE_BOTTOM;
  const cardsY = lifelineBottom + 0.15;
  return { participantBoxes, centerX, messageYs, lifelineBottom, cardsY };
}

const BOUNDARY_PAD = 0.3;

export function computeArchitectureLayout(raw) {
  const boxes = {};
  for (const c of raw.components) {
    const [x, y] = c.pos;
    const [w, h] = c.size || [1.8, 0.95];
    boxes[c.id] = { x, y, w, h };
  }

  const boundaryBoxes = (raw.boundaries || []).map((b) => {
    const wrapped = b.wraps.map((id) => {
      if (!boxes[id]) throw new Error(`Boundary "${b.label}" wraps unknown component "${id}"`);
      return boxes[id];
    });
    const minX = Math.min(...wrapped.map((n) => n.x)) - BOUNDARY_PAD;
    const minY = Math.min(...wrapped.map((n) => n.y)) - BOUNDARY_PAD;
    const maxX = Math.max(...wrapped.map((n) => n.x + n.w)) + BOUNDARY_PAD;
    const maxY = Math.max(...wrapped.map((n) => n.y + n.h)) + BOUNDARY_PAD;
    return { label: b.label, color: b.color, box: { x: minX, y: minY, w: maxX - minX, h: maxY - minY } };
  });

  return { boxes, boundaryBoxes };
}
