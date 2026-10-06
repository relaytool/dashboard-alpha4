(() => {
  "use strict";

  const ORDER = ["RECEIVED", "SENT", "DISCARD"];

  function canonicalMovement(value) {
    const movement = String(value || "").trim().toUpperCase();
    if (["RECEIVED","INBOUND","IN","RECEIVE"].includes(movement)) return "RECEIVED";
    if (["SENT","OUTBOUND","OUT","SEND"].includes(movement)) return "SENT";
    if (["DISCARDED","DISCARD"].includes(movement)) return "DISCARD";
    return movement || "OTHER";
  }

  function movementLabel(value) {
    const movement = canonicalMovement(value);
    return movement === "RECEIVED" ? "Received" : movement === "SENT" ? "Sent" : movement === "DISCARD" ? "Discarded" : String(value || "Other");
  }

  function movementClass(value) {
    const movement = canonicalMovement(value);
    return movement === "RECEIVED" ? "movement-received" : movement === "SENT" ? "movement-sent" : movement === "DISCARD" ? "movement-discard" : "movement-other";
  }

  function timeKey(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return `raw:${String(value || "").trim()}`;
    return `date:${date.getTime()}`;
  }

  function groupTransactions(transactions = []) {
    const groups = new Map();
    (transactions || []).forEach((item, index) => {
      const client = String(item?.client || "Unknown client").trim() || "Unknown client";
      const key = `${String(client).toLowerCase()}\u0000${timeKey(item?.timestamp)}`;
      if (!groups.has(key)) {
        groups.set(key, {
          key,
          client,
          timestamp: item?.timestamp || "",
          items: [],
          movements: { RECEIVED: [], SENT: [], DISCARD: [], OTHER: [] },
          totalLines: 0,
          totalQuantity: 0,
          image: "",
          comment: "",
          originalIndexes: []
        });
      }
      const group = groups.get(key);
      const movement = canonicalMovement(item?.movement);
      const quantity = Number(item?.quantity) || 0;
      group.items.push(item);
      group.originalIndexes.push(index);
      group.totalLines += 1;
      group.totalQuantity += quantity;
      if (!group.image && item?.image) group.image = item.image;
      if (!group.comment && item?.comment) group.comment = item.comment;
      if (!group.movements[movement]) group.movements[movement] = [];
      group.movements[movement].push({
        asset: String(item?.asset || "Unknown asset").trim() || "Unknown asset",
        quantity,
        raw: item
      });
    });

    return [...groups.values()].sort((a, b) => {
      const da = new Date(a.timestamp).getTime();
      const db = new Date(b.timestamp).getTime();
      if (Number.isFinite(db) && Number.isFinite(da) && db !== da) return db - da;
      if (Number.isFinite(db) !== Number.isFinite(da)) return Number.isFinite(db) ? -1 : 1;
      return a.client.localeCompare(b.client);
    });
  }

  function movementSections(group, options = {}) {
    const includeEmpty = Boolean(options.includeEmpty);
    const sections = [];
    const addSection = movement => {
      const items = Array.isArray(group?.movements?.[movement]) ? group.movements[movement] : [];
      if (!items.length && !includeEmpty) return;
      const totalQuantity = items.reduce((sum, item) => sum + (Number(item?.quantity) || 0), 0);
      sections.push({
        movement,
        key: movement,
        label: movementLabel(movement),
        items,
        totalQuantity
      });
    };
    ORDER.forEach(addSection);
    addSection("OTHER");
    return sections;
  }

  function aggregateAssets(items = []) {
    const totals = new Map();
    items.forEach(item => {
      const asset = String(item.asset || "Unknown asset").trim() || "Unknown asset";
      totals.set(asset, (totals.get(asset) || 0) + (Number(item.quantity) || 0));
    });
    return [...totals.entries()].map(([asset, quantity]) => ({asset, quantity})).sort((a,b) => a.asset.localeCompare(b.asset));
  }

  function cardClass(group) {
    const movements = movementSections(group).map(section => section.movement);
    if (movements.length === 1) return `movement-card-${String(movements[0] ?? "other").toLowerCase()}`;
    if (movements.length > 1) return "movement-card-mixed";
    return "movement-card-other";
  }

  function movementLabels(group) {
    return movementSections(group).map(section => movementLabel(section.movement));
  }

  function assetSummary(group) {
    const assets = aggregateAssets(group?.items || []);
    return assets.map(item => item.asset).join(", ");
  }

  function formatLineCount(group) {
    const count = Number(group?.totalLines || 0);
    return `${count.toLocaleString()} ledger line${count === 1 ? "" : "s"}`;
  }

  window.FM_TRANSACTION_DISPLAY = {
    canonicalMovement,
    movementLabel,
    movementClass,
    groupTransactions,
    movementSections,
    aggregateAssets,
    cardClass,
    movementLabels,
    assetSummary,
    formatLineCount
  };
})();
