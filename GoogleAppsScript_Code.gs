// ============================================================================
// SHIITAKE FARM OS v9 — GOOGLE SHEETS CENTRAL BACKEND SCRIPT (Code.gs)
// Includes LockService concurrency protection & automatic BatchCode deduplication
// ============================================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    'Batches': ['ID', 'BatchCode', 'Strain', 'Room', 'InitialBlocks', 'DiscardedBlocks', 'SoldRtfBlocks', 'TransferredBlocks', 'InoculationDate', 'SampleWeightKg', 'Stage', 'UpdatedBy', 'UpdatedAt'],
    'PolyBatches': ['ID', 'PolyBatchCode', 'IncubationBatchCode', 'PolyhouseRoom', 'BlocksTransferred', 'DiscardedBlocks', 'TransferDate', 'Status', 'LoggedBy', 'UpdatedAt'],
    'SoakLogs': ['ID', 'PolyBatchCode', 'TargetFlush', 'SoakDate', 'EstPinningDate', 'EstHarvestDate', 'MethodNotes', 'Status', 'LoggedBy', 'UpdatedAt'],
    'ClimateLogs': ['ID', 'Timestamp', 'Room', 'TempC', 'RHPct', 'Fogger', 'Ventilation', 'LoggedBy', 'CreatedAt'],
    'HarvestLogs': ['ID', 'Date', 'BatchCode', 'Flush', 'GradeAKg', 'GradeBKg', 'CullKg', 'TotalFreshKg', 'Notes', 'LoggedBy', 'CreatedAt'],
    'InventorySalesLedger': ['ID', 'Date', 'ActionType', 'ProductCategory', 'QtyKg', 'OutputQtyKg', 'RatePerKg', 'RevenueINR', 'ReferenceNotes', 'LoggedBy', 'CreatedAt']
  };
  for (let name in sheets) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(sheets[name]);
      sh.getRange(1, 1, 1, sheets[name].length).setFontWeight('bold').setBackground('#d9ead3');
    }
  }
}

function doGet(e) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const bRows = ss.getSheetByName('Batches').getDataRange().getValues().slice(1);
  const bMap = {};
  bRows.filter(r => r[0] && r[1]).forEach(r => {
    const code = String(r[1]).trim();
    bMap[code.toUpperCase()] = {
      id: String(r[0]), code: code, strain: String(r[2]), room: String(r[3]),
      initialBlocks: Number(r[4] || 500),
      discardedBlocks: Number(r[5] || 0),
      soldRtfBlocks: Number(r[6] || 0),
      transferredBlocks: Number(r[7] || 0),
      date: String(r[8]), sampleWt: Number(r[9]), stage: String(r[10]), by: String(r[11] || 'Admin')
    };
  });
  const batches = Object.keys(bMap).map(k => bMap[k]);

  const pbRows = ss.getSheetByName('PolyBatches').getDataRange().getValues().slice(1);
  const pbMap = {};
  pbRows.filter(r => r[0] && r[1]).forEach(r => {
    const polyCode = String(r[1]).trim();
    pbMap[polyCode.toUpperCase()] = {
      id: String(r[0]), polyCode: polyCode, incubationBatch: String(r[2]),
      room: String(r[3]), blocksTransferred: Number(r[4]), discardedBlocks: Number(r[5] || 0),
      transferDate: String(r[6]), status: String(r[7]), by: String(r[8] || '')
    };
  });
  const polyBatches = Object.keys(pbMap).map(k => pbMap[k]);

  const sRows = ss.getSheetByName('SoakLogs').getDataRange().getValues().slice(1);
  const soakLogs = sRows.filter(r => r[0]).map(r => ({
    id: String(r[0]), polyCode: String(r[1]), flush: String(r[2]), soakDate: String(r[3]),
    estPinDate: String(r[4]), estHarvestDate: String(r[5]), notes: String(r[6]), status: String(r[7]), by: String(r[8] || '')
  }));

  const cRows = ss.getSheetByName('ClimateLogs').getDataRange().getValues().slice(1);
  const climateLogs = cRows.filter(r => r[0]).map(r => ({
    id: String(r[0]), time: String(r[1]), room: String(r[2]), temp: Number(r[3]),
    rh: Number(r[4]), fogger: String(r[5]), vent: String(r[6]), by: String(r[7] || '')
  })).reverse();

  const hRows = ss.getSheetByName('HarvestLogs').getDataRange().getValues().slice(1);
  const harvests = hRows.filter(r => r[0]).map(r => ({
    id: String(r[0]), date: String(r[1]), batch: String(r[2]), flush: String(r[3]),
    gradeA: Number(r[4]), gradeB: Number(r[5]), cull: Number(r[6]),
    notes: String(r[8] || ''), by: String(r[9] || '')
  })).reverse();

  const invRows = ss.getSheetByName('InventorySalesLedger').getDataRange().getValues().slice(1);
  const salesLedger = invRows.filter(r => r[0]).map(r => ({
    id: String(r[0]), date: String(r[1]), actionType: String(r[2]), product: String(r[3]),
    qtyKg: Number(r[4]), outputKg: Number(r[5]), rate: Number(r[6]), revenue: Number(r[7]),
    notes: String(r[8] || ''), by: String(r[9] || '')
  })).reverse();

  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok', batches, polyBatches, soakLogs, climateLogs, harvests, salesLedger
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    setupSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const payload = JSON.parse(e.postData.contents);
    const now = new Date().toISOString();

    if (payload.action === 'SYNC_ALL') {
      const bSheet = ss.getSheetByName('Batches');
      bSheet.clearContents();
      bSheet.appendRow(['ID', 'BatchCode', 'Strain', 'Room', 'InitialBlocks', 'DiscardedBlocks', 'SoldRtfBlocks', 'TransferredBlocks', 'InoculationDate', 'SampleWeightKg', 'Stage', 'UpdatedBy', 'UpdatedAt']);
      const seenB = {};
      (payload.state.batches || []).forEach(b => {
        const key = String(b.code || '').trim().toUpperCase();
        if (!key || seenB[key]) return;
        seenB[key] = true;
        bSheet.appendRow([b.id, String(b.code).trim(), b.strain, b.room, b.initialBlocks, b.discardedBlocks || 0, b.soldRtfBlocks || 0, b.transferredBlocks || 0, b.date, b.sampleWt, b.stage, b.by || 'Admin', now]);
      });

      const pbSheet = ss.getSheetByName('PolyBatches');
      pbSheet.clearContents();
      pbSheet.appendRow(['ID', 'PolyBatchCode', 'IncubationBatchCode', 'PolyhouseRoom', 'BlocksTransferred', 'DiscardedBlocks', 'TransferDate', 'Status', 'LoggedBy', 'UpdatedAt']);
      const seenPB = {};
      (payload.state.polyBatches || []).forEach(pb => {
        const key = String(pb.polyCode || '').trim().toUpperCase();
        if (!key || seenPB[key]) return;
        seenPB[key] = true;
        pbSheet.appendRow([pb.id, String(pb.polyCode).trim(), pb.incubationBatch, pb.room, pb.blocksTransferred, pb.discardedBlocks || 0, pb.transferDate, pb.status, pb.by || '', now]);
      });

      const sSheet = ss.getSheetByName('SoakLogs');
      sSheet.clearContents();
      sSheet.appendRow(['ID', 'PolyBatchCode', 'TargetFlush', 'SoakDate', 'EstPinningDate', 'EstHarvestDate', 'MethodNotes', 'Status', 'LoggedBy', 'UpdatedAt']);
      (payload.state.soakLogs || []).forEach(s => {
        sSheet.appendRow([s.id, s.polyCode, s.flush, s.soakDate, s.estPinDate, s.estHarvestDate, s.notes, s.status, s.by || '', now]);
      });

      const cSheet = ss.getSheetByName('ClimateLogs');
      cSheet.clearContents();
      cSheet.appendRow(['ID', 'Timestamp', 'Room', 'TempC', 'RHPct', 'Fogger', 'Ventilation', 'LoggedBy', 'CreatedAt']);
      (payload.state.climateLogs || []).slice().reverse().forEach(c => {
        cSheet.appendRow([c.id, c.time, c.room, c.temp, c.rh, c.fogger, c.vent, c.by || '', now]);
      });

      const hSheet = ss.getSheetByName('HarvestLogs');
      hSheet.clearContents();
      hSheet.appendRow(['ID', 'Date', 'BatchCode', 'Flush', 'GradeAKg', 'GradeBKg', 'CullKg', 'TotalFreshKg', 'Notes', 'LoggedBy', 'CreatedAt']);
      (payload.state.harvests || []).slice().reverse().forEach(h => {
        const tot = Number(h.gradeA) + Number(h.gradeB) + Number(h.cull);
        hSheet.appendRow([h.id, h.date, h.batch, h.flush, h.gradeA, h.gradeB, h.cull, tot, h.notes || '', h.by || '', now]);
      });

      const invSheet = ss.getSheetByName('InventorySalesLedger');
      invSheet.clearContents();
      invSheet.appendRow(['ID', 'Date', 'ActionType', 'ProductCategory', 'QtyKg', 'OutputQtyKg', 'RatePerKg', 'RevenueINR', 'ReferenceNotes', 'LoggedBy', 'CreatedAt']);
      (payload.state.salesLedger || []).slice().reverse().forEach(tx => {
        invSheet.appendRow([tx.id, tx.date, tx.actionType, tx.product, tx.qtyKg, tx.outputKg || 0, tx.rate || 0, tx.revenue || 0, tx.notes || '', tx.by || '', now]);
      });
    } else if (payload.action === 'ADD_CLIMATE') {
      const c = payload.item;
      ss.getSheetByName('ClimateLogs').appendRow([c.id, c.time, c.room, c.temp, c.rh, c.fogger, c.vent, c.by, now]);
    } else if (payload.action === 'ADD_HARVEST') {
      const h = payload.item;
      const tot = Number(h.gradeA) + Number(h.gradeB) + Number(h.cull);
      ss.getSheetByName('HarvestLogs').appendRow([h.id, h.date, h.batch, h.flush, h.gradeA, h.gradeB, h.cull, tot, h.notes || '', h.by, now]);
    } else if (payload.action === 'ADD_SALES_TX') {
      const tx = payload.item;
      ss.getSheetByName('InventorySalesLedger').appendRow([tx.id, tx.date, tx.actionType, tx.product, tx.qtyKg, tx.outputKg || 0, tx.rate || 0, tx.revenue || 0, tx.notes || '', tx.by, now]);
    } else if (payload.action === 'ADD_SOAK_LOG') {
      const s = payload.item;
      ss.getSheetByName('SoakLogs').appendRow([s.id, s.polyCode, s.flush, s.soakDate, s.estPinDate, s.estHarvestDate, s.notes, s.status, s.by, now]);
    } else if (payload.action === 'ADD_POLY_BATCH') {
      const pb = payload.item;
      ss.getSheetByName('PolyBatches').appendRow([pb.id, pb.polyCode, pb.incubationBatch, pb.room, pb.blocksTransferred, pb.discardedBlocks || 0, pb.transferDate, pb.status, pb.by, now]);
    } else if (payload.action === 'UPDATE_POLY_DISCARDED') {
      const pbSheet = ss.getSheetByName('PolyBatches');
      const data = pbSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          pbSheet.getRange(i + 1, 6).setValue(payload.discardedBlocks || 0);
          pbSheet.getRange(i + 1, 9).setValue(payload.by || '');
          pbSheet.getRange(i + 1, 10).setValue(now);
          break;
        }
      }
    } else if (payload.action === 'UPDATE_BATCH_DISPOSITION') {
      const bSheet = ss.getSheetByName('Batches');
      const data = bSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          if (payload.discardedBlocks !== undefined) bSheet.getRange(i + 1, 6).setValue(payload.discardedBlocks);
          if (payload.soldRtfBlocks !== undefined) bSheet.getRange(i + 1, 7).setValue(payload.soldRtfBlocks);
          if (payload.transferredBlocks !== undefined) bSheet.getRange(i + 1, 8).setValue(payload.transferredBlocks);
          if (payload.sampleWt !== undefined) bSheet.getRange(i + 1, 10).setValue(payload.sampleWt);
          if (payload.stage !== undefined) bSheet.getRange(i + 1, 11).setValue(payload.stage);
          bSheet.getRange(i + 1, 12).setValue(payload.by || '');
          bSheet.getRange(i + 1, 13).setValue(now);
          break;
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', syncedAt: now }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
