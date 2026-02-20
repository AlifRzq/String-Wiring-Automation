/**
 * STRING_PLAN(total, minLen, maxLen)
 *
 * Returns a 1-row horizontal array of string lengths that sums EXACTLY to total.
 *
 * Rules:
 * - Allowed lengths: minLen..maxLen
 * - No singleton: any used length cannot appear exactly once
 * - Priority: larger lengths preferred (maximize maxLen, then maxLen-1, ...)
 *
 * Examples:
 * =STRING_PLAN(A1,16,18)
 * =STRING_PLAN(A1,17,20)
 * =STRING_PLAN(A1,16,19)
 */
function STRING_PLAN(total, minLen, maxLen) {
  total = toInt_(total);
  minLen = (minLen === undefined || minLen === "") ? 16 : toInt_(minLen);
  maxLen = (maxLen === undefined || maxLen === "") ? 18 : toInt_(maxLen);

  if (!isFinite(total) || total <= 0) return [[""]];
  if (minLen > maxLen) return [["IMPOSSIBLE"]];

  const lens = [];
  for (let L = maxLen; L >= minLen; L--) lens.push(L);

  const maxCounts = {};
  lens.forEach(L => maxCounts[L] = Math.floor(total / L));

  let best = null;

  function dfs(idx, sum, counts) {
    if (sum > total) return;
    if (idx === lens.length) {
      if (sum !== total) return;

      // no-singleton rule
      for (const L of lens) {
        if (counts[L] === 1) return;
      }

      const nStrings = Object.values(counts).reduce((a,b)=>a+b,0);

      // score: prefer larger strings, then fewer total strings
      let score = nStrings;
      let weight = 1e9;
      for (const L of lens) {
        score -= counts[L] * weight;
        weight /= 1000;
      }

      if (!best || score < best.score) {
        best = {counts: {...counts}, score};
      }
      return;
    }

    const L = lens[idx];
    for (let k = 0; k <= maxCounts[L]; k++) {
      counts[L] = k;
      dfs(idx + 1, sum + k * L, counts);
    }
    counts[L] = 0;
  }

  dfs(0, 0, {});

  if (!best) return [["IMPOSSIBLE"]];

  const out = [];
  for (const L of lens) {
    for (let i = 0; i < best.counts[L]; i++) out.push(L);
  }

  return [out];
}

function toInt_(x) {
  if (typeof x === "string") x = x.replace(",", ".");
  const n = Number(x);
  return Math.trunc(n);
}

/***************
 * PV Tools v5.4
 * - Smart Allocate: if allocate fails, auto rebalance (donate from donors) + repack donors + retry
 * - ALLOC_LOG has State column: ACTIVE / MOVED / REMOVED
 * - Used-set counts only ACTIVE, so moved strings become available again
 * - Repack donor inverter to avoid holes & match terminal pattern for new NstringsUsed
 * - Supports InvType: 150, 100, 30-50
 * - Hardcoded output stride = 3 rows per MPPT (template fixed)
 * - Orientation write is MERGE-SAFE: write only to MPPT top row
 ***************/

const DCAC_TOL = 0.06;
const STRIDE_ROWS_PER_MPPT = 3;

const COL_NSTRINGS_USED = 'NstringsUsed';
const COL_VMPDIFF_MAX = 'VmpDiffMax';

const LOG_HEADERS = ['Timestamp','InvID','Building','Orientation','StringLength','SourceKey','TargetCell','State'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PV Tools')
    .addItem('Allocate Next Inverter (Smart)', 'allocateNextInverterSmart')
    .addItem('Allocate Next Inverter (Normal)', 'allocateNextInverter')
    .addSeparator()
    .addItem('Smart Rebalance (for Pending)', 'smartRebalanceForPending')
    .addItem('Repack All Allocated Inverters', 'repackAllAllocatedInverters')
    .addSeparator()
    .addItem('Audit & Rebalance (Fill Empties, Finish if clean)', 'auditAndRebalance')
    .addSeparator()
    .addItem('Reset Allocation Log (DANGER)', 'resetAllocLog')
    .addSeparator()
    .addItem('RESET ALL (Clear outputs + log + statuses)', 'resetAllAllocations')
    .addToUi();
}

/* =========================================================
   Allocate - Normal
   ========================================================= */

function allocateNextInverter() {
  const ss = SpreadsheetApp.getActive();
  const result = allocateNextInverterInternal_(ss, {silent:false});
  if (!result.ok) {
    SpreadsheetApp.getUi().alert(result.message);
  }
}

/* =========================================================
   Allocate - Smart (Auto rebalance then retry)
   ========================================================= */

function allocateNextInverterSmart() {
  const ss = SpreadsheetApp.getActive();

  // Try normal first
  let r1 = allocateNextInverterInternal_(ss, {silent:true});
  if (r1.ok) {
    SpreadsheetApp.getUi().alert(r1.message);
    return;
  }

  // If failed, attempt smart rebalance
  const reb = smartRebalanceForPendingInternal_(ss, {maxMoves: 12, silent:true});
  if (reb.ok) {
    // Retry allocation
    const r2 = allocateNextInverterInternal_(ss, {silent:true});
    if (r2.ok) {
      SpreadsheetApp.getUi().alert(`Smart Rebalance succeeded ✅\n\n${reb.message}\n\nThen Allocation succeeded ✅\n\n${r2.message}`);
      return;
    }
    SpreadsheetApp.getUi().alert(`Smart Rebalance ran, but allocation still failed.\n\nRebalance:\n${reb.message}\n\nAllocate error:\n${r2.message}`);
    return;
  }

  SpreadsheetApp.getUi().alert(
    `Allocation failed, and Smart Rebalance could not find a safe move.\n\nAllocate error:\n${r1.message}\n\nRebalance info:\n${reb.message}`
  );
}

/* =========================================================
   Allocate internal
   ========================================================= */

function allocateNextInverterInternal_(ss, {silent}) {
  try {
    const cfgSh = mustGetSheet_(ss, 'INV_CONFIG');
    ensureInvConfigColumns_(cfgSh, [COL_NSTRINGS_USED, COL_VMPDIFF_MAX]);

    const cfg = readConfigFirstPending_(cfgSh);
    if (!cfg) {
      return {ok:false, message:'No PENDING inverter found in INV_CONFIG.'};
    }

    const inv = configToInv_(cfg);
    const logSh = ensureAllocLog_(ss);

    const moduleWp = getNamedNumber_(ss, 'MODULE_WP', 630);
    const moduleVmp = getNamedNumber_(ss, 'MODULE_VMP_ACTUAL', 38.98);

    const invData = readStringInventory_(ss);
    if (invData.groups.length === 0) {
      return {ok:false, message:'STRING inventory not found / empty. Expected sheet "STRING_INVENTORY" OR any sheet with header "String 1".'};
    }

    const usedKeySet = readUsedSetActiveOnly_(logSh);
    const availGroups = buildAvailableGroups_(invData, usedKeySet);
    if (availGroups.length === 0) {
      return {ok:false, message:'No available strings left (all ACTIVE used).'};
    }

    const patternMap = readTerminalPatternMap_(ss, inv.invType);

    const plan = findBestPlan_({
      inv,
      moduleWp,
      moduleVmp,
      availGroups,
      terminalPatternMap: patternMap
    });

    if (!plan) {
      return {
        ok:false,
        message:
          `No feasible allocation found for InvType=${inv.invType}.\n` +
          `Constraints: DC/AC within ±${DCAC_TOL}, VoltageDiff <= ${inv.vmpDiffMax}.\n` +
          `Tip: run Smart Allocate to attempt rebalancing.`
      };
    }

    writeOutputByTerminals_(ss, inv, plan, moduleWp);
    fillOrientationTopRows_(ss, inv);
    appendAllocLogActive_(logSh, inv.invId, plan);

    const statusCol = colIndexByHeader_(cfgSh, 'Status');
    const nUsedCol = colIndexByHeader_(cfgSh, COL_NSTRINGS_USED);
    cfgSh.getRange(inv.cfgRow, statusCol).setValue('ALLOCATED');
    cfgSh.getRange(inv.cfgRow, nUsedCol).setValue(plan.totalStrings);

    return {
      ok:true,
      message:
        `Allocated ${plan.totalStrings} strings to ${inv.invId} (InvType ${inv.invType}).\n` +
        `DC/AC=${plan.dcac.toFixed(3)} (Ref ${inv.refDcac}, tol ±${DCAC_TOL})\n` +
        `VoltageDiff=${plan.vmpDiff.toFixed(1)} V (limit ${inv.vmpDiffMax})`
    };
  } catch (e) {
    return {ok:false, message:String(e && e.message ? e.message : e)};
  }
}

function resetAllAllocations() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const res = ui.alert(
    'RESET ALL',
    'Ini akan menghapus SEMUA output inverter, ALLOC_LOG, UNUSED_STRINGS, dan set status inverter jadi PENDING.\n\nLanjut?',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const cfgSh = mustGetSheet_(ss, 'INV_CONFIG');
  ensureInvConfigColumns_(cfgSh, [COL_NSTRINGS_USED, COL_VMPDIFF_MAX]);

  // 1) Clear ALL inverter outputs (NoPV + Orientation)
  const cfgRows = readAllConfigRows_(cfgSh);
  let cleared = 0;

  for (const r of cfgRows) {
    const invId = String(r.InvID || '').trim();
    const invType = String(r.InvType || '').trim();
    const aNo = String(r.OutputAnchor_NoPV || '').trim();
    const aOr = String(r.OutputAnchor_Orientation || '').trim();
    if (!invId || !invType || !aNo || !aOr) continue;
    if (!isSupportedInvType_(invType)) continue;

    // Build inv object minimal for clearing
    const inv = configToInvFromRowObj_(r);
    clearInvOutput_(ss, inv);
    cleared++;
  }

  // 2) Reset ALLOC_LOG
  const logSh = ss.getSheetByName('ALLOC_LOG');
  if (logSh) {
    logSh.clearContents();
    logSh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  }

  // 3) Clear UNUSED_STRINGS sheet
  const unSh = ss.getSheetByName('UNUSED_STRINGS');
  if (unSh) {
    unSh.clearContents();
    unSh.getRange(1,1,1,4).setValues([['Building','Orientation','StringLength','SourceKey']]);
  }

  // 4) Set NstringsUsed = blank and Status -> PENDING (for ALLOCATED/FINISHED)
  const lastRow = cfgSh.getLastRow();
  if (lastRow >= 2) {
    const statusCol = colIndexByHeader_(cfgSh, 'Status');
    const nUsedCol = colIndexByHeader_(cfgSh, COL_NSTRINGS_USED);

    const statusRange = cfgSh.getRange(2, statusCol, lastRow - 1, 1);
    const nUsedRange = cfgSh.getRange(2, nUsedCol, lastRow - 1, 1);

    const statusVals = statusRange.getValues();
    const nUsedVals = nUsedRange.getValues();

    for (let i = 0; i < statusVals.length; i++) {
      const st = String(statusVals[i][0] || '').trim().toUpperCase();
      if (st === 'ALLOCATED' || st === 'FINISHED') statusVals[i][0] = 'PENDING';
      nUsedVals[i][0] = '';
    }

    statusRange.setValues(statusVals);
    nUsedRange.setValues(nUsedVals);
  }

  ui.alert(
    `RESET ALL selesai ✅\n` +
    `- Cleared output inverter: ${cleared}\n` +
    `- ALLOC_LOG reset\n` +
    `- UNUSED_STRINGS cleared\n` +
    `- Status ALLOCATED/FINISHED -> PENDING\n` +
    `- NstringsUsed cleared`
  );
}

/* =========================================================
   Smart Rebalance
   ========================================================= */

function smartRebalanceForPending() {
  const ss = SpreadsheetApp.getActive();
  const res = smartRebalanceForPendingInternal_(ss, {maxMoves: 12, silent:false});
  SpreadsheetApp.getUi().alert(res.ok ? `Smart Rebalance ✅\n\n${res.message}` : `Smart Rebalance failed ❌\n\n${res.message}`);
}

function smartRebalanceForPendingInternal_(ss, {maxMoves, silent}) {
  try {
    const cfgSh = mustGetSheet_(ss, 'INV_CONFIG');
    ensureInvConfigColumns_(cfgSh, [COL_NSTRINGS_USED, COL_VMPDIFF_MAX]);

    const pending = readConfigFirstPending_(cfgSh);
    if (!pending) return {ok:false, message:'No PENDING inverter found.'};

    const invTarget = configToInv_(pending);

    const logSh = ensureAllocLog_(ss);

    const moduleWp = getNamedNumber_(ss, 'MODULE_WP', 630);
    const moduleVmp = getNamedNumber_(ss, 'MODULE_VMP_ACTUAL', 38.98);

    const invData = readStringInventory_(ss);
    if (invData.groups.length === 0) return {ok:false, message:'Inventory not found.'};

    // If target is already solvable with current unused, no need rebalance
    const usedActive = readUsedSetActiveOnly_(logSh);
    const availGroups0 = buildAvailableGroups_(invData, usedActive);
    const patternT = readTerminalPatternMap_(ss, invTarget.invType);
    const plan0 = findBestPlan_({inv:invTarget, moduleWp, moduleVmp, availGroups:availGroups0, terminalPatternMap:patternT});
    if (plan0) return {ok:true, message:'Target pending inverter is already feasible with current unused strings (no move needed).'}; // harmless

    // Build donor list (ALLOCATED/FINISHED allowed as donors; we’ll only touch ALLOCATED by default)
    const cfgRows = readAllConfigRows_(cfgSh);
    const donors = cfgRows
      .filter(r => ['ALLOCATED','FINISHED'].includes(String(r.Status||'').trim().toUpperCase()))
      .filter(r => isSupportedInvType_(r.InvType))
      .map(r => configToInvFromRowObj_(r))
      .filter(inv => inv.invId !== invTarget.invId);

    if (donors.length === 0) return {ok:false, message:'No donor inverters found (ALLOCATED/FINISHED).'}; 

    // We will attempt a series of moves. After each move: repack donor(s), then check if target feasible.
    let movesDone = 0;
    let notes = [];

    for (let step=0; step<maxMoves; step++) {
      // Recompute availability each loop (because moves change ACTIVE set)
      const usedNow = readUsedSetActiveOnly_(logSh);
      const availGroups = buildAvailableGroups_(invData, usedNow);

      const planTry = findBestPlan_({inv:invTarget, moduleWp, moduleVmp, availGroups, terminalPatternMap:patternT});
      if (planTry) {
        return {ok:true, message: `Target became feasible after ${movesDone} move(s).\n` + notes.join('\n')};
      }

      // Need to free more “good” strings for target. Try donate from a donor.
      const move = findOneSafeDonationMove_({
        ss, logSh, invData, moduleWp, moduleVmp,
        invTarget,
        donors
      });

      if (!move) {
        return {
          ok:false,
          message:
            `No safe donation move found.\n` +
            `Tried donors: ${donors.map(d=>d.invId).join(', ')}\n` +
            `Reason: donor constraints or no movable MPPT bundles left.\n` +
            (notes.length ? `\nProgress notes:\n${notes.join('\n')}` : '')
        };
      }

      // Apply donation
      applyDonationMove_(ss, logSh, move);
      notes.push(`Move ${step+1}: Donor ${move.donor.invId} -> freed ${move.items.length} string(s) (len=${move.len}, ${move.building}/${move.orientation}).`);

      // Repack donor to remove holes & align terminal pattern to new NstringsUsed
      const rep = repackInverterByLog_(ss, cfgSh, logSh, move.donor.invId, moduleWp, moduleVmp);
      notes.push(`Repack donor ${move.donor.invId}: ${rep}`);

      movesDone++;
    }

    return {ok:false, message:`Reached maxMoves=${maxMoves}, target still not feasible.\n` + notes.join('\n')};
  } catch (e) {
    return {ok:false, message:String(e && e.message ? e.message : e)};
  }
}

/**
 * Find ONE safe donation move:
 * - We only donate a FULL donor MPPT bundle (so donor never leaves MPPT=1 for Inv150).
 * - Bundle must be uniform (same building+orientation+len) by design.
 * - Donation "frees" ACTIVE strings back to unused pool by marking MOVED & clearing donor output cells.
 */
function findOneSafeDonationMove_({ss, logSh, invData, moduleWp, moduleVmp, invTarget, donors}) {
  // Heuristic: prefer donating strings that are likely helpful for target (higher len usually helps VmpTarget and DC/AC).
  // We just try donors in order of "most slack" first: donors with DC/AC comfortably above lower bound.
  const donorStats = donors.map(d => {
    const activeItems = readActiveItemsForInv_(logSh, d.invId);
    const pdc = activeItems.reduce((s,it)=>s + it.len*moduleWp, 0)/1000.0;
    const dcac = d.pacKw ? (pdc/d.pacKw) : 0;
    const slack = (dcac - (d.refDcac - DCAC_TOL)); // positive means can lose some DC and still within tol
    return {inv:d, activeItemsCount:activeItems.length, dcac, slack};
  }).sort((a,b)=>b.slack - a.slack);

  for (const ds of donorStats) {
    const donor = ds.inv;
    if (ds.activeItemsCount === 0) continue;

    // Build donor MPPT bundles from output + terminal mapping + log
    const bundles = buildDonorMpptBundles_(ss, logSh, donor);

    // Sort bundles: donate smaller first (less risk), then by len (helpful), then by slack fit
    bundles.sort((a,b)=>{
      if (a.items.length !== b.items.length) return a.items.length - b.items.length;
      if (a.len !== b.len) return b.len - a.len;
      return 0;
    });

    for (const b of bundles) {
      // Quick donor feasibility check if remove this bundle:
      const activeItems = readActiveItemsForInv_(logSh, donor.invId);
      const remain = activeItems.length - b.items.length;
      if (remain <= 0) continue;

      const pdcBefore = activeItems.reduce((s,it)=>s + it.len*moduleWp, 0)/1000.0;
      const pdcAfter = pdcBefore - (b.items.length*b.len*moduleWp)/1000.0;
      const dcacAfter = pdcAfter / donor.pacKw;
      if (Math.abs(dcacAfter - donor.refDcac) > DCAC_TOL) continue;

      // VoltageDiff donor after (approx by remaining lens):
      const lensAfter = activeItems
        .filter(it => !b.items.some(x => x.sourceKey === it.sourceKey))
        .map(it => it.len);

      if (lensAfter.length === 0) continue;
      const vmpDiffAfter = (Math.max(...lensAfter) - Math.min(...lensAfter)) * moduleVmp;
      if (donor.vmpDiffMax && vmpDiffAfter > donor.vmpDiffMax) continue;

      // Inv150 rule safety: since we remove full MPPT bundle, donor will not leave a MPPT with 1 string.
      // But repack will be applied anyway. Still, we must ensure remain can be repacked to a valid Nstrings pattern.
      const patternMap = readTerminalPatternMap_(ss, donor.invType);
      if (!patternMap.has(remain)) continue;

      // Great: found safe donation move
      return {
        donor,
        mppt: b.mppt,
        len: b.len,
        building: b.building,
        orientation: b.orientation,
        items: b.items, // each has sourceKey, targetCell
      };
    }
  }

  return null;
}

/**
 * Build MPPT bundles from donor output:
 * - Determine donor terminal mapping for current NstringsUsed (from INV_CONFIG)
 * - For each MPPT, collect ACTIVE strings via log.TargetCell matching filled NoPV cells
 * - Require uniform length within MPPT and same orientation/building (from log)
 */
function buildDonorMpptBundles_(ss, logSh, donor) {
  const cfgSh = mustGetSheet_(ss, 'INV_CONFIG');
  const rowObj = readConfigRowByInvId_(cfgSh, donor.invId);
  if (!rowObj) return [];

  const nUsed = int_(rowObj[COL_NSTRINGS_USED]);
  if (!nUsed) return [];

  const patternMap = readTerminalPatternMap_(ss, donor.invType);
  const terminals = patternMap.get(nUsed);
  if (!terminals) return [];

  const out = readInvOutputBlock_(ss, donor);

  // Map targetCell -> log item
  const logItems = readActiveLogItemsForInv_(logSh, donor.invId); // includes rowIndex
  const byTarget = new Map();
  for (const li of logItems) byTarget.set(li.targetCell, li);

  const mpptBuckets = new Map(); // mppt -> items[]
  for (const t of terminals) {
    const {mppt, input} = terminalToMpptInputFor_(donor, t);
    const idx = (mppt-1)*STRIDE_ROWS_PER_MPPT + (input-1);
    const len = int_(out.noPvVals[idx]);
    if (len <= 0) continue;

    const targetCell = out.targetA1[idx]; // Sheet!A1
    const li = byTarget.get(targetCell);
    if (!li) continue; // if log missing, skip (conservative)

    if (!mpptBuckets.has(mppt)) mpptBuckets.set(mppt, []);
    mpptBuckets.get(mppt).push(li);
  }

  const bundles = [];
  for (const [mppt, items] of mpptBuckets.entries()) {
    if (!items.length) continue;

    // Uniform len check
    const lens = items.map(x=>x.len);
    const len0 = lens[0];
    if (!lens.every(L => L === len0)) continue;

    // Uniform building+orientation check (from log)
    const b0 = items[0].building;
    const o0 = items[0].orientation;
    if (!items.every(x => x.building === b0 && x.orientation === o0)) continue;

    // Donation safety: for Inv150, MPPT bundle must be size 2 or 3 (never 1)
    const invType = normalizeInvType_(donor.invType);
    if (invType === '150' && items.length === 1) continue;

    bundles.push({mppt, len:len0, building:b0, orientation:o0, items});
  }

  return bundles;
}

/**
 * Apply donation move:
 * - Mark donor log rows MOVED
 * - Clear donor output cells (NoPV) for those target cells
 * - Clear orientation for MPPT if MPPT becomes empty
 */
function applyDonationMove_(ss, logSh, move) {
  // 1) Mark log MOVED
  markLogRowsState_(logSh, move.items.map(it=>it.rowIndex), 'MOVED');

  // 2) Clear donor table cells
  const donor = move.donor;
  const out = readInvOutputBlock_(ss, donor);

  const targetSet = new Set(move.items.map(it=>it.targetCell));

  // clear NoPV for those cells
  for (let i=0;i<out.targetA1.length;i++) {
    if (targetSet.has(out.targetA1[i])) {
      out.noPvVals[i] = '';
    }
  }

  // For merge-safe orientation: if MPPT becomes empty, clear top orientation cell
  const mpptBase = (move.mppt-1)*STRIDE_ROWS_PER_MPPT;
  let anyLeft = false;
  for (let j=0;j<donor.inputsPerMppt;j++) {
    if (int_(out.noPvVals[mpptBase+j]) > 0) { anyLeft = true; break; }
  }
  if (!anyLeft) {
    out.orVals[mpptBase] = '';
  }

  flushInvOutputBlock_(out);
  fillOrientationTopRows_(ss, donor); // ensure merge-safe tidy
}

/* =========================================================
   Repack donor inverter (by ACTIVE log items)
   ========================================================= */

function repackAllAllocatedInverters() {
  const ss = SpreadsheetApp.getActive();
  const cfgSh = mustGetSheet_(ss, 'INV_CONFIG');
  const logSh = ensureAllocLog_(ss);

  const moduleWp = getNamedNumber_(ss, 'MODULE_WP', 630);
  const moduleVmp = getNamedNumber_(ss, 'MODULE_VMP_ACTUAL', 38.98);

  const cfgRows = readAllConfigRows_(cfgSh);
  const invIds = cfgRows
    .filter(r => ['ALLOCATED','FINISHED'].includes(String(r.Status||'').trim().toUpperCase()))
    .map(r => String(r.InvID||'').trim())
    .filter(Boolean);

  let msgs = [];
  for (const id of invIds) {
    try {
      msgs.push(`${id}: ${repackInverterByLog_(ss, cfgSh, logSh, id, moduleWp, moduleVmp)}`);
    } catch (e) {
      msgs.push(`${id}: ERROR ${String(e && e.message ? e.message : e)}`);
    }
  }
  SpreadsheetApp.getUi().alert(`Repack done.\n\n${msgs.join('\n')}`);
}

function repackInverterByLog_(ss, cfgSh, logSh, invId, moduleWp, moduleVmp) {
  const rowObj = readConfigRowByInvId_(cfgSh, invId);
  if (!rowObj) return 'Config row not found.';
  if (!isSupportedInvType_(rowObj.InvType)) return 'Unsupported InvType.';

  const inv = configToInvFromRowObj_(rowObj);

  // Gather ACTIVE items for this inverter
  const active = readActiveLogItemsForInv_(logSh, invId); // {rowIndex, building, orientation, len, sourceKey, targetCell}
  const n = active.length;

  if (n === 0) {
    // Clear output block for neatness
    clearInvOutput_(ss, inv);
    rowObj[COL_NSTRINGS_USED] = 0;
    writeConfigCellByInvId_(cfgSh, invId, COL_NSTRINGS_USED, 0);
    return 'No ACTIVE strings; cleared output.';
  }

  // Must have terminal pattern for n
  const patternMap = readTerminalPatternMap_(ss, inv.invType);
  const terminals = patternMap.get(n);
  if (!terminals) return `No terminal pattern for Nstrings=${n} (cannot repack).`;

  // Build "inventory" from active items (grouped)
  const groups = groupItems_(active); // {building, orientation, strings:[{len, sourceKey}]}

  // Build mpptSlots from terminals
  const mpptSlots = new Array(inv.mpptUsed).fill(0);
  for (const t of terminals) {
    const {mppt} = terminalToMpptInputFor_(inv, t);
    mpptSlots[mppt-1]++;
  }
  if (mpptSlots.some(k => k > inv.inputsPerMppt)) return 'Pattern violates InputsPerMPPT (unexpected).';
  if (inv.invType === '150' && mpptSlots.some(k => k === 1)) return 'Pattern has MPPT=1 which is forbidden for Inv150.';

  // Re-solve allocation using ONLY these items
  const attempt = allocateMpptUniformLen_WithInvRules_({
    inv,
    mpptSlots,
    availGroups: groups
  });
  if (!attempt) return 'Cannot repack with current remaining strings (rule conflict).';

  // Check donor DC/AC and voltage diff (must remain within tol)
  const totalWp = attempt.lens.reduce((s,L)=>s + L*moduleWp, 0);
  const pdcKw = totalWp / 1000.0;
  const dcac = pdcKw / inv.pacKw;
  if (Math.abs(dcac - inv.refDcac) > DCAC_TOL) return `Repack would violate donor DC/AC (dcac=${dcac.toFixed(3)} ref=${inv.refDcac}).`;

  const vmpDiff = (attempt.maxLen - attempt.minLen) * moduleVmp;
  if (inv.vmpDiffMax && vmpDiff > inv.vmpDiffMax) return `Repack would violate donor VmpDiff (diff=${vmpDiff.toFixed(1)} limit=${inv.vmpDiffMax}).`;

  // Prepare plan object compatible with writer
  const plan = {
    totalStrings: n,
    terminals,
    assigned: attempt.assigned,
    dcac,
    vmpDiff
  };

  // Clear output then write in new order
  clearInvOutput_(ss, inv);
  writeOutputByTerminals_(ss, inv, plan, moduleWp);
  fillOrientationTopRows_(ss, inv);

  // Update log target cells for ACTIVE rows
  // We matched by sourceKey: each item carries sourceKey & new targetA1
  const mapping = new Map(); // sourceKey -> newTargetCell
  for (const block of plan.assigned) {
    if (!block) continue;
    for (const it of block.items) mapping.set(it.sourceKey, it.targetA1);
  }

  updateLogTargetCellsBySourceKey_(logSh, invId, mapping);

  // Update config NstringsUsed
  writeConfigCellByInvId_(cfgSh, invId, COL_NSTRINGS_USED, n);

  return `Repacked OK (N=${n}, DC/AC=${dcac.toFixed(3)}, VmpDiff=${vmpDiff.toFixed(1)}V).`;
}

/* =========================================================
   Audit & Rebalance (fill empties only)
   ========================================================= */

function auditAndRebalance() {
  const ss = SpreadsheetApp.getActive();
  const cfgSh = mustGetSheet_(ss, 'INV_CONFIG');
  ensureInvConfigColumns_(cfgSh, [COL_NSTRINGS_USED, COL_VMPDIFF_MAX]);

  const logSh = ensureAllocLog_(ss);

  const moduleWp = getNamedNumber_(ss, 'MODULE_WP', 630);
  const moduleVmp = getNamedNumber_(ss, 'MODULE_VMP_ACTUAL', 38.98);

  const invData = readStringInventory_(ss);
  if (invData.groups.length === 0) throw new Error('Inventory not found.');

  const usedBefore = readUsedSetActiveOnly_(logSh);
  const leftoversBefore = computeUnused_(invData, usedBefore);

  const cfgRows = readAllConfigRows_(cfgSh);

  const pool = buildUnusedPool_(leftoversBefore);
  let added = 0;

  for (const row of cfgRows) {
    const status = String(row.Status || '').trim().toUpperCase();
    if (status !== 'ALLOCATED') continue;

    const invType = String(row.InvType || '').trim();
    if (!isSupportedInvType_(invType)) continue;

    const nUsed = int_(row[COL_NSTRINGS_USED]);
    if (!nUsed) continue;

    const inv = configToInvFromRowObj_(row);

    const patternMap = readTerminalPatternMap_(ss, inv.invType);
    const terminals = patternMap.get(nUsed);
    if (!terminals) continue;

    const out = readInvOutputBlock_(ss, inv);
    const empties = findEmptyOnTerminals_(out, terminals, inv);
    if (empties.length === 0) continue;

    const mpptInfo = inferMpptInfoFromLogAndOutput_(logSh, inv.invId, out, inv);

    const lensNow = out.noPvVals.map(v => int_(v)).filter(x => x > 0);
    if (lensNow.length === 0) continue;

    const pdcNowKw = computePdcFromOutput_(out.noPvVals, moduleWp);

    for (const e of empties) {
      const info = mpptInfo[e.mppt - 1];
      if (!info || !info.building || !info.orientation || !info.len) continue;

      const addPdcKw = (info.len * moduleWp) / 1000.0;
      const dcacAfter = (pdcNowKw + addPdcKw) / inv.pacKw;
      if (Math.abs(dcacAfter - inv.refDcac) > DCAC_TOL) continue;

      const maxLenAfter = Math.max(...lensNow, info.len);
      const minLenAfter = Math.min(...lensNow, info.len);
      const vmpDiffAfter = (maxLenAfter - minLenAfter) * moduleVmp;
      if (inv.vmpDiffMax && vmpDiffAfter > inv.vmpDiffMax) continue;

      const pick = popUnused_(pool, info.building, info.orientation, info.len);
      if (!pick) continue;

      writeOneTerminal_(out, e, info.len, info.orientation);
      flushInvOutputBlock_(out);

      lensNow.push(info.len);

      appendOneAllocLogActive_(logSh, inv.invId, pick.building, pick.orientation, pick.len, pick.sourceKey, e.targetA1);
      added++;
    }

    fillOrientationTopRows_(ss, inv);
  }

  const usedAfter = readUsedSetActiveOnly_(logSh);
  const leftoversAfter = computeUnused_(invData, usedAfter);
  writeUnusedSheet_(ss, leftoversAfter);

  if (leftoversAfter.length === 0) {
    markAllAllocatedFinished_(cfgSh);
    SpreadsheetApp.getUi().alert(
      `Audit & Rebalance done ✅\n` +
      `Before leftovers: ${leftoversBefore.length}\n` +
      `Added during rebalance: ${added}\n` +
      `After leftovers: 0\n` +
      `All ALLOCATED set to FINISHED`
    );
  } else {
    SpreadsheetApp.getUi().alert(
      `Audit & Rebalance done ✅\n` +
      `Before leftovers: ${leftoversBefore.length}\n` +
      `Added during rebalance: ${added}\n` +
      `After leftovers: ${leftoversAfter.length}\n` +
      `See UNUSED_STRINGS sheet.`
    );
  }
}

/* =========================================================
   Solver (same as v5.3c)
   ========================================================= */

function findBestPlan_({inv, moduleWp, moduleVmp, availGroups, terminalPatternMap}) {
  const candidates = Array.from(terminalPatternMap.keys()).sort((a,b)=>a-b);
  let best = null;

  for (const totalStrings of candidates) {
    const terminals = terminalPatternMap.get(totalStrings);
    if (!terminals || terminals.length !== totalStrings) continue;

    const mpptSlots = new Array(inv.mpptUsed).fill(0);
    for (const t of terminals) {
      const {mppt} = terminalToMpptInputFor_(inv, t);
      if (mppt >= 1 && mppt <= inv.mpptUsed) mpptSlots[mppt-1]++;
    }

    if (mpptSlots.some(k => k > inv.inputsPerMppt)) continue;
    if (inv.invType === '150' && mpptSlots.some(k => k === 1)) continue;

    const attempt = allocateMpptUniformLen_WithInvRules_({
      inv,
      mpptSlots,
      availGroups
    });
    if (!attempt) continue;

    const totalWp = attempt.lens.reduce((s,L)=>s + L*moduleWp, 0);
    const pdcKw = totalWp / 1000.0;
    const dcac = pdcKw / inv.pacKw;
    if (Math.abs(dcac - inv.refDcac) > DCAC_TOL) continue;

    const vmpDiff = (attempt.maxLen - attempt.minLen) * moduleVmp;
    if (inv.vmpDiffMax && vmpDiff > inv.vmpDiffMax) continue;

    const vmpStrings = attempt.lens.map(L => L * moduleVmp);

    const score = scorePlan_({
      dcac, ref: inv.refDcac,
      orphanCount: attempt.orphanCount,
      vmpStrings,
      vmpMinGood: inv.vmpMinGood,
      vmpMinWorst: inv.vmpMinWorst,
      vmpDiff
    });

    const candidate = {
      totalStrings,
      terminals,
      mpptSlots,
      assigned: attempt.assigned,
      orphanCount: attempt.orphanCount,
      dcac,
      vmpDiff,
      score
    };

    if (!best || candidate.score < best.score) best = candidate;
  }

  return best;
}

function allocateMpptUniformLen_WithInvRules_({inv, mpptSlots, availGroups}) {
  const groups = availGroups.map(g => {
    const buckets = new Map();
    for (const it of g.strings) {
      if (!buckets.has(it.len)) buckets.set(it.len, []);
      buckets.get(it.len).push({len: it.len, sourceKey: it.sourceKey});
    }
    const lensDesc = Array.from(buckets.keys()).sort((a,b)=>b-a);
    return {building: g.building, orientation: g.orientation, buckets, lensDesc};
  });

  const order = mpptSlots
    .map((k, idx)=>({k, idx}))
    .filter(x=>x.k>0)
    .sort((a,b)=>b.k-a.k);

  const assigned = new Array(inv.mpptUsed).fill(null);
  const lensAll = [];
  let orphanCount = 0;

  for (const req of order) {
    const k = req.k;
    let bestChoice = null;

    // Pass 1: avoid leftover==1 ONLY for Inv150
    for (let gi=0; gi<groups.length; gi++) {
      const g = groups[gi];
      for (const len of g.lensDesc) {
        const bucket = g.buckets.get(len);
        if (!bucket || bucket.length < k) continue;

        const after = bucket.length - k;
        if (inv.invType === '150' && after === 1) continue;

        const metric = (len * 100) + bucket.length;
        if (!bestChoice || metric > bestChoice.metric) bestChoice = {gi, len, metric, after};
      }
    }

    // Pass 2: allow leftover==1 (Inv150 penalized but allowed)
    if (!bestChoice) {
      for (let gi=0; gi<groups.length; gi++) {
        const g = groups[gi];
        for (const len of g.lensDesc) {
          const bucket = g.buckets.get(len);
          if (!bucket || bucket.length < k) continue;

          const after = bucket.length - k;
          const penalty = (inv.invType === '150' && after === 1) ? 5000 : 0;
          const metric = (len * 100) + bucket.length - penalty;

          if (!bestChoice || metric > bestChoice.metric) bestChoice = {gi, len, metric, after};
        }
      }
    }

    if (!bestChoice) return null;

    const g = groups[bestChoice.gi];
    const bucket = g.buckets.get(bestChoice.len);

    const before = bucket.length;
    const taken = bucket.splice(0, k);
    const after = before - k;

    if (inv.invType === '150' && after === 1) orphanCount++;

    if (bucket.length === 0) {
      g.buckets.delete(bestChoice.len);
      g.lensDesc = g.lensDesc.filter(x => x !== bestChoice.len);
    }

    assigned[req.idx] = {building: g.building, orientation: g.orientation, items: taken};
    for (const it of taken) lensAll.push(it.len);
  }

  if (lensAll.length === 0) return null;

  return {
    assigned,
    lens: lensAll,
    maxLen: Math.max(...lensAll),
    minLen: Math.min(...lensAll),
    orphanCount
  };
}

function scorePlan_({dcac, ref, orphanCount, vmpStrings, vmpMinGood, vmpMinWorst, vmpDiff}) {
  const diff = dcac - ref;
  let pen_dcac = Math.abs(diff);
  if (diff < 0) pen_dcac *= 3;

  let pen_vmp = 0;
  for (const vmp of vmpStrings) {
    if (vmp >= vmpMinGood) continue;
    if (vmp >= vmpMinWorst) pen_vmp += 1 + (vmpMinGood - vmp)/25;
    else pen_vmp += 30 + (vmpMinWorst - vmp)/5;
  }

  const pen_orphan = orphanCount * 50;
  const pen_vmpDiff = vmpDiff / 2000;

  return 10*pen_dcac + pen_vmp + pen_orphan + pen_vmpDiff;
}

/* =========================================================
   Output write/read (stride=3) + MERGE-SAFE ORIENTATION
   ========================================================= */

function writeOutputByTerminals_(ss, inv, plan, moduleWp) {
  const {sheetName: shNo, row: rNo, col: cNo} = parseA1WithSheet_(inv.outNoPvA1, ss);
  const {sheetName: shOr, row: rOr, col: cOr} = parseA1WithSheet_(inv.outOriA1, ss);
  if (shNo !== shOr || rNo !== rOr) throw new Error('Anchor mismatch (NoPV vs Orientation).');

  const sh = mustGetSheet_(ss, shNo);

  const totalRows = inv.mpptUsed * STRIDE_ROWS_PER_MPPT;
  sh.getRange(rNo, cNo, totalRows, 1).clearContent();
  sh.getRange(rOr, cOr, totalRows, 1).clearContent();

  const curs = plan.assigned.map(b => ({block: b, idx: 0}));

  for (const t of plan.terminals) {
    const {mppt, input} = terminalToMpptInputFor_(inv, t);
    const linearRow = (mppt-1)*STRIDE_ROWS_PER_MPPT + (input-1);

    const cursor = curs[mppt-1];
    const block = cursor.block;
    if (!block) continue;
    const item = block.items[cursor.idx];
    if (!item) continue;
    cursor.idx++;

    sh.getRange(rNo + linearRow, cNo).setValue(item.len);

    // MERGE SAFE: orientation to MPPT top row only
    const mpptTopRow = (mppt-1)*STRIDE_ROWS_PER_MPPT;
    sh.getRange(rOr + mpptTopRow, cOr).setValue(block.orientation);

    item.targetA1 = `${sh.getName()}!${a1_(rNo + linearRow, cNo)}`;
  }

  plan.totalWp = plan.assigned
    .flatMap(b => (b ? b.items : []))
    .reduce((s, it) => s + it.len * moduleWp, 0);
}

function fillOrientationTopRows_(ss, inv) {
  const {sheetName: shOr, row: rOr, col: cOr} = parseA1WithSheet_(inv.outOriA1, ss);
  const sh = mustGetSheet_(ss, shOr);

  const totalRows = inv.mpptUsed * STRIDE_ROWS_PER_MPPT;
  const orRange = sh.getRange(rOr, cOr, totalRows, 1);
  const orVals = orRange.getValues().map(r => r[0]);

  for (let m=0; m<inv.mpptUsed; m++) {
    const base = m * STRIDE_ROWS_PER_MPPT;

    let ori = null;
    for (let i=0;i<STRIDE_ROWS_PER_MPPT;i++) {
      const v = orVals[base+i];
      if (v !== '' && v !== null && v !== undefined) { ori = v; break; }
    }
    if (ori === null) continue;

    orVals[base] = ori;
  }

  orRange.setValues(orVals.map(v => [v]));
}

function clearInvOutput_(ss, inv) {
  const {sheetName: shNo, row: rNo, col: cNo} = parseA1WithSheet_(inv.outNoPvA1, ss);
  const {sheetName: shOr, row: rOr, col: cOr} = parseA1WithSheet_(inv.outOriA1, ss);
  if (shNo !== shOr || rNo !== rOr) throw new Error('Anchor mismatch.');

  const sh = mustGetSheet_(ss, shNo);
  const totalRows = inv.mpptUsed * STRIDE_ROWS_PER_MPPT;
  sh.getRange(rNo, cNo, totalRows, 1).clearContent();
  sh.getRange(rOr, cOr, totalRows, 1).clearContent();
}

/* =========================================================
   Terminal patterns
   ========================================================= */

function readTerminalPatternMap_(ss, invType) {
  const t = normalizeInvType_(invType);
  if (t === '150') return readTerminalPattern_(ss, 'INV150_TERMINALS');
  if (t === '100') return readTerminalPattern_(ss, 'INV100_TERMINALS');
  if (t === '30-50') return readTerminalPattern_(ss, 'INV30_50_TERMINALS');
  throw new Error(`No terminal pattern sheet defined for InvType=${invType}`);
}

function readTerminalPattern_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Missing sheet ${sheetName}. Must have headers: Nstrings, TerminalList`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) throw new Error(`${sheetName} is empty.`);

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(x=>String(x||'').trim());
  const idxN = headers.indexOf('Nstrings');
  const idxL = headers.indexOf('TerminalList');
  if (idxN < 0 || idxL < 0) throw new Error(`${sheetName} must have headers: Nstrings, TerminalList`);

  const data = sh.getRange(2,1,lastRow-1,lastCol).getValues();

  const map = new Map();
  for (const r of data) {
    const n = int_(r[idxN]);
    const listStr = String(r[idxL] || '').trim();
    if (!n || !listStr) continue;
    map.set(n, parseTerminalList_(listStr));
  }
  return map;
}

function parseTerminalList_(s) {
  const parts = s.split(',').map(x=>x.trim()).filter(Boolean);
  const nums = [];
  for (const p of parts) {
    const m = /^PV\s*(\d+)$/i.exec(p);
    if (!m) throw new Error(`Bad terminal token "${p}". Use PV<number>.`);
    nums.push(parseInt(m[1],10));
  }
  return Array.from(new Set(nums)).sort((a,b)=>a-b);
}

/* =========================================================
   Terminal -> (MPPT, input) mapping (logical)
   ========================================================= */

function terminalToMpptInputFor_(inv, t) {
  const invType = normalizeInvType_(inv.invType);

  // Inv100: 2 inputs per MPPT => PV1..PV20 -> MPPT1..10
  if (invType === '100') {
    const per = inv.inputsPerMppt; // 2
    const mppt = Math.floor((t - 1) / per) + 1;
    const input = ((t - 1) % per) + 1;
    return { mppt, input };
  }

  // Inv30-50: 2 inputs per MPPT => PV1..PV8 -> MPPT1..4
  if (invType === '30-50') {
    const per = inv.inputsPerMppt; // 2
    const mppt = Math.floor((t - 1) / per) + 1;
    const input = ((t - 1) % per) + 1;
    return { mppt, input };
  }

  // Default (Inv150): 3 inputs per MPPT
  const per = inv.inputsPerMppt; // 3
  const mppt = Math.floor((t - 1) / per) + 1;
  const input = ((t - 1) % per) + 1;
  return { mppt, input };
}

/* =========================================================
   INV_CONFIG parsing + validation
   ========================================================= */

function configToInv_(cfg) {
  const invTypeRaw = String(cfg.InvType).trim();
  const invType = normalizeInvType_(invTypeRaw);

  const inv = {
    invId: String(cfg.InvID),
    invType,
    pacKw: num_(cfg.Pac_kW),
    mpptUsed: int_(cfg.MPPT_used),
    inputsPerMppt: int_(cfg.InputsPerMPPT),

    refDcac: num_(cfg.RefDCAC),
    vmpTarget: num_(cfg.VmpTarget),
    vmpMinGood: num_(cfg.VmpTarget) + num_(cfg.VmpMinGoodDelta),
    vmpMinWorst: num_(cfg.VmpTarget) + num_(cfg.VmpMinWorstDelta),

    vmpDiffMax: num_(cfg[COL_VMPDIFF_MAX]),
    outNoPvA1: String(cfg.OutputAnchor_NoPV || '').trim(),
    outOriA1: String(cfg.OutputAnchor_Orientation || '').trim(),
    cfgRow: cfg.__row
  };

  if (!inv.outNoPvA1 || !inv.outOriA1) throw new Error('Empty anchor A1.');
  if (!isFinite(inv.vmpDiffMax) || inv.vmpDiffMax <= 0) {
    throw new Error(`VmpDiffMax missing/invalid for InvID=${inv.invId}. Fill ${COL_VMPDIFF_MAX} in INV_CONFIG.`);
  }

  if (invType === '150') {
    if (inv.mpptUsed !== 7 || inv.inputsPerMppt !== 3) throw new Error('Inv150 must be MPPT_used=7 and InputsPerMPPT=3');
  } else if (invType === '100') {
    if (inv.mpptUsed !== 10 || inv.inputsPerMppt !== 2) throw new Error('Inv100 must be MPPT_used=10 and InputsPerMPPT=2');
  } else if (invType === '30-50') {
    if (inv.mpptUsed !== 4 || inv.inputsPerMppt !== 2) throw new Error('Inv30-50 must be MPPT_used=4 and InputsPerMPPT=2');
  } else {
    throw new Error(`Unsupported InvType=${invTypeRaw}`);
  }

  return inv;
}

function configToInvFromRowObj_(row) {
  const invType = normalizeInvType_(String(row.InvType || '').trim());
  return {
    invId: String(row.InvID),
    invType,
    pacKw: num_(row.Pac_kW),
    mpptUsed: int_(row.MPPT_used),
    inputsPerMppt: int_(row.InputsPerMPPT),

    refDcac: num_(row.RefDCAC),
    vmpTarget: num_(row.VmpTarget),
    vmpMinGood: num_(row.VmpTarget) + num_(row.VmpMinGoodDelta),
    vmpMinWorst: num_(row.VmpTarget) + num_(row.VmpMinWorstDelta),

    vmpDiffMax: num_(row[COL_VMPDIFF_MAX]),
    outNoPvA1: String(row.OutputAnchor_NoPV || '').trim(),
    outOriA1: String(row.OutputAnchor_Orientation || '').trim(),
    cfgRow: row.__row
  };
}

function ensureInvConfigColumns_(cfgSh, columnNames) {
  const lastCol = cfgSh.getLastColumn();
  const headers = cfgSh.getRange(1,1,1,lastCol).getValues()[0].map(x=>String(x||'').trim());
  let changed = false;
  for (const name of columnNames) {
    if (headers.indexOf(name) === -1) { headers.push(name); changed = true; }
  }
  if (changed) cfgSh.getRange(1,1,1,headers.length).setValues([headers]);
}

function readConfigFirstPending_(cfgSh) {
  const lastRow = cfgSh.getLastRow();
  const lastCol = cfgSh.getLastColumn();
  if (lastRow < 2) return null;

  const headers = cfgSh.getRange(1,1,1,lastCol).getValues()[0].map(x=>String(x||'').trim());
  const idx = {};
  for (let i=0;i<headers.length;i++) idx[headers[i]] = i;

  const required = [
    'InvID','InvType','Pac_kW','MPPT_used','InputsPerMPPT',
    'RefDCAC','VmpTarget','VmpMinGoodDelta','VmpMinWorstDelta',
    'OutputAnchor_NoPV','OutputAnchor_Orientation','Status',
    COL_VMPDIFF_MAX
  ];
  for (const h of required) if (idx[h] === undefined) throw new Error(`Missing column "${h}" in INV_CONFIG header.`);

  const data = cfgSh.getRange(2,1,lastRow-1,lastCol).getValues();
  for (let r=0;r<data.length;r++) {
    const row = data[r];
    const status = String(row[idx['Status']] || '').trim().toUpperCase();
    if (status === '' || status === 'PENDING') {
      const cfg = {};
      for (const h of required) cfg[h] = row[idx[h]];
      if (idx[COL_NSTRINGS_USED] !== undefined) cfg[COL_NSTRINGS_USED] = row[idx[COL_NSTRINGS_USED]];
      cfg.__row = r+2;
      return cfg;
    }
  }
  return null;
}

function readAllConfigRows_(cfgSh) {
  const lastRow = cfgSh.getLastRow();
  const lastCol = cfgSh.getLastColumn();
  if (lastRow < 2) return [];

  const headers = cfgSh.getRange(1,1,1,lastCol).getValues()[0].map(x=>String(x||'').trim());
  const data = cfgSh.getRange(2,1,lastRow-1,lastCol).getValues();

  const rows = [];
  for (let i=0;i<data.length;i++) {
    const obj = {};
    for (let c=0;c<headers.length;c++) obj[headers[c]] = data[i][c];
    obj.__row = i+2;
    rows.push(obj);
  }
  return rows;
}

function markAllAllocatedFinished_(cfgSh) {
  const lastRow = cfgSh.getLastRow();
  if (lastRow < 2) return;

  const statusCol = colIndexByHeader_(cfgSh, 'Status');
  const statuses = cfgSh.getRange(2, statusCol, lastRow - 1, 1).getValues();

  for (let i=0;i<statuses.length;i++) {
    const v = String(statuses[i][0] || '').trim().toUpperCase();
    if (v === 'ALLOCATED') statuses[i][0] = 'FINISHED';
  }
  cfgSh.getRange(2, statusCol, lastRow - 1, 1).setValues(statuses);
}

function readConfigRowByInvId_(cfgSh, invId) {
  const rows = readAllConfigRows_(cfgSh);
  for (const r of rows) if (String(r.InvID||'') === String(invId)) return r;
  return null;
}

function writeConfigCellByInvId_(cfgSh, invId, headerName, value) {
  const row = readConfigRowByInvId_(cfgSh, invId);
  if (!row) return;
  const col = colIndexByHeader_(cfgSh, headerName);
  cfgSh.getRange(row.__row, col).setValue(value);
}

function colIndexByHeader_(sh, headerName) {
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
  const i = headers.indexOf(headerName);
  if (i < 0) throw new Error(`Header "${headerName}" not found on sheet ${sh.getName()}`);
  return i+1;
}

/* =========================================================
   Inventory
   ========================================================= */

function readStringInventory_(ss) {
  let sh = ss.getSheetByName('STRING_INVENTORY');
  if (!sh) {
    const sheets = ss.getSheets();
    for (const s of sheets) {
      const rng = s.getRange(1,1,1, Math.min(200, s.getLastColumn()));
      const headers = rng.getValues()[0].map(x=>String(x||'').trim());
      if (headers.some(h => /^String\s*1$/i.test(h))) { sh = s; break; }
    }
  }
  if (!sh) return {sheetName:null, groups:[]};

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return {sheetName: sh.getName(), groups:[]};

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(x=>String(x||'').trim());
  const idxBuilding = findHeaderIndex_(headers, ['Name of Building','Building','Name']);
  const idxOrientation = findHeaderIndex_(headers, ['Orientation','Ori']);
  const idxString1 = headers.findIndex(h => /^String\s*1$/i.test(h));
  if (idxBuilding < 0 || idxOrientation < 0 || idxString1 < 0) return {sheetName: sh.getName(), groups:[]};

  const data = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  const groups = [];

  for (let r=0;r<data.length;r++) {
    const row = data[r];
    const building = row[idxBuilding];
    const orientation = row[idxOrientation];
    if (building === '' || orientation === '') continue;

    const strings = [];
    for (let c=idxString1; c<lastCol; c++) {
      const len = int_(row[c]);
      if (len > 0) {
        const a1 = `${sh.getName()}!${a1_(r+2, c+1)}`;
        strings.push({len, sourceKey: a1});
      }
    }
    if (strings.length) groups.push({building, orientation, strings});
  }

  return {sheetName: sh.getName(), groups};
}

function buildAvailableGroups_(invData, usedKeySet) {
  return invData.groups.map(g => ({
    building: g.building,
    orientation: g.orientation,
    strings: g.strings.filter(it => !usedKeySet.has(it.sourceKey))
  })).filter(g => g.strings.length > 0);
}

/* =========================================================
   ALLOC_LOG with State
   ========================================================= */

function ensureAllocLog_(ss) {
  let sh = ss.getSheetByName('ALLOC_LOG');
  if (!sh) {
    sh = ss.insertSheet('ALLOC_LOG');
    sh.getRange(1,1,1,LOG_HEADERS.length).setValues([LOG_HEADERS]);
    return sh;
  }

  // Ensure headers contain State
  const lastCol = Math.max(sh.getLastColumn(), LOG_HEADERS.length);
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(x=>String(x||'').trim());
  let changed = false;
  for (const h of LOG_HEADERS) {
    if (headers.indexOf(h) === -1) {
      headers.push(h);
      changed = true;
    }
  }
  if (changed) sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}

function resetAllocLog() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('ALLOC_LOG');
  if (!sh) return;
  sh.clearContents();
  sh.getRange(1,1,1,LOG_HEADERS.length).setValues([LOG_HEADERS]);
}

function readUsedSetActiveOnly_(logSh) {
  const lastRow = logSh.getLastRow();
  if (lastRow < 2) return new Set();

  const headers = logSh.getRange(1,1,1,logSh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
  const idxSource = headers.indexOf('SourceKey');
  const idxState = headers.indexOf('State');

  const vals = logSh.getRange(2,1,lastRow-1,logSh.getLastColumn()).getValues();
  const used = new Set();
  for (const r of vals) {
    const key = String(r[idxSource] || '').trim();
    if (!key) continue;
    const st = idxState >= 0 ? String(r[idxState] || 'ACTIVE').trim().toUpperCase() : 'ACTIVE';
    if (st === 'ACTIVE') used.add(key);
  }
  return used;
}

function appendAllocLogActive_(logSh, invId, plan) {
  const ts = new Date();
  const rows = [];
  for (const block of plan.assigned) {
    if (!block) continue;
    for (const it of block.items) {
      rows.push([ts, invId, block.building, block.orientation, it.len, it.sourceKey, it.targetA1 || '', 'ACTIVE']);
    }
  }
  if (rows.length) logSh.getRange(logSh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
}

function appendOneAllocLogActive_(logSh, invId, building, orientation, len, sourceKey, targetCell) {
  logSh.appendRow([new Date(), invId, building, orientation, len, sourceKey, targetCell || '', 'ACTIVE']);
}

function readActiveItemsForInv_(logSh, invId) {
  const items = readActiveLogItemsForInv_(logSh, invId);
  return items.map(x => ({len:x.len, sourceKey:x.sourceKey}));
}

function readActiveLogItemsForInv_(logSh, invId) {
  const lastRow = logSh.getLastRow();
  if (lastRow < 2) return [];

  const headers = logSh.getRange(1,1,1,logSh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
  const idxInv = headers.indexOf('InvID');
  const idxB = headers.indexOf('Building');
  const idxO = headers.indexOf('Orientation');
  const idxL = headers.indexOf('StringLength');
  const idxS = headers.indexOf('SourceKey');
  const idxT = headers.indexOf('TargetCell');
  const idxState = headers.indexOf('State');

  const vals = logSh.getRange(2,1,lastRow-1,logSh.getLastColumn()).getValues();

  const out = [];
  for (let i=0;i<vals.length;i++) {
    const r = vals[i];
    if (String(r[idxInv]||'') !== String(invId)) continue;
    const st = idxState >= 0 ? String(r[idxState]||'ACTIVE').trim().toUpperCase() : 'ACTIVE';
    if (st !== 'ACTIVE') continue;
    out.push({
      rowIndex: i+2,
      invId: String(invId),
      building: String(r[idxB]||''),
      orientation: String(r[idxO]||''),
      len: int_(r[idxL]),
      sourceKey: String(r[idxS]||''),
      targetCell: String(r[idxT]||'')
    });
  }
  return out;
}

function markLogRowsState_(logSh, rowIndices, newState) {
  if (!rowIndices.length) return;

  const headers = logSh.getRange(1,1,1,logSh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
  const idxState = headers.indexOf('State');
  if (idxState < 0) throw new Error('ALLOC_LOG missing State column.');

  // Set individually (small count)
  for (const r of rowIndices) {
    logSh.getRange(r, idxState+1).setValue(String(newState).toUpperCase());
  }
}

function updateLogTargetCellsBySourceKey_(logSh, invId, sourceKeyToTargetCellMap) {
  const lastRow = logSh.getLastRow();
  if (lastRow < 2) return;

  const headers = logSh.getRange(1,1,1,logSh.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
  const idxInv = headers.indexOf('InvID');
  const idxSource = headers.indexOf('SourceKey');
  const idxTarget = headers.indexOf('TargetCell');
  const idxState = headers.indexOf('State');

  const vals = logSh.getRange(2,1,lastRow-1,logSh.getLastColumn()).getValues();

  let updates = 0;
  for (let i=0;i<vals.length;i++) {
    const r = vals[i];
    if (String(r[idxInv]||'') !== String(invId)) continue;
    const st = idxState >= 0 ? String(r[idxState]||'ACTIVE').trim().toUpperCase() : 'ACTIVE';
    if (st !== 'ACTIVE') continue;

    const src = String(r[idxSource]||'').trim();
    if (!src) continue;
    const newTarget = sourceKeyToTargetCellMap.get(src);
    if (!newTarget) continue;

    // Update cell directly
    logSh.getRange(i+2, idxTarget+1).setValue(newTarget);
    updates++;
  }
  return updates;
}

/* =========================================================
   Group items helper
   ========================================================= */

function groupItems_(items) {
  const map = new Map(); // key building||orientation -> {building,orientation,strings:[]}
  for (const it of items) {
    const k = `${it.building}||${it.orientation}`;
    if (!map.has(k)) map.set(k, {building: it.building, orientation: it.orientation, strings: []});
    map.get(k).strings.push({len: it.len, sourceKey: it.sourceKey});
  }
  return Array.from(map.values());
}

/* =========================================================
   Donor output parsing (read/write)
   ========================================================= */

function readInvOutputBlock_(ss, inv) {
  const {sheetName, row, col} = parseA1WithSheet_(inv.outNoPvA1, ss);
  const {sheetName: sh2, row: r2, col: c2} = parseA1WithSheet_(inv.outOriA1, ss);
  if (sheetName !== sh2 || row !== r2) throw new Error('Anchor mismatch (NoPV vs Orientation).');

  const sh = mustGetSheet_(ss, sheetName);
  const totalRows = inv.mpptUsed * STRIDE_ROWS_PER_MPPT;

  const noPvRange = sh.getRange(row, col, totalRows, 1);
  const orRange = sh.getRange(row, c2, totalRows, 1);

  const noPvVals = noPvRange.getValues().map(r=>r[0]);
  const orVals = orRange.getValues().map(r=>r[0]);

  const targetA1 = [];
  for (let i=0;i<totalRows;i++) targetA1.push(`${sh.getName()}!${a1_(row+i, col)}`);

  return {sh, startRow: row, totalRows, inv, noPvRange, orRange, noPvVals, orVals, targetA1};
}

function flushInvOutputBlock_(out) {
  out.noPvRange.setValues(out.noPvVals.map(v=>[v]));
  out.orRange.setValues(out.orVals.map(v=>[v]));
}

function findEmptyOnTerminals_(out, terminals, inv) {
  const empties = [];
  for (const t of terminals) {
    const {mppt, input} = terminalToMpptInputFor_(inv, t);
    const linearIndex = (mppt-1)*STRIDE_ROWS_PER_MPPT + (input-1);
    const val = out.noPvVals[linearIndex];
    if (val === '' || val === null || val === undefined) {
      empties.push({terminal:t, mppt, input, linearIndex, targetA1: out.targetA1[linearIndex]});
    }
  }
  return empties;
}

function writeOneTerminal_(out, empty, len, orientation) {
  const idx = empty.linearIndex;
  out.noPvVals[idx] = len;

  // MERGE SAFE: write orientation at MPPT top row
  const base = (empty.mppt-1)*STRIDE_ROWS_PER_MPPT;
  out.orVals[base] = orientation;
}

function computePdcFromOutput_(noPvVals, moduleWp) {
  let sumWp = 0;
  for (const v of noPvVals) {
    const len = int_(v);
    if (len > 0) sumWp += len * moduleWp;
  }
  return sumWp / 1000.0;
}

function inferMpptInfoFromLogAndOutput_(logSh, invId, out, inv) {
  const mpptInfo = new Array(inv.mpptUsed).fill(null).map(()=>({building:null, orientation:null, len:null}));

  for (let m=0;m<inv.mpptUsed;m++) {
    const base = m*STRIDE_ROWS_PER_MPPT;
    const lens = [];
    for (let i=0;i<inv.inputsPerMppt;i++) {
      const L = int_(out.noPvVals[base+i]);
      if (L>0) lens.push(L);
    }
    const ori = out.orVals[base] || null;
    mpptInfo[m].orientation = ori ? String(ori) : null;
    if (lens.length) mpptInfo[m].len = modeInt_(lens);
  }

  // Building comes from log TargetCell mapping
  const logItems = readActiveLogItemsForInv_(logSh, invId);
  const byTarget = new Map();
  for (const li of logItems) byTarget.set(li.targetCell, li.building);

  for (let i=0;i<out.targetA1.length;i++) {
    const building = byTarget.get(out.targetA1[i]);
    if (!building) continue;
    const mppt = Math.floor(i / STRIDE_ROWS_PER_MPPT) + 1;
    const idx = mppt - 1;
    if (!mpptInfo[idx].building) mpptInfo[idx].building = building;
  }

  return mpptInfo;
}

function modeInt_(arr) {
  const freq = new Map();
  for (const x of arr) freq.set(x, (freq.get(x) || 0) + 1);
  let best = arr[0], bestC = -1;
  for (const [k,v] of freq.entries()) {
    if (v > bestC) { bestC = v; best = k; }
  }
  return best;
}

/* =========================================================
   UNUSED sheet
   ========================================================= */

function computeUnused_(invData, usedKeySet) {
  const unused = [];
  for (const g of invData.groups) {
    for (const it of g.strings) {
      if (!usedKeySet.has(it.sourceKey)) unused.push({building:g.building, orientation:g.orientation, len:it.len, sourceKey:it.sourceKey});
    }
  }
  return unused;
}

function writeUnusedSheet_(ss, leftovers) {
  let sh = ss.getSheetByName('UNUSED_STRINGS');
  if (!sh) sh = ss.insertSheet('UNUSED_STRINGS');
  sh.clearContents();
  sh.getRange(1,1,1,4).setValues([['Building','Orientation','StringLength','SourceKey']]);
  if (leftovers.length) {
    sh.getRange(2,1,leftovers.length,4).setValues(leftovers.map(x=>[x.building,x.orientation,x.len,x.sourceKey]));
  }
}

function buildUnusedPool_(unusedList) {
  const pool = new Map();
  for (const u of unusedList) {
    const k = `${u.building}||${u.orientation}||${u.len}`;
    if (!pool.has(k)) pool.set(k, []);
    pool.get(k).push(u);
  }
  return pool;
}

function popUnused_(pool, building, orientation, len) {
  const k = `${building}||${orientation}||${len}`;
  const arr = pool.get(k);
  if (!arr || arr.length === 0) return null;
  return arr.pop();
}

/* =========================================================
   Inv type normalization
   ========================================================= */

function normalizeInvType_(invTypeRaw) {
  const s = String(invTypeRaw || '').trim().toLowerCase();
  if (s === '150' || s === 'inv150') return '150';
  if (s === '100' || s === 'inv100') return '100';
  if (s === '30-50' || s === '30/50' || s === '30_50' || s === '30' || s === '50') return '30-50';
  return String(invTypeRaw || '').trim();
}

function isSupportedInvType_(invTypeRaw) {
  const t = normalizeInvType_(invTypeRaw);
  return (t === '150' || t === '100' || t === '30-50');
}

/* =========================================================
   Named ranges
   ========================================================= */

function getNamedNumber_(ss, name, fallback) {
  const r = ss.getRangeByName(name);
  if (!r) return fallback;
  const n = num_(r.getValue());
  return isFinite(n) && n > 0 ? n : fallback;
}

/* =========================================================
   Sheet & A1 helpers
   ========================================================= */

function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function parseA1WithSheet_(a1, ss) {
  const s = String(a1||'').trim();
  if (!s) throw new Error('Empty anchor A1.');

  let sheetName, cellA1;
  if (s.includes('!')) {
    const parts = s.split('!');
    sheetName = parts[0];
    cellA1 = parts[1];
  } else {
    sheetName = ss.getActiveSheet().getName();
    cellA1 = s;
  }

  const m = /^([A-Z]+)(\d+)$/i.exec(cellA1);
  if (!m) throw new Error(`Invalid A1 reference: ${a1}`);
  const col = colToNum_(m[1].toUpperCase());
  const row = parseInt(m[2], 10);
  return {sheetName, row, col};
}

function a1_(row, col) {
  return `${numToCol_(col)}${row}`;
}

function colToNum_(letters) {
  let n = 0;
  for (let i=0;i<letters.length;i++) n = n*26 + (letters.charCodeAt(i) - 64);
  return n;
}

function numToCol_(n) {
  let s = '';
  while (n > 0) {
    const m = (n-1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n-1)/26);
  }
  return s;
}

function findHeaderIndex_(headers, candidates) {
  const lower = headers.map(h => h.toLowerCase());
  for (const c of candidates) {
    const i = lower.indexOf(String(c).toLowerCase());
    if (i >= 0) return i;
  }
  return -1;
}

function int_(x) {
  if (x === null || x === undefined || x === '') return 0;
  const n = Number(String(x).replace(',', '.'));
  return isFinite(n) ? Math.trunc(n) : 0;
}
function num_(x) {
  if (x === null || x === undefined || x === '') return NaN;
  const n = Number(String(x).replace(',', '.'));
  return isFinite(n) ? n : NaN;
}
