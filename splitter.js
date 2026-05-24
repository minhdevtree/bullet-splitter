// Browser port of split-bullets-surgical.js with runtime options.
// Same surgical algorithm: edits only the worksheet XMLs of the configured
// sheets plus xl/sharedStrings.xml. Every other byte of the workbook is
// preserved.

/* global JSZip */

// ---------- helpers ----------

function colToIdx(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function idxToCol(idx) {
  let s = '';
  while (idx > 0) {
    const r = (idx - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    idx = Math.floor((idx - 1) / 26);
  }
  return s;
}
function parseRef(ref) {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  return {
    left: colToIdx(m[1]), top: +m[2],
    right: colToIdx(m[3]), bottom: +m[4],
  };
}
function xmlEscape(s) {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
function xmlUnescape(s) {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

// Two flavours of list marker at the start of a trimmed line.
//   Unordered: - – — • * + · ▪ ▶ → ►   (one or more, e.g. "--" works)
//   Ordered:   1.  2)  10.  10)         |  a.  a)  I.
// Each followed by optional spaces/tabs.
const UNORDERED_RE = /^[-–—•*+·▪▶→►]+[ \t]*/;
const ORDERED_RE   = /^(?:[0-9]+[.)]|[A-Za-z][.)])[ \t]*/;
const BULLET_RE    = /^(?:[-–—•*+·▪▶→►]+|[0-9]+[.)]|[A-Za-z][.)])[ \t]*/;

function isBulletLine(s)    { return BULLET_RE.test(s); }
function isUnordered(s)     { return UNORDERED_RE.test(s); }
function isOrdered(s)       { return ORDERED_RE.test(s); }

// Pick the bullet group that dominates the text. When a list mixes a stray
// ordered marker (e.g. "16." as a student number) with proper unordered
// bullets, the minority group is treated as a preamble and dropped. Ties
// prefer unordered, since that's the overwhelmingly common Vietnamese
// school-report style.
function dominantBulletLines(lines) {
  const unordered = lines.filter(isUnordered);
  const ordered   = lines.filter(isOrdered);
  if (unordered.length === 0 && ordered.length === 0) return [];
  return unordered.length >= ordered.length ? unordered : ordered;
}

function makeBulletDetector(mode) {
  // 'strict': first trimmed line is a bullet, AND there's another bullet
  //           line further down. No preamble allowed.
  // 'tolerant': at least 2 lines are bullets anywhere in the text — allows
  //             a non-bullet preamble such as "26.\n- abc\n- def".
  if (mode === 'strict') {
    return text => {
      if (!text) return false;
      const lines = text.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      if (lines.length < 2) return false;
      return isBulletLine(lines[0]) && lines.slice(1).some(isBulletLine);
    };
  }
  return text => {
    if (!text) return false;
    const lines = text.split(/\r?\n/).map(s => s.trim());
    return dominantBulletLines(lines).length >= 2;
  };
}
function splitBullets(text, mode) {
  const lines = text.split(/\r?\n/).map(s => s.trim());
  if (mode === 'strict') {
    // Honour the user's choice: keep every non-empty line verbatim.
    return lines.filter(s => s.length > 0);
  }
  return dominantBulletLines(lines);
}
// Strip the list marker plus any spaces/tabs after it. Leaves non-bullet
// lines untouched (no-op when there's nothing to strip).
function stripLeadingBullet(line) {
  return line.replace(BULLET_RE, '');
}

// xlsx stores control characters like CR/LF inside <t> as the literal
// 7-char sequence "_x000D_" / "_x000A_". Decode all such _xHHHH_ escapes so
// line-splitting on \r?\n works the same regardless of how Excel wrote it.
function decodeXlsxEscapes(s) {
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex) =>
    String.fromCharCode(parseInt(hex, 16))
  );
}

function parseSharedStrings(xml) {
  const items = [];
  const re = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const inner = m[1];
    let text = '';
    const tre = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
    let tm;
    while ((tm = tre.exec(inner)) !== null) text += tm[1];
    items.push(decodeXlsxEscapes(xmlUnescape(text)));
  }
  return items;
}

// Strip a pair of matching surrounding quote characters (straight or curly,
// single or double). Some workbooks wrap multi-line Nhận xét content in
// outer quotes; we want the raw lines.
function stripSurroundingQuotes(s) {
  if (!s) return s;
  const t = s.trim();
  if (t.length < 2) return s;
  const first = t[0], last = t[t.length - 1];
  const pairs = { '"': '"', "'": "'", '“': '”', '‘': '’', '«': '»' };
  if (pairs[first] && pairs[first] === last) return t.slice(1, -1);
  return s;
}

// Fallback for text without bullet markers: when the number of non-empty
// lines exactly matches the merge's row count we treat each line as a
// per-row comment. Returns the lines array, or null when we shouldn't
// split.
function tryLineFallback(text, expectedRows) {
  if (!text || expectedRows < 2) return null;
  const stripped = stripSurroundingQuotes(text);
  const lines = stripped.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (lines.length === expectedRows && lines.length >= 2) return lines;
  return null;
}

function appendSharedStrings(xml, newStrings) {
  const sstOpenRe = /<sst\b([^>]*)>/;
  const sstMatch = sstOpenRe.exec(xml);
  if (!sstMatch) throw new Error('sharedStrings.xml missing <sst> root');
  const attrs = sstMatch[1];
  const getNum = (name) => {
    const re = new RegExp(`\\b${name}="(\\d+)"`);
    const mm = re.exec(attrs);
    return mm ? +mm[1] : null;
  };
  let count = getNum('count');
  let unique = getNum('uniqueCount');
  const addCount = newStrings.length;
  if (count != null) count += addCount;
  if (unique != null) unique += addCount;

  let newAttrs = attrs;
  if (count != null) newAttrs = newAttrs.replace(/\bcount="\d+"/, `count="${count}"`);
  if (unique != null) newAttrs = newAttrs.replace(/\buniqueCount="\d+"/, `uniqueCount="${unique}"`);

  const newBlocks = newStrings.map(s => {
    const needsPreserve = /^\s|\s$/.test(s);
    const tAttr = needsPreserve ? ' xml:space="preserve"' : '';
    return `<si><t${tAttr}>${xmlEscape(s)}</t></si>`;
  }).join('');

  let out = xml.replace(sstOpenRe, `<sst${newAttrs}>`);
  out = out.replace(/<\/sst>\s*$/, `${newBlocks}</sst>`);
  return out;
}

function findRowCell(sheetXml, colLetter, rowNum) {
  const ref = `${colLetter}${rowNum}`;
  const re = new RegExp(`<c\\s+r="${ref}"([^/>]*)(/>|>([\\s\\S]*?)</c>)`);
  const m = re.exec(sheetXml);
  if (!m) return null;
  const attrsPart = m[1];
  const isSelfClose = m[2] === '/>';
  const inner = isSelfClose ? '' : m[3];
  const styleM = /\bs="([^"]*)"/.exec(attrsPart);
  const typeM = /\bt="([^"]*)"/.exec(attrsPart);
  return {
    index: m.index,
    length: m[0].length,
    style: styleM ? styleM[1] : null,
    type: typeM ? typeM[1] : null,
    inner,
    isSelfClose,
  };
}

function findFullBoxXfs(stylesXml) {
  const bordersM = /<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(stylesXml);
  if (!bordersM) return null;
  const borders = [];
  const bRe = /<border>[\s\S]*?<\/border>|<border\s*\/>/g;
  let m;
  while ((m = bRe.exec(bordersM[1])) !== null) borders.push(m[0]);

  const sideHas = (borderXml, side) => {
    const re = new RegExp('<' + side + '\\b([^>/]*?)(/>|>[\\s\\S]*?</' + side + '>)');
    const mm = re.exec(borderXml);
    return !!(mm && /\bstyle="[^"]+"/.test(mm[1]));
  };
  const fullBoxBorderIds = new Set();
  for (let i = 0; i < borders.length; i++) {
    if (sideHas(borders[i], 'left') && sideHas(borders[i], 'right')
        && sideHas(borders[i], 'top') && sideHas(borders[i], 'bottom')) {
      fullBoxBorderIds.add(i);
    }
  }
  if (fullBoxBorderIds.size === 0) return null;

  const xfsM = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
  if (!xfsM) return null;
  const xfs = [];
  const xfRe = /<xf\b[^/]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g;
  while ((m = xfRe.exec(xfsM[1])) !== null) xfs.push(m[0]);

  let masterXf = null, innerXf = null;
  for (let i = 0; i < xfs.length; i++) {
    const xf = xfs[i];
    const bM = /\bborderId="(\d+)"/.exec(xf);
    if (!bM) continue;
    if (!fullBoxBorderIds.has(+bM[1])) continue;
    const aM = /<alignment\b([^/>]*)/.exec(xf);
    if (!aM) continue;
    const attrs = aM[1];
    const horiz = (/\bhorizontal="([^"]+)"/.exec(attrs) || [])[1];
    const vert = (/\bvertical="([^"]+)"/.exec(attrs) || [])[1];
    const wrap = /\bwrapText="1"/.test(attrs);
    if (!wrap) continue;
    if (!masterXf && (horiz === 'justify' || horiz === 'left') && vert === 'center') masterXf = i;
    if (!innerXf && horiz === 'center') innerXf = i;
  }
  if (!masterXf && innerXf != null) masterXf = innerXf;
  if (masterXf == null || innerXf == null) return null;
  return { masterXf, innerXf };
}

// ---------- main entry ----------

/**
 * @param {ArrayBuffer} arrayBuffer
 * @param {{
 *   sheetNames?: string[],
 *   bulletDetection?: 'strict'|'tolerant',
 *   applyFullBorder?: boolean,
 *   stripBulletPrefix?: boolean,
 * }} [opts]
 */
export async function processWorkbookBuffer(arrayBuffer, opts = {}) {
  const sheetNames = opts.sheetNames && opts.sheetNames.length
    ? opts.sheetNames
    : ['DKQHT', 'NLPC'];
  const mode = opts.bulletDetection === 'strict' ? 'strict' : 'tolerant';
  const applyFullBorder = opts.applyFullBorder !== false;
  const stripBullet = opts.stripBulletPrefix !== false; // default on
  const isBulletList = makeBulletDetector(mode);

  const zip = await JSZip.loadAsync(arrayBuffer);

  const stylesFile = zip.file('xl/styles.xml');
  const stylesXml = stylesFile ? await stylesFile.async('string') : '';
  const fullBox = applyFullBorder && stylesXml ? findFullBoxXfs(stylesXml) : null;

  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const sheetIdToName = {};
  const sheetRe = /<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/>/g;
  let mm;
  while ((mm = sheetRe.exec(wbXml)) !== null) sheetIdToName[mm[2]] = mm[1];

  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const ridToTarget = {};
  const relRe = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/>/g;
  while ((mm = relRe.exec(relsXml)) !== null) ridToTarget[mm[1]] = mm[2];

  const ssPath = 'xl/sharedStrings.xml';
  const ssFile = zip.file(ssPath);
  let ssXml = ssFile ? await ssFile.async('string') : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>';
  const ssList = parseSharedStrings(ssXml);
  const newStrings = [];
  const addString = (text) => {
    const newIdx = ssList.length + newStrings.length;
    newStrings.push(text);
    return newIdx;
  };

  const log = [];
  let totalSplits = 0;   // number of multi-row merges that got split
  let totalRows = 0;     // number of bullet lines we wrote (sum across merges)
  const sheetsProcessed = [];
  const sheetsMissing = [];

  const wantSet = new Set(sheetNames.map(s => s.trim()).filter(Boolean));
  for (const wanted of wantSet) {
    if (!Object.values(sheetIdToName).includes(wanted)) sheetsMissing.push(wanted);
  }

  for (const [rid, name] of Object.entries(sheetIdToName)) {
    if (!wantSet.has(name)) continue;
    const target = ridToTarget[rid];
    if (!target) continue;
    const sheetPath = `xl/${target}`;
    let sheetXml = await zip.file(sheetPath).async('string');

    const mergesBlockRe = /<mergeCells\b([^>]*)>([\s\S]*?)<\/mergeCells>/;
    const mb = mergesBlockRe.exec(sheetXml);
    if (!mb) {
      sheetsProcessed.push(name);
      continue;
    }
    const mergesAttrs = mb[1];
    const mergesBody = mb[2];
    const blockStart = mb.index;
    const blockEnd = mb.index + mb[0].length;

    const mergeEntries = [];
    const meRe = /<mergeCell\s+ref="([^"]+)"\s*\/>/g;
    let me;
    while ((me = meRe.exec(mergesBody)) !== null) {
      mergeEntries.push({ raw: me[0], ref: me[1] });
    }

    const toReplace = [];
    for (const entry of mergeEntries) {
      const r = parseRef(entry.ref);
      if (!r) continue;
      if (r.bottom <= r.top) continue;
      // Skip full-width content blocks (Khen thưởng, "5. Khen thưởng", etc).
      // Those always start at column A; real Nhận xét merges start at column
      // J or further right, never at A.
      if (r.left === 1) continue;
      const leftCol = idxToCol(r.left);

      const masterCellInfo = findRowCell(sheetXml, leftCol, r.top);
      if (!masterCellInfo) continue;
      let text = null;
      if (masterCellInfo.type === 's') {
        const vMatch = /<v>(\d+)<\/v>/.exec(masterCellInfo.inner);
        if (!vMatch) continue;
        text = ssList[+vMatch[1]];
      } else {
        const fMatch = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(masterCellInfo.inner);
        if (fMatch) {
          text = decodeXlsxEscapes(xmlUnescape(fMatch[1]));
        } else {
          const vMatch = /<v>([\s\S]*?)<\/v>/.exec(masterCellInfo.inner);
          if (vMatch && !/^#[A-Z]+!?\??$/.test(vMatch[1].trim())) {
            text = decodeXlsxEscapes(xmlUnescape(vMatch[1]));
          }
        }
      }
      if (text == null) continue;

      // Try bullet detection first. If that fails, fall back to splitting
      // by line when the line count matches the merge's row count exactly
      // — common in workbooks where teachers typed plain per-row comments
      // wrapped in outer quotes, no "-" prefix.
      const expectedRows = r.bottom - r.top + 1;
      let rawBullets;
      if (isBulletList(text)) {
        rawBullets = splitBullets(text, mode);
      } else {
        rawBullets = tryLineFallback(text, expectedRows);
        if (!rawBullets) continue;
      }
      const bullets = rawBullets.map(b => stripBullet ? stripLeadingBullet(b) : b);
      const rows = [];
      for (let i = 0; i < r.bottom - r.top + 1; i++) {
        const rowNum = r.top + i;
        const bullet = bullets[i] != null ? bullets[i] : '';
        const newIdx = addString(bullet);
        rows.push({ rowNum, newIdx });
      }
      const rightCol = idxToCol(r.right);
      const newRefs = rows.map(({ rowNum }) => `${leftCol}${rowNum}:${rightCol}${rowNum}`);

      toReplace.push({
        oldEntry: entry,
        newRefs,
        leftCol,
        rightCol,
        rows,
        masterStyle: masterCellInfo.style,
      });
    }

    sheetsProcessed.push(name);
    if (toReplace.length === 0) continue;

    // rewrite mergeCells block first
    const replacedOldRefs = new Set(toReplace.map(x => x.oldEntry.ref));
    const keptEntries = mergeEntries.filter(e => !replacedOldRefs.has(e.ref));
    const allNewMergeRefs = toReplace.flatMap(x => x.newRefs);
    const allRefs = [...keptEntries.map(e => e.ref), ...allNewMergeRefs];
    const newBody = allRefs.map(r => `<mergeCell ref="${r}"/>`).join('');
    let newAttrs = mergesAttrs;
    if (/\bcount="\d+"/.test(newAttrs)) {
      newAttrs = newAttrs.replace(/\bcount="\d+"/, `count="${allRefs.length}"`);
    } else {
      newAttrs = ` count="${allRefs.length}"`;
    }
    const newMergesBlock = `<mergeCells${newAttrs}>${newBody}</mergeCells>`;
    sheetXml = sheetXml.slice(0, blockStart) + newMergesBlock + sheetXml.slice(blockEnd);

    // cell rewrites
    for (const r of toReplace) {
      const leftIdx = colToIdx(r.leftCol);
      const rightIdx = colToIdx(r.rightCol);
      const masterStyleAttr = fullBox ? ` s="${fullBox.masterXf}"` : (r.masterStyle ? ` s="${r.masterStyle}"` : '');
      const innerStyleAttr = fullBox ? ` s="${fullBox.innerXf}"` : '';
      for (const { rowNum, newIdx } of r.rows) {
        {
          const info = findRowCell(sheetXml, r.leftCol, rowNum);
          if (info) {
            const styleAttr = masterStyleAttr || (info.style ? ` s="${info.style}"` : '');
            const repl = `<c r="${r.leftCol}${rowNum}"${styleAttr} t="s"><v>${newIdx}</v></c>`;
            sheetXml = sheetXml.slice(0, info.index) + repl + sheetXml.slice(info.index + info.length);
          }
        }
        if (fullBox) {
          for (let ci = leftIdx + 1; ci <= rightIdx; ci++) {
            const col = idxToCol(ci);
            const info = findRowCell(sheetXml, col, rowNum);
            if (!info) continue;
            const repl = `<c r="${col}${rowNum}"${innerStyleAttr}/>`;
            sheetXml = sheetXml.slice(0, info.index) + repl + sheetXml.slice(info.index + info.length);
          }
        }
      }
    }

    zip.file(sheetPath, sheetXml);

    for (const r of toReplace) {
      log.push(`${name}: ${r.oldEntry.ref} → ${r.newRefs.length} dòng`);
      totalSplits++;
      totalRows += r.newRefs.length;
    }
  }

  if (newStrings.length > 0) {
    ssXml = appendSharedStrings(ssXml, newStrings);
    zip.file(ssPath, ssXml);
  }

  const outBlob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  return {
    blob: outBlob,
    log,
    totalSplits,   // count of merge groups that were split
    totalRows,     // count of bullet lines written
    sheetsProcessed,
    sheetsMissing,
  };
}

/**
 * Apply a prefix/suffix to a filename while keeping the extension.
 */
export function decorateFilename(name, prefix, suffix) {
  if (!prefix && !suffix) return name;
  const dot = name.lastIndexOf('.');
  if (dot < 0) return `${prefix || ''}${name}${suffix || ''}`;
  return `${prefix || ''}${name.slice(0, dot)}${suffix || ''}${name.slice(dot)}`;
}
