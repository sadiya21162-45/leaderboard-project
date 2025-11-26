// ranker.js
// Usage: node ranker.js
// Requires: npm install exceljs chalk

const ExcelJS = require("exceljs");
const fs = require("fs");
const chalk = require("chalk");

const INPUT_PATH = "./data/leaderboard.xlsx";
const OUTPUT_JSON = "./ranked_leaderboard.json";

function toNumber(val) {
  if (val === null || val === undefined) return 0;
  if (typeof val === "number") return val;
  // strip commas, currency symbols, parentheses etc
  const s = String(val).replace(/[,£$]/g, "").trim();
  if (s === "-" || /^D\$\s*Q$/i.test(s) || /^D\$Q$/i.test(s) || s === "")
    return 0;
  const n = Number(s);
  return Number.isNaN(n) ? 0 : n;
}

function normalizePointsCell(val) {
  if (val === null || val === undefined) return 0;
  const s = String(val).trim();
  if (s === "-" || s.toUpperCase().includes("D$Q") || s.toUpperCase() === "DQ")
    return 0;
  // if it's numeric-like, parse
  const n = Number(s);
  return Number.isNaN(n) ? 0 : n;
}

// Countback comparator for two players a and b
// returns -1 if a should come before b (a ranks higher), 1 if b higher, 0 if equal
function countbackCompare(a, b) {
  // a.scores and b.scores are arrays of numbers (length = number of rounds)
  // Build union of unique scores (descending)
  const all = Array.from(new Set([...a.scores, ...b.scores]));
  // sort descending
  all.sort((x, y) => y - x);

  for (const score of all) {
    // occurrences of this score in each player's full scores
    const aCount = a.scores.filter((s) => s === score).length;
    const bCount = b.scores.filter((s) => s === score).length;

    // If one has the score and the other doesn't, the one who has it is higher
    if (aCount > 0 && bCount === 0) return -1;
    if (bCount > 0 && aCount === 0) return 1;

    // If both have it, compare the frequency: most occurrences ranks higher
    if (aCount > bCount) return -1;
    if (bCount > aCount) return 1;

    // else equal for this score, continue to next-highest
  }

  // nothing broke the tie
  return 0;
}

function comparePlayers(a, b) {
  // 1) Total points, descending
  if (a.totalPoints !== b.totalPoints) return b.totalPoints - a.totalPoints;

  // Mark both as tied by default; we'll clear if tie broken.
  a.tied = true;
  b.tied = true;

  // 2) Total spent, ascending (lower spending is better)
  if (a.totalSpent !== b.totalSpent) {
    // smaller spent -> rank higher (return negative if a spent less)
    return a.totalSpent - b.totalSpent;
  }

  // 3) Countback system
  const cb = countbackCompare(a, b);
  if (cb !== 0) {
    // clear tied if resolved
    a.tied = false;
    b.tied = false;
    return cb;
  }

  // 4) final alphabetical order (case-insensitive)
  // leaving tied true for both (so caller can highlight)
  return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
}

async function findHeaderRow(sheet, predicateFn) {
  // Look through first N rows (reasonable limit) for a header that matches predicate
  const MAX_ROWS = Math.min(30, sheet.rowCount);
  for (let r = 1; r <= MAX_ROWS; r++) {
    const row = sheet.getRow(r);
    const cells = [];
    for (let c = 1; c <= Math.min(40, row.cellCount); c++) {
      const v = row.getCell(c).value;
      if (v !== null && v !== undefined) cells.push(String(v).trim());
    }
    if (cells.length === 0) continue;
    if (predicateFn(cells)) return { rowIndex: r, cells };
  }
  return null;
}

async function loadWorkbookAndParse(path) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  const sheet = wb.worksheets[0];

  // Find points header: row that includes "Player" and at least one "Pts" or "R01"/"R02" etc
  const pointsHeader = await findHeaderRow(sheet, (cells) => {
    const lower = cells.map((c) => c.toLowerCase());
    return (
      lower.includes("player") &&
      (lower.includes("pts") || lower.some((x) => /^r\d{2}$/i.test(x)))
    );
  });
  if (!pointsHeader) {
    throw new Error(
      "Could not find points header row (must contain 'Player' and 'Pts' or R01/R02...)"
    );
  }

  // Find spending header: row that includes 'Player' and 'Spent' or '$m' or 'Budget'
  const spendHeader = await findHeaderRow(sheet, (cells) => {
    const joined = cells.join(" ").toLowerCase();
    return (
      joined.includes("player") &&
      (joined.includes("spent") ||
        joined.includes("$m") ||
        joined.includes("budget"))
    );
  });

  if (!spendHeader) {
    console.warn(
      chalk.yellow(
        "Warning: Spending header row not found automatically. Attempting to infer spending table below the points table."
      )
    );
  }

  // Parse points table
  const pHdrRow = pointsHeader.rowIndex;
  // build map of column index to round name starting from header row
  const pHdr = pointsHeader.cells; // array of header cell strings
  // Need to find the column index for "Player" in the sheet
  const sheetPRow = sheet.getRow(pHdrRow);
  let playerCol = null;
  // find first cell that equals 'Player' ignoring case
  for (let c = 1; c <= sheetPRow.cellCount; c++) {
    const raw = sheetPRow.getCell(c).value;
    if (raw && String(raw).trim().toLowerCase() === "player") {
      playerCol = c;
      break;
    }
  }
  if (!playerCol) {
    // fallback: player is column 2
    playerCol = 2;
  }

  // find columns that are rounds (cells after 'Player' column in header)
  const roundCols = [];
  for (let c = playerCol + 1; c <= sheetPRow.cellCount; c++) {
    const v = sheetPRow.getCell(c).value;
    if (!v) continue;
    const txt = String(v).trim();
    // Accept 'Pts', 'R01', 'R1', or any numeric header; but we want to include rounds until the totals column (if present)
    // We'll accept headers that look like R##, Pts, or numbers.
    if (
      /^r\d{1,2}$/i.test(txt) ||
      txt.toLowerCase().startsWith("pts") ||
      /^\d+$/.test(txt) ||
      /^r\d/i.test(txt)
    ) {
      roundCols.push({ col: c, name: txt });
    } else {
      // Some spreadsheets use 'Pts' repeated; also accept any non-empty until we hit a known totals header like 'Points Totals' or 'Points'
      // We'll break only if header contains "Total" or "Points Totals" or "Spent" or "Budget" which indicates end
      const low = txt.toLowerCase();
      if (
        low.includes("total") ||
        low.includes("points totals") ||
        low.includes("spent") ||
        low.includes("budget")
      ) {
        break;
      } else {
        // still accept as round header (defensive)
        roundCols.push({ col: c, name: txt });
      }
    }
  }

  // If no roundCols detected, attempt to take the next ~22 columns
  if (roundCols.length === 0) {
    for (let c = playerCol + 1; c < playerCol + 30; c++) {
      roundCols.push({ col: c, name: "R" + (c - playerCol) });
    }
  }

  // Read points rows until an empty Player cell is encountered (start from next row)
  const players = [];
  for (let r = pHdrRow + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const playerNameCell = row.getCell(playerCol).value;
    if (
      playerNameCell === null ||
      playerNameCell === undefined ||
      String(playerNameCell).trim() === ""
    ) {
      // Might be the separator between top and bottom table. Stop when we hit consecutive blank player rows.
      // We will break only if we encounter more than 4 consecutive blanks just to be safe.
      // Simple approach: break on first blank - because the points table is contiguous.
      break;
    }
    const name = String(playerNameCell).trim();
    // read round scores
    const scores = roundCols.map((rc) => {
      const cell = row.getCell(rc.col).value;
      return normalizePointsCell(cell);
    });
    const totalPoints = scores.reduce((s, x) => s + (Number(x) || 0), 0);
    players.push({
      name,
      scores,
      totalPoints,
      totalSpent: 0,
      spentPerPt: 0,
      tied: false,
    });
  }

  // Parse spending table
  let spendStartRow = spendHeader ? spendHeader.rowIndex : null;

  // If spending header not found, try to search below the points table for a header that contains "$m" or "Spent"
  if (!spendStartRow) {
    // Search rows after pHdrRow for a row with 'Spent' or 'Budget' or '$m'
    for (let r = pHdrRow + 1; r <= pHdrRow + 40 && r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      const txt = [];
      for (let c = 1; c <= Math.min(40, row.cellCount); c++) {
        const v = row.getCell(c).value;
        if (v) txt.push(String(v).toLowerCase());
      }
      const joined = txt.join(" ");
      if (
        joined.includes("$m") ||
        joined.includes("spent") ||
        joined.includes("budget") ||
        joined.includes("spent ($m)")
      ) {
        spendStartRow = r;
        break;
      }
    }
  }

  if (!spendStartRow) {
    console.warn(
      chalk.yellow(
        "Warning: Could not automatically locate spending table header. All player.totalSpent will remain 0."
      )
    );
  } else {
    // Locate player column and 'Spent ($m)' column in spending header row
    const sRow = sheet.getRow(spendStartRow);
    let spendPlayerCol = null;
    let totalSpentCol = null;
    const spendRoundCols = [];

    for (let c = 1; c <= Math.min(80, sRow.cellCount); c++) {
      const raw = sRow.getCell(c).value;
      if (!raw) continue;
      const txt = String(raw).trim().toLowerCase();
      if (txt === "player" && !spendPlayerCol) spendPlayerCol = c;
      if (txt.includes("spent") && totalSpentCol === null) {
        // header could be 'Spent ($m)' or 'Spent ($m) ' etc
        totalSpentCol = c;
      }
      if (
        txt.includes("$m") ||
        /^r\d{1,2}$/i.test(txt) ||
        txt.startsWith("r")
      ) {
        // treat columns with $m or R## as per-round spending
        spendRoundCols.push({ col: c, name: raw });
      }
    }

    // Fallbacks
    if (!spendPlayerCol) spendPlayerCol = 2;
    // Try to locate a "Spent ($m)" text anywhere in the header cells if not found by exact match
    if (!totalSpentCol) {
      for (let c = 1; c <= Math.min(80, sRow.cellCount); c++) {
        const raw = sRow.getCell(c).value;
        if (!raw) continue;
        const txt = String(raw).trim().toLowerCase();
        if (
          txt.includes("spent") ||
          txt.includes("spent ($m)") ||
          txt.includes("$m/pt") ||
          txt.includes("$m/pt")
        ) {
          totalSpentCol = c;
          break;
        }
      }
    }

    // Now read spending rows until blank Player encountered
    for (let r = spendStartRow + 1; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      const pCell = row.getCell(spendPlayerCol).value;
      if (!pCell || String(pCell).trim() === "") break;
      const name = String(pCell).trim();
      let totalSpent = 0;
      // prefer reading explicit totalSpent column if present
      if (totalSpentCol) {
        totalSpent = toNumber(row.getCell(totalSpentCol).value);
      } else {
        // sum per-round spends captured in spendRoundCols
        for (const sc of spendRoundCols) {
          totalSpent += toNumber(row.getCell(sc.col).value);
        }
      }

      // match player by name in players list
      const player = players.find(
        (p) => p.name.trim().toLowerCase() === name.trim().toLowerCase()
      );
      if (player) {
        player.totalSpent = totalSpent;
        // compute spend per pt guard against zero points
        player.spentPerPt =
          player.totalPoints > 0 ? totalSpent / player.totalPoints : Infinity;
      } else {
        // maybe spending table has players not in points table; ignore but optionally log
        // console.warn(`Player "${name}" found in spending table but not in points table.`);
      }
    }
  }

  // Any players without totalSpent set will remain 0
  // Now compute ranking
  players.sort((a, b) => {
    const cmp = comparePlayers(a, b);
    return cmp;
  });

  // After sort, mark groups that are tied (where comparePlayers returned 0 earlier would have left tied = true)
  // We'll also prepare output with rank numbers; for equal totalPoints/spent/countback results we'll assign same rank? Recruiter didn't require equal rank numbering; we will give sequential ranks but mark ties.
  const ranked = players.map((p, idx) => ({
    rank: idx + 1,
    name: p.name,
    totalPoints: p.totalPoints,
    totalSpent: Number(p.totalSpent.toFixed(2)),
    spentPerPt: Number(
      (p.spentPerPt === Infinity ? 0 : p.spentPerPt).toFixed(3)
    ),
    tied: !!p.tied,
    scores: p.scores,
  }));

  return ranked;
}

(async function main() {
  try {
    if (!fs.existsSync(INPUT_PATH)) {
      console.error(chalk.red(`Input file not found at path: ${INPUT_PATH}`));
      process.exit(1);
    }
    console.log(chalk.cyan("Parsing workbook..."));
    const ranked = await loadWorkbookAndParse(INPUT_PATH);

    // Save JSON
    fs.writeFileSync(OUTPUT_JSON, JSON.stringify(ranked, null, 2), "utf8");
    console.log(chalk.green(`Saved ranked leaderboard to ${OUTPUT_JSON}`));

    // Print nicely to console
    console.log(chalk.bold("\nFinal Ranked Leaderboard:"));
    console.log(
      chalk.bold("Rank | Name                     | Points | Spent($m) | Tied")
    );
    console.log(
      chalk.bold("-----|--------------------------|--------|-----------|-----")
    );
    ranked.forEach((r) => {
      const name = r.tied ? chalk.red(r.name.padEnd(24)) : r.name.padEnd(24);
      const tiedFlag = r.tied ? chalk.red("YES") : "NO";
      console.log(
        `${String(r.rank).padEnd(4)} | ${name} | ${String(r.totalPoints).padEnd(
          6
        )} | ${String(r.totalSpent).padEnd(9)} | ${tiedFlag}`
      );
    });

    console.log(chalk.green("\nRanking complete."));
  } catch (err) {
    console.error(chalk.red("Error:"), err);
    process.exit(1);
  }
})();
