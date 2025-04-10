const xlsx = require("xlsx");
const fs = require("fs");

const reel_counts = {
  Reel1: {
    el_01: 1,
    el_02: 1,
    el_03: 1,
    el_04: 1,
    el_wild: 1,
    el_scatter: 1,
    el_bonus: 20,
    DD: 20,
    FG: 20,
    AX: 25,
    BX: 25,
    W1: 2,
    W2: 0,
  },
  Reel2: {
    el_01: 1,
    el_02: 0,
    el_03: 1,
    el_04: 1,
    el_wild: 1,
    el_scatter: 1,
    el_bonus: 20,
    DD: 20,
    FG: 20,
    AX: 15,
    BX: 20,
    W1: 0,
    W2: 1,
  },
  Reel3: {
    el_01: 1,
    el_02: 1,
    el_03: 1,
    el_04: 1,
    el_wild: 1,
    el_scatter: 1,
    el_bonus: 5,
    DD: 5,
    FG: 5,
    AX: 4,
    BX: 20,
    W1: 2,
    W2: 0,
  },
  Reel4: {
    el_01: 0,
    el_02: 0,
    el_03: 0,
    el_04: 0,
    el_wild: 0,
    el_scatter: 0,
    el_bonus: 1,
    DD: 1,
    FG: 1,
    AX: 0,
    BX: 0,
    W1: 0,
    W2: 1,
  },
  Reel5: {
    el_01: 0,
    el_02: 1,
    el_03: 0,
    el_04: 0,
    el_wild: 0,
    el_scatter: 0,
    el_bonus: 1,
    DD: 1,
    FG: 1,
    AX: 1,
    BX: 1,
    W1: 3,
    W2: 0,
  },
};

const combinationTable = {
  1: {
    name: "el_01",
    points: [
      {
        count: 5,
        point: 320,
      },
      {
        count: 4,
        point: 80,
      },
      {
        count: 3,
        point: 40,
      },
    ],
  },
  2: {
    name: "el_02",
    points: [
      {
        count: 5,
        point: 320,
      },
      {
        count: 4,
        point: 80,
      },
      {
        count: 3,
        point: 40,
      },
    ],
  },
  3: {
    name: "el_03",
    points: [
      {
        count: 5,
        point: 320,
      },
      {
        count: 4,
        point: 80,
      },
      {
        count: 3,
        point: 40,
      },
    ],
  },
  4: {
    name: "el_04",
    points: [
      {
        count: 5,
        point: 320,
      },
      {
        count: 4,
        point: 80,
      },
      {
        count: 3,
        point: 40,
      },
    ],
  },
  5: {
    name: "el_05",
    points: [
      {
        count: 5,
        point: 800,
      },
      {
        count: 4,
        point: 160,
      },
      {
        count: 3,
        point: 80,
      },
    ],
  },
  6: {
    name: "el_06",
    points: [
      {
        count: 5,
        point: 960,
      },
      {
        count: 4,
        point: 200,
      },
      {
        count: 3,
        point: 80,
      },
    ],
  },
  7: {
    name: "el_07",
    points: [
      {
        count: 5,
        point: 1120,
      },
      {
        count: 4,
        point: 240,
      },
      {
        count: 3,
        point: 80,
      },
    ],
  },
  8: {
    name: "el_08",
    points: [
      {
        count: 5,
        point: 1280,
      },
      {
        count: 4,
        point: 320,
      },
      {
        count: 3,
        point: 80,
      },
    ],
  },
  9: {
    name: "el_09",
    points: [
      {
        count: 5,
        point: 1600,
      },
      {
        count: 4,
        point: 400,
      },
      {
        count: 3,
        point: 120,
      },
    ],
  },
  10: {
    name: "el_wild",
    points: [
      {
        count: 5,
        point: 1600,
      },
      {
        count: 4,
        point: 400,
      },
      {
        count: 3,
        point: 120,
      },
    ],
  },
  11: {
    name: "el_scatter",
    points: [
      {
        count: 5,
        point: 20000,
      },
      {
        count: 4,
        point: 2000,
      },
      {
        count: 3,
        point: 400,
      },
    ],
  },
  12: {
    name: "el_bonus",
    points: [{}],
  },
};

const symbolWithName = {
  el_01: "Line",
  el_02: "Line",
  el_03: "Line",
  el_04: "Line",
  el_05: "A",
  el_06: "K",
  el_07: "Q",
  el_08: "J",
  el_wild: "Wild",
  el_scatter: "Scatter",
  el_bonus: "Bonus",
};

const createReelData = (counts) => {
  let reel_data = [];
  for (let symbol in counts) {
    let count = counts[symbol];
    for (let i = 0; i < count; i++) {
      reel_data.push(symbol);
    }
  }
  return shuffleArray(reel_data);
};

const shuffleArray = (array) => {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
};

let reels = {
  Reel1: createReelData(reel_counts.Reel1).join(","),
  Reel2: createReelData(reel_counts.Reel2).join(","),
  Reel3: createReelData(reel_counts.Reel3).join(","),
  Reel4: createReelData(reel_counts.Reel4).join(","),
  Reel5: createReelData(reel_counts.Reel5).join(","),
};

const createExcelFromReels = (
  reels,
  outputFileName,
  combinationTable,
  symbolWithName
) => {
  // Delete the file if it already exists
  if (fs.existsSync(outputFileName)) {
    fs.unlinkSync(outputFileName);
    console.log(`Existing file ${outputFileName} deleted.`);
  }


  const reelKeys = Object.keys(reels);
  const processedReels = reelKeys.map((reelKey) => reels[reelKey].split(","));
  const maxReelLength = Math.max(...processedReels.map((reel) => reel.length));

  const worksheetData = [
  ["S.No", ...reelKeys],
  ...Array(maxReelLength).fill().map((_, rowIndex) => [
    rowIndex + 1,
    ...processedReels.map((reel) => reel[rowIndex] || ""),
  ]),
];

  const workbook = xlsx.utils.book_new();

  const reelSheet = xlsx.utils.aoa_to_sheet(worksheetData);
  xlsx.utils.book_append_sheet(workbook, reelSheet, "Reels");

  const comboData = [["Symbol", "Count", "Point"]];

  Object.values(combinationTable).forEach(({ name, points }) => {
    if (Array.isArray(points)) {
      points.forEach(({ count, point }) => {
        if (count && point) {
          comboData.push([name, count, point]);
        }
      });
    }
  });

  const comboSheet = xlsx.utils.aoa_to_sheet(comboData);
  xlsx.utils.book_append_sheet(workbook, comboSheet, "CombinationTable");

  const symbolWithKey = [["Symbol key", "symbol Name"]];
  for (const [key, value] of Object.entries(symbolWithName)) {
    symbolWithKey.push([key, value]);
  }

  const symbolKeySheet = xlsx.utils.aoa_to_sheet(symbolWithKey);
  xlsx.utils.book_append_sheet(workbook, symbolKeySheet, "SymbolWithKey");

  // Save the workbook
  xlsx.writeFile(workbook, outputFileName);
  console.log(`Excel file saved as ${outputFileName}`);
};

createExcelFromReels(reels, "reels.xlsx", combinationTable, symbolWithName);
