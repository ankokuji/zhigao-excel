const xlsx = require("node-xlsx");
const _ = require("lodash/fp");
const fs = require("fs");
// const path = require("path");
const [flightDisMap, sectionMap] = require("./config.json");

const sectionAndFlightDistanceSheetName = "片区与航距";

const sheets = xlsx.parse("./source/2018.xlsx");

const filterFlightDistanceSheet = _.filter(sheet => {
  // Filter `"片区与航距"`.
  return sheet.name !== sectionAndFlightDistanceSheetName;
  // Why deep clone?
});

function tap(a) {
  return a;
}

function generateWeekAccumTable() {
  const year = getYearOfTable(sheets)
  const weeklyProcessedSheets = _.compose(
    accumulateByWeek,
    filterFlightDistanceSheet
  )(sheets);
  const buffer = generateMainAndDistrictTable(weeklyProcessedSheets)
  fs.writeFileSync(`./output/output_weekly_${year}.xlsx`, buffer);
}

function getYearOfTable(sheets) {
  const getYearOfRow = _.compose(
    _.head,
    _.split("-"),
    formatDate,
    _.head,
  );

  return _.compose(getYearOfRow, _.nth(1), _.property("data"), _.head)(sheets)
}

function generateMainAndDistrictTable(sheets) {
  const table = generateTotalTable(_.cloneDeep(sheets), true);
  const districtTable = generateDistrictTable(_.cloneDeep(sheets), true);

  const transformDistrictTable = _.compose(
    _.map(([district, data]) => {
      return {
        name: district,
        data
      };
    }),
    _.toPairs
  )(districtTable);

  const combineTables = [
    {
      name: "总体",
      data: table
    },
    ...transformDistrictTable
  ];
  return writeXlsx(combineTables);
}

function main() {
  const buffer = generateMainAndDistrictTable(sheets)
  fs.writeFileSync("./output.xlsx", buffer);
}

// main();
generateWeekAccumTable();

function accumulateByWeek(sheets) {
  sheets = _.cloneDeep(sheets);
  sheets = _.compact(sheets);
  const tableTitle = getOriginalTableTitle(sheets);
  const tableWithoutTitle = _.map(sheet => {
    const processEachSheet = _.reduce(h)({});
    const sheetData = _.compose(
      _.concat([tableTitle]),
      _.map(_.identity),
      processEachSheet,
      _.tail
    )(sheet.data);
    return { name: sheet.name, data: sheetData };
  })(sheets);

  // const table = _.concat([tableTitle])(tableWithoutTitle);
  return tableWithoutTitle;

  function h(accum, row) {
    const rowName = _.head(row);
    if (rowName === "总合计") {
      // Do not handle specific sheets because they are not represent date.
      accum[rowName] = row;
      return accum;
    }
    const year = _.compose(
      _.head,
      _.split("-"),
      formatDate,
      _.head
    )(row);
    const weekNumString = _.compose(
      _.toString,
      getWeek,
      formatDate,
      _.head
    )(row);
    const identification = year + "-" + weekNumString;
    if (!(identification in accum)) {
      // Because
      row[0] = identification;
      accum[identification] = row;
    } else {
      accum[identification] = accumulate2ArrayWithTail(
        accum[identification],
        row
      );
    }
    return accum;
  }
}

/**
 * Sum 2 arrays position by position except head.
 *
 * @param {*} array1
 * @param {*} array2
 * @returns
 */
function accumulate2ArrayWithTail(array1, array2) {
  return _.zipWith((val1, val2) => {
    if (typeof val1 === "number" && typeof val2 === "number") {
      return val1 + val2;
    } else {
      return val1;
    }
  })(array1)(array2);
}

function generateDistrictTable(sheets, shouldHandleExcelDate) {
  const res = _.mapValues(handleEachDistrict)(sectionMap);

  return res;
  /**
   * Generate table from target trip of every district.
   *
   * @param {*} trips
   * @returns
   */
  function handleEachDistrict(trips) {
    const res = _.map(_.compose(extractTargetSheet))(trips);

    return generateTotalTable(res, shouldHandleExcelDate);

    function extractTargetSheet(tripOfDistance) {
      const out = _.find(sheet => {
        const trip = sheet.name;
        const flag = testTripName(trip, name => {
          return name === tripOfDistance;
        });
        return !!flag;
      })(sheets);

      if (!out) {
        // throw new Error("Can't find district!!")
        console.log(`Can't find district!! --- ${tripOfDistance}`);
      }
      return out;
    }
  }
}

/**
 * Generate a statitics table of all sheets with total counts.
 *
 * @param {*} sheets
 * @returns
 */
function generateTotalTable(sheets, shouldHandleExcelDate) {
  sheets = _.compact(sheets);
  const tableTitle = getTableTitleWithAdditionalItem(sheets);
  const accumAll = _.reduce(accumulate)({});
  let generateTableWithoutTitle = _.compose(
    // Delete last column because it was used only in calculating.
    _.map(_.initial),
    _.map(calculate),
    accumAll,
  );
  if (shouldHandleExcelDate) {
    generateTableWithoutTitle = _.compose(generateTableWithoutTitle, _.map(formatDateOfSheet))
  }
  const filterTitleOfSheet = function(sheet) {
    return {
      ...sheet,
      data: _.tail(sheet.data)
    }
  }
  const test = _.compose(
    generateTableWithoutTitle,
    // Filter title of each sheet.
    _.map(filterTitleOfSheet),
    filterFlightDistanceSheet
  )(sheets)
  const table = _.concat([tableTitle])(
    test
  );
  return table;
}

function writeXlsx(tables) {
  return xlsx.build(tables);
}

/**
 * Get table title of final output excel file.
 *
 * @param {*} sheets
 * @returns
 */
function getTableTitleWithAdditionalItem(sheets) {
  return _.compose(
    curryFlip2(_.concat)("航距总和"),
    curryFlip2(_.concat)("统计城市数目"),
    getOriginalTableTitle
  )(sheets);
}

function getOriginalTableTitle(sheets) {
  return _.compose(
    _.head,
    _.property("data"),
    _.head,
  )(sheets);
}

function calculate(sheet) {
  return _.compose(
    calculateRow,
    _.map(_.identity)
  )(sheet);

  function calculateRow(row) {
    //客座率=合计旅客数/提供座位数
    row[6] = row[3] / row[2];

    // 座公里=总合计收入/（提供座位数*航距）
    // row[10] = row[4] / (row[2] * _.last(row)) * (_.nth(-2)(row))
    row[10] = row[4] / _.last(row);
    return row;
  }
}

/**
 * Return a curried and argument reversed version of function `f`.
 *
 * @type :: (a -> b -> c) -> b -> a -> c
 * @param {*} f
 * @returns
 */
function curryFlip2(f) {
  return function(x) {
    return function(y) {
      return f(y)(x);
    };
  };
}

/**
 * Accumulate every sheet position by position group by day.
 *
 * @param {*} accum
 * @param {*} sheet
 * @returns
 */
function accumulate(accum, sheet) {
  // Because we are not using immutable data structure here,
  // usage of reduce is actually unnecessary.
  const accumRes = _.compose(_.reduce(accumEachRow)(accum))(sheet.data);

  function accumEachRow(innerAccum, row) {
    const [date] = row;

    if (!innerAccum[date]) {
      // One for flight count and one for all flight dis.
      const columnNum = row.length + 2;
      const newArray = _.times(_.constant(0))(columnNum);
      newArray[0] = date;
      innerAccum[date] = newArray;
    }

    const flightDistance = getFlightDistance(sheet.name);

    // Append flight count and flight distance to the end of row to accumulate.
    const finalRow = _.compose(
      _.concat(row),
      // Count trip.
      _.concat([1]),
      // 上海航距
      _.concat([flightDistance]),
      // 上海提供座位*上海航距
      _.concat(row[2] * flightDistance)
    )([]);

    const accumedArr = addWithEachPos(finalRow, innerAccum[date]);
    innerAccum[date] = accumedArr;

    return innerAccum;
  }

  return accumRes;
}

/**
 *
 *
 * @param {*} name
 * @param {*} f
 * @returns
 */
function testTripName(name, f) {
  name = treatSpecialName(name);

  const dashSplit = _.split("-")(name);
  let realFlightName = name;

  if (dashSplit.length !== 2) {
    realFlightName = "昆明-" + name;
  }

  let response = f(realFlightName);

  if (!response) {
    response = f(
      _.compose(
        _.join("-"),
        _.reverse,
        _.split("-")
      )(realFlightName)
    );
  }

  return response;
}

/**
 * Get flight distance of specific trip.
 *
 * @param {*} name
 * @returns
 */
function getFlightDistance(name) {
  name = treatSpecialName(name);

  const dashSplit = _.split("-")(name);
  let realFlightName = name;

  if (dashSplit.length !== 2) {
    realFlightName = "昆明-" + name;
  }

  let distance = flightDisMap[realFlightName];

  if (!distance) {
    distance =
      flightDisMap[
        _.compose(
          _.join("-"),
          _.reverse,
          _.split("-")
        )(realFlightName)
      ];
  }

  if (!distance) {
    throw new Error(`Can't get distance from map!!! --- ${realFlightName}`);
  }

  return distance;
}

function treatSpecialName(name) {
  return _.compose(
    _.join("-"),
    _.map(treatSingleCityName),
    _.split("-")
  )(name);

  function treatSingleCityName(name) {
    if (name === "版纳") {
      return "西双版纳";
    }
    return name;
  }
}

/**
 * Add the array from the second item by item.
 *
 * @param {array} xs
 * @param {array} ys
 * @returns
 */
function addWithEachPos(xs, ys) {
  return _.zipWith(sumIfNum)(xs)(ys);

  function sumIfNum(x, y) {
    if (typeof x === "number" && typeof y === "number") {
      return x + y;
    } else {
      return x;
    }
  }
}

/**
 * As the name of function indicates.
 *
 * @param {*} sheet
 * @returns
 */
function formatDateOfSheet(sheet) {
  function formatDateOfRow(row) {
    row[0] = formatDate(row[0]);
    return row;
  }
  const formatDateOfSheetTail = _.compose(
    _.map(formatDateOfRow),
  );

  return {
    name: sheet.name,
    data: formatDateOfSheetTail(sheet.data)
  };
}

/**
 * Format date from excel to "YYYY-MM-DD" pattern.
 *
 * @param {*} dateItem
 * @returns
 */
function formatDate(dateItem) {
  if (_.isString(dateItem)) {
    // Is a string of "YYYY-MM-DD总计" or "总合计".
    const slashSplit = _.split("-")(dateItem);
    if (slashSplit.length === 3) {
      return _.compose(
        _.join("-"),
        deleteLastChinese
      )(slashSplit);
    } else {
      return dateItem;
    }
  } else {
    // Is from excel date type and in js environment will be a number.
    return transformExcelDate(dateItem);
  }

  function deleteLastChinese(arr) {
    const deleted = _.compose(
      _.head,
      _.split("合计"),
      _.last
    )(arr);
    // _.last(arr)
    arr[arr.length - 1] = deleted;
    return arr;
  }
}

/**
 * Should minus 2. Why?
 * Have some bugs.
 *
 * @deprecated
 * @param {*} numb
 * @param {*} format
 * @returns
 */
function unstable_transformExcelDate(numb, format) {
  let time = new Date((numb - 2) * 24 * 3600000 + 1);
  time.setYear(time.getFullYear() - 70);
  let year = time.getFullYear() + "";
  let month = time.getMonth() + 1 + "";
  let date = time.getDate() + "";
  if (format && format.length === 1) {
    return year + format + month + format + date;
  }
  return (
    year +
    "-" +
    (month < 10 ? "0" + month : month) +
    "-" +
    (date < 10 ? "0" + date : date)
  );
}

/**
 * Use this instead.
 *
 * @param {*} excelDate
 * @returns
 */
function transformExcelDate(excelDate) {
  const time = excelDateToJSDate(excelDate);
  const year = time.getFullYear() + "";
  const month = time.getMonth() + 1 + "";
  const date = time.getDate() + "";
  return (
    year +
    "-" +
    (month < 10 ? "0" + month : month) +
    "-" +
    (date < 10 ? "0" + date : date)
  );
}

/**
 * Function from `https://cloud.tencent.com/developer/ask/194095`.
 *
 * @param {*} serial
 * @returns
 */
function excelDateToJSDate(serial) {
  var utc_days = Math.floor(serial - 25569);
  var utc_value = utc_days * 86400;
  var date_info = new Date(utc_value * 1000);
  var fractional_day = serial - Math.floor(serial) + 0.0000001;
  var total_seconds = Math.floor(86400 * fractional_day);
  var seconds = total_seconds % 60;
  total_seconds -= seconds;
  var hours = Math.floor(total_seconds / (60 * 60));
  var minutes = Math.floor(total_seconds / 60) % 60;

  return new Date(
    date_info.getFullYear(),
    date_info.getMonth(),
    date_info.getDate(),
    hours,
    minutes,
    seconds
  );
}

/**
 * Calculate the specified date in the which week of the year.
 *
 * @export
 * @param {string} dt
 * @returns
 */
function getWeek(dt) {
  let d1 = new Date(dt);
  let d2 = new Date(dt);
  d2.setMonth(0);
  d2.setDate(1);
  let rq = d1 - d2;
  let days = Math.ceil(rq / (24 * 60 * 60 * 1000));
  let num = Math.ceil(days / 7);
  return num;
}
