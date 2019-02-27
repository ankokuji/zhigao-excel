const xlsx = require("node-xlsx");
const _ = require("lodash/fp");
const fs = require("fs");
const path = require("path");

const sheets = xlsx.parse("./2016.xlsx");
const accumAll = _.reduce(accumulate)({});

const sectionAndFlightDistanceSheetName = "片区与航距";

const [flightDisMap, sectionMap] = generateFlightDistanceAndSectionMap(sheets);

/**
 *
 *
 * @param {*} sheets
 * @returns
 */
function generateFlightDistanceAndSectionMap(sheets) {
  const rows = _.find({ name: sectionAndFlightDistanceSheetName })(sheets).data;

  const flightDisMap = _.compose(
    _.fromPairs,
    _.tail,
    _.map(_.takeLast(2))
  )(rows);

  const sectionList = _.compose(
    validSectionList,
    _.map(_.head),
    _.tail
  )(rows);
  const sectionFlightList = _.compose(
    _.compact,
    _.map(_.nth(1)),
    _.tail
  )(rows);

  const sectionMap = _.compose(
    _.mapValues(_.compact),
    _.mapValues(_.map(_.nth(1))),
    _.groupBy(_.head)
  )(_.zip(sectionList, sectionFlightList));

  function validSectionList(list) {
    let currentSection;
    return _.map(section => {
      if (section) {
        currentSection = section;
      }
      return currentSection;
    })(list);
  }

  return [flightDisMap, sectionMap];
}

function tap(a) {
  debugger;
  return a
}

function main() {
  const tableTitle = getTableTitle(sheets);
  const res = _.compose(
    _.map(calculate),
    accumAll,
    _.map(formatDateOfSheet),
    tap,
    _.filter(sheet => {
      return sheet.name !== sectionAndFlightDistanceSheetName;
    })
  )(sheets);
  const table = _.concat([tableTitle])(res);
  const buffer = writeXlsx(table);
  fs.writeFileSync(path.join(__dirname, "output.xlsx"), buffer);
}

main();

function writeXlsx(table) {
  return xlsx.build([
    {
      name: "out",
      data: table
    }
  ]);
}

/**
 * Get table title of final output excel file.
 *
 * @param {*} sheets
 * @returns
 */
function getTableTitle(sheets) {
  return _.compose(
    curryFlip2(_.concat)("航距总和"),
    curryFlip2(_.concat)("统计城市数目"),
    _.head,
    _.property("data"),
    _.head
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
    row[10] = row[4] / (row[2] * _.last(row)) * (_.nth(-2)(row))
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
      _.concat([1]),
      _.concat([flightDistance])
    )([]);

    const accumedArr = addWithEachPos(finalRow, innerAccum[date]);
    innerAccum[date] = accumedArr;

    return innerAccum;
  }

  return accumRes;
}

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
    throw new Error("Can't get distance from map!!!");
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
  const formatDateOfDataOfSheet = _.compose(
    _.map(formatDateOfRow),
    _.tail
  );

  return {
    name: sheet.name,
    data: formatDateOfDataOfSheet(sheet.data)
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
