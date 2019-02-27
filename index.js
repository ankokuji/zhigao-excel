const xlsx = require('node-xlsx')
const _ = require("lodash/fp")
const fs = require("fs")
const path = require("path")

const sheets = xlsx.parse('./2016.xlsx')
const accumAll = _.reduce(accumulate)({})

function tap(a) {
  debugger
}

function main() {
  const tableTitle = getTableTitle(sheets)
  const res = _.compose(_.map(calculate), accumAll, _.map(formatDateOfSheet))(sheets)
  const table = _.concat([tableTitle])(res)
  const buffer = writeXlsx(table)
  fs.writeFileSync(path.join(__dirname, "output.xlsx"), buffer)
}

main()

function writeXlsx(table) {
  return xlsx.build([{
    name: "out",
    data: table
  }])
}

/**
 * Get table title of final output excel file.
 *
 * @param {*} sheets
 * @returns
 */
function getTableTitle(sheets) {
  return _.compose(curryFlip2(_.concat)("统计城市数目"), _.head, _.property("data"), _.head)(sheets)
}

function calculate(sheet) {
  return _.compose(calculateRow, _.map(_.identity))(sheet)

  function calculateRow(row) {

    //客座率=合计旅客数/提供座位数
    row[6] = row[3] / row[2]

    return row
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
  return function (x) {
    return function (y) {
      return f (y) (x)
    }
  }
}

function accumulate(accum, sheet) {
  // Because we are not using immutable data structure here,
  // usage of reduce is actually unnecessary.
  const accumRes = _.compose(_.reduce(accumEachRow)(accum))(sheet.data)

  function accumEachRow(innerAccum, row) {
    const [date] = row

    if (!innerAccum[date]) {
      const columnNum = row.length + 1
      const newArray = _.times(_.constant(0))(columnNum)
      newArray[0] = date
      innerAccum[date] = newArray
    }

    const accumedArr = addWithEachPos(_.concat(row)(1), innerAccum[date])
    innerAccum[date] = accumedArr

    return innerAccum
  }

  return accumRes
}

/**
 * Add the array from the second item by item.
 *
 * @param {array} xs
 * @param {array} ys
 * @returns
 */
function addWithEachPos(xs, ys) {
  return _.zipWith(sumIfNum)(xs)(ys)

  function sumIfNum(x, y) {
    if (typeof x === "number" && typeof y === "number") {
      return x + y
    } else {
      return x
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
    row[0] = formatDate(row[0])
    return row
  }
  const formatDateOfDataOfSheet = _.compose(_.map(formatDateOfRow), _.tail)

  return {
    data: formatDateOfDataOfSheet(sheet.data)
  }
}

/**
 * Format date from excel to "YYYY-MM-DD" pattern.
 *
 * @param {*} dateItem
 * @returns
 */
function formatDate(dateItem) {

  if(_.isString(dateItem)) {
    // Is a string of "YYYY-MM-DD总计" or "总合计".
    const slashSplit = _.split("-")(dateItem)
    if ((slashSplit).length === 3) {
      return _.compose(_.join("-"), deleteLastChinese)(slashSplit)
    } else {
      return dateItem
    }
  } else {
    // Is from excel date type and in js environment will be a number.
    return transformExcelDate(dateItem)
  }

  function deleteLastChinese(arr) {
    const deleted = _.compose(_.head, _.split("合计"), _.last)(arr)
    // _.last(arr)
    arr[arr.length - 1] = deleted
    return arr
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
  let time = new Date((numb - 2) * 24 * 3600000 + 1)
  time.setYear(time.getFullYear() - 70)
  let year = time.getFullYear() + ''
  let month = time.getMonth() + 1 + ''
  let date = time.getDate() + ''
  if (format && format.length === 1) {
    return year + format + month + format + date
  }
  return year + "-" + (month < 10 ? '0' + month : month) + "-" + (date < 10 ? '0' + date : date)
}

/**
 * Use this instead.
 *
 * @param {*} excelDate
 * @returns
 */
function transformExcelDate(excelDate) {
  const time = excelDateToJSDate(excelDate)
  const year = time.getFullYear() + ''
  const month = time.getMonth() + 1 + ''
  const date = time.getDate() + ''
  return year + "-" + (month < 10 ? '0' + month : month) + "-" + (date < 10 ? '0' + date : date)
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

  return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}