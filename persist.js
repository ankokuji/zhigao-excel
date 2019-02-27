const xlsx = require("node-xlsx");
const _ = require("lodash/fp");
const fs = require("fs");
const path = require("path");

const sheets = xlsx.parse("./2016.xlsx");

const sectionAndFlightDistanceSheetName = "片区与航距";

function persistFlightDistanceAndSectionMap(sheets) {
  const res = generateFlightDistanceAndSectionMap(sheets);
  fs.writeFileSync(path.join(__dirname, "./config.json"), JSON.stringify(res))
}
/**
 *
 *
 * @param {*} sheets
 * @returns
 */
function generateFlightDistanceAndSectionMap(sheets) {
  const rows = _.find({
    name: sectionAndFlightDistanceSheetName
  })(sheets).data;

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

persistFlightDistanceAndSectionMap(sheets)