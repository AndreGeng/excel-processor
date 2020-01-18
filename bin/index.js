#!/usr/bin/env node
const xlsx = require("xlsx")

const SHEET_ONE = {
  NAME_IDX: 0,
  VALUE_IDX: 1,
}
const SHEET_TWO = {
  FILENAME_IDX: 0,
  NAME_IDX: 1,
  VALUE_IDX: 2,
}

/**
 * 读取workbook中的第一个sheet, 并返回json数据
 */
const readFirstSheet = (book) => {
  const bookData = xlsx.readFile(book)
  if (bookData.SheetNames.length <= 0) {
    throw new Error(`${book}里没有任何sheets`)
  }
  const sheetRaw = bookData.Sheets[bookData.SheetNames[0]]
  return xlsx.utils.sheet_to_json(sheetRaw)
}
/**
 * 把表格数据转换成以headers中指定的key为键的对象
 * @param {object[]} arr sheet转换出的json对象
 * @param {string} keyHeader 要作为键的字段
 */
const normalizeRow = (arr, keyHeader) => {
  return arr.reduce((acc, row) => {
    const name = row[keyHeader]
    acc[name] = row
    return acc
  }, {})
}
/**
 * 读取指定column里的值
 */
const getValueByCol = (row, idx) => {
  const headers = Object.keys(row)
  const key = headers[idx]
  return row[key]
}
/**
 * 通过header名称，找到当前这列是表格中的第几列
 */
const getIdxByHeader = (arr, header) => {
  if (arr.length <= 0) {
    throw new Error("getIdxByHeader没有需处理的数据")
  }
  const headers = Object.keys(arr[0])
  return headers.indexOf(header)
}

const [book1, book2] = process.argv.slice(2)
const b1s1Json = readFirstSheet(book1)
if (b1s1Json.length <= 0) {
  throw new Error(`${book1}没有可处理的数据`)
}
const b1s1Headers = Object.keys(b1s1Json[0])
const {
  [SHEET_ONE.NAME_IDX]: b1s1NameHeader,
  [SHEET_ONE.VALUE_IDX]: b1s1ValueHeader,
} = b1s1Headers;
const b1s1NormalizedJson = normalizeRow(b1s1Json, b1s1NameHeader)
const b2s1Json = readFirstSheet(book2)
if (b2s1Json.length <= 0) {
  throw new Error(`${book2}没有可处理的数据`)
}
const b2s1Headers = Object.keys(b2s1Json[0])
const {
  [SHEET_TWO.FILENAME_IDX]: b2s1FilenameHeader,
  [SHEET_TWO.NAME_IDX]: b2s1NameHeader,
  [SHEET_TWO.VALUE_IDX]: b2s1ValueHeader,
} = b2s1Headers
const filename = b2s1Json[0][b2s1FilenameHeader]
const resultArr = []
b2s1Json.forEach(row => {
  const name = row[b2s1NameHeader]
  if (!name) {
    return;
  }
  let result;
  // 两个表格都存在的数据，表1减去表2的数值
  if (b1s1NormalizedJson[name]) {
    delete b1s1NormalizedJson[name]["__EMPTY"]
    result = Object.assign({
      [b2s1FilenameHeader]: row[b2s1FilenameHeader],
    }, b1s1NormalizedJson[name], {
      [b2s1ValueHeader]: row[b2s1ValueHeader],
      "结果": Number((getValueByCol(b1s1NormalizedJson[name], SHEET_ONE.VALUE_IDX) - row[b2s1ValueHeader]).toFixed(2)),
    })
    delete b1s1NormalizedJson[name]
  } else {
    // 只有表2存在的数据
    result = Object.assign({
      [b2s1FilenameHeader]: row[b2s1FilenameHeader],
      [b1s1NameHeader]: row[b2s1NameHeader],
      [b1s1ValueHeader]: 0,
    }, row, {
      "结果": Number((0 - row[b2s1ValueHeader]).toFixed(2)),
    })
    // 固定采用表1的name表头
    delete result[b2s1NameHeader]
  }
  resultArr.push(result)
})
// 只有表1存在数据
Object.keys(b1s1NormalizedJson)
  .forEach((name) => {
    const row = b1s1NormalizedJson[name]
    if (!row || !row[name]) {
      return;
    }
    const newRow = Object.assign({
      [b2s1FilenameHeader]: filename,
    }, row, {
      [b2s1ValueHeader]: 0,
      "结果": Number((row[b1s1ValueHeader] - 0).toFixed(2)),
    });
    resultArr.push(newRow)
  })
// 添加表末的计算结果，可用来校对计算是否正确
const dataLength = resultArr.length
const codeA = 'A'.charCodeAt(0)
const b1s1ValueColChar = String.fromCharCode(codeA + getIdxByHeader(resultArr, b1s1ValueHeader))
const b2s1VauleColChar = String.fromCharCode(codeA + getIdxByHeader(resultArr, b2s1ValueHeader))
const resultColChar = String.fromCharCode(codeA + getIdxByHeader(resultArr, "结果"))
resultArr.push({
  [b2s1FilenameHeader]: "原始数据求合相减",
  [b1s1ValueHeader]: {f: `SUM(${b1s1ValueColChar}2:${b1s1ValueColChar}${dataLength + 1})`},
  [b2s1ValueHeader]: {f: `SUM(${b2s1VauleColChar}2:${b2s1VauleColChar}${dataLength + 1})`},
  "结果": {f: `${b1s1ValueColChar}${dataLength + 2}-${b2s1VauleColChar}${dataLength + 2}`}
})
resultArr.push({
  [b2s1FilenameHeader]: "计算结果求合",
  "结果": {f: `SUM(${resultColChar}2:${resultColChar}${dataLength + 1})`}
})

const resultBook = xlsx.utils.book_new()
const sheet = xlsx.utils.json_to_sheet(resultArr)
xlsx.utils.book_append_sheet(resultBook, sheet)
xlsx.writeFile(resultBook, `${filename}.xlsx`)
