import xlsx from "xlsx"
/**
 * Parse the time to string
 * @param {(Object|string|number)} time
 * @param {string} cFormat
 * @returns {string | null}
 */
export function parseTime(time, cFormat) {
  if (arguments.length === 0 || !time) {
    return null
  }
  const format = cFormat || '{y}-{m}-{d} {h}:{i}:{s}'
  let date
  if (typeof time === 'object') {
    date = time
  } else {
    if ((typeof time === 'string')) {
      if ((/^[0-9]+$/.test(time))) {
        // support "1548221490638"
        time = parseInt(time)
      } else {
        // support safari
        // https://stackoverflow.com/questions/4310953/invalid-date-in-safari
        time = time.replace(new RegExp(/-/gm), '/')
      }
    }

    if ((typeof time === 'number') && (time.toString().length === 10)) {
      time = time * 1000
    }
    date = new Date(time)
  }
  const formatObj = {
    y: date.getFullYear(),
    m: date.getMonth() + 1,
    d: date.getDate(),
    h: date.getHours(),
    i: date.getMinutes(),
    s: date.getSeconds(),
    a: date.getDay(),
    S: date.getMilliseconds()
  }
  const time_timer = format.replace(/{([ymdhisSa])+}/g, (result, key) => {
    const value = formatObj[key]
    // Note: getDay() returns 0 on Sunday
    if (key === 'a') { return [vue.$i18n.t('week.sun'), vue.$i18n.t('week.mon'), vue.$i18n.t('week.tue'), vue.$i18n.t('week.wed'), vue.$i18n.t('week.thu'), vue.$i18n.t('week.fri'), vue.$i18n.t('week.sat')][value] }
    return value.toString().padStart(2, '0')
  })
  return time_timer
}

/**
 * 截取指定字符中间的数字
 * @param {*} str 
 * @param {*} startChar 
 * @returns 
 */
export const extractNumberFromStart = (str, startChar) => {
  const startIndex = str.indexOf(startChar) + 1;
  const numStr = str.slice(startIndex).match(/\d+/);
  return numStr ? numStr[0] : null;
}

/**
 * 截取指定字符去掉数字后的字符
 * @param {*} str 
 * @param {*} startChar 
 * @returns 
 */
export const extractStr = (str, startChar) => {
  const startIndex = str.indexOf(startChar) - 1;
  let sliceStr = str.slice(0, startIndex)
  if(str.indexOf('订单号:')) sliceStr += ')'
  return sliceStr
}

/**
 * 导出json excel
 * @param {*} array 
 * @param {*} sheetName 
 * @param {*} fileName 
 * @returns 
 */
export const exportExcelFile = (array, sheetName = '表1', fileName = 'example.xlsx') => {
  const jsonWorkSheet = xlsx.utils.json_to_sheet(array);
  const workBook = {
    SheetNames: [sheetName],
    Sheets: {
      [sheetName]: jsonWorkSheet,
    }
  };
  return xlsx.writeFile(workBook, fileName);
}

/**
 * 判断name in obj
 * @param {*} obj 
 * @param {*} name 
 * @returns 
 */
export const isAttr = (obj, name) => {
  return (name in obj) ? obj[name] : ''
}