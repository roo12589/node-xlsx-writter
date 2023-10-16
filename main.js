const Excel = require('exceljs')
const fs = require('fs')

// main()
const Mondays = getMondaysOfYear(new Date().getFullYear())
Mondays.forEach(monday => {
    main(monday)
})

/**
 * 
 * @param {Date?} mon 周一的日期
 * @returns {void}
 */
async function main(mon) {

    const data = fs.readFileSync('info.json', 'utf8')
    const info = JSON.parse(data)
    const {
        department,
        position,
        name,
        directLeader
    } = info

    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile('./default.xlsx')
    var worksheets = workbook.worksheets
    // console.log(workbook) 

    const Monday = mon || getCurrentMonday()

    worksheets.forEach((worksheet, index) => {
        worksheet.getRow(2).getCell(3).value = department
        worksheet.getRow(3).getCell(3).value = position
        worksheet.getRow(4).getCell(3).value = name
        worksheet.getRow(5).getCell(3).value = directLeader
        const targetDate = new Date(Monday)
        targetDate.setDate(Monday.getDate() + index)
        console.log(index, targetDate.getDate());
        // timezone offset
        worksheet.getRow(6).getCell(3).value = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 8, 0, 0, 0)
        worksheet.name = `${targetDate.getMonth() + 1}月${targetDate.getDate()}日`

        // worksheet.commit()
    })

    const Friday = new Date(Monday)
    Friday.setDate(Monday.getDate() + 4)
    // const filename = `${Monday.toLocaleDateString()}-${Friday.toLocaleDateString()}`.replace(/\//g, '').replace(new RegExp(new Date().getFullYear(), 'g'), '')

    const startStr = `${Monday.getMonth() + 1 < 10 ? '0' : ''}${Monday.getMonth() + 1}${Monday.getDate() < 10 ? '0' : ''}${Monday.getDate()}`
    const endStr = `${Friday.getMonth() + 1 < 10 ? '0' : ''}${Friday.getMonth() + 1}${Friday.getDate() < 10 ? '0' : ''}${Friday.getDate()}`
    const filename = `${startStr}-${endStr}`
    return workbook.xlsx.writeFile(`./out/《岗位饱和度贡献度分析表-${name}》${filename}.xlsx`)
}

function getCurrentMonday() {
    const d = new Date()
    const date = d.getDate()
    const day = d.getDay() === 0 ? 7 : d.getDay()
    const targetMonday = new Date()
    targetMonday.setDate(date - day + 1)
    targetMonday.setHours(0, 0, 0, 0)
    return targetMonday
}

function getMondaysOfYear(year) {
    const mondays = [];

    // 创建一个 Date 对象，初始日期为所给年份的 1 月 1 日
    const date = new Date(year, 0, 1);

    // 循环遍历全年的日期
    while (date.getFullYear() === year) {
        // 如果当前日期是周一，将其加入结果数组
        if (date.getDay() === 1) {
            mondays.push(new Date(date));
        }

        // 增加一天
        date.setDate(date.getDate() + 1);
    }

    return mondays;
}
