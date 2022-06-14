const xlsx = require("xlsx");
const fs = require("fs")

//賞与に対する源泉徴収税額の算出率の表の読み込み
const book = xlsx.readFile('incomeTaxofBonus.xls');
const sheet = book.Sheets['賞与'];

//賞与の金額に乗ずべき率
let rate = [0]
for (let index = 9; index < 36; index++) {
    const address = xlsx.utils.encode_cell({ r: index, c: 1 })
    const cell = sheet[address]

    if (cell !== undefined) {
        rate.push(cell.v)
    }
}

//前月の社会保険料等控除後の給与等の金額(甲)
let amountRange0Person = [{ start: 0, end: 68000 }]
let amountRange1Person = [{ start: 0, end: 94000 }]
let amountRange2Person = [{ start: 0, end: 133000 }]
let amountRange3Person = [{ start: 0, end: 171000 }]
let amountRange4Person = [{ start: 0, end: 210000 }]
let amountRange5Person = [{ start: 0, end: 243000 }]
let amountRange6Person = [{ start: 0, end: 275000 }]
let amountRange7Person = [{ start: 0, end: 308000 }]

const amountRangeListOfTypeA = [amountRange0Person, amountRange1Person, amountRange2Person, amountRange3Person, amountRange4Person, amountRange5Person, amountRange6Person, amountRange7Person]
let amounRangePersonNumber = 0
for (let colIndex = 3; colIndex < 19; colIndex = colIndex + 2) {
    const amountRangePerson = amountRangeListOfTypeA[amounRangePersonNumber]
    for (let rowIndex = 9; rowIndex < 36; rowIndex++) {
        const address1 = xlsx.utils.encode_cell({ r: rowIndex, c: colIndex })
        const address2 = xlsx.utils.encode_cell({ r: rowIndex, c: colIndex + 1 })

        const cell1 = sheet[address1]
        const cell2 = sheet[address2]

        if (cell1 !== undefined && cell2 !== undefined) {
            const vaue1 = cell1.v * 1000
            const vaue2 = Number.isInteger(cell2.v) ? cell2.v * 1000 : Infinity
            amountRangePerson.push({ start: vaue1, end: vaue2 })
        }
    }
    amounRangePersonNumber++;
}

//前月の社会保険料等控除後の給与等の金額(乙)
let amountRangeOfTypeB = [undefined]
for (let rowIndex = 9; rowIndex < 36; rowIndex++) {
    const address0 = xlsx.utils.encode_cell({ r: rowIndex, c: 1 })
    const address1 = xlsx.utils.encode_cell({ r: rowIndex, c: 19 })
    const address2 = xlsx.utils.encode_cell({ r: rowIndex, c: 20 })

    const cell0 = sheet[address0]
    const cell1 = sheet[address1]
    const cell2 = sheet[address2]

    if (cell0 !== undefined && cell1 !== undefined && cell2 !== undefined) {
        const vaue1 = cell1.v * 1000
        const vaue2 = Number.isInteger(cell2.v) ? cell2.v * 1000 : Infinity
        amountRangeOfTypeB.push({ start: vaue1, end: vaue2 })
    } else if (cell0 !== undefined && cell1 === undefined && cell2 === undefined) {
        amountRangeOfTypeB.push(undefined)
    } else if (cell0 !== undefined && cell1 === undefined && cell2 !== undefined) {
        const vaue2 = cell2.v * 1000
        amountRangeOfTypeB.push({start:0, end:vaue2})
    }
}

//ファイル書き込み
const incomeTaxOfBonusJson = {
    rate,
    amountRangeListOfTypeA,
    amountRangeOfTypeB
}
fs.writeFileSync("./incomeTaxOfBonus.json",JSON.stringify(incomeTaxOfBonusJson))