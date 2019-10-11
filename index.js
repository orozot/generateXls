const Excel = require("exceljs");
const moment = require("moment");
const workbook = new Excel.Workbook();

moment.locale("en");

const dateIndexList = ["C2", "E2", "G2", "I2", "K2", "M2", "O2"];

workbook.xlsx.readFile("./weeklyReport.xlsx").then(() => {
    const lastSheetName = workbook["_worksheets"].slice(-1)[0]["name"];
    let lastSheetIndex = 1 + Number(lastSheetName.replace('W',''));

    let wsObj  = workbook.getWorksheet(lastSheetName);
    let newSheet = workbook.addWorksheet(`W${lastSheetIndex}`);

    newSheet.model = Object.assign(wsObj.model, {mergeCells: wsObj.model.merges});
    newSheet.name = `W${lastSheetIndex}`;

    dateIndexList.forEach(item => {
        let dateCellVal = newSheet.getCell(item).value;
        const [,weekday, date] = dateCellVal.match(/([^(]+)\s\(([^)]+)/);

        let dateList = date.split("-");
        let newDate = moment().month(dateList[0]).date(Number(dateList[1])).add(7, "d").format("ll");
        newDate = newDate.split(",")[0].replace(" ", "-");
        newSheet.getCell(item).value = weekday + ` (${newDate})`;
    });

    return workbook.xlsx.writeFile('new.xlsx');
});
