const Excel = require("exceljs");
const moment = require("moment");
const template = require("./template.json");
const workbook = new Excel.Workbook();

moment.locale("en");

const dateIndexList = ["C2", "E2", "G2", "I2", "K2", "M2", "O2"];
const workDetailMap = {
    "Monday": {
        "Workload Brief": "C",
        "Hours": "D"
    },
    "Tuesday": {
        "Workload Brief": "E",
        "Hours": "F"
    },
    "Wednesday": {
        "Workload Brief": "G",
        "Hours": "H"
    },
    "Thursday": {
        "Workload Brief": "I",
        "Hours": "J"
    },
    "Friday": {
        "Workload Brief": "K",
        "Hours": "L"
    },
    "Saturday": {
        "Workload Brief": "M",
        "Hours": "N"
    },
    "Sunday": {
        "Workload Brief": "O",
        "Hours": "P"
    },
};

let workBriefTemplate = function(role, sprint){
    let template = `Sprint ${sprint} techinal development`;
    if(role === "BA"){
        template = `Sprint ${sprint} business analysis`;
    }
   return template;
};


function changeWorkHour(newSheet){
    const nameList = newSheet.getColumn("A").values;
    let totalHours = 0;
    nameList.forEach((item, index) => {
        let memberObj = template.member[item];
        if(memberObj !== undefined) {
            for(let i in memberObj) {
                const workBriefIndex = workDetailMap[i]["Workload Brief"] + index;
                const workHourIndex = workDetailMap[i]["Hours"] + index;
                if(memberObj[i] !== "0") {
                    const role = newSheet.getCell(`B${index}`).value;
                    workBriefTemplate(role, template.sprint);
                    newSheet.getCell(workBriefIndex).value = workBriefTemplate(role, template.sprint);
                }else{
                    newSheet.getCell(workBriefIndex).value = "Break";
                }
                totalHours = totalHours + Number(memberObj[i]);
                newSheet.getCell(workHourIndex).value = memberObj[i];
                newSheet.getCell("R9").value = totalHours;
            }
        }
    });
    return newSheet;
}

function changeSummary(newSheet){
    const summaryList = newSheet.getColumn("B").values;
    summaryList.forEach((item, index) => {
        let summaryIndex = `B${index}`;
        newSheet.getCell(summaryIndex).value = item.replace(/\d+/, template.sprint);
    });
    return newSheet;
}

function changeDateCell(newSheet){
    dateIndexList.forEach(item => {
        let dateCellVal = newSheet.getCell(item).value;
        newSheet.getCell(item).value = calculateNextWeekDate(dateCellVal);
    });

    return newSheet;
}

function calculateNextWeekDate(str){
    const [,weekday, date] = str.match(/([^(]+)\s\(([^)]+)/);
    let dateList = date.split("-");
    let newDate = moment().month(dateList[0]).date(Number(dateList[1])).add(7, "d").format("ll");
    newDate = newDate.split(",")[0].replace(" ", "-");
    return weekday + ` (${newDate})`
}

workbook.xlsx.readFile("./weeklyReport.xlsx").then(() => {
    const lastSheetName = workbook["_worksheets"].slice(-1)[0]["name"];
    let lastSheetIndex = 1 + Number(lastSheetName.replace('W',''));

    let wsObj  = workbook.getWorksheet(lastSheetName);
    let newSheet = workbook.addWorksheet(`W${lastSheetIndex}`);

    newSheet.model = Object.assign(wsObj.model, {mergeCells: wsObj.model.merges});
    newSheet.name = `W${lastSheetIndex}`;

    newSheet = changeDateCell(newSheet);
    newSheet = changeWorkHour(newSheet);
    newSheet = changeSummary(newSheet);

    return workbook.xlsx.writeFile('weeklyReport.xlsx');
});
