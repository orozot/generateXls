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
    let totalHours = 0;
    let memberList = Object.keys(template.member);
    memberList.forEach((name,index) => {
        let _index = index + 5;

        if(newSheet.getCell(`A${_index}`).value === "Delivery brief summary"){
            newSheet.spliceRows(_index,0,[]);
            newSheet.getCell(`B${_index+1}`).value = null;
            newSheet.getCell(`C${_index+1}`).value = null;

            newSheet.getRow(5).eachCell((cell)=>{
                let newRowCellIndex = cell._address.split("")[0]+_index;
                newSheet.getRow(_index).height = newSheet.getRow(5).height;
                newSheet.getCell(newRowCellIndex).style = cell.style;
            })
        }


        let memberData = template.member[name];
        const workData = memberData.workData;
        newSheet.getCell(`A${_index}`).value = name;
        newSheet.getCell(`B${_index}`).value = memberData.role;
        for(let i in workData) {
            const workBriefIndex = workDetailMap[i]["Workload Brief"] + _index;
            const workHourIndex = workDetailMap[i]["Hours"] + _index;
            if(workData[i] !== "0") {
                const role = memberData.role;
                workBriefTemplate(role, template.sprint);
                newSheet.getCell(workBriefIndex).value = workBriefTemplate(role, template.sprint);
            }else{
                newSheet.getCell(workBriefIndex).value = "Break";
            }
            totalHours = totalHours + Number(workData[i]);
            newSheet.getCell(workHourIndex).value = workData[i];
        }

    });
    newSheet.getCell("R9").value = totalHours;
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
