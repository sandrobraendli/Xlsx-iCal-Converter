// noinspection EqualityComparisonWithCoercionJS

const monthColumn = 0;
// noinspection JSNonASCIINames
const months = {
    "Januar": 1,
    "Februar": 2,
    "März": 3,
    "April": 4,
    "Mai": 5,
    "Juni": 6,
    "Juli": 7,
    "August": 8,
    "September": 9,
    "Oktober": 10,
    "November": 11,
    "Dezember": 12
}

const abbrevRow = 2;
const abbrevStartColumn = 11;

const startRow = abbrevRow + 1;

const dayColumn = 3;

document.getElementById('year').value = new Date().getFullYear();
document.getElementById('fileInput')
    .addEventListener('change', function (e) {
        const reader = new FileReader();
        const file = e.target.files[0];
        reader.readAsArrayBuffer(file);
        reader.onload = (e) => {
            const events = parseRawData(
                firstSheetToArray(e.target.result),
                document.getElementById('year').value,
                document.getElementById('abbrev').value,
                JSON.parse(document.getElementById('shiftDefinitions').value)
            );

            createCalendar(events).download();
        };
    });

function firstSheetToArray(workbookRawData) {
    const workbook = XLSX.read(workbookRawData, {type: 'binary'});
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const csv = XLSX.utils.sheet_to_csv(sheet);
    return CSV.parse(csv);
}

function parseRawData(rawData, year, abbrev, shiftDefinitions) {
    const relevantColumn = getAbbrevColumn(abbrev, rawData);
    const events = [];
    let currentMonth = undefined;

    for (let i = startRow; i < rawData.length; i++) {
        const row = rawData[i];
        const newMonth = months[row[monthColumn]];

        if (newMonth != undefined && newMonth !== currentMonth) {
            currentMonth = newMonth;
        }
        const day = row[dayColumn];
        const shiftType = row[relevantColumn];
        if (currentMonth != undefined && day != undefined && shiftType != undefined) {
            const shiftDefinition = shiftDefinitions[shiftType.toLowerCase()];

            if (shiftDefinition == undefined) {
                console.log("Shift definition not found for " + shiftType);
                continue;
            } else {
                events.push({
                    title: shiftDefinition.name,
                    date: new Date(year, currentMonth - 1, day),
                    start: shiftDefinition.start,
                    end: shiftDefinition.end,
                    overnight: shiftDefinition.overnight
                });
            }
        }

        if (currentMonth === 12 && day === 31) {
            break;
        }
    }
    return events;
}

function getAbbrevColumn(abbrev, rawData) {
    for (let i = abbrevStartColumn; i < rawData.length; i++) {
        if (rawData[abbrevRow][i] === abbrev) {
            return i;
        }
    }
}

function createCalendar(events) {
    const cal = ics();
    events.forEach(event => {
        const startDateString = toIsoString(event.date);
        const endDateString = !event.overnight ? startDateString : toIsoString(new Date(event.date.getTime() + 24 * 60 * 60 * 1000));

        cal.addEvent(event.title, '', '', `${startDateString}T${event.start}`, `${endDateString}T${event.end}`);
    });
    return cal;
}

function toIsoString(date) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}
