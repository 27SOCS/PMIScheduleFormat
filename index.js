function loadFileAsText() {
  var fileToLoad = document.getElementById("fileToLoad").files[0];

  var fileReader = new FileReader();
  fileReader.onload = function (fileLoadedEvent) {
    var textFromFileLoaded = fileLoadedEvent.target.result;
    parsePMI(textFromFileLoaded);
  };

  fileReader.readAsText(fileToLoad, "UTF-8");
}

let formattedData = [];
let returnCarReg = /\r/g;
let newLineReg = /\n\s*\n/g;
let julianReg = /(\(\s?\d.+\))/g;
let extraInfoReg =
  /^PREPARED.+|BY.+|INPUT.+|PCN.+$|.+INSPECTIONS.+|.+ESR REPORTABLE.+|.+JOB STANDARD.+|.+LOCATION.+WCE.+|.+THIS PMI.+$/gm;

const unitTest = new RegExp(/(UNIT)/);
const equipTest = new RegExp(/(EQUIP DESIG)/);

parsePMI = (PMIData) => {
  // Removing extra info and blank space
  PMIData = PMIData.replace(returnCarReg, "").replace(extraInfoReg, "").replace(newLineReg, "\n");

  // Splitting into array to do more parsing
  PMIData = PMIData.trim().split(/\n/);
  let currentUnit = "";
  let currentEquip = "";
  let currentJob = "";
  let currentLoc = "";

  let u = false;
  let e = false;
  let j = false;
  let l = false;

  for (let i = 0; i < PMIData.length; i++) {
    PMIData[i] = PMIData[i].trim();

    let currentParse = stateMachine(u, e, j, l, PMIData[i]);

    // Uses state machine to determine which parser to choose
    // Updates state afterwards
    switch (currentParse) {
      case 1: // Unit
        currentUnit = unitParse(PMIData[i]);
        u = true;
        e = true;
        j = false;
        l = false;
        break;
      case 2: // Equipment
        currentEquip = equipParse(PMIData[i]);
        u = false;
        e = false;
        j = true;
        l = false;
        break;
      case 3: // Job
        currentJob = jobParse(PMIData[i]);
        u = false;
        e = false;
        j = false;
        l = true;
        break;
      case 4: // Location
        currentLoc = locationParse(PMIData[i]);
        u = true;
        e = true;
        j = true;
        l = false;

        fullJobLine = {
          ...currentUnit,
          ...currentEquip,
          ...currentJob,
          ...currentLoc
        };

        formattedData.push(fullJobLine);
        break;
      default:
        console.log("Must be a bug");
    }
  }

  for (let i = 0; i < formattedData.length; i++) {
    formattedData[i].JOB = formattedData[i].JOB + " " + formattedData[i].JOB2;
    delete formattedData[i].JOB2;
  }

  downloadXLSXFromJson(formattedData, "./PMISchedule.xlsx");
};

stateMachine = (u, e, j, l, data) => {
  // 1 state active, or 0 when program initiating
  if (u + e + j + l <= 1) {
    if (e == 1) return 2;
    if (j == 1) return 3;
    if (l == 1) return 4;
    return 1;
  } else {
    // Multiple states active, which means a location is the last completed line
    if (unitTest.test(data)) return 1;
    if (equipTest.test(data)) return 2;
    return 3;
  }
};

const unitParse = (x) => {
  x = x.replace(julianReg, "").trim().split(" ");
  date = `${x[x.length - 3]} ${x[x.length - 2]} ${x[x.length - 1]}`;

  return { UNIT: x[1], ORGID: x[3], WORKCENTER: x[5], DUEDATE: date };
};

const equipParse = (x) => {
  x = x.replace("EQUIP DESIG:", "").trim();

  return { EQUIPDESIG: x };
};

const jobParse = (x) => {
  equipID = x.split(" ")[0];
  x = x
    .replace(equipID, "")
    .replace(/[-]{2,}/g, "")
    .trim();
  description = x.split("  ")[0];
  x = x.replace(description, "").trim();
  workcenter = x.split(" ")[0];
  x = x.replace(workcenter, "").trim();
  workUnitCode = x.split(" ")[0];
  x = x.replace(workUnitCode, "").trim();
  jst = x.split(" ")[0];
  x = x.replace(jst, "").replace(/(- -)/g, "").trim().split("  ");

  return {
    EQID: equipID,
    DESCRIPTION: description,
    WORKCENTER: workcenter,
    WUC: workUnitCode,
    JST: jst,
    JOB: x[0],
    SYSTEM: x[x.length - 1].trim()
  };
};

const locationParse = (x) => {
  let loc = "";
  let job = "";

  if (x.split("  ").length > 1) {
    loc = x.split("  ")[0];
    last = x.split("  ").length - 1;
    job = x.split("  ")[last].trim();
  } else {
    loc = x.trim();
    job = "";
  }

  return { LOCATION: loc, JOB2: job };
};

downloadXLSXFromJson = (data, name) => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Jobs");

  XLSX.utils.sheet_add_aoa(
    worksheet,
    [["UNIT", "ORG-ID", "WORKCENTER", "DUE DATE", "EQUIP DESIG", "EQUIP-ID"]],
    { origin: "A1" }
  );

  XLSX.writeFile(workbook, name);
};
