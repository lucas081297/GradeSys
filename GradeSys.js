//EXTERNAL LIBS
const gauth = require('google-auth-library')
const gspread = require('google-spreadsheet')
const keys = require ('./keys.json')

//MAX PERCENTAGE OF ABSENCE (CAN BE CHANGE EASILY)
const MAX_PERCENT_ABS = 25

//AUTH KEYS
const serviceAccountAuth = new gauth.JWT({
    email: keys.client_email,
    key: keys.private_key,
    scopes: keys.scopes,
})

//ACCESS THE GSPREAD DOC WITH KEYS
const doc = new gspread.GoogleSpreadsheet(keys.table_id,serviceAccountAuth);

//SIMPLE FUNCTION TO CALCULE THE AVERAGE BETWEEN 3 GRADES
function Average (p1,p2,p3){
    return Math.ceil((parseInt(p1) + parseInt(p2) + parseInt(p3))/30)
}

//SIMPLE FUNCTION TO RETURN THE 'NAF' ACCORDING TO AVERAGE
function Naf (m){
    return 10-m
}

//FUNCTION TO LOAD THE USEFUL RANGE OF SPREADSHEET
async function loadSheetCells(range) {
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];
    await sheet.loadCells(range);
    return sheet;
}

//FUNCTION TO GET THE MAX LIMIT ACCORDING TO THE NUMBER OF TOTAL CLASSES (TOTAL CLASS ON SPREADSHEET CAN BE CHANGE EASILY)
const getAbsLimit = async () => {
    const sheet = await loadSheetCells('A2');
    const totalCell = sheet.getCellByA1('A2');
    const total = parseInt(totalCell.value.match(/\d+/g)) / 4;
    return total;
}

//FUNCTION RETURN THE NUMBER OF FIRST AND LAST ROW THAT HAVE A STUDENT DATA
const getDataRange = async () => {
    const sheet = await loadSheetCells('A:A');
    const allACells = await sheet.getCellsInRange('A:A');
    let count = 1;
    for (let cell of allACells) {
        if (cell[0] === 'Matricula') {
            break;
        } else {
            count++;
        }
    }
    return { firstRow: count + 1, lastRow: allACells.length };
}

/* 
MAIN FUNCTION: GET ALL STUDENT DATA, CREATE A STUDENT OBJECT, LOOP THE STUDENTS,
CALCULATE: ABSENCE, AVERAGE AND ‘NAF’, THEN WRITE THE VALUES AND SAVE IN THE SPREADSHEET.
NOTE: FOR BEST PERFORMANCE, I SAVE 15 STUDENTS PER LOOP
*/ 
const calculateAverage = async () =>{
    const DataRange = await getDataRange();
    await doc.loadInfo()
    const sheet = doc.sheetsByIndex[0];
    const absLimit = await getAbsLimit();
    await sheet.loadCells(`A${DataRange.firstRow}:H${DataRange.lastRow}`)
    const studentsGrade = await sheet.getCellsInRange(`A${DataRange.firstRow}:H${DataRange.lastRow}`)
    for (let position in studentsGrade){
        const situationCell =  sheet.getCellByA1(`G${parseInt(position)+DataRange.firstRow}`)
        const nafCell = sheet.getCellByA1(`H${parseInt(position)+DataRange.firstRow}`)
        const abs = studentsGrade[position][2]
        if(abs>absLimit) {
            situationCell.value = 'Reprovado por Falta'
            nafCell.value = '-'
            continue
        }
        const grades = studentsGrade[position].slice(3,6) 
        const avg = Average(...grades)
        if (avg < 5) situationCell.value = 'Reprovado por Nota'
        else if(avg < 7){ 
            situationCell.value = 'Exame Final'
            nafCell.value = Naf(avg)
        }
        else if(avg >= 7) {
            situationCell.value ='Aprovado'
            nafCell.value = '-'
    }
        if(parseInt(position)%15==0 && parseInt(position)!=0) await sheet.saveUpdatedCells();
    }
    await sheet.saveUpdatedCells();
    console.log('Alunos Atualizados!')

}
calculateAverage();