// const xlsx = require('xlsx');
// const { find, readXLSX, getDataFromWb } = require('./utils');

const excel = require('exceljs');

const STAFF_NAME_CELL = 'B17';
const CONTENT_ROW_START = 19;
const EQ_ID_COL = 'A';
const STUDENT_NAME_COL = 'B';
const YEAR_LEVEL_COL = 'C';
const SUBJECT_COL = 'D';
const CATEGORY_COL = 'E';
const LEVEL_COL = 'F';
const CONTENT_COL = 'G';
const PROCESS_COL = 'H';
const PRODUCT_COL = 'I';
const ENVIRONMENT_COL = 'J';

async function generateSheets(data) {
    // let template = await readXLSX('./template.xlsx');
    // template_sheet = getDataFromWb(template);
    
    // for (teacher in data) {
    //     if(!template_sheet[STAFF_NAME_CELL]) template_sheet[STAFF_NAME_CELL] = {};
    //     template_sheet[STAFF_NAME_CELL].t = 's';
    //     template_sheet[STAFF_NAME_CELL].v = `${data[teacher].f_name} ${data[teacher].l_name}`;
    //     xlsx.writeFile(template, `./sheets/${data[teacher].f_name} ${data[teacher].l_name} - NCCD.xlsx`);
    // }

    let sheetCount = 0;
    
    for (const teacher in data) {
        let workbook = new excel.Workbook();
        workbook = await workbook.xlsx.readFile('./template.xlsx');
        let worksheet = workbook.getWorksheet('Sheet1');
        worksheet.getCell(STAFF_NAME_CELL).value = `${data[teacher].f_name} ${data[teacher].l_name}`;
        let i = CONTENT_ROW_START;
        let studentsByTeacher = {};
        for (const c in data[teacher].classes) {
            for (const student in data[teacher].classes[c]) {
                if (data[teacher].classes[c][student] !== undefined) {
                    // worksheet.getCell(`${EQ_ID_COL}${i}`).value = student;
                    // worksheet.getCell(`${STUDENT_NAME_COL}${i}`).value = data[teacher].classes[c][student][0];
                    // worksheet.getCell(`${YEAR_LEVEL_COL}${i}`).value = data[teacher].classes[c][student][1];
                    // worksheet.getCell(`${SUBJECT_COL}${i}`).value = c;
                    // i++;
                    studentsByTeacher[student] ? studentsByTeacher[student].classes.push(c) : studentsByTeacher[student] = {
                        'name': data[teacher].classes[c][student][0],
                        'year': data[teacher].classes[c][student][1],
                        'classes': [c]
                    }
                } else {
                    studentsByTeacher[student] ? studentsByTeacher[student].classes.push(c) : studentsByTeacher[student] = {
                        'name': '',
                        'year': '',
                        'classes': [c]
                    }
                }
            }
        }

        if (Object.keys(studentsByTeacher).length > 0) {
            for (const student in studentsByTeacher) {
                worksheet.getCell(`${EQ_ID_COL}${i}`).value = student;
                worksheet.getCell(`${STUDENT_NAME_COL}${i}`).value = studentsByTeacher[student].name;
                worksheet.getCell(`${YEAR_LEVEL_COL}${i}`).value = studentsByTeacher[student].year;
                worksheet.getCell(`${SUBJECT_COL}${i}`).value = studentsByTeacher[student].classes.join(', ');
                worksheet.getCell(`${CATEGORY_COL}${i}`).dataValidation = {
                    type: 'list',
                    allowBlank: true,
                    operator: 'equal',
                    formulae: ['"Physical,Sensory,Cognitive,Social/Emotional"']
                };
                worksheet.getCell(`${LEVEL_COL}${i}`).dataValidation = {
                    type: 'list',
                    allowBlank: true,
                    operator: 'equal',
                    formulae: ['"Differentiated,Supplementary,Substantial"']
                };

                for (const category of [ CONTENT_COL, PROCESS_COL, PRODUCT_COL, ENVIRONMENT_COL ]) {
                    worksheet.getCell(`${category}${i}`).dataValidation = {
                        type: 'list',
                        allowBlank: true,
                        operator: 'equal',
                        formulae: ['"Yes,No"']
                    };
                }

                i++;
            }
    
            workbook.xlsx.writeFile(`./sheets/${data[teacher].f_name} ${data[teacher].l_name} - NCCD.xlsx`);
            sheetCount++;
        } else {
            console.log(':(')
        } 
    }

    console.log(`Generated ${sheetCount} sheets.`);
}

module.exports = { generateSheets }