const xlsx = require('xlsx');
const { find, readXLSX, getDataFromWb } = require('./utils');

async function parse() {
    try {
        console.log('Parsing data...');
        let studentList = await find('DynamicStudentList');
        let teacherList = await find('ExportTimetable');
        let classList = await find('ExportStudentClass');

        console.log('...');
    
        studentList = await readXLSX(studentList);
        teacherList = await readXLSX(teacherList);
        classList = await readXLSX(classList);

        console.log('...');
    
        const students = await parseStudents(studentList);
        const teachers = await parseTeachers(teacherList);
        const classes = await parseClasses(classList);

        console.log('...\n');

        const data = await synthesizeData(students, teachers, classes);

        console.log(`Successfully collected ${Object.keys(students).length} students into ${Object.keys(classes).length} classes, sorted into ${Object.keys(teachers).length} teachers!`);

        return data;
    } catch (error) {
        throw (error);
    }
}

function synthesizeData(students, teachers, classes) {
    return new Promise((resolve, reject) => {

        
        for (const c in classes) {
            let temp_class_students = classes[c];
            classes[c] = {};
            
            
            temp_class_students.forEach(v => {
                classes[c][v] = students[v];
            })
        }
        
        
        for (const teacher in teachers) {
            let doesTeacherHaveStudents = false;
            for (const c in teachers[teacher].classes) {
                teachers[teacher].classes[c] = classes[c];
                if (classes[c]) { doesTeacherHaveStudents = true };
            }

            if (!doesTeacherHaveStudents) { 
                delete teachers[teacher];
            }
        }

        resolve(teachers);
    })
}

function parseStudents(workbook) {
    return new Promise((resolve, reject) => {
        const sheetList = workbook.SheetNames;
        const data = workbook.Sheets[sheetList[0]];
        const parsed = xlsx.utils.sheet_to_json(data, {header: 
            [ "Student_Name", "EQ_ID", "Roll_Class", "Year" ]});

        if (!parsed.length) {
            reject(new Error('Couldn\'t resolve Student data.'));
        }
    
        let eqids = {};
    
        // Start off reading from the 5th line
        for (i = 5; i < parsed.length; i++) {
            eqids[[parsed[i].EQ_ID]] = [parsed[i].Student_Name, parsed[i].Year];
        }
    
        resolve(eqids);
    })
}

function parseTeachers(workbook) {
    return new Promise((resolve, reject) => {
        const sheetList = workbook.SheetNames;
        const data = workbook.Sheets[sheetList[0]];
        const parsed = xlsx.utils.sheet_to_json(data, {header: 
            [ "Class_Name",
            "Subject_Name",
            "Subject_Level_Code",
            "Day",
            "Date",
            "Period",
            "Start_Time",
            "End_Time",
            "Staff_Id",
            "Staff_MISID",
            "Staff_Code",
            "Staff_Last_Name",
            "Staff_First_Name",
            "Facility_Code",
            "Facility_Name" ]});

        if (!parsed.length) {
            reject(new Error('Couldn\'t resolve Teacher data.'));
        }
    
        let teachers = {};
    
        for (i = 1; i < parsed.length; i++) {
            if (parsed[i].Staff_MISID === undefined) {
                continue
            }
            if (!teachers[parsed[i].Staff_MISID]) {
                teachers[parsed[i].Staff_MISID] = {
                    classes: {},
                    f_name: parsed[i].Staff_First_Name,
                    l_name: parsed[i].Staff_Last_Name,
                    email: `${parsed[i].Staff_MISID}@eq.edu.au`
                }
            } else {
                teachers[parsed[i].Staff_MISID].classes[parsed[i].Class_Name] = {};
            }

        }
    
        resolve(teachers);
    })
}

function parseClasses(workbook) {
    return new Promise((resolve, reject) => {
        const data = getDataFromWb(workbook);
        const parsed = xlsx.utils.sheet_to_json(data, {header: 
            [ "EQ_ID", "Class_Name" ]});

        if (!parsed.length) {
            reject(new Error('Couldn\'t resolve Class data.'));
        }
    
        let classes = {};
    
        for (i = 1; i < parsed.length; i++) {
            if (!classes[parsed[i].Class_Name]) {
                classes[parsed[i].Class_Name] = [parsed[i].EQ_ID];
            } else {
                classes[parsed[i].Class_Name].push(parsed[i].EQ_ID);
            }
        }
    
        resolve(classes);
    })
}

module.exports = { parse }