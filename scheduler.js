const XLSX = require('xlsx');

// Load and parse the teacher info Excel sheet
function parseTeacherInfo(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet);
}

// Load and parse the date sheet Excel file
function parseDateSheet(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    //console.log(XLSX.utils.sheet_to_json(sheet, { header: 3 }))
    return XLSX.utils.sheet_to_json(sheet,{ range: 2, header:1 });
}

// Check if a teacher is assigned to the course they are teaching
function isTeachingCourse(teacher, courseCode) {
    return teacher.CoursesTaught && teacher.CoursesTaught.split(',').includes(courseCode);
}

// Track the number of duties assigned per day for each teacher
function hasReachedDailyLimit(teacherAssignments, teacherName, date) {
    if (!teacherAssignments[teacherName]) {
        teacherAssignments[teacherName] = {};
    }
    const dailyDuties = teacherAssignments[teacherName][date] || 0;
    return dailyDuties >= 2; // Assuming a daily limit of 2 duties
}

// Increment the number of duties assigned to a teacher on a specific day
function incrementDailyDuties(teacherAssignments, teacherName, date) {
    if (!teacherAssignments[teacherName]) {
        teacherAssignments[teacherName] = {};
    }
    teacherAssignments[teacherName][date] = (teacherAssignments[teacherName][date] || 0) + 1;
}

function assignDuties(teachers, dateSheet) {
    const dutiesPerRank = {
        'Professor': 2,
        'Professor & HOD': 2,
        'Associate Professor': 3,
        'Assistant Professor': 4,
        'Others': 5
    };

    const assignments = [];
    const teacherAssignments = {}; // Keeps track of daily duties assigned per teacher

    console.log(dateSheet.slice(1,2));
    teachers.forEach(teacher => {
        const rank = teacher.Rank;
        const maxDuties = dutiesPerRank[rank] || dutiesPerRank['Others'];
        let dutiesAssigned = 0;

        const timelist=[];
        dateSheet.forEach((row, index) => {
            if (index === 0){
                timelist[0]=row[3]
                timelist[1]=row[5]
                return; // Skip header rows
            }

            const day = row[0]; // Day
            const date = row[1]; // Date
            const courseCode1 = row[2]; // 09:00 - 12:00 Code
            const courseName1 = row[3]; // 09:00 - 12:00 Course Name
            const courseCode2 = row[4]; // 1:00 - 4:00 Code
            const courseName2 = row[5]; // 1:00 - 4:00 Course Name
            const section = row[6] || ''; // Section
            const venue = row[7] || ''; // Venue
            const timeSlots = [
                { code: courseCode1, name: courseName1, time: timelist[0] },
                { code: courseCode2, name: courseName2, time: timelist[1] }
            ];
            
            timeSlots.forEach(slot => {
                const courseCode = slot.code;
                const courseName = slot.name;
                const time = slot.time;
                
                // Skip assignment if teacher teaches the course or has reached daily duty limit
                if (
                    dutiesAssigned < maxDuties &&
                    courseCode && // Ensure course code is not empty
                    !isTeachingCourse(teacher, courseCode) &&
                    !hasReachedDailyLimit(teacherAssignments, teacher.Name, date)
                ) {
                    assignments.push({
                        SrNo: assignments.length + 1,
                        Day: day,
                        Date: date,
                        CourseCode: courseCode,
                        CourseName: courseName,
                        Section: section,
                        No: '', // Empty column for No
                        Venue: venue,
                        Time: time,
                        Invigilator: teacher.Name
                    });

                    dutiesAssigned++;
                    incrementDailyDuties(teacherAssignments, teacher.Name, date);
                }
            });
        });
    });

    return assignments;
}

// Write assignments to an Excel file
function writeAssignmentsToExcel(assignments, filePath) {
    // Title row and header row
    const worksheetData = [
        ['Invigilation List For Courses and Lab Final Exam Spring 2024'],
        [], // Empty row for spacing
        ['Sr. No.', 'Day', 'Date', 'Course Code', 'Course Name', 'Sec.', 'No', 'Venue', 'Time', 'Invigilator']
    ];

    // Add assignment data
    assignments.forEach(assignment => {
        worksheetData.push([
            assignment.SrNo, // Sr. No.
            assignment.Day, // Day
            assignment.Date, // Date
            assignment.CourseCode, // Course Code
            assignment.CourseName, // Course Name
            assignment.Section, // Sec.
            assignment.No, // No
            assignment.Venue, // Venue
            assignment.Time, // Time
            assignment.Invigilator // Invigilator
        ]);
        worksheetData.push([assignment.Invigilator]); // Add Invigilator name on a new line
    });

    // Create worksheet and workbook
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Invigilation List');

    // Write to file
    XLSX.writeFile(workbook, filePath);
}

// Main function to execute the script
function main() {
    const teacherFilePath = './teacher_info.xlsx';
    const dateSheetFilePath = './date_sheet.xlsx';
    const outputFilePath = './assignments.xlsx';

    const teachers = parseTeacherInfo(teacherFilePath);
    //console.log(teachers)
    const dateSheet = parseDateSheet(dateSheetFilePath);
   // console.log(dateSheet.slice(0, 10));
    const assignments = assignDuties(teachers, dateSheet);

    // Output assignments to an Excel file
    writeAssignmentsToExcel(assignments, outputFilePath);
    console.log('Assignments written to Excel successfully!');
}

// Execute the main function
main();
