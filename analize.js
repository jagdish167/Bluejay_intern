// const XLSX = require('xlsx');
// const workbook = XLSX.readFile('C:\Users\DELL\OneDrive\Desktop\WD\bluejay_intern\Assignment_Timecard.xlsx');
// const sheetName = workbook.SheetNames[0];
// const worksheet = workbook.Sheets[sheetName];

// const data = XLSX.utils.sheet_to_json(worksheet);

// // Your analysis logic here
// // ...

// // Print the results
// console.log("Employees who have worked for 7 consecutive days:");
// // Print relevant data

// console.log("\nEmployees with less than 10 hours between shifts:");
// // Print relevant data

// console.log("\nEmployees who have worked for more than 14 hours in a single shift:");
// // Print relevant data
const XLSX = require('xlsx');
const workbook = XLSX.readFile('C:\Users\DELL\OneDrive\Desktop\WD\bluejay_intern\Assignment_Timecard.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

const data = XLSX.utils.sheet_to_json(worksheet);

// Function to calculate the time difference between two shifts in hours
function calculateTimeDifference(start, end) {
    const startTime = new Date(`01/01/2022 ${start}`);
    const endTime = new Date(`01/01/2022 ${end}`);
    const diffInMilliseconds = endTime - startTime;
    return diffInMilliseconds / (1000 * 60 * 60);
}

// Function to check if an employee has worked for 7 consecutive days
function hasWorkedFor7ConsecutiveDays(employee) {
    const consecutiveDays = employee['Timecard Hours (as Time)'].filter(day => day > 0).length;
    return consecutiveDays === 7;
}

// Function to check if an employee has less than 10 hours between shifts
function hasLessThan10HoursBetweenShifts(employee) {
    const timecardHours = employee['Timecard Hours (as Time)'];
    for (let i = 1; i < timecardHours.length; i++) {
        const timeDifference = calculateTimeDifference(timecardHours[i - 1], timecardHours[i]);
        if (timeDifference < 10 && timeDifference > 1) {
            return true;
        }
    }
    return false;
}

// Function to check if an employee has worked for more than 14 hours in a single shift
function hasWorkedMoreThan14Hours(employee) {
    const timecardHours = employee['Timecard Hours (as Time)'];
    for (const time in timecardHours) {
        const timeWorked = parseFloat(time);
        if (timeWorked > 14) {
            return true;
        }
    }
    return false;
}

// Iterate through employees and perform analysis
for (const employee of data) {
    // Check for each criterion and print relevant information
    if (hasWorkedFor7ConsecutiveDays(employee)) {
        console.log(`Name: ${employee['Employee Name']}, Position: ${employee['Position ID']} - Worked for 7 consecutive days`);
    }

    if (hasLessThan10HoursBetweenShifts(employee)) {
        console.log(`Name: ${employee['Employee Name']}, Position: ${employee['Position ID']} - Less than 10 hours between shifts`);
    }

    if (hasWorkedMoreThan14Hours(employee)) {
        console.log(`Name: ${employee['Employee Name']}, Position: ${employee['Position ID']} - Worked more than 14 hours in a single shift`);
    }
}
