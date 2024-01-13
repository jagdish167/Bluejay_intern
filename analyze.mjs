// import { readFile, utils } from 'xlsx';
import pkg from 'xlsx';
const { readFile, utils } = pkg;

// Main function to analyze the Excel file
function analyzeExcelFile() {
    try {
        // Read Excel file
        const workbook = readFile('C:/Users/DELL/OneDrive/Desktop/WD/bluejay_intern/Assignment_Timecard.xlsx');
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = utils.sheet_to_json(worksheet);

        if (!data || data.length === 0) {
            console.error('Error: No data found in the Excel file.');
            return;
        }

        // Assuming the first row is headers
        const headers = Object.keys(data[0]);

        // Find the indices of the relevant columns
        const nameIndex = headers.indexOf('Employee Name');
        const positionIndex = headers.indexOf('Position ID');
        const timecardHoursIndex = headers.indexOf('Timecard Hours (as Time)');

        if (nameIndex === -1 || positionIndex === -1 || timecardHoursIndex === -1) {
            console.error('Error: Required columns not found in the Excel file.');
            return;
        }

        // Iterate through employees and perform analysis
        for (const employee of data) {
            const name = employee['Employee Name'];
            const position = employee['Position ID'];
            const timecardHours = employee['Timecard Hours (as Time)'];

            // Check for each criterion and print relevant information
            if (hasWorkedFor7ConsecutiveDays(employee)) {
                console.log(`Name: ${name}, Position: ${position} - Worked for 7 consecutive days`);
            }

            if (hasLessThan10HoursBetweenShifts(employee)) {
                console.log(`Name: ${name}, Position: ${position} - Less than 10 hours between shifts`);
            }

            if (hasWorkedMoreThan14Hours(employee)) {
                console.log(`Name: ${name}, Position: ${position} - Worked more than 14 hours in a single shift`);
            }
        }
    } catch (error) {
        console.error('Error analyzing the Excel file:', error.message);
    }
}

// Function to calculate the time difference between two shifts in hours
function calculateTimeDifference(start, end) {
    const startTime = new Date(`09/10/2023 ${start}`);
    const endTime = new Date(`09/23/2023 ${end}`);
    const diffInMilliseconds = endTime - startTime;
    return diffInMilliseconds / (1000 * 60 * 60);
}

// Function to convert time in the format hh:mm to total hours
function convertTimeToHours(time) {
    const [hours, minutes] = time.split(':').map(Number);
    return hours + minutes / 60;
}

// Function to check if an employee has worked for 7 consecutive days
function hasWorkedFor7ConsecutiveDays(employee) {
    const timecardHours = employee['Timecard Hours (as Time)'];

    if (typeof timecardHours === 'string') {
        const consecutiveDays = timecardHours.split(',').filter(day => convertTimeToHours(day) > 0).length;
        return consecutiveDays === 7;
    }

    return false;
}

// Function to check if an employee has less than 10 hours between shifts
function hasLessThan10HoursBetweenShifts(employee) {
    const timecardHours = employee['Timecard Hours (as Time)'];
    const convertedHours = timecardHours.split(',').map(convertTimeToHours);

    for (let i = 1; i < convertedHours.length; i++) {
        const timeDifference = calculateTimeDifference(convertedHours[i - 1], convertedHours[i]);
        if (timeDifference < 10 && timeDifference > 1) {
            return true;
        }
    }

    return false;
}

// Function to check if an employee has worked for more than 14 hours in a single shift
function hasWorkedMoreThan14Hours(employee) {
    const timecardHours = employee['Timecard Hours (as Time)'];
    const convertedHours = timecardHours.split(',').map(convertTimeToHours);

    for (const timeWorked of convertedHours) {
        if (timeWorked > 14) {
            return true;
        }
    }

    return false;
}

// Call the main function to start the analysis
analyzeExcelFile();
