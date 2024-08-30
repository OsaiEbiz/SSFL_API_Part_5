import { fileURLToPath } from 'url';
import path from 'path';
import XLSX from 'xlsx';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const excelBufferToJson = (buffer) => {
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    return jsonData;
};

const ensureDirectoryExists = async (dirPath) => {
    await fs.promises.mkdir(dirPath, { recursive: true });
};

const logTime = async (logFilePath, fileName, startTime, endTime) => {
    const logMessage = `File: ${fileName}, Start Time: ${startTime}, End Time: ${endTime}\n`;
    await fs.promises.appendFile(logFilePath, logMessage);
};

export const createContact = async () => {
    try {
        const uploadFolderPath = path.join(__dirname, 'upload');
        const files = await fs.promises.readdir(uploadFolderPath);

        const outputFolderPath = path.join(__dirname, 'doneData');
        await ensureDirectoryExists(outputFolderPath);

        const successfulDataFile = path.join(outputFolderPath, 'successfulData.json');
        const failedDataFile = path.join(outputFolderPath, 'failedData.xlsx');
        const logFile = path.join(outputFolderPath, 'log.txt');
        const timeLogFile = path.join(outputFolderPath, 'timeLog.txt');

        for (let file of files) {
            const startTime = new Date().toISOString();
            console.log(`Processing file: ${file}`);
            const fullFilePath = path.join(uploadFolderPath, file);
            const fileBuffer = await fs.promises.readFile(fullFilePath);
            const data = excelBufferToJson(fileBuffer);

            if (data.length === 0) {
                console.log(`No data found in file: ${file}`);
                continue;
            } else {
                console.log(`Data from file ${file}`);
            }

            for (let element of data) {
                //await fs.promises.appendFile(successfulDataFile, JSON.stringify({ CustomerId: element.CustomerId }) + '\n');
                const gender = element.Gender === "M" ? "Male" : element.Gender === "F" ? "Female" : "";
                const maritalStatus = element.MaritalStatus === "M" ? "Married" : element.MaritalStatus === "W" ? "Widow" : element.MaritalStatus === "S" ? "Single" : "";

                const payload = {
                    "address": element.address,
                    "mobile": parseInt(element.MobileNo) || null,
                    "name": element.CustomerName,
                    "custom_fields": {
                        "customer_id": parseInt(element.CustomerId) || null,
                        "village": element.village,
                        "pincode": parseInt(element.PinCode) || null,
                        "state": element.State,
                        "district": element.District,
                        "client_age": parseInt(element.ClientAge) || null,
                        "kyc_1": element.kyc_1,
                        "id_1": element.Id_1,
                        "kcy_2": element.Kyc_2,
                        "id_2": element.Id_2,
                        "gender": gender,
                        "dob": element.DateOfBirth,
                        "marital_status": maritalStatus,
                        "father_spouse_name": element.FatherSpouseName,
                        "nominee_name": element.NomineeName,
                        "nominee_age": parseInt(element.NomineeAge) || null,
                        "nominee_kyc_id": element.NomineeKycId,
                        "center_name": element.CenterName,
                        "center_id": parseInt(element.CenterId) || null,
                        "group_name": element.GroupName,
                        "group_id": parseInt(element.GroupId) || null,
                        "staff_id": parseInt(element.StaffId) || null,
                        "house_hold_exp": parseInt(element.HouseholdExp) || null,
                        "household_income": parseInt(element.HouseholdIncome) || null,
                        "bank_account_number": element.BankAccountNumber,
                        "bank_name": element.BankName
                    }
                };
                
                try {
                    const response = await fetch('https://spandanasphoorty.freshdesk.com/api/v2/contacts', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'Authorization': 'Basic ' + Buffer.from('nmNdZMqExqsXEVzES:X').toString('base64')
                        },
                        body: JSON.stringify(payload)
                    });
                    const logMessage = `Contact with CustomerId ${element.CustomerId} ${response.ok ? 'created successfully' : 'failed to create'}\n`;
                    await fs.promises.appendFile(logFile, logMessage);
                    if (!response.ok) {
                        await ensureDirectoryExists(path.dirname(failedDataFile));
                        let existingFailedData = [];
                        if (fs.existsSync(failedDataFile)) {
                            const workbook = XLSX.readFile(failedDataFile);
                            const firstSheetName = workbook.SheetNames[0];
                            const worksheet = workbook.Sheets[firstSheetName];
                            existingFailedData = XLSX.utils.sheet_to_json(worksheet);
                        }
                        const newFailedData = [...existingFailedData, element];
                        const newWorkbook = XLSX.utils.book_new();
                        const newWorksheet = XLSX.utils.json_to_sheet(newFailedData);
                        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "FailedData");
                        XLSX.writeFile(newWorkbook, failedDataFile);
                    } else {
                        await ensureDirectoryExists(path.dirname(successfulDataFile));
                        await fs.promises.appendFile(successfulDataFile, JSON.stringify({ CustomerId: element.CustomerId }) + '\n');
                    }
                } catch (error) {
                    console.error(`Error creating contact for CustomerId ${element.CustomerId}: ${error.message}`);
                }
            };
            await fs.promises.unlink(fullFilePath);

            const endTime = new Date().toISOString();
            await logTime(timeLogFile, file, startTime, endTime);
        }
        console.log("Contacts processed and files removed successfully");
        return;
    } catch (error) {
        console.error(`Error in creating contacts: ${error.message}`);
    }
};
