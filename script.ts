import * as ExcelJS from 'exceljs';
import { PrismaClient } from '@prisma/client';

// Define the structure of the Excel data
interface ExcelDataRow {
  channel: number;
  frequency: number;
}

async function readExcelFile(filePath: string): Promise<ExcelDataRow[]> {
    // Create a new ExcelJS Workbook
    const workbook = new ExcelJS.Workbook();
  
    try {
      // Read the Excel file specified by 'filePath'
      await workbook.xlsx.readFile(filePath);
      // Get the first worksheet from the workbook
      const worksheet = workbook.getWorksheet(1);
  
      if (!worksheet) {
        console.error('Error: Worksheet is undefined.');
        return [];
      }
  
      // Initialize an array to store the extracted Excel data
      const excelData: ExcelDataRow[] = [];
  
      // Loop through each row in the worksheet
      worksheet.eachRow((row, rowNumber) => {
            // Extract the cell value from the first column (channel)
            const channelCellValue = row.getCell(1).value;

            // Check if it's not a string // Check if it doesn't start with 'א' // Check if 'א' is not in the string // Check if it doesn't match the pattern 'א' followed by a number
            if (typeof channelCellValue !== 'string' || !channelCellValue.startsWith('א') || channelCellValue.indexOf('א') === -1 || !isValidChannelFormat(channelCellValue)) {
                // Handle the error or invalid case
                console.error(`Error: Invalid channel value at row ${rowNumber}. Skipping this row.`);
                return;
            }
              
            // Extract the channel as a number
            const indexOfAlef = channelCellValue.indexOf('א');
            const channel = parseInt(channelCellValue.substring(indexOfAlef + 1));

            // Extract the frequency cell value and handle empty cells
            const frequencyCell = row.getCell(2);
            if (frequencyCell.type !==( ExcelJS.ValueType.Null)) {
              // Check if the frequency cell is not empty
              const frequency = parseFloat(frequencyCell.text);

              if (!isNaN(frequency) && (Number.isInteger(frequency) || !Number.isNaN(frequency)) ) {
                    // Check if it's a valid frequency (just a number )
                    excelData.push({ channel, frequency });
              } else {
                    // Log an error and skip this row if the frequency is invalid
                    console.error(`Error: Invalid frequency value at row ${rowNumber}. Skipping this row.`);
              }
            } else {
              // Log an error and skip this row if the frequency cell is empty
              console.error(`Error: Empty frequency cell at row ${rowNumber}. Skipping this row.`);
            }
  

        });
  
      // Return the extracted Excel data
      return excelData;
    } catch (error) {
      console.error('Error reading the Excel file:', error);
      return [];
    }
}

// Function to check if the string contains only 'א' and a number after it
function isValidChannelFormat(input: string): boolean {
    // Check if the string contains only 'א' and a number after it
    let foundAleph = false; // A flag to track if 'א' has been found
    for (let i = 0; i < input.length; i++) {
      if (input.charAt(i) === 'א') {
        foundAleph = true; // 'א' has been found
      } else if (foundAleph) {
        if (input.charAt(i) === ' ' || !isNaN(Number(input.charAt(i)))) {
          return false; // Invalid format: contains a space or a valid number after 'א'
        } else {
          return false; // Invalid format: contains other characters after 'א'
        }
      }
    }
    return false; // Invalid format: doesn't contain a number after 'א'
  }
  
  
// Function to upload Excel data to a database using Prisma
async function uploadData(prisma: PrismaClient, excelData: ExcelDataRow[]) {
  try {
    // Loop through the extracted Excel data
    for (const { channel, frequency } of excelData) {
      // Check if a record with the same channel already exists in the database
      const existingRecord = await prisma.excelData.findFirst({
        where: { channel:channel,},
      });

      if (existingRecord) {
        // Update the database record with the new frequency if it exists
        await prisma.excelData.update({
          where: { channel },
          data: { frequency },
        });
      } else {
        // Create a new record in the database if the channel doesn't exist
        await prisma.excelData.create({
          data: { channel, frequency },
        });
      }
    }

    // Log a success message when data is uploaded to the database
    console.log('Data uploaded to the database.');
  } catch (error) {
    // Log an error message if there are issues with data uploading
    console.error('Error updating data:', error);
  }
}

// Function to fetch and display data from the database
async function fetchData(prisma: PrismaClient) {
  try {
    // Fetch all data from the database using Prisma
    const data = await prisma.excelData.findMany();

    // Log the fetched data
    console.log('Fetched Data:', data);
  } catch (error) {
    // Log an error message if there are issues with data fetching
    console.error('Error fetching data:', error);
  }
}

// Main function to orchestrate the program
async function main() {
  const prisma = new PrismaClient();

  try {
    // Read data from the Excel file
    const excelData = await readExcelFile('D:/ExcelToTable.xlsx');
    // Upload Excel data to the database
    await uploadData(prisma, excelData);

    // Fetch and display data from the database
    await fetchData(prisma);

    console.log('The program successfully completed.');
  } catch (error) {
    console.error('Error:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Call the main function to start the program
main();
