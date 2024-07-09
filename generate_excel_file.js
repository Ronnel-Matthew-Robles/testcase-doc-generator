import excel from 'exceljs';
import path from 'path';

const issues = [{
    userStoryNumber: "WCX-5630",
    userStoryTitle: "[WCX-NICE] [Permissions] [Role Permission] Parameters",
    testCases: {
        "epicNumber": "WCX-5313",
        "userStoryNumber": "WCX-5630",
        "testCases": [
          {
            "title": "Verify Role Name and ASG dropdowns appear in Role Permission pop-up",
            "steps": [
              "Login to WCX as a valid user.",
              "Navigate to Permissions section.",
              "Click on Role Permission.",
              "Observe the Role Permission pop-up."
            ],
            "expectedResults": "Role Permission pop-up should display dropdowns for Role Name and Advisor Scheduling Group.",
            "type": "positive"
          },
          {
            "title": "Verify Role Name dropdown placeholder and ality",
            "steps": [
              "Login to WCX as a valid user.",
              "Navigate to Permissions section.",
              "Click on Role Permission.",
              "Observe the Role Name dropdown."
            ],
            "expectedResults": "Role Name dropdown should have a placeholder text 'Select' and an edit icon.",
            "type": "positive"
          },
          {
            "title": "Verify Role Name dropdown displays existing roles",
            "steps": [
              "Login to WCX as a valid user.",
              "Navigate to Permissions section.",
              "Click on Role Permission.",
              "Click on the Role Name dropdown."
            ],
            "expectedResults": "Role Name dropdown should display a list of all existing roles.",
            "type": "positive"
          },
          {
            "title": "Verify ASG dropdown default option and ality",
            "steps": [
              "Login to WCX as a valid user.",
              "Navigate to Permissions section.",
              "Click on Role Permission.",
              "Observe the Advisor Scheduling Group dropdown."
            ],
            "expectedResults": "Advisor Scheduling Group dropdown should have 'All' as a default option.",
            "type": "positive"
          },
          {
            "title": "Verify ASG dropdown displays existing options",
            "steps": [
              "Login to WCX as a valid user.",
              "Navigate to Permissions section.",
              "Click on Role Permission.",
              "Click on the Advisor Scheduling Group dropdown."
            ],
            "expectedResults": "Advisor Scheduling Group dropdown should display a list of all existing options.",
            "type": "positive"
          },
          {
            "title": "Verify no s are added to the grid and add feature",
            "steps": [
              "Login to WCX as a valid user.",
              "Navigate to Permissions section.",
              "Click on Role Permission.",
              "Observe the grid area in the Role Permission pop-up."
            ],
            "expectedResults": "No new s should be added to the grid and add feature.",
            "type": "positive"
          }
        ]
      }
}];

class TestCaseDocGenerator {
    constructor(issues) {
        this.issues = issues;
        this.workbook = new excel.Workbook();
        this.logoId = this.workbook.addImage({
            filename: path.resolve('concentrix_logo.png'),
            extension: 'png',
        });
    }

    async generateTestCaseDoc(filename) {
        this.createGuidelineWorksheet();
        this.createCoverPageWorksheet();
        this.createTestSummaryWorksheet();
        issues.forEach((issue) => {
            this.createUserStorySheet(issue);
        });
        this.createTemplateRevisionHistoryWorksheet();
        this.createFormulasSheet();

        // Save the workbook
        this.workbook.xlsx.writeFile(`output/${filename}`)
        .then(() => {
            console.log("File saved successfully");
        })
        .catch((error) => {
            console.error("Error saving file:", error);
        });
    }

     createGuidelineWorksheet() {
        const guidelineWorksheet = this.workbook.addWorksheet('Guideline', {views: [{showGridLines: false}]});
    
        // Set column widths
        guidelineWorksheet.columns = [
            { width: 2 },  // A
        ];
        
        // Add title
        guidelineWorksheet.mergeCells('B2:J2');
        const guidelineTitleCell = guidelineWorksheet.getCell('B2');
        guidelineTitleCell.value = 'Guideline on Defect Severity and Defect Type';
        guidelineTitleCell.font = { bold: true };
        guidelineTitleCell.alignment = { horizontal: 'center', wrapText: true };
        guidelineTitleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'BFBFBF' }
        };
        
        // Add Defect Severity section
        guidelineWorksheet.mergeCells('B3:J3');
        guidelineWorksheet.getCell('B3').value = 'Defect Severity';
        guidelineWorksheet.getCell('B3').alignment = {wrapText: true}
        guidelineWorksheet.getCell('B3').font = {bold: true}
        
        const severityData = [
            ['Critical', 'Priority 1 defects with no workaround which will prevent further processing of the application and need immediate action.'],
            ['Major', 'Priority 2 defects with an available workaround but affecting a major al requirement.'],
            ['Medium', 'Priority 3 defects which do not impact the ality and have less impact.'],
            ['Minor', 'Priority 4 defects which are cosmetic in nature like UI, Typo, logo or images wrongly placed, internal links not working, etc.']
        ];
        
        guidelineWorksheet.mergeCells('B4:J11');
        
        const richTextValues = severityData.map((item) => ({
            // Each item in the array is a rich text object
            richText: [
                { font: { bold: true }, text: `${item[0]}: ` }, // Severity level in bold
                { font: { bold: false }, text: item[1] } // Description in regular font
            ]
        }));
        
        // Since the cells are merged from B4 to J11, you only need to set the value for the top-left cell of the merged range
        guidelineWorksheet.getCell('B4').value = {
          richText: [
            ...richTextValues[0].richText,
            { text: '\n' }, // New line between each severity level and its description
            ...richTextValues[1].richText,
            { text: '\n' },
            ...richTextValues[2].richText,
            { text: '\n' },
            ...richTextValues[3].richText
          ]
        };
        
        // Ensure the cell where the rich text is set has wrapText enabled to display multi-line text correctly
        guidelineWorksheet.getCell('B4').alignment = { wrapText: true, vertical: 'top' };
            
        // Add Defect Type section
        guidelineWorksheet.mergeCells('B12:J12');
        guidelineWorksheet.getCell('B12').value = 'Defect Type';
        guidelineWorksheet.getCell('B12').alignment = {wrapText: true}
        guidelineWorksheet.getCell('B12').font = {bold: true}
        
        const typeData = [
            ['Unclear Requirements', 'defects related to ambiguous requirements i.e. requirement stated has more than one interpretation.'],
            ['Incomplete Requirements', 'defects caused due to insufficient information related to requirements.'],
            ['Incorrect Requirements', 'defects caused due to requirements being documented incorrectly.'],
            ['Functional', 'defects related to functionality'],
            ['Logic', 'logical defects caused in the system that makes the system to operate incorrectly, but not to terminate abnormally.'],
            ['Design Issues', 'defects caused due to architecture/design issues, incorrect architecture/design solution.'],
            ['Incomplete Test Cases', 'Incomplete requirements coverage, Test cases does not cover non-al requirements.'],
            ['Incorrect Test Cases', 'defects related to faulty test cases.'],
            ['User Interface', 'defects related to characteristics for the human/computer interaction i.e. screen format, validation for user input,  availability, page layout, etc.'],
            ['Coding Standards', 'defects related to non compliance of coding standards.'],
            ['Data Type', 'defects caused due to data inconsistency, data mismatch etc'],
            ['Configuration', 'defects related to version control, change control or any configuration related defect.'],
            ['Incorrect Input/Output Messages', 'defects related to text messages, warning messages , drop down values etc.'],
            ['Documentation', 'Defects related to incorrect documentation'],
            ['Non-al', 'Defects related to non-al requirements such as performance, etc.'],
            ['Others', 'defects that do not fit in the above categories.']
        ];
        
        typeData.forEach((item, index) => {
            // Calculate the row number starting from 13
            const rowNumber = 13 + index;
        
            // Merge cells from B to J for the current row
            guidelineWorksheet.mergeCells(`B${rowNumber}:J${rowNumber}`);
        
            // Set the richText value for the merged cell
            guidelineWorksheet.getCell(`B${rowNumber}`).value = {
                richText: [
                { font: { bold: true, size: 11 }, text: `${item[0]}: ` }, // Defect type in bold
                { font: { bold: false, size: 11 }, text: item[1] } // Description in regular font
                ]
            };
        
            // Ensure the cell has wrapText enabled to display multi-line text correctly
            guidelineWorksheet.getCell(`B${rowNumber}`).alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
        
            if ((`${item[0]}${item[1]}`.length + index) > 100) {
                guidelineWorksheet.getRow(rowNumber).height = 30;
            }
        });
        
        // Add borders
        const borderStyle = { style: 'thick', color: { argb: '000000' } };
        const thinBorderStyle = { style: 'thin', color: { argb: '000000' } };
        
        guidelineWorksheet.getCell('B2').border = {
            top: borderStyle,
            left: borderStyle,
            right: borderStyle
        };
        
        for (let row = 3; row <= 11; row++) {
        ['B', 'C'].forEach(col => {
            guidelineWorksheet.getCell(`${col}${row}`).border = {
            top: thinBorderStyle,
            left: borderStyle,
            bottom: thinBorderStyle,
            right: borderStyle
            };
        });
        }
        
        guidelineWorksheet.getCell('B12').border = {
            top: thinBorderStyle,
            left: borderStyle,
            bottom: borderStyle,
            right: borderStyle
        };
        
        // Apply the thick border to the outer edges of the range from B13 to J28
        // Left and right borders for all rows
        for (let row = 13; row <= 28; row++) {
            guidelineWorksheet.getCell(`B${row}`).border = { left: borderStyle, right: borderStyle };
        }
        
        // Bottom border for the last row
        guidelineWorksheet.getCell('B28').border = {
            left: borderStyle,
            bottom: borderStyle,
            right: borderStyle
        };
    
        return guidelineWorksheet;
    }
    
     createCoverPageWorksheet() {
        const coverPageSheet = this.workbook.addWorksheet('Cover Page', {views: [{showGridLines: false}]});
    
        // Set column widths
        coverPageSheet.columns = [
          { width: 1.5 },
          { width: 9 }, 
          { width: 13 }, 
          { width: 70 },
          { width: 20 }, 
          { width: 20 }, 
          { width: 20 }, 
          { width: 15 }
        ];
        
        // Add CITS/EN/TMP07/V1.0
        coverPageSheet.mergeCells('B2:C4');
        coverPageSheet.getCell('B2').value = 'CITS/EN/TMP07/V1.0';
        coverPageSheet.getCell('B2').font = { bold: true };
        coverPageSheet.getCell('B2').alignment = { vertical: 'middle' };
        coverPageSheet.getCell('B2').border = {
          top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
        
        // Add Test Case Document title
        coverPageSheet.mergeCells('D2:F4');
        const titleCell = coverPageSheet.getCell('D2');
        titleCell.value = 'Test Case Document';
        titleCell.font = { bold: true, size: 16 };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        titleCell.border = {
          top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
    
        coverPageSheet.mergeCells('G2:H4');
        coverPageSheet.getCell('G2').border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
          };
        
        // Add Project Name
        coverPageSheet.mergeCells('B6:C6');
        coverPageSheet.getCell('B6').value = 'Project Name:';
        coverPageSheet.getCell('B6').font = { bold: true };
        coverPageSheet.getCell('B6').border = {
          top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
        coverPageSheet.getCell('B6').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'D9D9D9' }
          };
    
        coverPageSheet.mergeCells('D6:E6');
        const projectNameCell = coverPageSheet.getCell('D6');
        projectNameCell.value = 'WorkforceCX';
        projectNameCell.alignment = { horizontal: 'center', vertical: 'middle' };
        projectNameCell.border = {
          top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
        
        // Add Revision History
        coverPageSheet.mergeCells('B8:G8');
        coverPageSheet.getCell('B8').value = 'Revision History';
        coverPageSheet.getCell('B8').font = { bold: true };
        coverPageSheet.getCell('B8').border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
          };
        coverPageSheet.getCell('B8').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'D9D9D9' }
          };
        
        // Add table headers
        const headers = ['Version', 'Date', 'Change Description', 'Prepared By', 'Reviewed By', 'Approved By'];
        headers.forEach((header, index) => {
          const cell = coverPageSheet.getCell(9, index + 2);
          cell.value = header;
          cell.font = { bold: true };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' } // Yellow background
          };
          cell.border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
          };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        coverPageSheet.getRow(9).height = 25;
        
        // Add table data
        // const tableData = ['1.0', '13-Jun-24', '[WCX-NICE][Timeoff Approval] Disable Buttons When Action is Used', 'Likhitha V', 'Mylie', 'Mylie'];
        const tableData = [];
        tableData.forEach((value, index) => {
          const cell = coverPageSheet.getCell(10, index + 2);
          cell.value = value;
          cell.border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
          };
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        });
        coverPageSheet.getRow(10).height = 25;
        
        // Adjust alignment for the change description cell
        coverPageSheet.getCell('C10').alignment = { horizontal: 'left', vertical: 'middle' };
        
        coverPageSheet.addImage(this.logoId, {
          tl: { col: 6, row: 1 },
          br: { col: 8, row: 4 }
        });
    
        return coverPageSheet;
    }
    
     createTestSummaryWorksheet() {
        const testSummarySheet = this.workbook.addWorksheet('Test Summary', {views: [{showGridLines: false}]});
    
        // Set column widths
        testSummarySheet.columns = [
            { width: 3 }, { width: 70 }, { width: 20 }, { width: 20 }, 
            { width: 85 }, { width: 25 }, { width: 15 }, { width: 15 },
            { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 },
            { width: 15 }
        ];
      
        // Add title
        testSummarySheet.mergeCells('B3:M3');
        testSummarySheet.getCell('B3').value = 'Test Summary Report';
        testSummarySheet.getCell('B3').font = { bold: true };
        testSummarySheet.getCell('B3').alignment = { horizontal: 'center', vertical: 'middle' };
        testSummarySheet.getCell('B3').border = {
            top: {style:'thick'}, left: {style:'thick'}, bottom: {style:'thick'}, right: {style:'thick'}
        };
    
        testSummarySheet.getRow(3).height = 21;
        
        testSummarySheet.addImage(this.logoId, {
            tl: { col: 8, row: 4 },
            br: { col: 10, row: 6 }
        });
    
        testSummarySheet.getCell('B6').value = "Type of Test:";
        testSummarySheet.getCell('B6').font = { bold: true };
        testSummarySheet.getCell('B6').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'BFBFBF' }
        };
        testSummarySheet.getCell('B6').border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
    
        testSummarySheet.getCell('C6').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'BFBFBF' }
        };
        testSummarySheet.getCell('C6').border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
    
        testSummarySheet.getCell('D6').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'BFBFBF' }
        };
        testSummarySheet.getCell('D6').border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
      
        // Add al header
        testSummarySheet.mergeCells('E6:G6');
        testSummarySheet.getCell('E6').value = 'Functional';
        testSummarySheet.getCell('E6').border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
        
        // Add Round headers
        const roundHeaders = ['Round 1', 'Round 2', 'Round 3'];
        roundHeaders.forEach((header, index) => {
            const cell = testSummarySheet.getCell(7, 8 + index * 2);
            testSummarySheet.mergeCells(7, 8 + index * 2, 7, 9 + index * 2);
            cell.value = header;
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: index === 0 ? 'D8E4BC' : (index === 1 ? 'CCC0DA' : 'B7DEE8') }
            };
            cell.font = { bold: true };
            cell.border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
      
        // Add table headers
        const testSummarySheetHeaders = ['Test Case Reference No', 'Epic#', 'UserStory#', 'Test Description', 'Positive/Negative', 'No of Test Steps Planned*', 'No of Test Steps Executed*', 'No of Test Steps Failed*', 'No of Test Steps Executed*', 'No of Test Steps Failed*', 'No of Test Steps Executed*', 'No of Test Steps Failed*'];
        testSummarySheetHeaders.forEach((header, index) => {
            const cell = testSummarySheet.getCell(8, index + 2);
            cell.value = header;
            cell.font = { bold: true };
            let color;
            switch (index) {
                case 6: // Column H
                case 7: // Column I
                    color = 'D8E4BC';
                    break;
                case 8: // Column J
                case 9: // Column K
                    color = 'CCC0DA';
                    break;
                case 10: // Column L
                case 11: // Column M
                    color = 'B7DEE8';
                    break;
                default:
                    color = 'BFBFBF';
            }
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: color }
            };
            cell.border = {
                top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        });
    
        testSummarySheet.getRow(8).height = 75;
        
        // Add data rows
        let rowIndex = 9;
        issues.forEach(issue => {
            testSummarySheet.mergeCells(`B${rowIndex}:M${rowIndex}`);
            testSummarySheet.getCell(`B${rowIndex}`).value = issue.userStoryTitle || 'USER STORY TITLE';
            testSummarySheet.getCell(`B${rowIndex}`).font = { bold: true };
            testSummarySheet.getCell(`B${rowIndex}`).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF00' }
            };
            testSummarySheet.getCell(`B${rowIndex}`).border = {
                top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
            };
    
            rowIndex++;
            issue.testCases.testCases.forEach((testCase, index) => {
                const row = [
                    `TC00${index + 1}`,
                    issue.testCases.epicNumber,
                    issue.testCases.userStoryNumber,
                    testCase.title,
                    testCase.type.charAt(0).toUpperCase() + testCase.type.slice(1),
                    testCase.steps.length,
                    '', '', '', '', '', ''
                ];
                const dataRow = testSummarySheet.getRow(rowIndex);
                row.forEach((value, colIndex) => {
                    const cell = dataRow.getCell(colIndex + 2);
                    cell.value = value;
                    cell.border = {
                        top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
                    };
                    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                    if (colIndex == 3) {
                        cell.alignment = { horizontal: "left" };
                    }
                });
                rowIndex++;
            });
        });
    
        return testSummarySheet;
    }
    
     createUserStorySheet(issue) {
        const sheetName = issue.testCases.userStoryNumber;
        const sheet = this.workbook.addWorksheet(sheetName);
    
        // Set column widths
        sheet.columns = [
            { width: 1.5 }, { width: 10 }, { width: 55 }, { width: 80 },
            { width: 45 }
        ];
    
        let rowIndex = 1;
    
        issue.testCases.testCases.forEach((testCase, testCaseIndex) => {
            // Add Test Case information
            sheet.mergeCells(`B${rowIndex}:C${rowIndex}`);
            sheet.getCell(`B${rowIndex}`).value = 'Test Case No:';
            sheet.getCell(`B${rowIndex}`).font = { bold: true };
            sheet.getCell(`B${rowIndex}`).fill = {  type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
            sheet.getCell(`D${rowIndex}`).value = `TC00${testCaseIndex + 1}`;
            sheet.getCell(`D${rowIndex}`).fill = {  type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
            rowIndex++;
    
            sheet.mergeCells(`B${rowIndex}:C${rowIndex}`);
            sheet.getCell(`B${rowIndex}`).value = 'Test Case Name:';
            sheet.getCell(`B${rowIndex}`).font = { bold: true };
            sheet.getCell(`B${rowIndex}`).fill = {  type: 'pattern', pattern: 'solid', fgColor: { argb: 'c0c0c0' } };
            sheet.getCell(`D${rowIndex}`).value = testCase.title;
            rowIndex++;
    
            sheet.mergeCells(`B${rowIndex}:C${rowIndex}`);
            sheet.getCell(`B${rowIndex}`).value = 'Test Description:';
            sheet.getCell(`B${rowIndex}`).font = { bold: true };
            sheet.getCell(`B${rowIndex}`).fill = {  type: 'pattern', pattern: 'solid', fgColor: { argb: 'c0c0c0' } };
            sheet.getCell(`D${rowIndex}`).value = testCase.title;
            rowIndex++;
    
            sheet.addImage(this.logoId, {
                tl: { col: 4, row: rowIndex - 4 },
                br: { col: 5, row: rowIndex - 1 }
            });
    
            // Add Preconditions
            sheet.mergeCells(`B${rowIndex}:C${rowIndex}`);
            sheet.getCell(`B${rowIndex}`).value = 'Preconditions:';
            sheet.getCell(`B${rowIndex}`).font = { bold: true };
            sheet.getCell(`B${rowIndex}`).fill = {  type: 'pattern', pattern: 'solid', fgColor: { argb: 'c0c0c0' } };
            sheet.getCell(`B${rowIndex}`).alignment = { vertical: 'middle' };
            // sheet.mergeCells(`D${rowIndex}:Q${rowIndex + 4}`);
            sheet.getCell(`D${rowIndex}`).value = testCase.steps.join('\n');
            sheet.getCell(`D${rowIndex}`).alignment = { wrapText: true, vertical: 'top' };
            rowIndex++;
    
            // Add Test Steps header
            sheet.mergeCells(`B${rowIndex}:E${rowIndex}`);
            sheet.getCell(`B${rowIndex}`).value = 'Test Steps';
            sheet.getCell(`B${rowIndex}`).font = { bold: true };
            sheet.getCell(`B${rowIndex}`).alignment = { horizontal: 'center', vertical: 'middle' };
            sheet.getCell(`B${rowIndex}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B7DEE8' } };
    
            sheet.mergeCells(`F${rowIndex}:S${rowIndex}`);
            sheet.getCell(`F${rowIndex}`).value = 'Test Results';
            sheet.getCell(`F${rowIndex}`).font = { bold: true };
            sheet.getCell(`F${rowIndex}`).alignment = { horizontal: 'center', vertical: 'middle' };
            sheet.getCell(`F${rowIndex}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FCD5B4' } };
    
            rowIndex++;
    
            // Add Test Steps table headers
            const headers = ['Step No', 'Action', 'Expected Result', 'Actual Results', 'Round 1', 'Round 2', 'Round 3', 'Defect ID', 'Remarks'];
    
            const merges = [
                { start: 6, end: 7, message: 'Executed By:' },
                { start: 8, end: 9, message: 'Date Executed' },
                { start: 10, end: 11, message: 'Executed By:' },
                { start: 12, end: 13, message: 'Date Executed' },
                { start: 14, end: 15, message: 'Executed By:' },
                { start: 16, end: 17, message: 'Date Executed' },
            ];
    
            const bottomMerges = [
                {column: 6, message: 'Status'},
                {column: 7, message: 'Stage Injected'},
                {column: 8, message: 'Defect Severity'},
                {column: 9, message: 'Defect Type'},
                {column: 10, message: 'Status'},
                {column: 11, message: 'Stage Injected'},
                {column: 12, message: 'Defect Severity'},
                {column: 13, message: 'Defect Type'},
                {column: 14, message: 'Status'},
                {column: 15, message: 'Stage Injected'},
                {column: 16, message: 'Defect Severity'},
                {column: 17, message: 'Defect Type'},
                
            ]; 
    
            [1,2,3].forEach(round => {
                const cell = sheet.getCell(rowIndex, 2 + round * 4);
                cell.value = `Round ${round}`
                cell.font = { bold: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC000' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                sheet.mergeCells(rowIndex, 2 + round * 4, rowIndex + 1, 5 + round * 4);
            });
            
            merges.forEach(({ start, end, message }) => {
                const cell = sheet.getCell(rowIndex + 2, start);
                cell.value = message;
                cell.font = { bold: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC000' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                sheet.mergeCells(rowIndex + 2, start, rowIndex + 2, end);
            });
    
            bottomMerges.forEach(({ column, message }) => {
                const cell = sheet.getCell(rowIndex + 3, column);
                cell.value = message;
                cell.font = { bold: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC000' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            });
    
            sheet.getRow(rowIndex + 2).height = 30;
            sheet.getRow(rowIndex + 3).height = 30;
    
            headers.forEach((header, index) => {
                let columnIndex = index + 2; // Original column index calculation
    
                if (index > 3 && index <= 6) {
                    // Adjust columnIndex based on specific conditions
                    switch (index) {
                        case 4:
                            columnIndex = 6;
                            break;
                        case 5:
                            columnIndex = 10;
                            break;
                        case 6:
                            columnIndex = 14;
                            break;
                    }
                } else if (index > 6) {
                    // Further adjustments for indexes beyond 6
                    switch (index) {
                        case 7:
                            columnIndex = 18;
                            break;
                        case 8:
                            columnIndex = 19;
                            break;
                    }
                    sheet.mergeCells(rowIndex, columnIndex, rowIndex + 3, columnIndex);
                } else {
                    sheet.mergeCells(rowIndex, columnIndex, rowIndex + 3, columnIndex);
                }
    
                const cell = sheet.getCell(rowIndex, columnIndex);
                cell.value = header;
                cell.font = { bold: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC000' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            });
    
            rowIndex += 4;
    
            testCase.steps.forEach((step, index) => {
                const row = sheet.getRow(rowIndex);
                row.getCell(2).value = index + 1;
                row.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };
                row.getCell(3).value = step;
                row.getCell(3).alignment = { vertical: 'middle', horizontal: 'center' };
    
                if (index == testCase.steps.length - 1) {
                    row.getCell(4).value = testCase.expectedResults;
                    row.getCell(4).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                } else {
                    row.getCell(4).value = "";
                }
    
                // Define the border style
                const borderStyle = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
    
                // Apply the border to cells from 5 to 19
                for (let i = 5; i <= 19; i++) {
                    const cell = row.getCell(i);
                    cell.border = borderStyle;
                }
    
                row.height = 30;
                rowIndex++;
            });
    
            // Add some space between test cases
            rowIndex += 2;
        });
    
        // Apply borders to all cells
        sheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                cell.border = {
                    top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
                };
            });
        });
    }
    
     createTemplateRevisionHistoryWorksheet() {
        // Add Template Revision History sheet
        const templateRevisionHistorySheet = this.workbook.addWorksheet('Template Revision History', {views: [{showGridLines: false}]});
    
        // Set column widths
        templateRevisionHistorySheet.columns = [
            { width: 5 }, { width: 12 }, { width: 12 }, { width: 100 }, 
            { width: 20 }, { width: 20 }, { width: 25 }
        ];
    
        // Add header text
        templateRevisionHistorySheet.getCell('B5').value = '<This is template revision history. Delete this sheet while using the document.>';
        templateRevisionHistorySheet.getCell('B5').font = { color: { argb: '8DB4E2' }, italic: false };
    
        templateRevisionHistorySheet.addImage(this.logoId, {
            tl: { col: 6, row: 2 },
            br: { col: 7, row: 4 }
        });
    
        // Add Revision History table
        templateRevisionHistorySheet.mergeCells('B6:H6');
        templateRevisionHistorySheet.getCell('B6').value = 'Revision History';
        templateRevisionHistorySheet.getCell('B6').alignment = { horizontal: 'center', vertical: 'middle' };
        templateRevisionHistorySheet.getCell('B6').font = { bold: true };
        templateRevisionHistorySheet.getCell('B6').border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
    
        const templateRevisionHistorySheetHeaders = ['Revision', 'Date', 'Change Description', 'Prepared By', 'ReviewedBy', 'Approved By', 'Remarks'];
        templateRevisionHistorySheetHeaders.forEach((header, index) => {
        const cell = templateRevisionHistorySheet.getCell(7, index + 2);
        cell.value = header;
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
        });
    
        // Add "All previous version history is managed in master list" row
        templateRevisionHistorySheet.mergeCells('B8:H8');
        templateRevisionHistorySheet.getCell('B8').value = 'All previous version history is managed in master list';
        templateRevisionHistorySheet.getCell('B8').alignment = { horizontal: 'center', vertical: 'middle' };
        templateRevisionHistorySheet.getCell('B8').border = {
        top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
        templateRevisionHistorySheet.getRow(8).height = 30;
    
        // Add data rows
        const data = [
        ['1.0', '13-Jun-24', '[WCX-NICE][Timeoff Approval] Disable Buttons When Action is Used', 'Likhitha V', 'Mylie', 'Mylie', ''],
        ['1.0', '17-Jun-24', '[WCX-NICE] Thai Part 4', 'Li Zhao', 'Mylie', 'Mylie', '']
        ];
    
        data.forEach((row, rowIndex) => {
            row.forEach((value, colIndex) => {
                const cell = templateRevisionHistorySheet.getCell(9 + rowIndex, colIndex + 2);
                cell.value = value;
                cell.alignment = { vertical: 'middle', wrapText: true };
                cell.border = {
                    top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
                };
                cell.alignment.horizontal = 'left';
    
            });
            templateRevisionHistorySheet.getRow(9 + rowIndex).height = 25;
        });
    }
    
     createFormulasSheet() {
        const formulasSheet = this.workbook.addWorksheet('Formulas');
    
       // Initialize counters
        let positiveCount = 0;
        let negativeCount = 0;
    
        // Iterate through issues to count positive and negative test cases
        issues.forEach(issue => {
            issue.testCases.testCases.forEach(testCase => {
                if (testCase.type === 'positive') {
                    positiveCount++;
                } else if (testCase.type === 'negative') {
                    negativeCount++;
                }
            });
        });
    
        // Write summary to Excel
        formulasSheet.addRow(['±ve Test Case Count', '']);
        formulasSheet.addRow(['+ve Test Case', positiveCount]);
        formulasSheet.addRow(['-ve Test Case', negativeCount]);
        formulasSheet.addRow([]);
        formulasSheet.addRow(['Total Test Cases', positiveCount + negativeCount]);
    
        // Apply basic styling
        formulasSheet.columns.forEach(column => {
            column.width = column.header === '±ve Test Case Count' ? 25 : 15;
        });
        formulasSheet.getRow(1).font = { bold: true };
        formulasSheet.getRow(5).font = { bold: true }; 
    }
}

export default TestCaseDocGenerator;