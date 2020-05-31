'use strict';

// Wrap everything in an anonymous function to avoid polluting the global namespace
(function() {
    // Use the jQuery document ready signal to know when everything has been initialized
    $(document).ready(function() {
        // Tell Tableau we'd like to initialize our extension
        tableau.extensions.initializeAsync().then(function() {
            showChooseSheetDialog();
        });
    });

    /**
     * Shows the choose sheet UI. Once a sheet is selected, the data table for the sheet is shown
     */
    function showChooseSheetDialog() {
        // Clear out the existing list of sheets
        $('#choose_sheet_buttons').empty();

        // Set the dashboard's name in the title
        const dashboardName = tableau.extensions.dashboardContent.dashboard.name;
        $('#choose_sheet_title').text(dashboardName);

        // The first step in choosing a sheet will be asking Tableau what sheets are available
        const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;

        var checkBoxes = [];

        worksheets.forEach(function(worksheet){
            
            const checkBox = createCheckBox(worksheet.name);
            checkBoxes.push(checkBox);
            
            // Add checkboxes for the list of worksheets to choose from
            $('#choose_sheets_checkboxes').append(checkBox.HTML);
        })
        
        const saveButton = createButton("Save Sheets");

        $('#choose_sheets_checkboxes').append(saveButton);

        saveButton.click(function(){

            // Loop on worksheet checkboxes and get selected worksheets
            var selectedSheets = []

            checkBoxes.forEach(function(checkBox){
                var sheetTitle = checkBox.title;
                if (document.getElementById(sheetTitle + '_checkbox').checked == true){
                    selectedSheets.push(sheetTitle);
                };
            })
            // Load and export data for selected worksheets
            loadSelectedMarks(selectedSheets);
        })
    }

    function createButton(buttonTitle) {
        const button =
            $(`<button type='button' class='btn btn-default btn-block'>${buttonTitle}</button>`);
        return button;
    }

    function createCheckBox(checkBoxTitle) {
        const checkBoxHTML = $(`<input type="checkbox" id="${checkBoxTitle}_checkbox">
        <label for="${checkBoxTitle}_checkbox">${checkBoxTitle}</label><br>`);
        const checkBox = {'HTML': checkBoxHTML, 'title': checkBoxTitle};
        return checkBox;
    }

    // This variable will save off the function we can call to unregister listening to marks-selected events
    let unregisterEventHandlerFunction;
    
    function loadSelectedMarks(worksheetNames) {
        // Remove any existing event listeners
        if (unregisterEventHandlerFunction) {
            unregisterEventHandlerFunction();
        }
        
        // Get the worksheet object we want to get the selected marks for
        const worksheets = worksheetNames.map(function(worksheetName){
            return getSelectedSheet(worksheetName);
        });

        // Set our title to an appropriate value
        // $('#selected_marks_title').text(worksheet.name);

        // Call to get the selected marks for our sheet
        var getSummaryDataPromises = worksheets.map(function(worksheet){
            return worksheet.getSummaryDataAsync();
        })

        Promise.all(getSummaryDataPromises).then(function(sumDataArray){

            var workbook = XLSX.utils.book_new();

            sumDataArray.forEach(function(sumData, index) {
                // Get the first DataTable for our selected marks (usually there is just one)
                const worksheetData = sumData;
                var worksheetName = worksheetNames[index];

                // Map our data into the format which the data table component expects it
                const data = worksheetData.data.map(function(row, index) {
                    const rowData = row.map(function(cell) {
                        return cell.formattedValue;
                    });

                    return rowData;
                });

                const columns = worksheetData.columns.map(function(column) {
                    return {
                        title: column.fieldName
                    };
                });

                var excelSheet = createAndFillWorksheet(data, columns);

                XLSX.utils.book_append_sheet(workbook, excelSheet, worksheetName);
            });
            
            // Output the Excel workbook
            XLSX.writeFile(workbook, 'output.xlsx');
        });
    }

    function createAndFillWorksheet(data, columns) {

        var sheetData = [];

        sheetData.push(columns.map(function(column){
            return column.title;
        }));

        data.forEach(function(row){
            sheetData.push(row);
        })


        var worksheet = XLSX.utils.aoa_to_sheet(sheetData);

        return worksheet;
    }

    // Save the columns we've applied filters to so we can reset them
    let filteredColumns = [];

    function getSelectedSheet(worksheetName) {
        if (!worksheetName) {
            worksheetName = tableau.extensions.settings.get('sheet');
        }

        // Go through all the worksheets in the dashboard and find the one we want
        return tableau.extensions.dashboardContent.dashboard.worksheets.find(function(sheet) {
            return sheet.name === worksheetName;
        });
    }
})();