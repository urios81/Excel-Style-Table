// Use an Immediately Invoked Function Expression (IIFE) to avoid polluting global state.
(function() {
    const eventHandlers = [];           // Store event handlers so they can be un-wired later.
    let unmounting = false;             // Used to cancel the checkElement search if the template is being removed from the DOM.
    
    // List to contain dropdowns
    let dropdownList = [];
    const rowsPerPage = 25;  // Number of rows per page
    
    const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    
    // Contain values for popup for each dropdown
    let searchTextList = [];
    let containsButtonList = [];
    let notContainsButtonList = [];
    let selectButtonsList = [];
    let datesList = [];
    let clearFilterButtons = [];
    
    /** Checks if the top-level element exists in the DOM. The DOM will be checked every 100ms.
     *   
     *  resolve -- This function is called once the top-level element is located.
     */
    function checkElement(elementId, resolve) {
        const element = document.getElementById(elementId);
        
        if (element) {
          resolve($(element));
        }
        else if (!unmounting) {
           setTimeout(() => checkElement("Excel_Style_Table", resolve), 100);
        }
    }
    
    /** Wrapper function to wait for the top-level element to exist in the DOM. */
    function waitForTopLevelElement() {
        return new Promise(resolve => checkElement("Excel_Style_Table", resolve));
    }
    
    /** Should be used to add an event handler so that same handler can be un-wired later. */
    function addEventHandler(element, event, handler) {
        $(element).on(event, handler);
        eventHandlers.push({ element, event, handler });
    }
    
    /** Uses a `MutationObserver` to watch for this same script to be removed from the DOM. Once it's removed,
     *  the event handlers which were created are then un-wired.
     */
    function configureCleanup() {
        const observer = new MutationObserver(mutations => {
            const wasScriptRemoved = mutations
                .filter(o => o.type === "childList")
                .flatMap(o => Array.from(o.removedNodes))
                .some(o => o.nodeName === "SCRIPT" && o.getAttribute("id") === "Excel_Style_Table_script");
            
            if (wasScriptRemoved) {
                eventHandlers.forEach(({ element, event, handler }) => {
                    $(element).off(event, handler);
                });
                
                observer.disconnect();
                unmounting = true;
            }
        });
        
        observer.observe(document.body, { childList: true });
    }
    
    function populateLists() {
         dropdownList = Array.from(document.querySelectorAll("#Excel_Style_Table .dropdown"));
        
        for (let i = 0; i < dropdownList.length; i++) {
            let searchText = dropdownList[i].querySelector(".dropdown-menu .filter-section .form-control");
            searchTextList.push(searchText);
        
            let radioButtons = dropdownList[i].querySelectorAll(".dropdown-menu .filter-section input[type='radio']");
            let containsButton = radioButtons[0];
            let notContainsButton = radioButtons[1];
            
            containsButtonList.push(containsButton);
            notContainsButtonList.push(notContainsButton);
            
            let clearFilterButton = dropdownList[i].querySelector("[id*=ClearFilter]");
            clearFilterButtons.push(clearFilterButton);
            
            let selectButtons = dropdownList[i].querySelectorAll(".filter-section .select-dropdown");
            selectButtonsList.push(selectButtons);
            
            // Case where list of buttons for date column
            if (selectButtons[1].classList.contains("parent-dropdown")) {
                let dateMap = new Map();
                
                let yearButton = null;
                let monthButton = null;
                selectButtons.forEach((selectButton, index) => {
                    // Select button is a year
                    if (selectButton.classList.contains("parent-dropdown")) {
                        yearButton = selectButton;
                        
                        dateMap.set(yearButton, new Map());
                    // If month button, add to map
                    } else if (selectButton.classList.contains("first-child-dropdown")) {
                        monthButton = selectButton;
                        
                        dateMap.get(yearButton).set(monthButton, []);
                    // If day button, add to list
                    } else if (selectButton.classList.contains("second-child-dropdown")) {
                        let currentDay = selectButton;
                        
                        dateMap.get(yearButton).get(monthButton).push(currentDay);
                    // Case when select button is (blanks)
                    } else if (index != 0) {
                        dateMap.set(selectButton, '');
                    }
                });
                datesList.push(dateMap);
                
            } else {
                datesList.push(null);
            }
        }
    }
    
    function sortForward() {
        // Get list of rows
        const tbody = document.querySelector(`#Excel_Style_Table_TableBody`);
        const rows = Array.from(tbody.querySelectorAll('tr'));
        
        let currentDropdownId = this.parentNode.parentNode.parentNode.id;
        let index = dropdownList.findIndex(dropdown => dropdown.id === currentDropdownId);
        
        const cellList = rows[0].querySelectorAll("td");
        const cellTextList = Array.from(cellList).map(cell => cell.textContent.trim());
        
        let columnClassList = selectButtonsList[index][1].classList; // Test second select button of column dropdown
        
        rows.sort((rowsA, rowsB) => {
            let cellA = rowsA.cells[index].textContent.trim();
            let cellB = rowsB.cells[index].textContent.trim();
            
            if (cellA == '' && cellB != '') return 1;
            if (cellB == '' && cellA != '') return -1;
            
            // Column is number
            if (Number(cellTextList[index]) || ((cellTextList[index][0] == '$') && Number(cellTextList[index].slice(1)))) {
                cellA = parseFloat(cellA) || parseFloat(cellA.slice(1)) || 0;
                cellB = parseFloat(cellB) || parseFloat(cellB.slice(1)) || 0;
            } else if (columnClassList.contains("parent-dropdown")) { // Column is date column
                cellA = new Date(cellA);
                cellB = new Date(cellB);
            }
            
            return cellA > cellB ? 1 : cellA < cellB ? -1 : 0;
        });
        
        // Add sorted rows to tbody
        rows.forEach(row => tbody.appendChild(row));
        
        // Get first pagination button of table
        let button = document.querySelector("#Excel_Style_Table .pager .btn");
        
        // Call the function to create the new page/tab icons, go to first page
        updatePagination(button);
    }
    
    function sortReverse() {
        // Get list of rows
        const tbody = document.querySelector(`#Excel_Style_Table_TableBody`);
        const rows = Array.from(tbody.querySelectorAll('tr'));
        
        let currentDropdownId = this.parentNode.parentNode.parentNode.id;
        let index = dropdownList.findIndex(dropdown => dropdown.id === currentDropdownId);
        
        const cellList = rows[0].querySelectorAll("td");
        const cellTextList = Array.from(cellList).map(cell => cell.textContent.trim());
        
        let columnClassList = selectButtonsList[index][1].classList; // Test second select button of column dropdown
        
        rows.sort((rowsA, rowsB) => {
            let cellA = rowsA.cells[index].textContent.trim();
            let cellB = rowsB.cells[index].textContent.trim();
            
            if (cellA == '' && cellB != '') return 1;
            if (cellB == '' && cellA != '') return -1;
            
            // Column is number
            if (Number(cellTextList[index]) || ((cellTextList[index][0] == '$') && Number(cellTextList[index].slice(1)))) {
                cellA = parseFloat(cellA) || parseFloat(cellA.slice(1)) || 0;
                cellB = parseFloat(cellB) || parseFloat(cellB.slice(1)) || 0;
            } else if (columnClassList.contains("parent-dropdown")) { // Column is date column
                cellA = new Date(cellA);
                cellB = new Date(cellB);
            }
            
            return cellA < cellB ? 1 : cellA > cellB ? -1 : 0;
        });
        
        // Add sorted rows to tbody
        rows.forEach(row => tbody.appendChild(row));
        
        // Get first pagination button of table
        let button = document.querySelector("#Excel_Style_Table .pager .btn");
        
        // Call the function to create the new page/tab icons, go to first page
        updatePagination(button);
    }
    
    // Uncheck the other filter type button
    function changeFilterType() {
        let buttonId = this.id;
        
        let currentDropdownId = this.parentNode.parentNode.parentNode.parentNode.id;
        let index = dropdownList.findIndex(dropdown => dropdown.id === currentDropdownId);
        
        if (buttonId.includes('Not_Contains')) {
            containsButtonList[index].checked = false;
        } else {
            notContainsButtonList[index].checked = false;
        } 
        
        filterSelectButtons(searchTextList[index]);
    }
    
    function selectFilter() {
    
        let buttonId = this.id;
        let currentSelectButton = this.parentNode;
        let dropdown = currentSelectButton.parentNode.parentNode.parentNode.parentNode;
        let index = dropdownList.indexOf(dropdown);
        
        // boxes only
        let inputBoxList = Array.from(currentSelectButton.parentNode.querySelectorAll(".input-box"));
        
        // Uncheck all boxes if 'Select All' isn't checked
        if (buttonId.includes('Select_All') && !this.checked) {
            
            inputBoxList.forEach(inputBox => {
                inputBox.checked = false;
            });
            
            return; // Don't filter if we only unchecked all
            
        // Check all boxes if 'Select All' button is checked
        } else if (buttonId.includes('Select_All') && this.checked) {
            
            inputBoxList.forEach(inputBox => {
                inputBox.checked = true;
            });
            
        // Get column node ID and uncheck select all button when other submission buttons are not checked
        } else if (!buttonId.includes('Select_All') && !this.checked) {
                
            if (inputBoxList[0].checked) {
                inputBoxList[0].checked = false;
            }
            
        // Get column node and check select all button when *all* other submission buttons are checked
        } else if (!buttonId.includes('Select_All') && this.checked) {
            
            visibleInputBoxList = inputBoxList.filter(inputBox => !inputBox.parentNode.classList.contains("hide-button") && !inputBox.parentNode.classList.contains("filter-button"));
            // Exclude first 'Select All' button
            if (visibleInputBoxList.slice(1).every(inputBox => inputBox.checked)) {
                // Check first 'Select All' button
                inputBoxList[0].checked = true;
            }
        }
        
        // Check if current column is date column
        if (datesList[index] instanceof Map) {
            let parentDiv = currentSelectButton.parentNode;
            let selectButtons = Array.from(parentDiv.querySelectorAll(".dropdown-item")).filter(selectButton => !selectButton.classList.contains("hide-button") && !selectButton.classList.contains("filter-button"));
            
            let yearButton = currentSelectButton;
            
            if (!currentSelectButton.classList.contains("parent-dropdown")) {
                let currentIndex = selectButtons.indexOf(currentSelectButton);
                let elementsList = selectButtons.slice(0, currentIndex);
                
                let yearList = elementsList.filter(selectButton => selectButton.classList.contains("parent-dropdown"));
                
                yearButton = yearList[yearList.length - 1];
            }
            
            let monthButton = currentSelectButton;
            
            if (currentSelectButton.classList.contains("second-child-dropdown")) {
                let currentIndex = selectButtons.indexOf(currentSelectButton);
                let elementsList = selectButtons.slice(0, currentIndex);
                
                let monthList = elementsList.filter(selectButton => selectButton.classList.contains("first-child-dropdown"));
                monthButton = monthList[monthList.length - 1];
            }
            
            let dayButton = currentSelectButton;
            
            // Check all if parent div is checked
            if (currentSelectButton.classList.contains("parent-dropdown") && this.checked) {
                let monthMap = datesList[index].get(yearButton);
                
                for (let [month, dayList] of monthMap) {
                    month.querySelector(".input-box").checked = true;
                    
                    for (let day of dayList) {
                        day.querySelector(".input-box").checked = true;
                    }
                }
            // Uncheck all if parent div is not checked
            } else if (currentSelectButton.classList.contains("parent-dropdown") && !this.checked) {
                let monthMap = datesList[index].get(yearButton);
                
                for (let [month, dayList] of monthMap) {
                    month.querySelector(".input-box").checked = false;
                    
                    for (let day of dayList) {
                        day.querySelector(".input-box").checked = false;
                    }
                }
            // Check all days if month div is checked
            } else if (currentSelectButton.classList.contains("first-child-dropdown") && this.checked) {
                let dayList = datesList[index].get(yearButton).get(monthButton);
                
                for (let day of dayList) {
                    day.querySelector(".input-box").checked = true;
                }
                
                visibleMonths = Array.from(datesList[index].get(yearButton).keys()).filter(month => !month.classList.contains("hide-button") && !month.classList.contains("filter-button"));
                if (visibleMonths.every(month => month.querySelector(".input-box").checked)) {
                    yearButton.querySelector(".input-box").checked = true;
                }
            // uncheck all days if month div is not checked
            } else if (currentSelectButton.classList.contains("first-child-dropdown") && !this.checked) {
                let dayList = datesList[index].get(yearButton).get(monthButton);
                
                for (let day of dayList) {
                    day.querySelector(".input-box").checked = false;
                }
                
                yearButton.querySelector(".input-box").checked = false;
            // Check if all days are checked
            } else if (currentSelectButton.classList.contains("second-child-dropdown") && this.checked) {
                let dayList = datesList[index].get(yearButton).get(monthButton);
                
                // Check relevant month box if all days checked
                visibleDays = dayList.filter(day => !day.classList.contains("hide-button") && !day.classList.contains("filter-button"));
                if (visibleDays.every(day => day.querySelector(".input-box").checked)) {
                    monthButton.querySelector(".input-box").checked = true;
                }
                
                // Check relevant year box if all months checked
                visibleMonths = Array.from(datesList[index].get(yearButton).keys()).filter(month => !month.classList.contains("hide-button") && !month.classList.contains("filter-button"));
                if (visibleMonths.every(month => month.querySelector(".input-box").checked)) {
                    yearButton.querySelector(".input-box").checked = true;
                }
            // Uncheck month/year
            } else if (currentSelectButton.classList.contains("second-child-dropdown") && !this.checked) {
                monthButton.querySelector(".input-box").checked = false;
                yearButton.querySelector(".input-box").checked = false;
            }
            
            // Exclude first 'Select All' button
            if (inputBoxList.slice(1).every(selectButton => selectButton.checked)) {
                // Check first 'Select All' button if all buttons are checked
                inputBoxList[0].checked = true;
            }
        }
        
        applyFilter();
    }
    
    function getValues() {
        // List to contain values to be displayed in select check boxes of each column
        let approvedValuesList = [];
        let checkedValuesList = []; // list of lists
        
        // Iterate through lists of select buttons
        selectButtonsList.forEach((selectButtons, index) => {
            
            // Case where list of buttons for date column
            if (datesList[index] instanceof Map) {
                let checkedDateMap = new Map(); // Hashmap of selected date buttons
                
                for (let [year, monthMap] of datesList[index]) {
                    // Only want unfiltered years
                    if (!year.classList.contains('filter-button')) {
                        let yearText = year.textContent.trim();
                    
                        if (monthMap instanceof Map) {
                            
                            let yearMap = new Map();
                            for (let [month, dayList] of monthMap) {
                                // Only want unfiltered months
                                if (!month.classList.contains('filter-button')) {
                                    let monthText = month.textContent.trim();
                                        
                                    let daysArray = Array.from(dayList).filter(day =>
                                        !day.classList.contains('filter-button') && day.querySelector(".input-box").checked
                                    ).map(day => day.textContent.trim());
                                    
                                    if (daysArray.length > 0) {
                                        yearMap.set(monthText, daysArray);
                                    }
                                }
                            }
                            if (yearMap.size > 0) {
                                checkedDateMap.set(yearText, yearMap);
                            }
                        } else {
                            if (year.querySelector(".input-box").checked) {
                                checkedDateMap.set("(Blanks)", '');
                            }
                        }
                    }
                }
                checkedValuesList.push(checkedDateMap);
            } else {
                // List of checked boxes text
                let checkedOptionsList = Array.from(selectButtons).filter(selectButton =>
                    selectButton.querySelector('.input-box').checked && !selectButton.classList.contains('filter-button')
                ).map(button => button.textContent.trim());
    
                if (checkedOptionsList[checkedOptionsList.length - 1] == '(Blanks)') {
                    checkedOptionsList[checkedOptionsList.length - 1] = '';
                }
                
                checkedValuesList.push(checkedOptionsList);
            }
            
            approvedValuesList.push(new Set());
        })
        
        return [approvedValuesList, checkedValuesList];
    }
    
    function updateSelectButtons(approvedValuesList, searchText) {
        const monthMapping = {
            "January": "01",
            "February": "02",
            "March": "03",
            "April": "04",
            "May": "05",
            "June": "06",
            "July": "07",
            "August": "08",
            "September": "09",
            "October": "10",
            "November": "11",
            "December": "12"
        }
        
        // Show/Hide select buttons based on approved values list
        selectButtonsList.forEach((selectButtons, index) => {
            
            // Approved values were dates for this column
            if (datesList[index] instanceof Map) {
                
                for (let [year, monthMap] of datesList[index]) {
                    // Check if 'monthMap' is a map (else it's '' because of a blank cell value)
                    if (monthMap instanceof Map) {
                        let yearText = year.textContent.trim();
                        
                        // Show/Hide Dates buttons
                        for (let [month, dayList] of monthMap) {
                            let monthText = monthMapping[month.textContent.trim()];
                            
                            // Show/Hide Day Buttons
                            dayList.forEach(dayButton => {
                                let date = yearText + '-' + monthText + '-' + dayButton.textContent.trim();
                                
                                // Check approved values list
                                if (approvedValuesList[index].has(date)) {
                                    dayButton.classList.remove('hide-button')
                                } else {
                                    dayButton.classList.add('hide-button');
                                }
                            });
                            
                            // Show/Hide Month Buttons
                            dayVisibleCount = datesList[index].get(year).get(month).filter(dayButton => {
                                return !dayButton.classList.contains('hide-button') && !dayButton.classList.contains('filter-button');
                            }).length;
                            
                            if (dayVisibleCount > 0) {
                                month.classList.remove('hide-button');
                            } else {
                                month.classList.add('hide-button');
                            }
                        }
                        
                        // Show/Hide Year buttons
                        monthVisibleCount = Array.from(datesList[index].get(year).keys()).filter(monthButton => {
                            return !monthButton.classList.contains('hide-button') && !monthButton.classList.contains('filter-button');
                        }).length;
                        
                        if (monthVisibleCount > 0) {
                            year.classList.remove('hide-button');
                        } else {
                            year.classList.add('hide-button');
                        }
                    } else {
                        // Show/Hide (Blanks) button
                        if (approvedValuesList[index].has('')) {
                            year.classList.remove('hide-button');
                        } else {
                            year.classList.add('hide-button');
                        }
                        
                        continue;
                    }
                }
            } else {
                selectButtons.forEach((selectButton, index2) => {
                    let selectButtonText = selectButton.textContent.trim();
                    if (selectButtonText == '(Blanks)') { selectButtonText = ''; }
                    
                    if (approvedValuesList[index].has(selectButtonText)) {
                        selectButton.classList.remove('hide-button'); // show row if conditionals true
                    } else {
                        selectButton.classList.add('hide-button');
                    }
                });
            }
            
            // Remove select all if there's no approved values
            if (approvedValuesList[index].size == 0) {
                selectButtons[0].classList.add('hide-button'); 
            } else {
                selectButtons[0].classList.remove('hide-button'); // Show
            }
            
            let visibleSelectButtons = Array.from(selectButtons).filter(selectButton => !selectButton.classList.contains('hide-button'));
            let checkedVisibleSelectButtons = visibleSelectButtons.filter(selectButton => selectButton.querySelector(".input-box").checked);
            
            if ((visibleSelectButtons.length != checkedVisibleSelectButtons.length) || (searchTextList[index].value != '')) {
                clearFilterButtons[index].removeAttribute("disabled");
            } else {
                clearFilterButtons[index].setAttribute("disabled", "");
            }
        });
        
        let enabledFilterButtons = clearFilterButtons.filter(clearFilterButton => !clearFilterButton.hasAttribute('disabled'));
        if ((enabledFilterButtons.length != 0) || (searchText != "")) {
            document.querySelector("#Excel_Style_Table_ClearAllFilters").removeAttribute("disabled");
        } else {
            document.querySelector("#Excel_Style_Table_ClearAllFilters").setAttribute("disabled", "");
        }
    }
    
    function getDateConditional(columnCell, checkedDatesMap) {
        // Check if cell is '' and (blanks) is not selected
        if (columnCell == '') {
            if (!checkedDatesMap.has("(Blanks)")) {
                return false;
            }
            return true; // (blanks) was selected
        }
        
        let currentDate = columnCell.split("-");
            
        let currentYear = currentDate[0];
        let currentMonth = monthNames[parseInt(currentDate[1]) - 1];
        let currentDay = currentDate[2];
        
        // Check if current year/month combo is not in map
        if (!checkedDatesMap.has(currentYear) || !checkedDatesMap.get(currentYear).has(currentMonth)) {
            return false;
        // Check if current day is not in map
        } else if (!checkedDatesMap.get(currentYear).get(currentMonth).includes(currentDay)) { // was in map, check for day
            return false;
        } else {
            return true;
        }
    }
    
    function applyFilter() {
        
        let values = getValues();
        
        // List to contain values to be displayed in select check boxes of each column
        let approvedValuesList = values[0];
        let checkedValuesList = values[1];
        
        const searchText = document.querySelector("#searchBar_Excel_Style_Table").value.toLowerCase();
        
        let rows = document.querySelectorAll(`#Excel_Style_Table_TableBody tr`);
        let dropdownLength = dropdownList.length;
        
        // Conditional if nested search boxes have text
        let searchContainsList = [];
        searchTextList.forEach(searchBox => {
            searchContainsList.push(searchBox.value.trim() != '')
        });
        
        // List of all select all buttons input boxes
        let selectAllButtonsList = [];
        selectButtonsList.forEach(selectButtons => {
            selectAllButtonsList.push(selectButtons[0].querySelector(".input-box"))
        });
        
        // Change visibility of rows based on filters applied
        rows.forEach((row, index) => {
            
            // current row cells list
            const rowCells = row.querySelectorAll("td");
            
            // table cell values
            let cellContentList = Array.from(rowCells).map(cell => cell.textContent.trim());
            
            // Main search check
            let searchConditional = row.textContent.toLowerCase().includes(searchText);
            
            // assume conditional is true
            let selectConditional = true;
            
            if (searchConditional == true) {
                for (let i = 0; i < dropdownLength; i++) {
                    
                    // Select buttons filter check
                    if (!selectAllButtonsList[i].checked || (searchContainsList[i])) { // Select All button not checked (means a filter was applied)
                        let columnCell = cellContentList[i];

                        // Current column is date column
                        if (datesList[i] instanceof Map) {
                            
                            // Check if there is selectButton combination with date in cell
                            selectConditional = getDateConditional(columnCell, checkedValuesList[i]);
                            if (!selectConditional) {
                                break;
                            }
                        }
                        // Didn't find selectButton with value in cell & something was selected
                        else if (!checkedValuesList[i].includes(columnCell)) {
                            selectConditional = false;
                            break;
                        }
                    }
                }
                
                // Filter select check boxes
                for (let i = 0; i < dropdownLength; i++) {
                    
                    let currentCell = cellContentList[i];
                    let containsConditional = true;
                    
                    // Case when there are values before current column
                    if (i != 0) {
                        for (let j = 0; j < i; j++) {
                            let columnCell = cellContentList[j];
                            
                            if (datesList[j] instanceof Map) {
                                
                                // Check if there is selectButton combination with date in cell
                                containsConditional = getDateConditional(columnCell, checkedValuesList[j]);
                                if (!containsConditional) {
                                    break;
                                }
                            }
                            
                            // Current column checked list didn't contain current column value
                            // Current cell was also not in approved values
                            else if (!checkedValuesList[j].includes(columnCell)) {
                                containsConditional = false;
                                break;
                            } 
                        }
                    }
                    
                    // Case when there are values after current column
                    if (containsConditional && (i != dropdownLength - 1)) {
                        for (let j = i + 1; j < dropdownLength; j++) {
                            let columnCell = cellContentList[j];
                            
                            if (datesList[j] instanceof Map) {
                                
                                // Check if there is selectButton combination with date in cell
                                containsConditional = getDateConditional(columnCell, checkedValuesList[j]);
                                if (!containsConditional) {
                                    break;
                                }
                            }
                            
                            // Current column checked list didn't contain current column value
                            // Current cell was also not in approved values
                            else if (!checkedValuesList[j].includes(columnCell)) {
                                containsConditional = false;
                                break;
                            } 
                        }
                    }
                    
                    if (containsConditional) {
                        approvedValuesList[i].add(currentCell); // Will add only approved values to relevant set
                    }
                }
            }
            
            // Toggle visibility of current row based on conditionals
            if (searchConditional && selectConditional) {
                
                row.classList.remove('filtered-out'); // show row if conditionals true
            } else {
                row.classList.add('filtered-out');
            }
        });
        
        // Update visibility of select buttons
        updateSelectButtons(approvedValuesList, searchText);
        
        // Get first pagination button of table
        let button = document.querySelector("#Excel_Style_Table .pager .btn");
        
        // Call the function to create the new page/tab icons, go to first page
        updatePagination(button);
    }
    
    function filterSelectButtons(searchBox) {
        let index = searchTextList.indexOf(searchBox);
        let searchText = searchBox.value.toLowerCase();
        let containsButtonChecked = containsButtonList[index].checked;
        
        let filterConditional = false;
        
        if (datesList[index] instanceof Map) {
        
            for (let [year, monthMap] of datesList[index]) {
                
                if (monthMap instanceof Map) {
                    
                    if (containsButtonChecked) {
                        if (year.textContent.toLowerCase().includes(searchText)) {
                            year.classList.remove("filter-button")
                            
                            for (let [month, dayList] of monthMap) {
                                month.classList.remove("filter-button");
                                for (let day of dayList) {
                                    day.classList.remove("filter-button");
                                    
                                    if (!day.classList.contains('hide-button')) {
                                        filterConditional = true;
                                    }
                                }
                            }
                            continue;
                        }
                        // Year did not contain searchText
                        year.classList.add("filter-button");
                    } else {
                        if (year.textContent.toLowerCase().includes(searchText) && searchText != '') {
                            year.classList.add("filter-button");
                            
                            for (let [month, dayList] of monthMap) {
                                month.classList.add("filter-button");
                                for (let day of dayList) {
                                    day.classList.add("filter-button");
                                }
                            }
                            continue;
                        }
                        // Year did not contain searchText
                        year.classList.remove("filter-button");
                    }
                    
                    for (let [month, dayList] of monthMap) {
                            
                        if (containsButtonChecked) {
                            if (month.textContent.toLowerCase().includes(searchText)) {
                                year.classList.remove("filter-button");
                                month.classList.remove("filter-button");
                                
                                for (let day of dayList) {
                                    day.classList.remove("filter-button");
                                    
                                    if (!day.classList.contains('hide-button')) {
                                        filterConditional = true;
                                    }
                                }
                                continue;
                            } 
                            // Month did not contain searchText
                            month.classList.add("filter-button");
                        } else {
                            if (month.textContent.toLowerCase().includes(searchText) && searchText != '') {
                                month.classList.add("filter-button");
                                
                                for (let day of dayList) {
                                    day.classList.add("filter-button");
                                }
                                continue;
                            } 
                            // Month did not contain searchText
                            month.classList.remove("filter-button");
                        }
                        
                        for (let day of dayList) {
                                
                            if (containsButtonChecked) {
                                if (day.textContent.toLowerCase().includes(searchText)) {
                                    year.classList.remove("filter-button");
                                    month.classList.remove("filter-button");
                                    day.classList.remove("filter-button");
                                    
                                    if (!day.classList.contains('hide-button')) {
                                        filterConditional = true;
                                    }
                                    continue;
                                } 
                                // Day did not contain searchText
                                day.classList.add("filter-button");
                            } else {
                                if (day.textContent.toLowerCase().includes(searchText) && searchText != '') {
                                    day.classList.add("filter-button");
                                    continue;
                                }
                                // Day did not contain searchText
                                day.classList.remove("filter-button");
                                
                                if (!day.classList.contains('hide-button')) {
                                    filterConditional = true;
                                }
                            }
                        }
                    }
                } else {
                    if (containsButtonChecked) {
                        if (year.textContent.toLowerCase().includes(searchText)) {
                            year.classList.remove('filter-button');
                            
                            if (!year.classList.contains('hide-button')) {
                                filterConditional = true;
                            }
                            continue;
                        }
                        
                        // (Blanks) did not contain searchText
                        year.classList.add('filter-button');
                    } else {
                        if (year.textContent.toLowerCase().includes(searchText) && searchText != '') {
                            year.classList.add('filter-button');
                            continue;
                        }
                        // (Blanks) did not contain searchText
                        year.classList.remove('filter-button');
                        
                        if (!year.classList.contains('hide-button')) {
                            filterConditional = true;
                        }
                    }
                }
            }
            
            let filterOptions = selectButtonsList[index][1].parentNode;
        
            // Expand or collapse date dropdowns only when search is used
            if (searchTextList[index].value != '') {
                // Open date icons
                while (filterOptions.querySelector(".expanded")) {
                    updateIcon(filterOptions.querySelector(".expanded"));
                }
            } else if (searchTextList[index].value == '') {
                // Collapse date icons
                while (filterOptions.querySelector(".expand-icon:not(.expanded)")) {
                    updateIcon(filterOptions.querySelector(".expand-icon:not(.expanded)"));
                }
            }
        } else {
            for (let [index2, selectButton] of selectButtonsList[index].entries()) {
                if (index2 != 0) { // Skip (Select All) button
                    let selectButtonText = selectButton.querySelector(".menu-label").textContent.toLowerCase();
                    
                    if (containsButtonChecked) {
                        if (selectButtonText.includes(searchText)) {
                            selectButton.classList.remove('filter-button');
                            
                            if (!selectButton.classList.contains('hide-button')) {
                                filterConditional = true;
                            }
                            continue;
                        }
                        
                        selectButton.classList.add('filter-button');
                    } else {
                        if (selectButtonText.includes(searchText) && searchText != '') {
                            selectButton.classList.add('filter-button');
                            continue;
                        }
                        selectButton.classList.remove('filter-button');
                        
                        if (!selectButton.classList.contains('hide-button')) {
                            filterConditional = true;
                        }
                    }
                }
            }
        }
        if (filterConditional) {
            selectButtonsList[index][0].classList.remove('filter-button');
        } else {
            selectButtonsList[index][0].classList.add('filter-button');
        }
            
        applyFilter();
    }
    
    // Function to show/hide the rows that passed the filter criteria & based on their page
    function pageData(pageIdx) {
        
        // Use page index to determine the first & last row to be displayed in the viz
        const startEntry = ((pageIdx - 1) * rowsPerPage);
        const finalEntry = (startEntry + rowsPerPage);
        
        const rows = Array.from(document.querySelectorAll("#Excel_Style_Table .rows")).filter(row => !row.classList.contains('filtered-out'));
        
        rows.forEach((row, index) => {
            if (index >= startEntry && index < finalEntry) {
                // hide the rows that didn't pass the page criteria
                row.style.display = '';
            } else {
                row.style.display = 'none';
            }
        });
    }
    
    function updatePagination(selectedButton) {
        
        let buttons = Array.from(document.querySelectorAll("#Excel_Style_Table .pager .btn"));
        
        const indexButton = buttons.indexOf(selectedButton);
        let currentButtonNumber;
        
        const rows = Array.from(document.querySelectorAll("#Excel_Style_Table .rows")).filter(row => !row.classList.contains('filtered-out'));
        let pageCount = Math.ceil(rows.length / rowsPerPage); // At least 1 page
        
        // Case where no visible rows
        if (pageCount == 0) {
            pageCount += 1;
        }
        
        if (indexButton == 0) {
            currentButtonNumber = 1;
        } else if ((indexButton == 1)) {
            let activeButton = buttons.filter(button => button.classList.contains("active"))[0];
            currentButtonNumber = parseInt(activeButton.textContent) - 1;
        } else if (indexButton == buttons.length - 2) {
            let activeButton = buttons.filter(button => button.classList.contains("active"))[0];
            currentButtonNumber = parseInt(activeButton.textContent) + 1;
        } else if (indexButton == buttons.length - 1) {
            currentButtonNumber = pageCount;
        } else {
            currentButtonNumber = parseInt(selectedButton.textContent);
        }
        
        let disabledText = '';
        if (currentButtonNumber == 1) {
            disabledText += ' disabled = ""';
        } 
        
        let displayButtonsHTML = '';
        displayButtonsHTML += '<button class="btn"' + disabledText + '>' +
                                  '<svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="angles-left" class="svg-inline--fa fa-angles-left " role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">' +
                                      '<path fill="currentColor" d="M41.4 233.4c-12.5 12.5-12.5 32.8 0 45.3l160 160c12.5 12.5 32.8 12.5 45.3 0s12.5-32.8 0-45.3L109.3 256 246.6 118.6c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0l-160 160zm352-160l-160 160c-12.5 12.5-12.5 32.8 0 45.3l160 160c12.5 12.5 32.8 12.5 45.3 0s12.5-32.8 0-45.3L301.3 256 438.6 118.6c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0z"></path>' +
                                  '</svg>' +
                              '</button>' +
                              '<button class="btn"' + disabledText + '>' +
                                  '<svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="angle-left" class="svg-inline--fa fa-angle-left " role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 512">' +
                                      '<path fill="currentColor" d="M41.4 233.4c-12.5 12.5-12.5 32.8 0 45.3l160 160c12.5 12.5 32.8 12.5 45.3 0s12.5-32.8 0-45.3L109.3 256 246.6 118.6c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0l-160 160z"></path>' +
                                  '</svg>' +
                              '</button>';
        
        if (pageCount > 7) {
            let numbersList = [];
            if (currentButtonNumber < 5) {
                numbersList = [1, 2, 3, 4, 5, 6, 7]
            } else if (currentButtonNumber > pageCount - 4) {
                numbersList = [pageCount - 6, pageCount - 5, pageCount - 4, pageCount - 3, pageCount - 2, pageCount - 1, pageCount];
            } else {
                numbersList = [currentButtonNumber - 3, currentButtonNumber - 2, currentButtonNumber - 1, currentButtonNumber, currentButtonNumber + 1, currentButtonNumber + 2, currentButtonNumber + 3];
            }
                                  
            numbersList.forEach(number => {
                if (number == currentButtonNumber) {
                    displayButtonsHTML += '<button class="btn num active" disabled=""><span> ' + number + ' </span></button>';
                } else {
                    displayButtonsHTML += '<button class="btn num"><span> ' + number + ' </span></button>';
                }
            });
        } else {
            for (let i = 1; i <= pageCount; i++) {
                if (i == currentButtonNumber) {
                    displayButtonsHTML += '<button class="btn num active" disabled=""><span> ' + i + ' </span></button>';
                } else {
                    displayButtonsHTML += '<button class="btn num"><span> ' + i + ' </span></button>';
                }
            }
        }
        
        disabledText = '';
        if (currentButtonNumber == pageCount) {
            disabledText += ' disabled = ""';
        } 
        
        displayButtonsHTML += '<button class="btn"' + disabledText + '>' +
                                  '<svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="angle-right" class="svg-inline--fa fa-angle-right " role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 512">' +
                                      '<path fill="currentColor" d="M278.6 233.4c12.5 12.5 12.5 32.8 0 45.3l-160 160c-12.5 12.5-32.8 12.5-45.3 0s-12.5-32.8 0-45.3L210.7 256 73.4 118.6c-12.5-12.5-12.5-32.8 0-45.3s32.8-12.5 45.3 0l160 160z"></path>' +
                                  '</svg>' +
                              '</button>' +
                              '<button class="btn"' + disabledText + '>' +
                                  '<svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="angles-right" class="svg-inline--fa fa-angles-right " role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">' +
                                      '<path fill="currentColor" d="M470.6 278.6c12.5-12.5 12.5-32.8 0-45.3l-160-160c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3L402.7 256 265.4 393.4c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0l160-160zm-352 160l160-160c12.5-12.5 12.5-32.8 0-45.3l-160-160c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3L210.7 256 73.4 393.4c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0z"></path>' +
                                  '</svg>' +
                              '</button>';
        
        let pager = document.querySelector("#Excel_Style_Table .pager");
        
        if ((currentButtonNumber >= 97) && !pager.classList.contains("pager-w3")) {
            pager.classList.remove("pager-w1");
            pager.classList.remove("pager-w2");
            
            pager.classList.add("pager-w3");
        } else if ((currentButtonNumber >= 7) && (currentButtonNumber <= 96) && (pageCount >= 10) && !pager.classList.contains("pager-w2")) {
            
            pager.classList.remove("pager-w1");
            pager.classList.remove("pager-w3");
            
            pager.classList.add("pager-w2");
        } else if (((pageCount <= 9) || (currentButtonNumber <= 6)) && !pager.classList.contains("pager-w1")) {
            pager.classList.remove("pager-w2");
            pager.classList.remove("pager-w3");
            
            pager.classList.add("pager-w1");
        }
        
        pager.innerHTML = displayButtonsHTML;
        addEventHandler($("#Excel_Style_Table .pager .btn"), "click", function() {
            updatePagination(this);
            
            // Scroll to top
            $(`#Excel_Style_Table`).get(0).scrollIntoView({ behavior: "smooth" });
        });
        
        pageData(currentButtonNumber);
        
        // Update current page text
        let currentPage = document.querySelector("#Excel_Style_Table .currentPage");
        currentPage.textContent = "Page " + currentButtonNumber + " of " + pageCount;
    }
    
    function stopPropagation(event) {
        event.stopPropagation();
    }
    
    function adjustPopup() {
        let parentDiv = this.parentNode;
        const dropdownRect = parentDiv.getBoundingClientRect();
        
        let dropdownMenu = parentDiv.querySelector(".dropdown-menu");
        
        if (!dropdownMenu.classList.contains('show')) {
        
            dropdownMenu.classList.add('dropdown-left');
            dropdownMenu.classList.remove('dropdown-right');
            
            if ((dropdownRect.right + parentDiv.offsetWidth / 2) > window.innerWidth) {
                dropdownMenu.classList.remove('dropdown-left');
                dropdownMenu.classList.add('dropdown-right');
            }
        }
    }
    
    function updateIcon(icon) {
        icon.classList.toggle('expanded');
        currentSelectButton = icon.parentNode;
        
        let dropdown = currentSelectButton.parentNode.parentNode.parentNode.parentNode;
        let index = dropdownList.indexOf(dropdown);
        
        let parentDiv = icon.parentNode.parentNode;
        let selectButtons = Array.from(parentDiv.querySelectorAll(".dropdown-item"));
            
        let yearButton = currentSelectButton;
        
        if (currentSelectButton.classList.contains("first-child-dropdown")) {
            let currentIndex = selectButtons.indexOf(currentSelectButton);
            let elementsList = selectButtons.slice(0, currentIndex);
            
            let yearList = elementsList.filter(selectButton => selectButton.classList.contains("parent-dropdown"));
            
            yearButton = yearList[yearList.length - 1];
        }
        
        let monthButton = currentSelectButton;
        
        if (currentSelectButton.classList.contains("parent-dropdown")) {
            let yearMap = datesList[index].get(yearButton);
            
            if (!icon.classList.contains("expanded")) {
                Array.from(yearMap.keys()).forEach(month => {
                    month.style.display = ""
                });
            } else {
                for (let [month, dayList] of yearMap) {
                    month.style.display = "none"
                    month.querySelector(".expand-icon").classList.add('expanded');
                    
                    dayList.forEach(day => {
                        day.style.display = "none";
                    });
                }
            }
        } else if (currentSelectButton.classList.contains("first-child-dropdown")) {
            let dayList = datesList[index].get(yearButton).get(monthButton);
            
            if (!icon.classList.contains("expanded")) {
                dayList.forEach(day => {
                    day.style.display = "";
                });
            } else {
                dayList.forEach(day => {
                    day.style.display = "none";
                });
            }
        }
    }
    
    // Function that clears all the filters
    function clearFilters(button) {
        let index = clearFilterButtons.indexOf(button);
        
        // Clear search filter and check relevant select buttons
        let selectButtons = Array.from(selectButtonsList[index]);
        selectButtons.forEach(selectButton => {
            selectButton.classList.remove('filter-button');
            let inputBox = selectButton.querySelector(".input-box");
            inputBox.checked = true;
        });
        
        // Collapse visible date icons
        if (datesList[index] instanceof Map) {
            let dropdown = selectButtons[1].parentNode;
            
            // Keep closing dates until none are found
            while (dropdown.querySelector(".expand-icon:not(.expanded)")) {
                updateIcon(dropdown.querySelector(".expand-icon:not(.expanded)"));
            }
        }
        
        // Default check 'contains' filter type
        containsButtonList[index].checked = true;
        
        // Uncheck 'not contains' filter type
        notContainsButtonList[index].checked = false;
        
        searchTextList[index].value = "";
    }
    
    function clearAllFilters() {
        
        clearFilterButtons.forEach((clearFilterButton, index) => {
            // Check ALL select buttons
            Array.from(selectButtonsList[index]).forEach(selectButton => {
                selectButton.classList.remove('filter-button');
                let inputBox = selectButton.querySelector(".input-box");
                inputBox.checked = true;
            });
            
            // Collapse date icons
            if (datesList[index] instanceof Map) {
                let dropdown = selectButtonsList[index][1].parentNode;
                
                while (dropdown.querySelector(".expand-icon:not(.expanded)")) {
                    updateIcon(dropdown.querySelector(".expand-icon:not(.expanded)"));
                }
            }
            
            // Default check 'contains' filter type
            containsButtonList[index].checked = true;
            
            // Uncheck 'not contains' filter type
            notContainsButtonList[index].checked = false;
            
            searchTextList[index].value = "";
        });
        
        // Set the main search bar value to blank
        document.querySelector("#searchBar_Excel_Style_Table").value = "";
        
        // Refilter the table so all rows appear again
        applyFilter();
    }

    function fetchData() {
        return fetch('data/placeholderData.json')
            .then(function (response) {
                if (response.status === 200) {
                    return response.json();
                }
            })
            .catch(function (error) {
                console.log(error);
            });
    }

    function createTable(rawData) {
        const projectNamesMap = {};
        const countryNamesMap = {};
        const priceValuesMap = {};
        const dateMap = {};

        for (let submission of rawData) {
            projectNamesMap[submission.Project_Name] = "";
            countryNamesMap[submission.Country] = "";
            priceValuesMap["$" + submission.Price] = "";
            
            const dateSplit = submission.Date.split("-");
            const year = dateSplit[0];
            const month = dateSplit[1];
            const day = dateSplit[2];

            dateMap[year] = dateMap[year] || {month: {day: ""}};
            dateMap[year][month] = dateMap[year][month] || {day: ""};
            dateMap[year][month][day] = "";
        }

        const columnFields = [Object.keys(projectNamesMap).sort(), Object.keys(countryNamesMap).sort(), Object.keys(priceValuesMap).sort(), dateMap];
        
        const headerIds = ["Project Name", "Country", "Price", "Date"];

        const maxEntriesPerPage = 25;
        const totalNumPages = Math.ceil(rawData.length / maxEntriesPerPage);

        const columnWidth = headerIds.length;

        document.querySelector("#Excel_Style_Table").insertAdjacentHTML('afterBegin', `<input type="text" id="searchBar_Excel_Style_Table" class="form-control" placeholder="Search..">
            <div style = "display: flex; margin: 5px 0px;">
                <button type="button" class="btn btn-sm subscribe-btn" id="Excel_Style_Table_ClearAllFilters" disabled = "">
                    Clear All Filters
                </button>
                
                <div class = "currentPage" style = "margin-left: auto; margin-top: auto">
                    Page 1 of ${totalNumPages}
                </div>
            </div>`);

        let tableHeadHTML = `<th class="basic-table-header" colspan="${columnWidth}">
                    Table Name
                </th>
                <tr>`;

        headerIds.forEach((header, index) => {
            let blanksConditional = false;

            tableHeadHTML += `<th class="basic-table-header" style="width: calc(100% / ${columnWidth});">
                        <div class="dropdown" id="Excel_Style_Table_dropDown_${index}">
                            <button id="${index}Btn_Excel_Style_Table" class="btn header-dropdown" type="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                                ${header}
                                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="currentColor" class="bi bi-caret-down-square" viewBox="0 0 16 16">
                                    <path d="M3.626 6.832A.5.5 0 0 1 4 6h8a.5.5 0 0 1 .374.832l-4 4.5a.5.5 0 0 1-.748 0l-4-4.5z"/>
                                    <path d="M0 2a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V2zm15 0a1 1 0 0 0-1-1H2a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V2z"/>
                                </svg>
                            </button>
                            <div id="Excel_Style_Table_${index}_Filter_Dropdown" class="dropdown-menu multi-dropdown" aria-labelledby="dropdownMenuButton" style='overflow: auto;'>
                                
                                <div class="sort-section">
                                    <p class="section-title">Sort</p>
                                    <a id="Excel_Style_Table_${index}_Sort_Forward" class="dropdown-item select-dropdown" href="#">
                                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="currentColor" class="bi bi-sort-alpha-down" viewBox="0 0 16 16">
                                            <path fill-rule="evenodd" d="M10.082 5.629 9.664 7H8.598l1.789-5.332h1.234L13.402 7h-1.12l-.419-1.371h-1.781zm1.57-.785L11 2.687h-.047l-.652 2.157h1.351z"/>
                                            <path d="M12.96 14H9.028v-.691l2.579-3.72v-.054H9.098v-.867h3.785v.691l-2.567 3.72v.054h2.645V14zM4.5 2.5a.5.5 0 0 0-1 0v9.793l-1.146-1.147a.5.5 0 0 0-.708.708l2 1.999.007.007a.497.497 0 0 0 .7-.006l2-2a.5.5 0 0 0-.707-.708L4.5 12.293V2.5z"/>
                                        </svg>
                                        Ascending
                                    </a>
                                    
                                    <a id="Excel_Style_Table_${index}_Sort_Reverse" class="dropdown-item select-dropdown" href="#">
                                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="currentColor" class="bi bi-sort-alpha-down-alt" viewBox="0 0 16 16">
                                            <path d="M12.96 7H9.028v-.691l2.579-3.72v-.054H9.098v-.867h3.785v.691l-2.567 3.72v.054h2.645V7z"/>
                                            <path fill-rule="evenodd" d="M10.082 12.629 9.664 14H8.598l1.789-5.332h1.234L13.402 14h-1.12l-.419-1.371h-1.781zm1.57-.785L11 9.688h-.047l-.652 2.156h1.351z"/>
                                            <path d="M4.5 2.5a.5.5 0 0 0-1 0v9.793l-1.146-1.147a.5.5 0 0 0-.708.708l2 1.999.007.007a.497.497 0 0 0 .7-.006l2-2a.5.5 0 0 0-.707-.708L4.5 12.293V2.5z"/>
                                        </svg>
                                        Descending
                                    </a>
                                </div>
                                
                                <hr class="solid dropdown-section-divider">
                                
                                <div class="filter-section">
                                    <p class="section-title">Filter</p>
                                    <div class='container-fluid' id='Excel_Style_Table_${index}_Includes_Btns' style="margin-top: 7px;">
                                        <input type="radio" id="Excel_Style_Table_${index}_Contains" name="${index}ContainsView" value="Contains" style='cursor: pointer;'>
                                        <label class="header-dropdown-contains" for="Excel_Style_Table_${index}_Contains">Contains</label>
                                        <br>
                                        <input type="radio" id="Excel_Style_Table_${index}_Not_Contains" name="${index}NotContainView" value="Not_Contain" style='cursor: pointer;'>
                                        <label class="header-dropdown-not-contains" for="Excel_Style_Table_${index}_Not_Contains">Does Not Contain</label>
                                    </div>
                                    
                                    <input class='form-control header-dropdown-search' id='Excel_Style_Table_FilterBar_${index}' type='text' placeholder='Filter...'>
                                    
                                    <div id="Excel_Style_Table_${index}_Filter_Options" class="header-dropdown-options-box">
                                        
                                        <a class="dropdown-item select-dropdown" href="#">
                                            <input class="input-box" type="checkbox" id="Select_All_${index}" placeholder="Yes" value='Select_All' style="cursor: pointer;">
                                            <label class="menu-label header-dropdown-option" for="Select_All_${index}">(Select All)</label>
                                        </a>`;
            // List
            if (columnFields[index].length) {
                columnFields[index].forEach(selectOption => {
                    if (selectOption != "") {
                        tableHeadHTML += `<a class="dropdown-item select-dropdown" href="#">
                                        <input class="input-box" type="checkbox" id="${selectOption}" placeholder="Yes" value='${selectOption}' style="cursor: pointer;">
                                        <label class="menu-label header-dropdown-option" for="${selectOption}">${selectOption}</label>
                                    </a>`;
                    } else {
                        blanksConditional = true;
                    }
                });

                if (blanksConditional == true) {
                    tableHeadHTML += `<a class="dropdown-item select-dropdown" href="#">
                                    <input class="input-box" type="checkbox" id="(Blanks)" placeholder="Yes" value='(Blanks)' style="cursor: pointer;">
                                    <label class="menu-label header-dropdown-option" for="(Blanks)">(Blanks)</label>
                                </a>`;
                }
            // Object
            } else {
                const monthMap = {
                    "01": "January",
                    "02": "February",
                    "03": "March",
                    "04": "April",
                    "05": "May",
                    "06": "June",
                    "07": "July",
                    "08": "August",
                    "09": "September",
                    "10": "October",
                    "11": "November",
                    "12": "December"
                };

                const dateYears = Array.from(Object.keys(columnFields[index])).sort().reverse();
                Array.from(dateYears).forEach(year => {
                    if (year != "") {
                        tableHeadHTML += `<a class="dropdown-item select-dropdown parent-dropdown" href="#">
                                        <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="caret-down" class="svg-inline--fa fa-caret-down expand-icon expanded" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 512">
                                            <path fill="currentColor" d="M137.4 374.6c12.5 12.5 32.8 12.5 45.3 0l128-128c9.2-9.2 11.9-22.9 6.9-34.9s-16.6-19.8-29.6-19.8L32 192c-12.9 0-24.6 7.8-29.6 19.8s-2.2 25.7 6.9 34.9l128 128z"></path>
                                        </svg>
                                        <input class="input-box" type="checkbox" id="${year}" placeholder="Yes" value='${year}' style="cursor: pointer;">
                                        <label class="menu-label header-dropdown-option" for="${year}">${year}</label>
                                    </a>`;

                        const dateMonths = Object.keys(columnFields[index][year]).sort();
                        Array.from(dateMonths).forEach(month => {
                            tableHeadHTML += `<a class="dropdown-item select-dropdown first-child-dropdown" style = "display: none" href="#">
                                            &nbsp;&nbsp;&nbsp;
                                            <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="caret-down" class="svg-inline--fa fa-caret-down expand-icon expanded" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 512">
                                                <path fill="currentColor" d="M137.4 374.6c12.5 12.5 32.8 12.5 45.3 0l128-128c9.2-9.2 11.9-22.9 6.9-34.9s-16.6-19.8-29.6-19.8L32 192c-12.9 0-24.6 7.8-29.6 19.8s-2.2 25.7 6.9 34.9l128 128z"></path>
                                            </svg>
                                            <input class="input-box" type="checkbox" id="${monthMap[month]}" placeholder="Yes" value='${monthMap[month]}' style="cursor: pointer;">
                                            <label class="menu-label header-dropdown-option" for="${monthMap[month]}">${monthMap[month]}</label>
                                        </a>`;

                            const dateDayList = Object.keys(columnFields[index][year][month]).sort();
                            Array.from(dateDayList).forEach(day => {
                                tableHeadHTML += `<a class="dropdown-item select-dropdown second-child-dropdown" style = "display: none" href="#">
                                                <span>
                                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                </span>
                                                <input class="input-box" type="checkbox" id="${day}" placeholder="Yes" value='${day}' style="cursor: pointer;">
                                                <label class="menu-label header-dropdown-option" for="${day}">${day}</label>
                                            </a>`;
                            });
                        });
                    } else {
                        blanksConditional = true;
                    }
                });

                if (blanksConditional = true) {
                    tableHeadHTML += `<a class="dropdown-item select-dropdown" href="#">
                                    <input class="input-box" type="checkbox" id="(Blanks)" placeholder="Yes" value='(Blanks)' style="cursor: pointer;">
                                    <label class="menu-label header-dropdown-option" for="(Blanks)">(Blanks)</label>
                                </a>`;
                }
            }

            tableHeadHTML += `</div>
                                <hr class="solid dropdown-section-divider">
                                <div class = "text-right" style = "margin-right: 5%">
                                    <button type="button" class="btn btn-sm subscribe-btn" id="Excel_Style_Table_${index}_ClearFilter" disabled = "">
                                        Clear Filter
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </th>`;
        });

        tableHeadHTML += `</tr>`;

        document.querySelector("#Excel_Style_Table table thead").innerHTML = tableHeadHTML;

        let tableBodyHTML = "";
        rawData.forEach((entry, index) => {
            const rowStyle = (index > maxEntriesPerPage - 1) ? ' style = "display: none"' : '';
            tableBodyHTML += `<tr class = "rows"` + rowStyle + `>
                <td style = "width: calc(100% / ${columnWidth});">
                    <div style = "min-height: 24.5px;">
                        ${entry.Project_Name}
                    </div>
                </td>
                <td style = "width: calc(100% / ${columnWidth});"> 
                    <div style = "min-height: 24.5px;">
                        ${entry.Country}
                    </div>
                </td>
                <td style = "width: calc(100% / ${columnWidth});"> 
                    <div style = "min-height: 24.5px;">
                        $${entry.Price}
                    </div>
                </td>
                <td style = "width: calc(100% / ${columnWidth});"> 
                    <div style = "min-height: 24.5px;">
                        ${entry.Date}
                    </div>
                </td>
            </tr>`;
        });

        document.querySelector("#Excel_Style_Table table tbody").innerHTML = tableBodyHTML;

        let paginationHTML = `<button class="btn" disabled = "">
            <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="angles-left" class="svg-inline--fa fa-angles-left " role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
                <path fill="currentColor" d="M41.4 233.4c-12.5 12.5-12.5 32.8 0 45.3l160 160c12.5 12.5 32.8 12.5 45.3 0s12.5-32.8 0-45.3L109.3 256 246.6 118.6c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0l-160 160zm352-160l-160 160c-12.5 12.5-12.5 32.8 0 45.3l160 160c12.5 12.5 32.8 12.5 45.3 0s12.5-32.8 0-45.3L301.3 256 438.6 118.6c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0z"></path>
            </svg>
        </button><button class="btn" disabled = "">
            <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="angle-left" class="svg-inline--fa fa-angle-left " role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 512">
                <path fill="currentColor" d="M41.4 233.4c-12.5 12.5-12.5 32.8 0 45.3l160 160c12.5 12.5 32.8 12.5 45.3 0s12.5-32.8 0-45.3L109.3 256 246.6 118.6c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0l-160 160z"></path>
            </svg>
        </button><button class="btn num active" disabled = "">
            <span>
                1
            </span>
        </button>`;
            
        const maxButtons = Math.min(totalNumPages, 7);
        const disabledAttribute = (totalNumPages == 1) ? ` disabled = ""` : ""
        
        for (let i = 2; i <= maxButtons; i++) {
            paginationHTML += `<button class="btn num"><span>` + i + `</span></button>`
        }
        
        paginationHTML += `<button class="btn"` + disabledAttribute + `>
            <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="angle-right" class="svg-inline--fa fa-angle-right " role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 512">
                <path fill="currentColor" d="M278.6 233.4c12.5 12.5 12.5 32.8 0 45.3l-160 160c-12.5 12.5-32.8 12.5-45.3 0s-12.5-32.8 0-45.3L210.7 256 73.4 118.6c-12.5-12.5-12.5-32.8 0-45.3s32.8-12.5 45.3 0l160 160z"></path>
            </svg>
        </button><button class="btn"` + disabledAttribute + `>
            <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="angles-right" class="svg-inline--fa fa-angles-right " role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
                <path fill="currentColor" d="M470.6 278.6c12.5-12.5 12.5-32.8 0-45.3l-160-160c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3L402.7 256 265.4 393.4c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0l160-160zm-352 160l160-160c12.5-12.5 12.5-32.8 0-45.3l-160-160c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3L210.7 256 73.4 393.4c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0z"></path>
            </svg>
        </button>`;
            
        document.querySelector("#Excel_Style_Table .pager").innerHTML = paginationHTML;
    }
    
    /** This function is called to start the entire process, it should...
     *    1. Wait for the top-level element to be in the DOM.
     *    2. Wire-up any events that are needed.
     *    3. Do any other logic that the template requires.
     */
    async function init() {
        const topLevelElement = await waitForTopLevelElement();

        const rawData = await fetchData();
        createTable(rawData);
        
        // Get list of all column filters
        populateLists();
        
        addEventHandler($("#Excel_Style_Table [id*=Sort_Forward]"), "click", sortForward);
        addEventHandler($("#Excel_Style_Table [id*=Sort_Reverse]"), "click", sortReverse);
        
        // Default check 'contains' filter type buttons
        let containsButtons = document.querySelectorAll("#Excel_Style_Table input[value='Contains']");
        containsButtons.forEach(containsButton => containsButton.checked = true);
        
        // Contains/Not Contains Buttons
        addEventHandler($("#Excel_Style_Table input[type='radio']"), 'click', changeFilterType);
        
        addEventHandler($("#searchBar_Excel_Style_Table"), 'keyup', applyFilter);
        
        addEventHandler($("#Excel_Style_Table .header-dropdown-search"), 'keyup', function() {
            filterSelectButtons(this)
        });
        
        // check all select buttons
        let inputBoxList = document.querySelectorAll("#Excel_Style_Table .input-box");
        inputBoxList.forEach(inputBox => inputBox.checked = true);
        
        addEventHandler($("#Excel_Style_Table .input-box"), "click", selectFilter);
        
        addEventHandler($("#Excel_Style_Table [id*=Filter_Dropdown]"), "click", stopPropagation);
        
        addEventHandler($("#Excel_Style_Table .header-dropdown"), "click", adjustPopup);
        
        addEventHandler($("#Excel_Style_Table .expand-icon"), "click",  function() {
            updateIcon(this);
        });
        
        addEventHandler($("#Excel_Style_Table .pager .btn"), "click", function() {
            updatePagination(this);
            
            // Scroll to top
            $(`#Excel_Style_Table`).get(0).scrollIntoView({ behavior: "smooth" });
        });
        
        addEventHandler($("#Excel_Style_Table [id*=ClearFilter]"), "click", function() {
            clearFilters(this);
            
            // Refilter the table so all rows appear again
            applyFilter();
        });
        
        addEventHandler($("#Excel_Style_Table_ClearAllFilters"), "click", clearAllFilters);
        
        configureCleanup();
    }
    
    init();
})();