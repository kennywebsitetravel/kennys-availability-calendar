document.addEventListener('DOMContentLoaded', function () {
  // Global variables to hold the current preset (if any) and its sorting method.
  let currentPreset = null;
  let currentSortingMethod = "alphabetical"; // default if none chosen

  // Utility function to log messages.
  const logMessage = message => console.log(message);

  // Spinner Functions.
  const showSpinner = () => {
    document.getElementById('spinner').style.display = 'block';
    document.getElementById('legend').style.display = 'none';
    document.getElementById('calendarControls').style.display = 'none';
  };

  const hideSpinner = () => {
    document.getElementById('spinner').style.display = 'none';
    document.getElementById('legend').style.display = 'block';
    document.getElementById('calendarControls').style.display = 'block';
  };

  // Populate the month/year dropdown for 18 consecutive months (starting with the current month).
  function populateMonthYearDropdown() {
    const select = document.getElementById('monthYearDropdown');
    select.innerHTML = "";
    const currentDate = new Date();
    const monthNames = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    for (let i = 0; i < 18; i++) {
      const date = new Date(currentDate.getFullYear(), currentDate.getMonth() + i, 1);
      const monthName = monthNames[date.getMonth()];
      const year = date.getFullYear();
      const option = document.createElement('option');
      option.value = `${monthName} ${year}`;
      option.textContent = `${monthName} ${year}`;
      select.appendChild(option);
    }
    const currentMonthName = monthNames[currentDate.getMonth()];
    select.value = `${currentMonthName} ${currentDate.getFullYear()}`;
    logMessage("Dropdown populated with 18 months.");
  }

  // Parse a month/year string into { monthIndex, year }.
  function parseMonthYear(value) {
    const [monthName, yearStr] = value.split(" ");
    const monthNames = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    return { monthIndex: monthNames.indexOf(monthName), year: parseInt(yearStr, 10) };
  }

  // Update the dropdown to a new month/year.
  function updateDropdownTo(monthIndex, year) {
    const monthNames = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    const newValue = `${monthNames[monthIndex]} ${year}`;
    document.getElementById('monthYearDropdown').value = newValue;
  }

  // Change month by delta (-1 for previous, +1 for next) and reload calendar data.
  function changeMonth(delta) {
    const currentValue = document.getElementById('monthYearDropdown').value;
    let { monthIndex, year } = parseMonthYear(currentValue);
    monthIndex += delta;
    if (monthIndex < 0) {
      monthIndex = 11;
      year--;
    } else if (monthIndex > 11) {
      monthIndex = 0;
      year++;
    }
    updateDropdownTo(monthIndex, year);
    loadCalendarData();
  }

  // Event Listeners for Arrow Buttons.
  const prevMonthBtn = document.getElementById('prevMonth');
  const nextMonthBtn = document.getElementById('nextMonth');
  if (prevMonthBtn) {
    prevMonthBtn.addEventListener('click', () => {
      logMessage("Previous month button clicked.");
      changeMonth(-1);
    });
  }
  if (nextMonthBtn) {
    nextMonthBtn.addEventListener('click', () => {
      logMessage("Next month button clicked.");
      changeMonth(1);
    });
  }

  // ------------------------
  // TOOLTIP FUNCTIONS
  // ------------------------

  // 1. First Column Link Tooltip.
  function showLinkTooltip(event, product) {
    const tooltip = document.createElement('div');
    tooltip.classList.add('tooltip');
    tooltip.innerHTML = `<span style="font-size:15px;"><strong>${product.supplierName || "Supplier"}</strong> - ${product.name}</span>`;
    // Position and style the tooltip.
    tooltip.style.position = 'absolute';
    tooltip.style.background = '#FECB00';
    tooltip.style.color = 'black';
    tooltip.style.pointerEvents = 'none';
    tooltip.style.padding = '5px';
    tooltip.style.border = '1px solid #333';
    tooltip.style.zIndex = '9999';
    const x = event.pageX + 10;
    const y = event.pageY + 10;
    tooltip.style.left = x + 'px';
    tooltip.style.top = y + 'px';
    document.body.appendChild(tooltip);
    event.currentTarget._tooltip = tooltip;
  }

  // 2. Daily Availability Cell Tooltip.
  function showCellTooltip(event, responses) {
    const tooltip = document.createElement('div');
    tooltip.classList.add('tooltip');
    tooltip.style.position = 'absolute';
    tooltip.style.background = '#FECB00';
    tooltip.style.color = '#FFFFFF';
    tooltip.style.pointerEvents = 'none';
    tooltip.style.padding = '5px';
    tooltip.style.border = '1px solid #333';
    tooltip.style.zIndex = '9999';

    // Define error messages to check.
    const errorMessages = [
      "Caching not enabled for this fare",
      "Wrong Season",
      "No BookingSystem",
      "Wrong Season / No BookingSystem",
      "Wrong Season / Caching not enabled for this fare"
    ];
    // If all responses are errors from the above list, show a simple text tooltip.
    const allErrors = responses.every(resp => resp.Error && errorMessages.some(err => resp.Error.includes(err)));
    if (allErrors) {
      tooltip.textContent = responses[0].Error;
    } else {
      // Otherwise, build a table of FareName and NumAvailable.
      const table = document.createElement('table');
      table.style.borderCollapse = 'collapse';
      table.style.fontSize = '12px';
      table.style.color = 'black';
      const headerRow = document.createElement('tr');
      const thFare = document.createElement('th');
      thFare.textContent = "Fare Name";
      thFare.style.border = '1px solid #333';
      thFare.style.padding = '2px 4px';
      const thAvail = document.createElement('th');
      thAvail.textContent = "Available";
      thAvail.style.border = '1px solid #333';
      thAvail.style.padding = '2px 4px';
      headerRow.appendChild(thFare);
      headerRow.appendChild(thAvail);
      table.appendChild(headerRow);
      responses.forEach(resp => {
        const row = document.createElement('tr');
        const tdFare = document.createElement('td');
        tdFare.textContent = resp.FareName || "";
        tdFare.style.border = '1px solid #333';
        tdFare.style.padding = '2px 4px';
        const tdAvail = document.createElement('td');
        tdAvail.textContent = (resp.NumAvailable !== undefined) ? resp.NumAvailable : "";
        tdAvail.style.border = '1px solid #333';
        tdAvail.style.padding = '2px 4px';
        row.appendChild(tdFare);
        row.appendChild(tdAvail);
        table.appendChild(row);
      });
      tooltip.appendChild(table);
    }
    const x = event.pageX + 10;
    const y = event.pageY + 10;
    tooltip.style.left = x + 'px';
    tooltip.style.top = y + 'px';
    document.body.appendChild(tooltip);
    event.currentTarget._tooltip = tooltip;
  }

  // 3. Hide Tooltip.
  function hideTooltip(event) {
    if (event.currentTarget._tooltip) {
      event.currentTarget._tooltip.remove();
      event.currentTarget._tooltip = null;
    }
  }

  // ------------------------
  // API & DATA FUNCTIONS
  // ------------------------

  // Fetch the ephemeral token from your Google Apps Script.
  async function getAccessTokenFromAppsScript() {
    const scriptUrl = "https://script.google.com/macros/s/AKfycbwe_T8l2Br5so4estFaBU9TgJLWUko_hz3dxPtDvW0_gj_Oo84OVozcaeS6iHF42evoJg/exec";
    try {
      const response = await fetch(scriptUrl, { method: "GET" });
      const data = await response.json();
      if (data.status === "error") {
        console.error("Apps Script error:", data.message);
        return null;
      }
      return data.accessToken;
    } catch (error) {
      console.error("Error fetching token from Apps Script:", error);
      return null;
    }
  }

  // API function to check availability for a given product.
  async function checkAvailabilityForProduct(product, accessToken, selectedMonthYear) {
    const [monthName, yearStr] = selectedMonthYear.split(" ");
    const year = parseInt(yearStr, 10);
    const monthIndex = new Date(Date.parse(monthName + " 1, 2020")).getMonth();
    const formattedMonth = (monthIndex + 1).toString().padStart(2, '0');
    const fullDaysInMonth = new Date(year, monthIndex + 1, 0).getDate();
    let startDay = 1;
    let days;
    const now = new Date();
    if (now.getFullYear() === year && now.getMonth() === monthIndex) {
      startDay = now.getDate() + 1;
      if (startDay > fullDaysInMonth) {
        startDay = fullDaysInMonth;
      }
      days = fullDaysInMonth - now.getDate();
      logMessage(`API Call: ${year}-${formattedMonth}-${String(startDay).padStart(2, '0')}, days = ${days}`);
    } else {
      startDay = 1;
      days = fullDaysInMonth;
      logMessage(`API Call: ${year}-${formattedMonth}-01, days = ${days}`);
    }
    const startDate = `${year}-${formattedMonth}-${String(startDay).padStart(2, '0')}`;
    const params = new URLSearchParams();
    params.append("startDate", startDate);
    params.append("days", days.toString());
    if (Array.isArray(product.faresprices)) {
      product.faresprices.forEach(fare => {
        params.append("productPricesDetailsIds[]", fare.productPricesDetailsId);
      });
    }
    const availabilityUrl = `https://tdms.websitetravel.com/apiv1/checkAvailability?${params.toString()}`;
    logMessage("API Call: " + availabilityUrl);
    try {
      const availResponse = await fetch(availabilityUrl, {
        method: 'GET',
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json",
          "Authorization": `Bearer ${accessToken}`
        }
      });
      logMessage("Availability response for " + product.name + ": " + availResponse.status);
      if (!availResponse.ok) {
        return {};
      }
      const availData = await availResponse.json();
      return availData;
    } catch (err) {
      console.error("Error in checkAvailabilityForProduct:", err);
      return {};
    }
  }

  // Main function to load calendar data.
  async function loadCalendarData() {
    logMessage("loadCalendarData triggered with productIds: " + document.getElementById('productIds').value);
    showSpinner();
    document.getElementById('results').innerHTML = '<p>Loading products...</p>';
    document.getElementById('availabilityResults').innerHTML = "";
    document.getElementById('progressUpdate').textContent = "";
    
    const productIds = document.getElementById('productIds').value.split(',')
      .map(id => id.trim()).filter(id => id !== '');
    const selectedMonthYear = document.getElementById('monthYearDropdown').value;
    const parsed = parseMonthYear(selectedMonthYear);
    const selectedYear = parsed.year;
    const selectedMonth = parsed.monthIndex;
    const formattedMonth = (selectedMonth + 1).toString().padStart(2, '0');
    const daysInMonth = new Date(selectedYear, selectedMonth + 1, 0).getDate();
    const today = new Date();
    if (selectedYear < today.getFullYear() ||
       (selectedYear === today.getFullYear() && selectedMonth < today.getMonth())) {
      document.getElementById('availabilityResults').innerHTML = "<p>No future dates available for the selected month.</p>";
      hideSpinner();
      return;
    }
    let startDay = 1;
    if (selectedYear === today.getFullYear() && selectedMonth === today.getMonth()) {
      startDay = today.getDate() + 1;
      if (startDay > daysInMonth) {
        startDay = daysInMonth;
      }
    }
    try {
      const accessToken = await getAccessTokenFromAppsScript();
      if (!accessToken) {
        throw new Error("No access token returned from Apps Script.");
      }
      const productsResponse = await fetch(`https://tdms.websitetravel.com/apiv1/product/${productIds.join(',')}`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Authorization': `Bearer ${accessToken}`
        }
      });
      if (!productsResponse.ok) throw new Error('Products fetch failed.');
      const productsData = await productsResponse.json();
      let products = productsData.results;
      if (!Array.isArray(products) || products.length === 0) {
        document.getElementById('availabilityResults').innerHTML = '<p>No matching products found.</p>';
        hideSpinner();
        return;
      }
      
      // Apply sorting based on the current preset’s sorting method.
      if (currentSortingMethod === "alphabetical") {
        products.sort((a, b) => a.name.localeCompare(b.name));
      } else if (currentSortingMethod === "list" && currentPreset) {
        const presetOrder = currentPreset.productIds.split(',').map(id => id.trim());
        products.sort((a, b) => {
          const indexA = presetOrder.indexOf(String(a.productId));
          const indexB = presetOrder.indexOf(String(b.productId));
          return indexA - indexB;
        });
      }
      
     // Build products table.
      const fragProducts = document.createDocumentFragment();
      const productsTable = document.createElement('table');
      productsTable.innerHTML = `<thead>
        <tr>
          <th>Product Name</th>
          <th>Duration</th>
          <th>Product Prices Details IDs</th>
        </tr>
      </thead>`;
      const ptbody = document.createElement('tbody');
      products.forEach(product => {
        const row = document.createElement('tr');
        row.innerHTML = `<td><a href="https://tdms.websitetravel.com/#search/text/${product.productId}" target="_blank">${product.name}</a></td>
                         <td>${product.durationDays || "0"}/${product.durationNight || "0"}</td>
                         <td>${(Array.isArray(product.faresprices) ? product.faresprices.map(f => f.productPricesDetailsId).join(', ') : "N/A")}</td>`;
        ptbody.appendChild(row);
      });
      productsTable.appendChild(ptbody);
      fragProducts.appendChild(productsTable);
      document.getElementById('results').innerHTML = "";
      document.getElementById('results').appendChild(fragProducts);
      
      // Build availability table.
      const fragAvail = document.createDocumentFragment();
      const availTable = document.createElement('table');
      availTable.classList.add("availability-table");
      const thead = document.createElement('thead');
      const weekdayRow = document.createElement('tr');
      let emptyTh = document.createElement('th');
      emptyTh.colSpan = 2;
      weekdayRow.appendChild(emptyTh);
      for (let d = startDay; d <= daysInMonth; d++) {
        const dayDate = new Date(selectedYear, selectedMonth, d);
        const weekdayShort = dayDate.toLocaleDateString('en-US', { weekday: 'short' }).substring(0,2);
        let th = document.createElement('th');
        th.textContent = weekdayShort;
        th.style.fontSize = '0.8em';
        weekdayRow.appendChild(th);
      }
      thead.appendChild(weekdayRow);
      const headerRow = document.createElement('tr');
      let nameTh = document.createElement('th');
      nameTh.textContent = "Product Name";
      headerRow.appendChild(nameTh);
      let durationTh = document.createElement('th');
      durationTh.textContent = "Duration";
      headerRow.appendChild(durationTh);
      for (let d = startDay; d <= daysInMonth; d++) {
        let th = document.createElement('th');
        th.textContent = d;
        headerRow.appendChild(th);
      }
      thead.appendChild(headerRow);
      availTable.appendChild(thead);
      const tbodyFragment = document.createDocumentFragment();
      for (const product of products) {
        if (!Array.isArray(product.faresprices) || product.faresprices.length === 0) {
          continue;
        }
        const availData = await checkAvailabilityForProduct(product, accessToken, selectedMonthYear);
        const row = document.createElement('tr');
        // First column: include product name and supplier name.
        row.innerHTML = `<td><a href="https://tdms.websitetravel.com/#search/text/${product.productId}" target="_blank" data-product="${product.name}" data-supplier="${product.supplierName || ''}">${product.name} - ${product.supplierName || ''}</a></td>
                         <td>${product.durationDays || "0"}/${product.durationNight || "0"}</td>`;
        // Add tooltip event listeners for the first-column link.
        const productLink = row.querySelector('td a');
        productLink.addEventListener('mouseenter', function(e) {
          showLinkTooltip(e, product);
        });
        productLink.addEventListener('mouseleave', function(e) {
          hideTooltip(e);
        });
        // Build one cell per day.
        for (let d = startDay; d <= daysInMonth; d++) {
          const cell = document.createElement('td');
          const dayStr = d.toString().padStart(2, '0');
          const dateKey = `${selectedYear}-${formattedMonth}-${dayStr}`;
          if (availData && availData.hasOwnProperty(dateKey)) {
            const responses = availData[dateKey];
            let cellColor = "";
            let cellText = "";
            // Check numeric availability.
            const numAvailabilities = responses.filter(item => typeof item.NumAvailable === 'number').map(item => item.NumAvailable);
            if (numAvailabilities.length > 0) {
              if (numAvailabilities.some(n => n >= 1)) {
                cellColor = 'green';
              } else if (numAvailabilities.some(n => n === 0)) {
                cellColor = '#cc6666'; // red for 0
              } else if (numAvailabilities.every(n => n < 0)) {
                cellColor = 'grey';
              }
            } else {
              // Check error messages.
              const errors = responses.filter(item => item.Error).map(item => item.Error);
              if (errors.length > 0) {
                if (errors.every(e => e.includes("Caching not enabled for this fare") || e.includes("No BookingSystem"))) {
                  cellColor = 'grey';
                } else if (errors.some(e => e.includes("Wrong Season"))) {
                  cellColor = '#cccc66'; // yellow
                } else {
                  cellText = errors[0];
                }
              } else {
                cellText = "N/A";
              }
            }
            cell.textContent = cellText;
            if (cellColor) cell.style.backgroundColor = cellColor;
            cell.dataset.responses = JSON.stringify(responses);
            cell.style.cursor = 'pointer';
            cell.addEventListener('mouseenter', function (e) {
              showCellTooltip(e, responses);
            });
            cell.addEventListener('mouseleave', function (e) {
              hideTooltip(e);
            });
          } else {
            cell.textContent = "N/A";
          }
          row.appendChild(cell);
        }
        tbodyFragment.appendChild(row);
      }
      const tbody = document.createElement('tbody');
      tbody.appendChild(tbodyFragment);
      availTable.appendChild(tbody);
      fragAvail.appendChild(availTable);
      document.getElementById('availabilityResults').innerHTML = "";
      document.getElementById('availabilityResults').appendChild(fragAvail);
      hideSpinner();
    } catch (error) {
      console.error("Error in loadCalendarData:", error);
      document.getElementById('availabilityResults').innerHTML = `<p>Error: ${error.message}</p>`;
      hideSpinner();
    }
  }

  // ------------------------
  // PRESET BUTTONS (Dynamic from external JSON)
  // ------------------------

  function createPresetButtons(presets) {
    const presetsContainer = document.getElementById('presetButtonsContainer');
    presetsContainer.innerHTML = ""; // Clear any existing buttons.
    presets.forEach(preset => {
      const btn = document.createElement('button');
      btn.id = preset.id;
      btn.textContent = preset.name;
      btn.addEventListener('click', function () {
        logMessage(`${preset.name} button clicked.`);
        document.getElementById('productIds').value = preset.productIds;
        currentPreset = preset;
        currentSortingMethod = preset.sorting;
        loadCalendarData();
      });
      presetsContainer.appendChild(btn);
    });
    // Always add the manual load button.
    const manualBtn = document.createElement('button');
    manualBtn.id = "manualLoadButton";
    manualBtn.textContent = "Manual load";
    manualBtn.addEventListener('click', function () {
      logMessage("Manual load button clicked. Revealing input row.");
      document.querySelector('.input-row').style.display = 'flex';
    });
    presetsContainer.appendChild(manualBtn);
  }

  // Fetch presets JSON from an external file (do not hardcode presets in this script).
  fetch('presets.json')
    .then(response => response.json())
    .then(data => {
      createPresetButtons(data.presets);
    })
    .catch(error => console.error("Error fetching presets JSON:", error));

  // Initialize dropdown and attach the manual submit event.
  populateMonthYearDropdown();
  document.getElementById('loadCalendars').addEventListener('click', loadCalendarData);
});
