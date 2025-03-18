document.addEventListener('DOMContentLoaded', function () {
  // Utility function to log key messages.
  const logMessage = message => console.log(message);

  // New functions for product tooltip.
  function showProductTooltip(event, product) {
    const tooltip = document.createElement('div');
    tooltip.classList.add('tooltip');
    // Use fallback if supplierName is missing.
    const supplier = product.supplierName ? product.supplierName : "No supplier";
    tooltip.textContent = product.name + " - " + supplier;
    document.body.appendChild(tooltip);
    const x = event.pageX + 10;
    const y = event.pageY + 10;
    tooltip.style.left = x + 'px';
    tooltip.style.top = y + 'px';
    event.currentTarget._tooltip = tooltip;
  }

  function hideProductTooltip(event) {
    const tooltip = event.currentTarget._tooltip;
    if (tooltip) {
      tooltip.remove();
      event.currentTarget._tooltip = null;
    }
  }

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

  // Populate the month/year dropdown for 18 consecutive months (starting with current month).
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

  // Parse a month/year string into { monthIndex, year } (monthIndex is 0-based).
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
  } else {
    console.error("Previous month button not found.");
  }
  if (nextMonthBtn) {
    nextMonthBtn.addEventListener('click', () => {
      logMessage("Next month button clicked.");
      changeMonth(1);
    });
  } else {
    console.error("Next month button not found.");
  }

  // Tooltip Functions for availability cells.
  function showTooltip(event, responses) {
    if (responses.every(resp => resp.Error === "Caching not enabled for this fare")) {
      const tooltip = document.createElement('div');
      tooltip.classList.add('tooltip');
      tooltip.textContent = "Caching not enabled for this fare";
      document.body.appendChild(tooltip);
      const x = event.pageX + 10;
      const y = event.pageY + 10;
      tooltip.style.left = x + 'px';
      tooltip.style.top = y + 'px';
      event.currentTarget._tooltip = tooltip;
      return;
    }
    if (responses.every(resp => resp.Error &&
         (resp.Error === "Wrong Season" ||
          resp.Error === "No BookingSystem" ||
          resp.Error === "Wrong Season / No BookingSystem" ||
          resp.Error === "Unable to fetch cached availability. Use checkavailabilityrange to get availability."))) {
      const tooltip = document.createElement('div');
      tooltip.classList.add('tooltip');
      tooltip.textContent = "Wrong Season";
      document.body.appendChild(tooltip);
      const x = event.pageX + 10;
      const y = event.pageY + 10;
      tooltip.style.left = x + 'px';
      tooltip.style.top = y + 'px';
      event.currentTarget._tooltip = tooltip;
      return;
    }
    let validResponses = responses.filter(resp => resp.Error !== "No BookingSystem" && resp.Error !== "Wrong Season");
    const uniqueResponses = {};
    validResponses.forEach(resp => {
      const id = resp.ProductPricesDetailsId;
      if (!(id in uniqueResponses)) {
        uniqueResponses[id] = resp;
      } else {
        if (resp.NumAvailable !== undefined && resp.NumAvailable > 0 &&
            (uniqueResponses[id].NumAvailable === undefined || uniqueResponses[id].NumAvailable <= 0)) {
          uniqueResponses[id] = resp;
        }
      }
    });
    validResponses = Object.values(uniqueResponses);
    if (validResponses.length === 0) return;

    const tooltip = document.createElement('div');
    tooltip.classList.add('tooltip');

    if (responses[0] && responses[0].ProductName) {
      const prodNameEl = document.createElement('div');
      prodNameEl.textContent = responses[0].ProductName;
      prodNameEl.style.fontWeight = 'bold';
      prodNameEl.style.marginBottom = '5px';
      tooltip.appendChild(prodNameEl);
    }
    const table = document.createElement('table');
    table.classList.add('tooltip-table');
    const headerRow = document.createElement('tr');
    const fareNameTh = document.createElement('th');
    fareNameTh.textContent = "Fare Name";
    const numAvailableTh = document.createElement('th');
    numAvailableTh.textContent = "Available";
    headerRow.appendChild(fareNameTh);
    headerRow.appendChild(numAvailableTh);
    table.appendChild(headerRow);

    validResponses.forEach(resp => {
      const row = document.createElement('tr');
      const fareNameTd = document.createElement('td');
      fareNameTd.textContent = resp.FareName || "";
      const numAvailableTd = document.createElement('td');
      numAvailableTd.textContent = (resp.NumAvailable !== undefined ? resp.NumAvailable : "");
      row.appendChild(fareNameTd);
      row.appendChild(numAvailableTd);
      table.appendChild(row);
    });
    tooltip.appendChild(table);
    document.body.appendChild(tooltip);
    const x = event.pageX + 10;
    const y = event.pageY + 10;
    tooltip.style.left = x + 'px';
    tooltip.style.top = y + 'px';
    tooltip.style.pointerEvents = 'none';
    event.currentTarget._tooltip = tooltip;
  }

  function hideTooltip(event) {
    const tooltip = event.currentTarget._tooltip;
    if (tooltip) {
      tooltip.remove();
      event.currentTarget._tooltip = null;
    }
  }

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

  // API Function using the ephemeral token.
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

  // Main Function to load calendar data.
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
      products.sort((a, b) => a.name.localeCompare(b.name));

      // Build the (hidden) products table.
      const fragProducts = document.createDocumentFragment();
      const productsTable = document.createElement('table');
      // Create table header.
      const thead = document.createElement('thead');
      const headerRow = document.createElement('tr');
      const thName = document.createElement('th');
      thName.textContent = "Product Name";
      const thDuration = document.createElement('th');
      thDuration.textContent = "Duration";
      const thPrices = document.createElement('th');
      thPrices.textContent = "Product Prices Details IDs";
      headerRow.appendChild(thName);
      headerRow.appendChild(thDuration);
      headerRow.appendChild(thPrices);
      thead.appendChild(headerRow);
      productsTable.appendChild(thead);

      const ptbody = document.createElement('tbody');
      products.forEach(product => {
        const row = document.createElement('tr');

        // Product Name column with tooltip on hover.
        const tdName = document.createElement('td');
        const link = document.createElement('a');
        link.href = `https://tdms.websitetravel.com/#search/text/${product.productId}`;
        link.target = "_blank";
        link.textContent = product.name;
        link.addEventListener('mouseenter', (e) => {
          showProductTooltip(e, product);
        });
        link.addEventListener('mouseleave', hideProductTooltip);
        tdName.appendChild(link);

        // Duration column.
        const tdDuration = document.createElement('td');
        tdDuration.textContent = (product.durationDays || "0") + "/" + (product.durationNight || "0");

        // Product Prices Details IDs column.
        const tdPrices = document.createElement('td');
        if (Array.isArray(product.faresprices)) {
          tdPrices.textContent = product.faresprices.map(f => f.productPricesDetailsId).join(', ');
        } else {
          tdPrices.textContent = "N/A";
        }

        row.appendChild(tdName);
        row.appendChild(tdDuration);
        row.appendChild(tdPrices);
        ptbody.appendChild(row);
      });
      productsTable.appendChild(ptbody);
      fragProducts.appendChild(productsTable);
      document.getElementById('results').innerHTML = "";
      document.getElementById('results').appendChild(fragProducts);

      // Build the Availability Table.
      const fragAvail = document.createDocumentFragment();
      const availTable = document.createElement('table');
      availTable.classList.add("availability-table");
      const availThead = document.createElement('thead');

      // First header row: Weekday initials from startDay to daysInMonth.
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
      availThead.appendChild(weekdayRow);

      // Second header row: "Product Name", "Duration", then day numbers.
      const headerRow2 = document.createElement('tr');
      let nameTh2 = document.createElement('th');
      nameTh2.textContent = "Product Name";
      headerRow2.appendChild(nameTh2);
      let durationTh2 = document.createElement('th');
      durationTh2.textContent = "Duration";
      headerRow2.appendChild(durationTh2);
      for (let d = startDay; d <= daysInMonth; d++) {
        let th = document.createElement('th');
        th.textContent = d;
        headerRow2.appendChild(th);
      }
      availThead.appendChild(headerRow2);
      availTable.appendChild(availThead);

      // Set up progress update.
      const totalProducts = products.length;
      let progressCount = 0;
      document.getElementById('progressUpdate').textContent = progressCount + " / " + totalProducts + " products fetched";

      // Build table rows and update progress after each product is processed.
      const tbodyFragment = document.createDocumentFragment();
      for (const product of products) {
        if (!Array.isArray(product.faresprices) || product.faresprices.length === 0) {
          progressCount++;
          document.getElementById('progressUpdate').textContent = progressCount + " / " + totalProducts + " products fetched";
          continue;
        }
        const availData = await checkAvailabilityForProduct(product, accessToken, selectedMonthYear);
        const row = document.createElement('tr');
        row.innerHTML = `<td><a href="https://tdms.websitetravel.com/#search/text/${product.productId}" target="_blank">${product.name}</a></td>
                         <td>${product.durationDays || "0"}/${product.durationNight || "0"}</td>`;
        for (let d = startDay; d <= daysInMonth; d++) {
          const cell = document.createElement('td');
          const dayStr = d.toString().padStart(2, '0');
          const dateKey = `${selectedYear}-${formattedMonth}-${dayStr}`;
          if (availData && availData.hasOwnProperty(dateKey)) {
            const responses = availData[dateKey];
            if (responses.some(item => item.NumAvailable !== undefined && item.NumAvailable >= 1)) {
              cell.textContent = "";
              cell.style.backgroundColor = 'green';
            } else if (responses.some(item => item.NumAvailable !== undefined) &&
                       responses.every(item => item.NumAvailable === 0 || item.NumAvailable === -1)) {
              cell.textContent = "";
              cell.style.backgroundColor = '#cc6666';
            } else if (responses.every(item => item.Error === "Caching not enabled for this fare")) {
              cell.textContent = "";
              cell.style.backgroundColor = '#cc6666';
            } else if (responses.every(item => item.Error && (
                         item.Error === "Wrong Season" ||
                         item.Error === "No BookingSystem" ||
                         item.Error === "Wrong Season / No BookingSystem" ||
                         item.Error === "Unable to fetch cached availability. Use checkavailabilityrange to get availability."))) {
              cell.textContent = "";
              cell.style.backgroundColor = '#cccc66';
            } else {
              const validResponse = responses.find(item => item.NumAvailable !== undefined);
              if (validResponse) {
                cell.textContent = (validResponse.NumAvailable === 0 || validResponse.NumAvailable === -1)
                  ? ""
                  : validResponse.NumAvailable;
                if (validResponse.NumAvailable === 0 || validResponse.NumAvailable === -1) {
                  cell.style.backgroundColor = '#cc6666';
                }
              } else {
                cell.textContent = responses[0].Error || "N/A";
              }
            }
            cell.dataset.responses = JSON.stringify(responses);
            cell.addEventListener('mouseenter', function(e) {
              showTooltip(e, responses);
            });
            cell.addEventListener('mouseleave', function(e) {
              hideTooltip(e);
            });
          } else {
            cell.textContent = "N/A";
          }
          cell.style.cursor = 'pointer';
          cell.addEventListener('click', function() {
            window.open(`https://tdms.websitetravel.com/#search/text/${product.productId}`, '_blank');
          });
          row.appendChild(cell);
        }
        tbodyFragment.appendChild(row);
        progressCount++;
        document.getElementById('progressUpdate').textContent = progressCount + " / " + totalProducts + " products fetched";
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

  // Preset Buttons Event Listeners.
  const presetWhitsunday = document.getElementById('presetWhitsunday');
  if (presetWhitsunday) {
    presetWhitsunday.addEventListener('click', function() {
      logMessage("Preset Whitsunday button clicked.");
      document.getElementById('productIds').value = "66086,66080,66083,43544,49872,3542,54377,68213,3549,17400,54332,66095,66969,66092,54367,54584,54533,54317,54530,54289,54284,67804,54775,54342,54254,54244,54347,54890,54956,54587,54294,66248,54269";
      loadCalendarData();
    });
  } else {
    console.error("Preset Whitsunday button not found.");
  }

  const presetGreatBarrierReef = document.getElementById('presetGreatBarrierReef');
  if (presetGreatBarrierReef) {
    presetGreatBarrierReef.addEventListener('click', function() {
      logMessage("Preset Great Barrier Reef button clicked.");
      document.getElementById('productIds').value = "66086,66080";
      loadCalendarData();
    });
  } else {
    console.error("Preset Great Barrier Reef button not found.");
  }

  const presetOtherTours = document.getElementById('presetOtherTours');
  if (presetOtherTours) {
    presetOtherTours.addEventListener('click', function() {
      logMessage("Preset Other Tours button clicked.");
      document.getElementById('productIds').value = "20633,66468,4308,52087,52107,18541,52092,68248,68251,66006,49688,49672";
      loadCalendarData();
    });
  } else {
    console.error("Preset Other Tours button not found.");
  }

  const manualLoadButton = document.getElementById('manualLoadButton');
  if (manualLoadButton) {
    manualLoadButton.addEventListener('click', function() {
      logMessage("Manual load button clicked. Revealing input row.");
      document.querySelector('.input-row').style.display = 'flex';
    });
  } else {
    console.error("Manual load button not found.");
  }

  // Initialize dropdown and attach Submit event.
  populateMonthYearDropdown();
  document.getElementById('loadCalendars').addEventListener('click', loadCalendarData);
});
