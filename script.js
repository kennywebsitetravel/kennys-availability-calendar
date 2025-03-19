document.addEventListener('DOMContentLoaded', function () {
  // Utility function to log messages
  const logMessage = message => console.log(message);

  // Global variable to hold current sorting method; default is alphabetical.
  let currentSorting = "alphabetical";

  // Load presets from an external JSON file (presets.json)
  function loadPresets() {
    fetch('presets.json')
      .then(response => response.json())
      .then(presetsData => {
        const presetButtonsContainer = document.getElementById('presetButtonsContainer');
        presetButtonsContainer.innerHTML = ""; // Clear existing buttons

        presetsData.presets.forEach(preset => {
          const btn = document.createElement('button');
          btn.id = preset.id;
          btn.textContent = preset.name;
          btn.addEventListener('click', function() {
            logMessage(`${preset.name} button clicked.`);
            document.getElementById('productIds').value = preset.productIds;
            // Save the sorting method globally for later use
            currentSorting = preset.sorting;
            loadCalendarData();
          });
          presetButtonsContainer.appendChild(btn);
        });

        // Always add the Manual load button (static)
        const manualBtn = document.createElement('button');
        manualBtn.id = "manualLoadButton";
        manualBtn.textContent = "Manual load";
        manualBtn.addEventListener('click', function() {
          logMessage("Manual load button clicked. Revealing input row.");
          document.querySelector('.input-row').style.display = 'flex';
        });
        presetButtonsContainer.appendChild(manualBtn);
      })
      .catch(error => {
        console.error('Error loading presets:', error);
      });
  }

  loadPresets();

  // Spinner Functions
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

  // Populate the month/year dropdown for 18 consecutive months (starting with current month)
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

  // Parse month/year string into { monthIndex, year }
  function parseMonthYear(value) {
    const [monthName, yearStr] = value.split(" ");
    const monthNames = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    return { monthIndex: monthNames.indexOf(monthName), year: parseInt(yearStr, 10) };
  }

  // Update the dropdown to a new month/year
  function updateDropdownTo(monthIndex, year) {
    const monthNames = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    const newValue = `${monthNames[monthIndex]} ${year}`;
    document.getElementById('monthYearDropdown').value = newValue;
  }

  // Change month by delta and reload calendar data.
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

  // Tooltip functions for availability cells
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

  // Tooltip functions for first column links (showing supplierName - productName)
  function showLinkTooltip(event, product) {
    const tooltip = document.createElement('div');
    tooltip.classList.add('link-tooltip');
    // supplierName in bold, then product name normal
    tooltip.innerHTML = `<strong>${product.supplierName}</strong> - ${product.name}`;
    document.body.appendChild(tooltip);
    const x = event.pageX + 10;
    const y = event.pageY + 10;
    tooltip.style.left = x + 'px';
    tooltip.style.top = y + 'px';
    event.currentTarget._linkTooltip = tooltip;
  }

  function hideLinkTooltip(event) {
    const tooltip = event.currentTarget._linkTooltip;
    if (tooltip) {
      tooltip.remove();
      event.currentTarget._linkTooltip = null;
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

  // API call to check availability for a given product.
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
      // Sort products according to the current preset's "sorting" value.
      if (currentSorting === "alphabetical") {
        products.sort((a, b) => a.name.localeCompare(b.name));
      } else if (currentSorting === "list") {
        // sort products in the order provided in the preset (using the productIds list order)
        const presetProductIds = document.getElementById('productIds').value.split(',').map(id => id.trim());
        products.sort((a, b) => presetProductIds.indexOf(a.productId.toString()) - presetProductIds.indexOf(b.productId.toString()));
      }

      // Build the product table (for reference)
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
        // Add tooltip for the first column link
        const linkElement = row.querySelector('td a');
        if (linkElement) {
          linkElement.addEventListener('mouseenter', function(e) {
            showLinkTooltip(e, product);
          });
          linkElement.addEventListener('mouseleave', function(e) {
            hideLinkTooltip(e);
          });
        }
        ptbody.appendChild(row);
      });
      productsTable.appendChild(ptbody);
      fragProducts.appendChild(productsTable);
      document.getElementById('results').innerHTML = "";
      document.getElementById('results').appendChild(fragProducts);

      // Build the availability table
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
      const totalProducts = products.length;
      let progressCount = 0;
      document.getElementById('progressUpdate').textContent = progressCount + " / " + totalProducts + " products fetched";
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
        // Add tooltip for first column link in the availability table
        const availLinkElement = row.querySelector('td a');
        if (availLinkElement) {
          availLinkElement.addEventListener('mouseenter', function(e) {
            showLinkTooltip(e, product);
          });
          availLinkElement.addEventListener('mouseleave', function(e) {
            hideLinkTooltip(e);
          });
        }
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
                       responses.every(item => item.NumAvailable === 0)) {
              cell.textContent = "";
              cell.style.backgroundColor = '#cc6666';
            } else if (responses.every(item => item.Error === "Caching not enabled for this fare" ||
                                                 item.Error === "No BookingSystem" ||
                                                 item.Error === "Wrong Season / Caching not enabled for this fare" ||
                                                 item.Error === "Wrong Season / No BookingSystem")) {
              cell.textContent = "";
              cell.style.backgroundColor = 'grey';
            } else if (responses.every(item => item.Error &&
                       (item.Error === "Wrong Season" ||
                        item.Error === "Unable to fetch cached availability. Use checkavailabilityrange to get availability."))) {
              cell.textContent = "";
              cell.style.backgroundColor = '#cccc66';
            } else {
              const validResponse = responses.find(item => item.NumAvailable !== undefined);
              if (validResponse) {
                cell.textContent = (validResponse.NumAvailable === 0)
                  ? ""
                  : validResponse.NumAvailable;
                if (validResponse.NumAvailable === 0) {
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

  populateMonthYearDropdown();
  document.getElementById('loadCalendars').addEventListener('click', loadCalendarData);
});
