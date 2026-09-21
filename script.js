(function () {
    "use strict";

    var CCL_PROGRAM = "mrd_lookahead_new_report:group1";
    var allCycles = [];
    var currentSort = {
        key: "name",
        direction: "asc"
    };

    var AUTO_REFRESH_INTERVAL = 60000;
    var autoRefreshTimer = null;

    var searchBox = document.getElementById("searchBox");
    var startDate = document.getElementById("startDate");
    var endDate = document.getElementById("endDate");
    var applyButton = document.getElementById("applyButton");
    var refreshButton = document.getElementById("refreshButton");
    var tableBody = document.getElementById("cycleTableBody");
    var emptyState = document.getElementById("emptyState");
    var message = document.getElementById("message");
    var resultCount = document.getElementById("resultCount");
    var dateRangeLabel = document.getElementById("dateRangeLabel");
    var providerSelect = document.getElementById("providerSelect");

    var lastRefreshed = document.getElementById("lastRefreshed");

    function updateLastRefreshedTime() {
    var now = new Date();

    lastRefreshed.textContent =
        "Last refreshed: " +
        now.toLocaleTimeString("en-CA", {
            hour: "numeric",
            minute: "2-digit",
            second: "2-digit"
        });
    }


    function pad(value) {
        return value < 10 ? "0" + value : String(value);
    }

    function toInputDate(date) {
        return date.getFullYear() + "-" + pad(date.getMonth() + 1) + "-" + pad(date.getDate());
    }

    function setDefaultDateRange() {
        var today = new Date();
        today.setHours(0, 0, 0, 0);

        var daysUntilMonday = (8 - today.getDay()) % 7;
        var monday = new Date(today);
        monday.setDate(today.getDate() + daysUntilMonday);

        var saturday = new Date(monday);
        saturday.setDate(monday.getDate() + 5);

        startDate.value = toInputDate(monday);
        endDate.value = toInputDate(saturday);
    }

    function toCclDate(inputValue) {
        var months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
        var parts = inputValue.split("-");

        return parts[2] + "-" + months[Number(parts[1]) - 1] + "-" + parts[0];
    }

    function displayInputDate(inputValue) {
        var parts = inputValue.split("-");
        var date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));

        return date.toLocaleDateString("en-CA", {
            year: "numeric",
            month: "short",
            day: "numeric"
        });
    }

    function setMessage(text, type) {
        message.textContent = text || "";
        message.className = text ? "message " + type : "message";
    }

    function setLoading(isLoading) {
        applyButton.disabled = isLoading;
        refreshButton.disabled = isLoading;

        if (isLoading) {
            setMessage("Loading pending cycles...", "loading");
        }
    }

    function normalizeId(value) {
        if (value === null || value === undefined || value === "") {
            return "0";
        }

        return String(value).replace(/\.0+$/, "");
    }

    function createCell(row, text, className) {
        var cell = document.createElement("td");
        cell.textContent = text || "";

        if (className) {
            cell.className = className;
        }

        row.appendChild(cell);
        return cell;
    }

    function createLinkCell(row, text, clickHandler) {
        var cell = document.createElement("td");
        var link = document.createElement("a");

        link.href = "#";
        link.className = "action-link";
        link.textContent = text || "";
        link.addEventListener("click", function (event) {
            event.preventDefault();
            clickHandler();
        });

        cell.appendChild(link);
        row.appendChild(cell);
    }

    function launchPatientChart(cycle) {
        var personId = normalizeId(cycle.personId);
        var encntrId = normalizeId(cycle.encntrId);

        if (typeof APPLINK !== "function") {
            setMessage("Patient chart launch is available only inside PowerChart.", "error");
            return;
        }

        APPLINK(
            0,
            "Powerchart.exe",
            "/PERSONID=" + personId + " /ENCNTRID=" + encntrId
        );
    }

    function launchPatientOrders(cycle) {
    var personId = normalizeId(cycle.personId);
    var encntrId = normalizeId(cycle.encntrId);

    if (!personId || personId === "0" ||
        !encntrId || encntrId === "0") {

        setMessage(
            "Missing patient or encounter identifier.",
            "error"
        );

        return;
    }

    if (typeof APPLINK !== "function") {
        setMessage(
            "PowerOrders is available only inside PowerChart.",
            "error"
        );

        return;
    }

    APPLINK(
    0,
    "Powerchart.exe",
    "/PERSONID=" + personId +
    " /ENCNTRID=" + encntrId +
    " /FIRSTTAB=^Orders^"
);
}

function launchPatientschedule(cycle) {
    var personId = normalizeId(cycle.personId);
    var encntrId = normalizeId(cycle.encntrId);

    if (!personId || personId === "0" ||
        !encntrId || encntrId === "0") {

        setMessage(
            "Missing patient or encounter identifier.",
            "error"
        );

        return;
    }

    if (typeof APPLINK !== "function") {
        setMessage(
            "ONC Pt Schedule is available only inside PowerChart.",
            "error"
        );

        return;
    }

APPLINK(
    0,
    "Powerchart.exe",
    "/PERSONID=" + personId +
    " /ENCNTRID=" + encntrId +
    " /FIRSTTAB=^ONC Pt Schedule^"
);
    }

    async function launchAppointmentDetails(cycle) {
    var schEventId = Number(
        normalizeId(cycle.schEventId)
    );

    var scheduleId = Number(
        normalizeId(
            cycle.scheduleId || cycle.scheduledId
        )
    );

    if (!schEventId || !scheduleId) {
        setMessage(
            "Missing appointment identifiers.",
            "error"
        );
        return;
    }

    try {
        var schedulingActions =
            await window.external.DiscernObjectFactory(
                "PEXSCHEDULINGACTIONS"
            );

        schedulingActions.ShowHistoryView(
            schEventId,
            scheduleId
        );

        setMessage("", "");
    } catch (error) {
        console.error(
            "Unable to open Appointment History View:",
            error
        );

        setMessage(
            "Unable to open Appointment History View.",
            "error"
        );
    }
}

    function renderRows(cycles) {
        tableBody.textContent = "";

        cycles.forEach(function (cycle) {
            var row = document.createElement("tr");

            createLinkCell(row, cycle.name, function () {
                launchPatientChart(cycle);
            });
            createCell(row, cycle.mrn);
            createCell(row, cycle.regimenName);
            createLinkCell(row, cycle.cycleDisplay, function () {
                launchPatientOrders(cycle);
            });
            createLinkCell(row, cycle.effectiveStartDate, function () {
            launchPatientschedule(cycle);
            });

            createLinkCell(row, cycle.appointmentType, function () {
                launchAppointmentDetails(cycle);
            });
            createCell(row, cycle.attending);

            var statusCell = createCell(row, "");
            var badge = document.createElement("span");
            badge.className = "status-badge";
            badge.textContent = cycle.schState || "";
            statusCell.appendChild(badge);

            tableBody.appendChild(row);
        });

        emptyState.hidden = cycles.length !== 0;
        resultCount.textContent = cycles.length + (cycles.length === 1 ? " cycle" : " cycles");
    }

    function dateSortValue(value) {
        var months = {
            JAN: 0,
            FEB: 1,
            MAR: 2,
            APR: 3,
            MAY: 4,
            JUN: 5,
            JUL: 6,
            AUG: 7,
            SEP: 8,
            OCT: 9,
            NOV: 10,
            DEC: 11
        };
        var parts = String(value || "").toUpperCase().split("-");

        if (parts.length !== 3 || months[parts[1]] === undefined) {
            return 0;
        }

        return new Date(
            Number(parts[0]),
            months[parts[1]],
            Number(parts[2])
        ).getTime();
    }

    function sortCycles(cycles) {
        var sortedCycles = cycles.slice();
        var multiplier = currentSort.direction === "asc" ? 1 : -1;

        sortedCycles.sort(function (first, second) {
            if (currentSort.key === "effectiveStartDate") {
                return (dateSortValue(first.effectiveStartDate) -
                    dateSortValue(second.effectiveStartDate)) * multiplier;
            }

            return String(first.name || "").localeCompare(
                String(second.name || ""),
                undefined,
                { sensitivity: "base" }
            ) * multiplier;
        });

        return sortedCycles;
    }

    function updateSortHeaders() {
        var headers = document.querySelectorAll("th.sortable");

        Array.prototype.forEach.call(headers, function (header) {
            var isActive = header.getAttribute("data-sort-key") === currentSort.key;

            header.classList.remove("sort-asc", "sort-desc");
            header.setAttribute(
                "aria-sort",
                isActive ?
                    (currentSort.direction === "asc" ? "ascending" : "descending") :
                    "none"
            );

            if (isActive) {
                header.classList.add("sort-" + currentSort.direction);
            }
        });
    }

    function filterRows() {
        var searchValue = searchBox.value.toLowerCase().trim();

        if (!searchValue) {
            renderRows(sortCycles(allCycles));
            return;
        }

        var filteredCycles = allCycles.filter(function (cycle) {
            return String(cycle.name || "").toLowerCase().indexOf(searchValue) !== -1 ||
                String(cycle.mrn || "").toLowerCase().indexOf(searchValue) !== -1;
        });

        renderRows(sortCycles(filteredCycles));
    }

    function initializeSorting() {
        var headers = document.querySelectorAll("th.sortable");

        Array.prototype.forEach.call(headers, function (header) {
            header.addEventListener("click", function () {
                var sortKey = header.getAttribute("data-sort-key");

                if (currentSort.key === sortKey) {
                    currentSort.direction = currentSort.direction === "asc" ? "desc" : "asc";
                } else {
                    currentSort.key = sortKey;
                    currentSort.direction = "asc";
                }

                updateSortHeaders();
                filterRows();
            });
        });

        updateSortHeaders();
    }

    function initializeColumnResizing() {
        var headers = document.querySelectorAll("thead th");

        Array.prototype.forEach.call(headers, function (header) {
            var handle = document.createElement("span");
            handle.className = "resize-handle";
            handle.setAttribute("aria-hidden", "true");

            handle.addEventListener("click", function (event) {
                event.stopPropagation();
            });

            handle.addEventListener("mousedown", function (event) {
                event.preventDefault();
                event.stopPropagation();

                var startX = event.clientX;
                var startWidth = header.offsetWidth;

                handle.classList.add("resizing");
                document.body.classList.add("resizing-columns");

                function resizeColumn(moveEvent) {
                    var newWidth = Math.max(80, startWidth + moveEvent.clientX - startX);
                    header.style.width = newWidth + "px";
                    header.style.minWidth = newWidth + "px";
                }

                function stopResizing() {
                    handle.classList.remove("resizing");
                    document.body.classList.remove("resizing-columns");
                    document.removeEventListener("mousemove", resizeColumn);
                    document.removeEventListener("mouseup", stopResizing);
                }

                document.addEventListener("mousemove", resizeColumn);
                document.addEventListener("mouseup", stopResizing);
            });

            header.appendChild(handle);
        });
    }

    function extractReply(responseText) {
    var parsed = JSON.parse(responseText);

        return parsed.reply || parsed.REPLY || {};
    }

    function asArray(value) {
        if (!value) {
        return [];
    }

        return Array.isArray(value) ? value : [value];
    }

    function populateProviderDropdown(
    providers,
    loggedInUser
) {
    var selectedValue =
        providerSelect.value || "0";

    var providerList =
        asArray(providers);

    providerSelect.textContent = "";

    var myPatientsOption =
        document.createElement("option");

    myPatientsOption.value = "0";

    myPatientsOption.textContent =
        loggedInUser
            ? "My Patients (" + loggedInUser + ")"
            : "My Patients";

    providerSelect.appendChild(
        myPatientsOption
    );

    providerList.forEach(function (provider) {
        var option =
            document.createElement("option");

        option.value = normalizeId(
            provider.personId ||
            provider.providerId
        );

        option.textContent =
            provider.name ||
            provider.providerName ||
            "";

        providerSelect.appendChild(option);
    });

    var previousOption =
        providerSelect.querySelector(
            'option[value="' +
            selectedValue +
            '"]'
        );

    if (previousOption) {
        providerSelect.value =
            selectedValue;
        }
    }


    function validateDateRange() {
        if (!startDate.value || !endDate.value) {
            setMessage("Please select both a start date and an end date.", "error");
            return false;
        }

        if (startDate.value > endDate.value) {
            setMessage("The start date cannot be after the end date.", "error");
            return false;
        }

        return true;
    }

    function scheduleNextAutoRefresh() {
    if (autoRefreshTimer) {
        clearTimeout(autoRefreshTimer);
    }

    autoRefreshTimer = setTimeout(function () {
        loadCycles();
    }, AUTO_REFRESH_INTERVAL);
}


function loadCycles() {
    if (autoRefreshTimer) {
        clearTimeout(autoRefreshTimer);
        autoRefreshTimer = null;
    }

    if (!validateDateRange()) {
        return;
    }

    setLoading(true);

    dateRangeLabel.textContent =
        displayInputDate(startDate.value) +
        " – " +
        displayInputDate(endDate.value);

    if (typeof XMLCclRequest !== "function") {
        setLoading(false);

        setMessage(
            "The CCL request can be run only inside the Cerner MPage environment.",
            "error"
        );

        renderRows([]);
        scheduleNextAutoRefresh();
        return;
    }

    var request = new XMLCclRequest();

    request.onreadystatechange = function () {
        if (request.readyState !== 4) {
            return;
        }

        setLoading(false);
        scheduleNextAutoRefresh();

        if (request.status !== 200) {
            setMessage(
                "Unable to load the pending cycles. Please try again.",
                "error"
            );

            renderRows([]);
            return;
        }

        try {
            var reply = extractReply(request.responseText);

            allCycles = asArray(
                reply.cycle || reply.CYCLE
            );

            populateProviderDropdown(
                reply.provider || reply.PROVIDER,
                reply.loggedInUser ||
                reply.LOGGED_IN_USER
            );

            setMessage("", "");
            filterRows();
            updateLastRefreshedTime();

        } catch (error) {
            console.error(
                "Unable to read pending-cycle data:",
                error
            );

            setMessage(
                "The pending-cycle data could not be read.",
                "error"
            );

            renderRows([]);
        }
    };

    request.open(
        "GET",
        CCL_PROGRAM,
        true
    );

    request.send(
        "^MINE^,^" +
        toCclDate(startDate.value) +
        "^,^" +
        toCclDate(endDate.value) +
        "^," +
        (providerSelect.value || "0")
    );
}
    searchBox.addEventListener("input", filterRows);
    applyButton.addEventListener("click", loadCycles);
    refreshButton.addEventListener("click", loadCycles);
    providerSelect.addEventListener("change", loadCycles);

    initializeSorting();
    initializeColumnResizing();
    setDefaultDateRange();
    loadCycles();
    
}());
