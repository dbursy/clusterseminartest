/* script.js -- loads seminars.xlsx and builds the DOM dynamically.
   Requirements: place seminars.xlsx in same directory as index.html.
   Dependencies: SheetJS (xlsx.full.min.js) and AddEvent script/link in index.html.
*/

/* ------------------ Dark mode helpers (same behavior as before) ------------------ */
function updateToggleIcon() {
  const btn = document.querySelector(".dark-toggle");
  if (btn) btn.textContent = document.body.classList.contains("dark-mode") ? "☀️" : "🌙";
}

function switchCalendarTheme(darkMode) {
  const themeLink = document.getElementById('addevent-theme');
  if (!themeLink) return;
  themeLink.href = darkMode
    ? "https://cdn.addevent.com/libs/atc/themes/000-theme-1/theme.css"
    : "https://cdn.addevent.com/libs/atc/themes/fff-theme-1/theme.css";
}

function toggleDarkMode() {
  document.body.classList.toggle("dark-mode");
  const isDark = document.body.classList.contains("dark-mode");
  localStorage.setItem("darkMode", isDark);
  updateToggleIcon();
  switchCalendarTheme(isDark);
}

/* ------------------ Abstract toggle ------------------ */
function toggleAbstract(id) {
  const abstract = document.getElementById(id);
  if (!abstract) return;
  abstract.classList.toggle("expanded");
  const btn = abstract.querySelector(".read-more-btn");
  if (btn) btn.textContent = abstract.classList.contains("expanded") ? " Show less" : " Show more";
}

/* ------------------ Utilities for date/time formatting ------------------ */
function formatDisplayDate(isoDate) {
  // show as DD.MM.YYYY for German readers
  const d = new Date(isoDate + "T00:00:00");
  if (isNaN(d)) return isoDate;
  return d.toLocaleDateString("de-DE");
}

function toUSDateString(isoDate) {
  // returns MM/DD/YYYY used by AddEvent examples
  const d = new Date(isoDate + "T00:00:00");
  if (isNaN(d)) return isoDate;
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
}

function formatTime12h(timeStr) {
  // input: "09:00" or "09:00 AM" / "17:30"
  // output: "h:mm AM/PM"
  if (!timeStr) return "";
  // accept "HH:MM" (24h)
  const m = timeStr.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return timeStr;
  let hh = parseInt(m[1], 10);
  const mm = m[2];
  const ampm = hh >= 12 ? "PM" : "AM";
  hh = hh % 12;
  if (hh === 0) hh = 12;
  return `${hh}:${mm} ${ampm}`;
}

/* ------------------ AddEvent re-init helper ------------------ */
function reinstateAddEventScript() {
  // Re-inserting the AddEvent script tag forces the library to pick up newly injected buttons.
  // Remove existing script tag (but keep the theme link) and add a fresh one.
  const existing = document.getElementById("addevent-script");
  if (existing) existing.remove();
  const s = document.createElement("script");
  s.id = "addevent-script";
  s.type = "text/javascript";
  s.src = "https://cdn.addevent.com/libs/atc/1.6.1/atc.min.js";
  s.async = true;
  s.defer = true;
  document.head.appendChild(s);
}

/* ------------------ Build DOM from parsed rows ------------------ */
function createSeminarEntry(sheetSlug, row, index) {
  // row expected keys: Date (YYYY-MM-DD), Speaker, Title, "Abstract (short)", "Abstract (long)", StartTime, EndTime, Location
  const container = document.createElement("div");
  container.className = "seminar-entry";
  const isoDate = row.Date ? String(row.Date).trim() : "";
  container.setAttribute("data-date", isoDate);

  const status = document.createElement("div");
  status.className = "seminar-status";
  container.appendChild(status);

  const header = document.createElement("div");
  header.innerHTML = `<strong>${isoDate ? formatDisplayDate(isoDate) : ""} — ${escapeHtml(row.Speaker || "")}</strong><br>
    <em>${escapeHtml(row.Title || "")}</em>`;
  container.appendChild(header);

  // Abstract / short + long
  const absId = `${sheetSlug.replace(/\s+/g, "_")}_${index}`;
  const p = document.createElement("p");
  p.className = "abstract";
  p.id = absId;

  const shortText = row["Abstract (short)"] || "";
  const longText = row["Abstract (long)"] || "";
  p.innerHTML = `${escapeHtml(shortText)}<span class="dots"> [...]</span><span class="more-text">${escapeHtml(longText)}</span>
    <span class="read-more-btn" onclick="toggleAbstract('${absId}')"> Show more</span>`;
  container.appendChild(p);

  // AddEvent button
  const atc = document.createElement("div");
  atc.className = "addeventatc";
  atc.setAttribute("title", "Add to Calendar");
  atc.setAttribute("data-styling", "none");

  // Determine start/end strings for AddEvent
  const startDateUS = isoDate ? toUSDateString(isoDate) : "";
  const startTime12 = row.StartTime ? formatTime12h(row.StartTime) : "";
  const endTime12 = row.EndTime ? formatTime12h(row.EndTime) : "";

  atc.innerHTML = `
    Add to Calendar
    <span class="addeventatc_icon"></span>
    <span class="start">${startDateUS} ${startTime12}</span>
    <span class="end">${startDateUS} ${endTime12}</span>
    <span class="timezone">Europe/Berlin</span>
    <span class="title">Cluster Seminar: ${escapeHtml(row.Speaker || "")}</span>
    <span class="description">${escapeHtml(row.Title || "")}</span>
    <span class="location">${escapeHtml(row.Location || "")}</span>
  `;

  container.appendChild(atc);

  return container;
}

/* ------------------ Main loader (reads seminars.xlsx) ------------------ */
async function loadSeminarsXLSX(xlsxPath = "seminars.xlsx") {
  try {
    const res = await fetch(xlsxPath);
    if (!res.ok) throw new Error(`Failed to fetch ${xlsxPath} (status ${res.status})`);
    const arrayBuffer = await res.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: "array" });

    const root = document.getElementById("seminars-root");
    root.innerHTML = ""; // clear placeholder

    // For ordering: sort sheet names alphabetically or preserve workbook order
    const sheetNames = wb.SheetNames;

    for (const sheetName of sheetNames) {
      // create details/summary block
      const details = document.createElement("details");
      details.className = "semester-block";
      const summary = document.createElement("summary");
      summary.textContent = sheetName;
      details.appendChild(summary);

      // parse rows as objects
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" }); // defval to avoid undefined

      if (rows.length === 0) {
        const p = document.createElement("p");
        p.style.opacity = 0.6;
        p.textContent = "No entries in this sheet.";
        details.appendChild(p);
      } else {
        rows.forEach((row, i) => {
          const entry = createSeminarEntry(makeSheetSlug(sheetName), row, i);
          details.appendChild(entry);
        });
      }

      root.appendChild(details);
    }

    // After injecting DOM, re-init AddEvent script so it picks up new buttons
    reinstateAddEventScript();

    // mark past/upcoming/future and auto-collapse semesters
    markEvents();

  } catch (err) {
    const root = document.getElementById("seminars-root");
    root.innerHTML = `<p style="color:crimson">Error loading seminars.xlsx: ${escapeHtml(String(err))}</p>`;
    console.error(err);
  }
}

/* ------------------ Mark events & auto-collapse logic (reused) ------------------ */
function markEvents() {
  const now = new Date();
  let firstUpcoming = null;

  document.querySelectorAll(".seminar-entry").forEach(entry => {
    const dateStr = entry.getAttribute("data-date");
    const statusEl = entry.querySelector(".seminar-status");
    // reset classes
    entry.classList.remove("past", "upcoming");
    if (!dateStr) {
      if (statusEl) statusEl.textContent = "";
      return;
    }
    // parse ISO-like date
    const date = new Date(dateStr + "T12:00:00"); // midday safe
    if (isNaN(date)) {
      if (statusEl) statusEl.textContent = "";
      return;
    }
    if (date < now) {
      entry.classList.add("past");
      if (statusEl) {
        statusEl.textContent = "Past Event";
        statusEl.classList.remove("upcoming");
        statusEl.classList.add("past");
      }
    } else {
      if (!firstUpcoming) {
        entry.classList.add("upcoming");
        firstUpcoming = entry;
        if (statusEl) {
          statusEl.textContent = "Upcoming Event";
          statusEl.classList.remove("past");
          statusEl.classList.add("upcoming");
        }
      } else {
        if (statusEl) {
          statusEl.textContent = "Future Event";
          statusEl.classList.remove("past","upcoming");
        }
      }
    }
  });

  // collapse all semesters, open only the one containing firstUpcoming
  document.querySelectorAll(".semester-block").forEach(d => d.removeAttribute("open"));
  if (firstUpcoming) {
    const sem = firstUpcoming.closest(".semester-block");
    if (sem) sem.setAttribute("open", "true");
    // optionally scroll into view:
    // firstUpcoming.scrollIntoView({behavior: "smooth", block: "center"});
  } else {
    // if all past, open the most recent semester (first one)
    const semesters = document.querySelectorAll(".semester-block");
    if (semesters.length > 0) semesters[0].setAttribute("open", "true");
  }
}

/* ------------------ helpers ------------------ */
function makeSheetSlug(name) {
  return name.replace(/[^\w\- ]+/g, "").replace(/\s+/g, "_");
}
function escapeHtml(unsafe) {
  if (unsafe === null || unsafe === undefined) return "";
  return String(unsafe)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

/* ------------------ boot ------------------ */
window.addEventListener("load", function () {
  // restore dark mode preference
  const isDark = localStorage.getItem("darkMode") === "true";
  if (isDark) document.body.classList.add("dark-mode");
  updateToggleIcon();
  switchCalendarTheme(isDark);

  // load seminars.xlsx from same folder
  loadSeminarsXLSX("seminars.xlsx");
});
