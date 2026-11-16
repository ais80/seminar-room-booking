/**
 * QE Pharmacy Seminar Room Booking System
 * Application Logic & UI Controller
 * University Hospitals Birmingham NHS Foundation Trust
 */

(function () {
  "use strict";

  // Application State
  const state = {
    selectedDate: new Date(),
    selectedSite: "QEH Heritage",
    selectedRoomId: "qeh-seminar-room",
    viewMode: "day", // "day" | "week" | "agenda"
    filterDepartment: "all",
    selectedDuration: 30, // 30, 60, 90, 120 minutes
    activeSlotStart: "09:00",
    editingBookingId: null,
    recurrence: null, // null if one-off, or recurrence config object
    isDarkMode: localStorage.getItem("qe_theme") === "dark",
    pendingPinAction: "cancel",
    pinFailedAttempts: 0
  };

  // DOM Elements Cache
  const elements = {};

  function initElements() {
    elements.roomStatusHero = document.getElementById("room-status-hero");
    elements.roomStatusText = document.getElementById("room-status-text");
    elements.roomStatusSub = document.getElementById("room-status-sub");
    elements.roomStatusMeta = document.getElementById("room-status-meta");
    elements.quickBookBtn = document.getElementById("quick-book-btn");
    elements.btnBookSession = document.getElementById("btn-book-session");
    elements.printDoorSignBtn = document.getElementById("print-door-sign-btn");
    elements.adminToggleBtn = document.getElementById("admin-toggle-btn");
    elements.adminActiveBar = document.getElementById("admin-active-bar");
    elements.themeToggleBtn = document.getElementById("theme-toggle-btn");
    elements.syncStatusBadge = document.getElementById("sync-status-badge");
    elements.syncDot = document.getElementById("sync-dot");
    elements.syncText = document.getElementById("sync-text");

    // Calendar Navigation
    elements.btnPrevDate = document.getElementById("btn-prev-date");
    elements.btnNextDate = document.getElementById("btn-next-date");
    elements.btnToday = document.getElementById("btn-today");
    elements.dateDisplayTitle = document.getElementById("date-display-title");
    elements.calendarDateInput = document.getElementById("calendar-date-input");
    elements.hiddenDatePicker = document.getElementById("hidden-date-picker");

    // Views
    elements.viewDayBtn = document.getElementById("view-day-btn");
    elements.viewWeekBtn = document.getElementById("view-week-btn");
    elements.viewAgendaBtn = document.getElementById("view-agenda-btn");
    elements.dayTimelineContainer = document.getElementById("day-timeline-container");
    elements.weekViewContainer = document.getElementById("week-view-container");
    elements.agendaViewContainer = document.getElementById("agenda-view-container");
    elements.departmentFilter = document.getElementById("department-filter");

    // Booking Modal
    elements.bookingModal = document.getElementById("booking-modal");
    elements.bookingModalTitle = document.getElementById("booking-modal-title");
    elements.bookingForm = document.getElementById("booking-form");
    elements.bookingDateInput = document.getElementById("booking-date");
    elements.bookingStartTimeInput = document.getElementById("booking-start-time");
    elements.bookingDurationInput = document.getElementById("booking-duration");
    elements.durationButtons = document.querySelectorAll(".duration-btn");
    elements.bookerNameInput = document.getElementById("booker-name");
    elements.bookerEmailInput = document.getElementById("booker-email");
    elements.bookerBleepInput = document.getElementById("booker-bleep");
    elements.bookerDeptInput = document.getElementById("booker-dept");
    elements.bookingTitleInput = document.getElementById("booking-title");
    elements.bookingAttendeesInput = document.getElementById("booking-attendees");
    elements.bookerPinInput = document.getElementById("booker-pin");
    elements.bookingSubmitBtn = document.getElementById("booking-submit-btn");

    // Live Availability Preview
    elements.checkerHeaderTitle = document.getElementById("checker-header-title");
    elements.checkerSlotCount = document.getElementById("checker-slot-count");
    elements.checkerSlotsGrid = document.getElementById("checker-slots-grid");
    elements.checkerConflictMsg = document.getElementById("checker-conflict-message");

    // Recurrence Section in Booking Modal
    elements.recurrenceSummaryText = document.getElementById("recurrence-summary-text");
    elements.btnOpenRecurrenceModal = document.getElementById("btn-open-recurrence-modal");

    // Outlook Recurrence Modal Elements
    elements.recurrenceModal = document.getElementById("recurrence-modal");
    elements.recTimeStart = document.getElementById("rec-time-start");
    elements.recTimeEnd = document.getElementById("rec-time-end");
    elements.recTimeDuration = document.getElementById("rec-time-duration");
    elements.recPatternRadios = document.querySelectorAll('input[name="rec-pattern-type"]');
    elements.recPanelWeekly = document.getElementById("rec-panel-weekly");
    elements.recPanelDaily = document.getElementById("rec-panel-daily");
    elements.recPanelMonthly = document.getElementById("rec-panel-monthly");
    elements.recPanelYearly = document.getElementById("rec-panel-yearly");
    elements.recWeeklyInterval = document.getElementById("rec-weekly-interval");
    elements.recDailyInterval = document.getElementById("rec-daily-interval");
    elements.recDailyRadios = document.querySelectorAll('input[name="rec-daily-mode"]');
    elements.recMonthlyDay = document.getElementById("rec-monthly-day");
    elements.recMonthlyInterval = document.getElementById("rec-monthly-interval");
    elements.recRangeStart = document.getElementById("rec-range-start");
    elements.recRangeRadios = document.querySelectorAll('input[name="rec-range-mode"]');
    elements.recRangeEndDate = document.getElementById("rec-range-end-date");
    elements.recRangeOccurrences = document.getElementById("rec-range-occurrences");
    elements.btnRecOk = document.getElementById("btn-rec-ok");
    elements.btnRecCancel = document.getElementById("btn-rec-cancel");
    elements.btnRecRemove = document.getElementById("btn-rec-remove");

    // Details Modal
    elements.detailsModal = document.getElementById("details-modal");
    elements.detailsBody = document.getElementById("details-modal-body");
    elements.btnOutlookInvite = document.getElementById("btn-outlook-invite");
    elements.btnEditBooking = document.getElementById("btn-edit-booking");
    elements.btnCancelBooking = document.getElementById("btn-cancel-booking");

    // Admin Modals
    elements.adminLoginModal = document.getElementById("admin-login-modal");
    elements.adminPinInput = document.getElementById("admin-pin-input");
    elements.adminLoginSubmit = document.getElementById("admin-login-submit");
    elements.adminLogoutBtn = document.getElementById("admin-logout-btn");
    elements.adminBlockBtn = document.getElementById("admin-block-btn");
    elements.adminExportBtn = document.getElementById("admin-export-btn");
    elements.adminBlockModal = document.getElementById("admin-block-modal");
    elements.adminBlockForm = document.getElementById("admin-block-form");

    // Door Sign Modal
    elements.doorSignModal = document.getElementById("door-sign-modal");
    elements.doorQrContainer = document.getElementById("door-qr-code");
    elements.doorQrUrlInput = document.getElementById("door-qr-url-input");
    elements.btnCopyDoorUrl = document.getElementById("btn-copy-door-url");
    elements.btnRefreshDoorQr = document.getElementById("btn-refresh-door-qr");
    elements.printSignActionBtn = document.getElementById("print-sign-action-btn");

    // Cloud Settings Modal
    elements.cloudSettingsModal = document.getElementById("cloud-settings-modal");
    elements.cloudUrlInput = document.getElementById("cloud-url-input");
    elements.saveCloudSettingsBtn = document.getElementById("save-cloud-settings-btn");

    // Cancellation PIN & MIB Neuralyzer Elements
    elements.neuralyzerFlash = document.getElementById("neuralyzer-flash");
    elements.cancelPinModal = document.getElementById("cancel-pin-modal");
    elements.cancelPinTitle = document.getElementById("cancel-pin-title");
    elements.cancelPinDesc = document.getElementById("cancel-pin-desc");
    elements.cancelSlotPin = document.getElementById("cancel-slot-pin");
    elements.cancelPinError = document.getElementById("cancel-pin-error");
    elements.btnVerifyCancelPin = document.getElementById("btn-verify-cancel-pin");
    elements.btnNotMyBooking = document.getElementById("btn-not-my-booking");
    elements.mibModal = document.getElementById("mib-modal");
    elements.mibAdminPassword = document.getElementById("mib-admin-password");
    elements.btnMibOverrideCancel = document.getElementById("btn-mib-override-cancel");
    elements.mibAuthStatus = document.getElementById("mib-auth-status");
    elements.btnMibDismiss = document.getElementById("btn-mib-dismiss");
    elements.mibGif = document.getElementById("mibGif");
    elements.lockoutSoundBtn = document.getElementById("lockoutSoundBtn");

    // Site & Room Navigation Elements
    elements.siteTabsContainer = document.getElementById("site-tabs-container");
    elements.roomSelectDropdown = document.getElementById("room-select-dropdown");
    elements.roomMetaLocation = document.getElementById("room-meta-location");
    elements.roomMetaCapacity = document.getElementById("room-meta-capacity");
    elements.roomMetaFacilities = document.getElementById("room-meta-facilities");
    elements.btnAdminAddRoomQuick = document.getElementById("btn-admin-add-room-quick");

    // Room Selection in Booking Modal
    elements.bookingRoom = document.getElementById("booking-room");
    elements.bookingRoomCapHint = document.getElementById("booking-room-cap-hint");
    elements.staffSuggestions = document.getElementById("staff-suggestions");
    elements.staffMatchBadge = document.getElementById("staff-match-badge");
    elements.bookerDeptSelect = elements.bookerDeptInput;

    // Staff Users Management Modal
    elements.manageUsersModal = document.getElementById("manage-users-modal");
    elements.adminManageUsersBtn = document.getElementById("admin-manage-users-btn");
    elements.userSearchInput = document.getElementById("user-search-input");
    elements.adminTotalUsersCount = document.getElementById("admin-total-users-count");
    elements.adminUsersTableBody = document.getElementById("admin-users-table-body");
    elements.adminAddUserForm = document.getElementById("admin-add-user-form");
    elements.editUserId = document.getElementById("edit-user-id");
    elements.userFormName = document.getElementById("user-form-name");
    elements.userFormEmail = document.getElementById("user-form-email");
    elements.userFormBleep = document.getElementById("user-form-bleep");
    elements.userFormDept = document.getElementById("user-form-dept");
    elements.userFormRole = document.getElementById("user-form-role");
    elements.userFormPin = document.getElementById("user-form-pin");
    elements.btnSaveUserSubmit = document.getElementById("btn-save-user-submit");
    elements.btnCancelEditUser = document.getElementById("btn-cancel-edit-user");
    elements.userCsvFileInput = document.getElementById("user-csv-file-input");
    elements.btnImportUsersCsv = document.getElementById("btn-import-users-csv");
    elements.btnExportUsersCsv = document.getElementById("btn-export-users-csv");

    // Door Sign Elements
    elements.doorSiteSelect = document.getElementById("door-site-select");
    elements.doorRoomSelect = document.getElementById("door-room-select");
    elements.doorSignRoomTitle = document.getElementById("door-sign-room-title");
    elements.doorSignRoomSubtitle = document.getElementById("door-sign-room-subtitle");
    elements.doorSignEtiquetteList = document.getElementById("door-sign-etiquette-list");

    // Rooms Management Modal
    elements.manageRoomsModal = document.getElementById("manage-rooms-modal");
    elements.adminTotalRoomsCount = document.getElementById("admin-total-rooms-count");
    elements.adminRoomsList = document.getElementById("admin-rooms-list");
    elements.adminAddRoomForm = document.getElementById("admin-add-room-form");
    elements.roomFormTitle = document.getElementById("room-form-title");
    elements.editRoomId = document.getElementById("edit-room-id");
    elements.roomFormSite = document.getElementById("room-form-site");
    elements.roomFormName = document.getElementById("room-form-name");
    elements.roomFormLocation = document.getElementById("room-form-location");
    elements.roomFormCapacity = document.getElementById("room-form-capacity");
    elements.roomFormDescription = document.getElementById("room-form-description");
    elements.btnSaveRoomSubmit = document.getElementById("btn-save-room-submit");
    elements.btnCancelEditRoom = document.getElementById("btn-cancel-edit-room");

    // Supabase Cloud Settings Elements
    elements.supabaseUrlInput = document.getElementById("supabase-url-input");
    elements.supabaseAnonKeyInput = document.getElementById("supabase-anon-key-input");
    elements.btnTestSupabase = document.getElementById("btn-test-supabase");
    elements.supabaseTestStatus = document.getElementById("supabase-test-status");
    elements.googleSheetUrlInput = document.getElementById("google-sheet-url-input");
    elements.btnTestGoogleSheet = document.getElementById("btn-test-google-sheet");
    elements.googleSheetTestStatus = document.getElementById("google-sheet-test-status");
    elements.btnInstallPwa = document.getElementById("btn-install-pwa");
  }

  // Formatting Utilities
  function formatDateHeader(d) {
    const today = new Date();
    const isToday = d.toDateString() === today.toDateString();
    const dayName = d.toLocaleDateString("en-GB", { weekday: "long" });
    const dateFormatted = d.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
    return isToday ? `Today (${dayName}, ${dateFormatted})` : `${dayName}, ${dateFormatted}`;
  }

  function getISODate(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  // Apply Theme
  function applyTheme() {
    document.documentElement.setAttribute("data-theme", state.isDarkMode ? "dark" : "light");
    if (elements.themeToggleBtn) {
      elements.themeToggleBtn.innerHTML = state.isDarkMode ? "?? Light" : "?? Dark";
    }
  }

  // Render Header Status & Live Room Countdown for Active Room
  function updateRoomStatus() {
    const now = new Date();
    const todayStr = getISODate(now);
    const currentMins = now.getHours() * 60 + now.getMinutes();

    const room = window.SyncEngine.getRoomById(state.selectedRoomId);
    const roomLabel = room ? `${room.name} (${room.site})` : "Room";
    const roomCapacity = room ? room.capacity : 16;
    const roomLocation = room ? room.location : "Level 1 Pharmacy";

    const todayBookings = window.SyncEngine.bookings.filter(b => {
      return b.date === todayStr && (b.roomId || 'qeh-seminar-room') === state.selectedRoomId;
    });
    
    // Find active booking right now
    const activeBooking = todayBookings.find(b => {
      const start = window.SyncEngine.timeToMinutes(b.startTime);
      const end = window.SyncEngine.timeToMinutes(b.endTime);
      return currentMins >= start && currentMins < end;
    });

    if (activeBooking) {
      elements.roomStatusHero.className = activeBooking.isBlocked ? "room-status-hero is-blocked" : "room-status-hero is-occupied";
      elements.roomStatusText.innerHTML = activeBooking.isBlocked 
        ? `🛑 ${room ? room.name.toUpperCase() : 'ROOM'} TEMPORARILY CLOSED (${room ? room.site : ''})` 
        : `🔴 ${room ? room.name.toUpperCase() : 'ROOM'} OCCUPIED UNTIL ${activeBooking.endTime}`;
      elements.roomStatusSub.textContent = `${activeBooking.title} • Booked by ${activeBooking.bookerName} (${activeBooking.department}) • ${roomLocation}`;

      const endMins = window.SyncEngine.timeToMinutes(activeBooking.endTime);
      const remainingMins = endMins - currentMins;
      elements.roomStatusMeta.innerHTML = `⏱️ ${remainingMins} mins remaining in session • Next available slot at ${activeBooking.endTime}`;
      elements.quickBookBtn.innerHTML = `⚡ Book from ${activeBooking.endTime}`;
      state.activeSlotStart = activeBooking.endTime;
    } else {
      elements.roomStatusHero.className = "room-status-hero";
      elements.roomStatusText.innerHTML = `🟢 CURRENTLY AVAILABLE • ${roomLabel}`;

      // Find next upcoming booking today
      const upcoming = todayBookings
        .filter(b => window.SyncEngine.timeToMinutes(b.startTime) > currentMins)
        .sort((a, b) => a.startTime.localeCompare(b.startTime));

      if (upcoming.length > 0) {
        const next = upcoming[0];
        elements.roomStatusSub.textContent = `Free now at ${room ? room.site : 'site'} • Next session at ${next.startTime}: '${next.title}' (${next.bookerName})`;
        const freeUntilMins = window.SyncEngine.timeToMinutes(next.startTime) - currentMins;
        elements.roomStatusMeta.innerHTML = `✅ Free for the next ${freeUntilMins} mins (Max Capacity: ${roomCapacity} people)`;
      } else {
        elements.roomStatusSub.textContent = `Free for the rest of today • ${roomLocation} • Ready for use.`;
        elements.roomStatusMeta.innerHTML = `✅ Free to book immediately (Max Capacity: ${roomCapacity} people)`;
      }

      // Round current time to nearest current or next 30-min block for quick booking
      const curM = now.getMinutes();
      const roundedMins = curM < 30 ? 30 : 60;
      let qHour = now.getHours();
      let qMin = roundedMins;
      if (qMin === 60) {
        qHour += 1;
        qMin = 0;
      }
      const qTimeStr = `${String(qHour).padStart(2, "0")}:${String(qMin).padStart(2, "0")}`;
      elements.quickBookBtn.innerHTML = `⚡ Quick Book 30 Mins (${qTimeStr})`;
      state.activeSlotStart = qTimeStr;
    }
  }

  // Render Day Timeline Grid
  function renderDayTimeline() {
    const container = elements.dayTimelineContainer;
    container.innerHTML = "";

    const dateStr = getISODate(state.selectedDate);
    const openH = window.SyncEngine.config.openHour || 8;
    const closeH = window.SyncEngine.config.closeHour || 19;
    const totalSlots = (closeH - openH) * 2; // 2 half-hour slots per hour

    const dateBookings = window.SyncEngine.bookings.filter(b => {
      if (b.date !== dateStr) return false;
      if (state.filterDepartment !== "all" && b.department !== state.filterDepartment) return false;
      return true;
    });

    const now = new Date();
    const isToday = getISODate(now) === dateStr;
    const nowMinutes = now.getHours() * 60 + now.getMinutes();

    const timelineWrapper = document.createElement("div");
    timelineWrapper.className = "timeline-slots-container";

    // Current Time Marker (if viewing today)
    if (isToday && nowMinutes >= openH * 60 && nowMinutes <= closeH * 60) {
      const startOfDayMins = openH * 60;
      const totalDayMins = (closeH - openH) * 60;
      const pct = ((nowMinutes - startOfDayMins) / totalDayMins) * 100;
      const marker = document.createElement("div");
      marker.className = "current-time-marker";
      marker.style.top = `${pct}%`;
      marker.title = `Current Time: ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
      timelineWrapper.appendChild(marker);
    }

    for (let i = 0; i < totalSlots; i++) {
      const slotStartMinutes = openH * 60 + i * 30;
      const slotEndMinutes = slotStartMinutes + 30;
      const timeStartStr = window.SyncEngine.minutesToTime(slotStartMinutes);
      const timeEndStr = window.SyncEngine.minutesToTime(slotEndMinutes);

      // Check if slot falls inside an existing booking
      const booking = dateBookings.find(b => {
        const bStart = window.SyncEngine.timeToMinutes(b.startTime);
        const bEnd = window.SyncEngine.timeToMinutes(b.endTime);
        return slotStartMinutes >= bStart && slotStartMinutes < bEnd;
      });

      const slotRow = document.createElement("div");
      slotRow.className = "slot-row";

      // Time Column
      const timeCol = document.createElement("div");
      timeCol.className = "slot-time-col";
      timeCol.innerHTML = `
        <span class="slot-time-start">${timeStartStr}</span>
        <span class="slot-time-end">${timeEndStr}</span>
      `;
      slotRow.appendChild(timeCol);

      // Content Column
      const contentCol = document.createElement("div");
      contentCol.className = "slot-content-col";

      if (booking) {
        if (booking.isBlocked) {
          slotRow.classList.add("is-blocked");
          contentCol.innerHTML = `
            <div class="blocked-card">
              <span>??</span>
              <div>
                <strong>${booking.title}</strong>
                <div style="font-size:0.72rem; opacity:0.85;">Scheduled by Directorate Admin</div>
              </div>
            </div>
          `;
          contentCol.onclick = () => openDetailsModal(booking);
        } else {
          slotRow.classList.add("is-booked");
          contentCol.innerHTML = `
            <div class="booked-card">
              <div class="booked-info-main">
                <div class="booked-title">
                  <span>${booking.title}</span>
                  <span class="tag-department">${booking.department}</span>
                </div>
                <div class="booked-meta">
                  <span>?? <strong>${booking.bookerName}</strong> (${booking.bleep})</span>
                  <span>?? ${booking.attendees} attendees</span>
                  <span>?? ${booking.startTime} - ${booking.endTime} (${booking.durationMinutes}m)</span>
                </div>
              </div>
              <div class="booked-actions-quick">
                <button class="btn-slot-action" title="Add to Outlook Calendar" onclick="event.stopPropagation(); window.SyncEngine.generateICS(window.currentClickedBooking = ${JSON.stringify(booking).replace(/"/g, '&quot;')})">
                  ?? Outlook
                </button>
              </div>
            </div>
          `;
          contentCol.onclick = () => openDetailsModal(booking);
        }
      } else {
        const isPastSlot = isToday ? (slotEndMinutes <= nowMinutes) : (dateStr < getISODate(now));
        if (isPastSlot) {
          slotRow.classList.add("is-past");
          contentCol.innerHTML = 
            <div class="past-slot-card">
              <span>⌛ Concluded Time Slot ( - )</span>
            </div>
          ;
          contentCol.onclick = () => {
            alert("This time slot has already concluded. Please choose an upcoming slot to book.");
          };
        } else {
          slotRow.classList.add("is-available");
          contentCol.innerHTML = 
            <div class="available-slot-card">
              <span>✨ 30-Minute Slot Available • Click to Reserve</span>
              <span class="badge-book-now">+ Book Slot</span>
            </div>
          ;
          contentCol.onclick = () => openBookingModal(dateStr, timeStartStr);
        }
      }

      slotRow.appendChild(contentCol);
      timelineWrapper.appendChild(slotRow);
    }

    container.appendChild(timelineWrapper);
  }

  // Render Week View
  function renderWeekView() {
    const container = elements.weekViewContainer;
    container.innerHTML = "";

    const curr = new Date(state.selectedDate);
    // Find Monday of current week
    const day = curr.getDay();
    const diff = curr.getDate() - day + (day === 0 ? -6 : 1);
    const monday = new Date(curr.setDate(diff));

    const grid = document.createElement("div");
    grid.className = "week-grid-container";

    // Corner cell
    const cornerCell = document.createElement("div");
    cornerCell.className = "week-header-cell";
    cornerCell.textContent = "Time";
    grid.appendChild(cornerCell);

    // 7 Day Headers
    const weekDates = [];
    const todayStr = getISODate(new Date());

    for (let i = 0; i < 7; i++) {
      const d = new Date(monday);
      d.setDate(monday.getDate() + i);
      const dStr = getISODate(d);
      weekDates.push(dStr);

      const headerCell = document.createElement("div");
      headerCell.className = "week-header-cell";
      if (dStr === todayStr) headerCell.classList.add("is-today");
      headerCell.innerHTML = `
        <div>${d.toLocaleDateString("en-GB", { weekday: "short" })}</div>
        <div style="font-size:0.95rem; font-weight:800;">${d.getDate()}</div>
      `;
      grid.appendChild(headerCell);
    }

    // Slots Rows (Hourly for week view)
    const openH = window.SyncEngine.config.openHour || 8;
    const closeH = window.SyncEngine.config.closeHour || 19;

    for (let h = openH; h < closeH; h++) {
      const timeStr = `${String(h).padStart(2, "0")}:00`;
      
      const timeCell = document.createElement("div");
      timeCell.className = "week-time-cell";
      timeCell.textContent = timeStr;
      grid.appendChild(timeCell);

      for (let i = 0; i < 7; i++) {
        const dStr = weekDates[i];
        const slotCell = document.createElement("div");
        slotCell.className = "week-slot-cell";

        const bookingsAtTime = window.SyncEngine.bookings.filter(b => {
          if (b.date !== dStr) return false;
          const bStartH = parseInt(b.startTime.split(":")[0], 10);
          return bStartH === h;
        });

        if (bookingsAtTime.length > 0) {
          slotCell.classList.add("is-booked");
          bookingsAtTime.forEach(b => {
            const badge = document.createElement("div");
            badge.className = "week-slot-badge";
            badge.textContent = `${b.startTime} ${b.title}`;
            badge.title = `${b.title} (${b.bookerName})`;
            slotCell.appendChild(badge);
          });
          slotCell.onclick = () => openDetailsModal(bookingsAtTime[0]);
        } else {
          const isPastCell = dStr < todayStr || (dStr === todayStr && (h + 1) * 60 <= nowMinutes);
          if (isPastCell) {
            slotCell.classList.add("is-past");
            slotCell.title = "Past time slot (Concluded)";
            slotCell.onclick = () => {
              alert("Cannot book slots in the past. Please select an upcoming slot.");
            };
          } else {
            slotCell.onclick = () => openBookingModal(dStr, timeStr);
          }
        }

        grid.appendChild(slotCell);
      }
    }

    container.appendChild(grid);
  }

  // Render Agenda / List View
  function renderAgendaView() {
    const container = elements.agendaViewContainer;
    container.innerHTML = "";

    const sortedBookings = [...window.SyncEngine.bookings].sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      return a.startTime.localeCompare(b.startTime);
    });

    if (sortedBookings.length === 0) {
      container.innerHTML = `<div style="text-align:center; padding:3rem; color:var(--text-muted);">No bookings found. Click 'Quick Book' to schedule a slot!</div>`;
      return;
    }

    // Group by Date
    const grouped = {};
    sortedBookings.forEach(b => {
      if (!grouped[b.date]) grouped[b.date] = [];
      grouped[b.date].push(b);
    });

    Object.keys(grouped).forEach(dateStr => {
      const dParts = dateStr.split("-");
      const dObj = new Date(dParts[0], dParts[1] - 1, dParts[2]);
      const dateTitle = formatDateHeader(dObj);

      const groupCard = document.createElement("div");
      groupCard.className = "agenda-day-group";

      const header = document.createElement("div");
      header.className = "agenda-day-header";
      header.innerHTML = `
        <span>?? ${dateTitle}</span>
        <span style="font-size:0.75rem; font-weight:600; opacity:0.8;">${grouped[dateStr].length} session(s)</span>
      `;
      groupCard.appendChild(header);

      grouped[dateStr].forEach(b => {
        const item = document.createElement("div");
        item.className = "agenda-booking-item";
        item.innerHTML = `
          <div>
            <div style="font-weight:700; color:var(--nhs-dark-blue); font-size:0.92rem;">
              ${b.isBlocked ? '?? ' : ''}${b.title}
            </div>
            <div style="font-size:0.78rem; color:var(--text-muted); margin-top:0.2rem;">
              ?? ${b.bookerName} (${b.department}) � ?? ${b.bleep} � ?? ${b.attendees} attendees
            </div>
          </div>
          <div style="text-align:right;">
            <div style="font-weight:700; font-size:0.88rem; color:var(--nhs-blue);">
              ${b.startTime} - ${b.endTime}
            </div>
            <div style="font-size:0.72rem; color:var(--text-light);">
              ${b.durationMinutes} mins
            </div>
          </div>
        `;
        item.onclick = () => openDetailsModal(b);
        groupCard.appendChild(item);
      });

      container.appendChild(groupCard);
    });
  }

  function renderCurrentView() {
    elements.dateDisplayTitle.textContent = formatDateHeader(state.selectedDate);
    if (elements.calendarDateInput) {
      elements.calendarDateInput.value = getISODate(state.selectedDate);
    }
    if (elements.hiddenDatePicker) {
      elements.hiddenDatePicker.value = getISODate(state.selectedDate);
    }

    // Switch view containers
    elements.dayTimelineContainer.style.display = state.viewMode === "day" ? "block" : "none";
    elements.weekViewContainer.style.display = state.viewMode === "week" ? "block" : "none";
    elements.agendaViewContainer.style.display = state.viewMode === "agenda" ? "block" : "none";

    elements.viewDayBtn.classList.toggle("active", state.viewMode === "day");
    elements.viewWeekBtn.classList.toggle("active", state.viewMode === "week");
    elements.viewAgendaBtn.classList.toggle("active", state.viewMode === "agenda");

    if (state.viewMode === "day") renderDayTimeline();
    else if (state.viewMode === "week") renderWeekView();
    else if (state.viewMode === "agenda") renderAgendaView();

    updateRoomStatus();
    renderAdminBar();
  }

  function renderAdminBar() {
    const isAdmin = window.SyncEngine.isAdmin;
    elements.adminActiveBar.style.display = isAdmin ? "flex" : "none";
    elements.adminToggleBtn.classList.toggle("admin-active", isAdmin);
    elements.adminToggleBtn.innerHTML = isAdmin ? "??? Superuser Active" : "?? Admin";
  }

  // Booking Modal Logic
  // ==========================================================================
  // Hospital Sites & Rooms Switcher & Manager
  // ==========================================================================

  function renderSiteRoomNavigator() {
    if (!elements.siteTabsContainer || !window.SyncEngine) return;
    const sites = window.SyncEngine.getSites();
    const rooms = window.SyncEngine.getRooms();

    // 1. Update site tabs counts & active status
    const siteTabButtons = elements.siteTabsContainer.querySelectorAll(".site-tab-pill");
    siteTabButtons.forEach(btn => {
      const site = btn.dataset.site;
      const count = rooms.filter(r => r.site === site).length;
      btn.classList.toggle("active", site === state.selectedSite);
      const countEl = btn.querySelector(".site-tab-count");
      if (countEl) {
        countEl.textContent = count === 1 ? "1 room" : `${count} rooms`;
      }
    });

    // 2. Populate Room Select Dropdown for the active site
    if (elements.roomSelectDropdown) {
      const siteRooms = window.SyncEngine.getRoomsBySite(state.selectedSite);
      if (!siteRooms.some(r => r.id === state.selectedRoomId)) {
        state.selectedRoomId = siteRooms.length > 0 ? siteRooms[0].id : "qeh-seminar-room";
      }

      elements.roomSelectDropdown.innerHTML = "";
      siteRooms.forEach(r => {
        const opt = document.createElement("option");
        opt.value = r.id;
        opt.textContent = `${r.name} (${r.location || 'Pharmacy'} • Cap: ${r.capacity})`;
        if (r.id === state.selectedRoomId) opt.selected = true;
        elements.roomSelectDropdown.appendChild(opt);
      });
    }

    // 3. Update Room Metadata Banner Pill
    const currentRoom = window.SyncEngine.getRoomById(state.selectedRoomId);
    if (currentRoom) {
      if (elements.roomMetaLocation) elements.roomMetaLocation.textContent = currentRoom.location || "Pharmacy Department";
      if (elements.roomMetaCapacity) elements.roomMetaCapacity.textContent = `👥 Max ${currentRoom.capacity} people`;
      if (elements.roomMetaFacilities) {
        const facilityMap = {
          wall_display: "📺 AV Display",
          owl_camera: "🦉 Teams Owl",
          whiteboard: "📋 Whiteboard",
          flipchart: "📄 Flipchart",
          polycom_phone: "📞 Conf Phone",
          pc_workstations: "💻 PCs"
        };
        const eq = Array.isArray(currentRoom.equipment) && currentRoom.equipment.length > 0
          ? currentRoom.equipment.map(k => facilityMap[k] || k.replace(/_/g, " ")).join(", ")
          : "Standard Presentation Facilities";
        elements.roomMetaFacilities.textContent = eq;
      }
    }

    // 4. Quick add room button visibility for admin
    if (elements.btnAdminAddRoomQuick) {
      elements.btnAdminAddRoomQuick.style.display = window.SyncEngine.isAdmin ? "inline-block" : "none";
      elements.btnAdminAddRoomQuick.textContent = `➕ Add Room to ${state.selectedSite}`;
    }
  }

  function selectSite(siteName, preferredRoomId = null) {
    state.selectedSite = siteName;
    const siteRooms = window.SyncEngine.getRoomsBySite(siteName);
    if (preferredRoomId && siteRooms.some(r => r.id === preferredRoomId)) {
      state.selectedRoomId = preferredRoomId;
    } else if (siteRooms.length > 0) {
      state.selectedRoomId = siteRooms[0].id;
    }
    renderSiteRoomNavigator();
    updateRoomStatus();
    renderCurrentView();
  }

  function selectRoom(roomId) {
    const room = window.SyncEngine.getRoomById(roomId);
    if (room) {
      state.selectedRoomId = room.id;
      if (room.site && room.site !== state.selectedSite) {
        state.selectedSite = room.site;
      }
    }
    renderSiteRoomNavigator();
    updateRoomStatus();
    renderCurrentView();
  }

  function populateRoomSelect(selectEl, siteName, selectedId = null) {
    if (!selectEl || !window.SyncEngine) return;
    const rooms = window.SyncEngine.getRoomsBySite(siteName);
    selectEl.innerHTML = "";
    rooms.forEach(r => {
      const opt = document.createElement("option");
      opt.value = r.id;
      opt.textContent = `${r.name} (${r.location || 'Department'} • Cap: ${r.capacity})`;
      if (selectedId && r.id === selectedId) opt.selected = true;
      selectEl.appendChild(opt);
    });
    if (!selectedId && rooms.length > 0) {
      selectEl.value = rooms[0].id;
    }
  }

  function updateBookingRoomCapacityHint() {
    if (!elements.bookingRoom) return;
    const rId = elements.bookingRoom.value || state.selectedRoomId;
    const room = window.SyncEngine.getRoomById(rId);
    if (room) {
      if (elements.bookingRoomCapHint) {
        elements.bookingRoomCapHint.textContent = `Cap: ${room.capacity} people`;
      }
      if (elements.bookingAttendeesInput) {
        elements.bookingAttendeesInput.max = room.capacity;
        if (parseInt(elements.bookingAttendeesInput.value, 10) > room.capacity) {
          elements.bookingAttendeesInput.value = room.capacity;
        }
      }
      if (elements.bookingModalTitle && !state.editingBookingId) {
        elements.bookingModalTitle.textContent = `Book ${room.name} (${room.site})`;
      }
    }
  }

  // Admin Superuser Hospital Rooms Directory & Add/Edit Manager
  function openManageRoomsModal() {
    if (!elements.manageRoomsModal) return;
    renderAdminRoomsDirectory();
    cancelEditRoom();
    if (elements.roomFormSite) elements.roomFormSite.value = state.selectedSite;
    elements.manageRoomsModal.classList.add("open");
  }

  function renderAdminRoomsDirectory() {
    if (!elements.adminRoomsList) return;
    const rooms = window.SyncEngine.getRooms();
    const sites = window.SyncEngine.getSites();

    if (elements.adminTotalRoomsCount) {
      elements.adminTotalRoomsCount.textContent = `${rooms.length} room(s) across 4 Trust campuses`;
    }

    elements.adminRoomsList.innerHTML = "";

    sites.forEach(siteName => {
      const siteRooms = rooms.filter(r => r.site === siteName);
      const siteGroup = document.createElement("div");
      siteGroup.className = "admin-site-room-group";

      const header = document.createElement("div");
      header.className = "admin-site-group-header";
      header.innerHTML = `
        <div style="font-weight:700; color:var(--nhs-dark-blue); font-size:0.9rem;">
          🏥 ${siteName} (${siteRooms.length} room${siteRooms.length === 1 ? '' : 's'})
        </div>
        <button class="btn-secondary btn-sm" type="button" style="font-size:0.75rem; padding:0.2rem 0.5rem;" onclick="window.prepareAddRoomForSite('${siteName}')">
          ➕ Add Room to ${siteName}
        </button>
      `;
      siteGroup.appendChild(header);

      if (siteRooms.length === 0) {
        const empty = document.createElement("div");
        empty.style.cssText = "padding:0.75rem 1rem; color:var(--text-muted); font-size:0.8rem; font-style:italic;";
        empty.textContent = "No rooms configured for this site yet. Click 'Add Room' to add one.";
        siteGroup.appendChild(empty);
      } else {
        const grid = document.createElement("div");
        grid.className = "admin-rooms-grid";

        siteRooms.forEach(r => {
          const card = document.createElement("div");
          card.className = "admin-room-item-card";
          if (r.isDefault) card.classList.add("is-default-card");

          const equipPills = (Array.isArray(r.equipment) ? r.equipment : [])
            .map(e => `<span class="tag-department">${e.replace(/_/g, ' ')}</span>`)
            .join(" ");

          card.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:flex-start;">
              <div>
                <strong style="color:var(--nhs-dark-blue); font-size:0.92rem;">${r.name}</strong>
                ${r.isDefault ? '<span class="default-room-badge">⭐ Core Seminar Room</span>' : ''}
                <div style="font-size:0.78rem; color:var(--text-muted); margin-top:2px;">
                  📍 ${r.location || 'Pharmacy Department'} • 👥 Max Capacity: <strong>${r.capacity}</strong>
                </div>
              </div>
              <div style="display:flex; gap:0.35rem;">
                <button class="btn-secondary btn-sm" type="button" title="Edit room details" onclick="window.startEditRoom('${r.id}')" style="font-size:0.75rem; padding:0.25rem 0.5rem;">
                  ✏️ Edit
                </button>
                ${!r.isBuiltIn && r.id !== 'qeh-seminar-room' ? `
                  <button class="btn-secondary btn-sm" type="button" title="Delete room" onclick="window.deleteRoomAction('${r.id}')" style="font-size:0.75rem; padding:0.25rem 0.5rem; color:#DA291C;">
                    🗑️ Delete
                  </button>
                ` : ''}
              </div>
            </div>
            ${r.description ? `<div style="font-size:0.78rem; color:var(--text-main); margin-top:0.4rem;">${r.description}</div>` : ''}
            <div style="margin-top:0.4rem; display:flex; gap:0.3rem; flex-wrap:wrap;">
              ${equipPills || '<span style="font-size:0.72rem; color:var(--text-light);">Standard presentation facilities</span>'}
            </div>
          `;
          grid.appendChild(card);
        });
        siteGroup.appendChild(grid);
      }

      elements.adminRoomsList.appendChild(siteGroup);
    });
  }

  window.prepareAddRoomForSite = function (siteName) {
    cancelEditRoom();
    if (elements.roomFormSite) elements.roomFormSite.value = siteName;
    if (elements.roomFormName) elements.roomFormName.focus();
  };

  window.startEditRoom = function (roomId) {
    const room = window.SyncEngine.getRoomById(roomId);
    if (!room) return;

    if (elements.editRoomId) elements.editRoomId.value = room.id;
    if (elements.roomFormTitle) elements.roomFormTitle.textContent = `✏️ Edit Room: ${room.name} (${room.site})`;
    if (elements.roomFormSite) {
      elements.roomFormSite.value = room.site;
      elements.roomFormSite.disabled = (room.id === "qeh-seminar-room");
    }
    if (elements.roomFormName) elements.roomFormName.value = room.name;
    if (elements.roomFormLocation) elements.roomFormLocation.value = room.location || "";
    if (elements.roomFormCapacity) elements.roomFormCapacity.value = room.capacity || 12;
    if (elements.roomFormDescription) elements.roomFormDescription.value = room.description || "";

    const equip = Array.isArray(room.equipment) ? room.equipment : [];
    document.querySelectorAll('input[name="room-form-equipment"]').forEach(cb => {
      cb.checked = equip.includes(cb.value);
    });

    if (elements.btnCancelEditRoom) elements.btnCancelEditRoom.style.display = "inline-block";
    if (elements.btnSaveRoomSubmit) elements.btnSaveRoomSubmit.textContent = "💾 Save Changes";
    if (elements.roomFormName) elements.roomFormName.focus();
  };

  function cancelEditRoom() {
    if (elements.editRoomId) elements.editRoomId.value = "";
    if (elements.roomFormTitle) elements.roomFormTitle.textContent = "➕ Add Additional Room";
    if (elements.roomFormSite) {
      elements.roomFormSite.disabled = false;
      elements.roomFormSite.value = state.selectedSite;
    }
    if (elements.adminAddRoomForm) elements.adminAddRoomForm.reset();
    if (elements.btnCancelEditRoom) elements.btnCancelEditRoom.style.display = "none";
    if (elements.btnSaveRoomSubmit) elements.btnSaveRoomSubmit.textContent = "➕ Add Room to Site";
  }

  window.deleteRoomAction = function (roomId) {
    const room = window.SyncEngine.getRoomById(roomId);
    if (!room) return;
    if (room.id === "qeh-seminar-room") {
      alert("The core Pharmacy Seminar Room under QEH Heritage cannot be deleted.");
      return;
    }
    const ok = confirm(`Are you sure you want to delete '${room.name}' at ${room.site}?\n\nAny historical bookings will remain preserved.`);
    if (!ok) return;

    try {
      window.SyncEngine.deleteRoom(roomId);
      if (state.selectedRoomId === roomId) {
        state.selectedRoomId = "qeh-seminar-room";
        state.selectedSite = "QEH Heritage";
      }
      renderAdminRoomsDirectory();
      renderSiteRoomNavigator();
      updateRoomStatus();
      renderCurrentView();
      alert(`✅ Room '${room.name}' successfully deleted.`);
    } catch (err) {
      alert("⚠️ Error: " + err.message);
    }
  };

  function handleAdminAddRoomSubmit(e) {
    e.preventDefault();
    const isEdit = Boolean(elements.editRoomId && elements.editRoomId.value);
    const site = elements.roomFormSite.value;
    const name = elements.roomFormName.value.trim();
    const location = elements.roomFormLocation.value.trim();
    const capacity = parseInt(elements.roomFormCapacity.value, 10) || 12;
    const description = elements.roomFormDescription ? elements.roomFormDescription.value.trim() : "";

    const selectedEquipment = [];
    document.querySelectorAll('input[name="room-form-equipment"]:checked').forEach(cb => {
      selectedEquipment.push(cb.value);
    });

    if (!name || !site) {
      alert("Room name and site are required.");
      return;
    }

    try {
      if (isEdit) {
        const id = elements.editRoomId.value;
        window.SyncEngine.updateRoom(id, {
          name: name,
          fullName: `${site} ${name}`,
          site: site,
          location: location,
          capacity: capacity,
          equipment: selectedEquipment,
          description: description
        });
        alert(`✅ Room '${name}' updated successfully!`);
      } else {
        const created = window.SyncEngine.addRoom({
          name: name,
          fullName: `${site} ${name}`,
          site: site,
          location: location,
          capacity: capacity,
          equipment: selectedEquipment,
          description: description
        });
        alert(`✅ Room '${created.name}' successfully added to ${created.site}!`);
        state.selectedSite = site;
        state.selectedRoomId = created.id;
      }

      cancelEditRoom();
      renderAdminRoomsDirectory();
      renderSiteRoomNavigator();
      updateRoomStatus();
      renderCurrentView();
    } catch (err) {
      alert("⚠️ Error: " + err.message);
    }
  }

  function updateDoorSignPreview() {
    const site = elements.doorSiteSelect ? elements.doorSiteSelect.value : state.selectedSite;
    const roomId = elements.doorRoomSelect ? elements.doorRoomSelect.value : state.selectedRoomId;
    const room = window.SyncEngine.getRoomById(roomId) || window.SyncEngine.getRoomsBySite(site)[0];
    if (!room) return;

    if (elements.doorSignTrustTitle) {
      elements.doorSignTrustTitle.textContent = `University Hospitals Birmingham NHS Foundation Trust • ${room.site}`;
    }
    if (elements.doorSignRoomTitle) {
      elements.doorSignRoomTitle.textContent = room.fullName || room.name;
    }
    if (elements.doorSignRoomSubtitle) {
      elements.doorSignRoomSubtitle.textContent = `${room.location || 'Pharmacy Department'} • ${room.site}`;
    }
    if (elements.doorSignEtiquetteList) {
      const facilityMap = {
        wall_display: "Wall Display",
        owl_camera: "Teams Owl 360 Camera",
        whiteboard: "Whiteboard",
        flipchart: "Flipchart",
        polycom_phone: "Conference Phone",
        pc_workstations: "PC Workstations"
      };
      const eqStr = (Array.isArray(room.equipment) && room.equipment.length > 0)
        ? room.equipment.map(k => facilityMap[k] || k.replace(/_/g, " ")).join(", ")
        : "Standard presentation facilities";

      elements.doorSignEtiquetteList.innerHTML = `
        <li>Maximum room capacity: <strong>${room.capacity} people</strong></li>
        <li>Equipped with ${eqStr}</li>
        <li>Please keep room tidy and wipe whiteboard after sessions</li>
        <li>Please cancel your slot in the app if your meeting is cancelled</li>
      `;
    }

    let baseUrl = window.location.href.split("?")[0];
    if (baseUrl.startsWith("file://")) baseUrl = "http://localhost:8089";
    const qrUrl = `${baseUrl}?site=${encodeURIComponent(room.site)}&room=${encodeURIComponent(room.id)}`;

    if (elements.doorQrUrlInput) {
      elements.doorQrUrlInput.value = qrUrl;
    }
    renderDoorQrCode(qrUrl);
  }
  function openBookingModal(dateStr, startTime) {
    state.editingBookingId = null;
    state.recurrence = null;
    elements.bookingForm.reset();
    if (elements.bookingRoom) {
      populateRoomSelect(elements.bookingRoom, state.selectedSite || "QEH Heritage", state.selectedRoomId || "qeh-seminar-room");
      updateBookingRoomCapacityHint();
    }

    const todayStr = getISODate(new Date());
    if (elements.bookingDateInput) {
      elements.bookingDateInput.min = todayStr;
    }

    let targetDate = dateStr || getISODate(state.selectedDate);
    if (targetDate < todayStr) {
      targetDate = todayStr;
    }
    elements.bookingDateInput.value = targetDate;

    const now = new Date();
    const nowMins = now.getHours() * 60 + now.getMinutes();
    let chosenStartTime = startTime || "09:00";
    if (targetDate === todayStr && window.SyncEngine.timeToMinutes(chosenStartTime) < nowMins) {
      chosenStartTime = state.activeSlotStart || window.SyncEngine.minutesToTime(Math.ceil(nowMins / 30) * 30);
    }
    elements.bookingStartTimeInput.value = chosenStartTime;
    setBookingDuration(30);

    if (elements.recurrenceSummaryText) {
      elements.recurrenceSummaryText.textContent = "One-off booking (Does not repeat)";
    }

    elements.bookerPinInput.value = "";
    if (elements.staffMatchBadge) elements.staffMatchBadge.style.display = "none";
    updateStaffSuggestions();
    elements.bookingModal.classList.add("open");

    updateLiveAvailabilityPreview();
  }

  function setBookingDuration(durationMinutes) {
    state.selectedDuration = durationMinutes;
    elements.bookingDurationInput.value = durationMinutes;
    elements.durationButtons.forEach(btn => {
      btn.classList.toggle("active", parseInt(btn.dataset.duration, 10) === durationMinutes);
    });
    updateLiveAvailabilityPreview();
  }

  // Live Room Availability Slot Checker for Selected Date
  function updateLiveAvailabilityPreview() {
    if (!elements.checkerSlotsGrid) return;
    const dateStr = elements.bookingDateInput.value;
    if (!dateStr) return;

    const parts = dateStr.split("-").map(Number);
    const d = new Date(parts[0], parts[1] - 1, parts[2]);
    const dayName = d.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "short", year: "numeric" });
    if (elements.checkerHeaderTitle) {
      elements.checkerHeaderTitle.textContent = `?? Room Availability for ${dayName}`;
    }

    // Get existing bookings for this date
    const dayBookings = (window.SyncEngine.bookings || []).filter(b => b.date === dateStr && b.id !== state.editingBookingId);

    // Build half-hour slots: 08:00 to 18:30
    const slots = [];
    for (let h = 8; h <= 18; h++) {
      slots.push(`${String(h).padStart(2, "0")}:00`);
      slots.push(`${String(h).padStart(2, "0")}:30`);
    }

    let freeCount = 0;
    elements.checkerSlotsGrid.innerHTML = "";

    const selectedStart = elements.bookingStartTimeInput.value;

    slots.forEach(slotStart => {
      const startMins = window.SyncEngine.timeToMinutes(slotStart);
      const endMins = startMins + 30;

      // Check if slot overlaps any booking
      const booked = dayBookings.find(b => {
        const bStart = window.SyncEngine.timeToMinutes(b.startTime);
        const bEnd = window.SyncEngine.timeToMinutes(b.endTime);
        return startMins < bEnd && endMins > bStart;
      });

      if (booked) {
        const pill = document.createElement("span");
        pill.className = "slot-pill booked";
        pill.title = `Booked (${booked.startTime} - ${booked.endTime}): ${booked.title} [${booked.bookerName}]`;
        pill.textContent = `${slotStart} ? Booked`;
        elements.checkerSlotsGrid.appendChild(pill);
      } else {
        freeCount++;
        const pill = document.createElement("button");
        pill.type = "button";
        pill.className = "slot-pill free";
        if (slotStart === selectedStart) {
          pill.classList.add("selected");
        }
        pill.textContent = `${slotStart} ? Free`;
        pill.title = `Click to choose ${slotStart} as start time`;
        pill.onclick = () => {
          elements.bookingStartTimeInput.value = slotStart;
          updateLiveAvailabilityPreview();
        };
        elements.checkerSlotsGrid.appendChild(pill);
      }
    });

    if (elements.checkerSlotCount) {
      elements.checkerSlotCount.textContent = `${freeCount} of ${slots.length} slots available`;
    }

    // Validate current selected Start Time and Duration
    const chosenStart = elements.bookingStartTimeInput.value;
    const chosenDuration = parseInt(elements.bookingDurationInput.value, 10) || 30;
    const chosenEnd = window.SyncEngine.minutesToTime(window.SyncEngine.timeToMinutes(chosenStart) + chosenDuration);
    const targetRoomId = (elements.bookingRoom && elements.bookingRoom.value) || state.selectedRoomId || "qeh-seminar-room";

    const conflict = window.SyncEngine.checkConflict(dateStr, chosenStart, chosenEnd, state.editingBookingId, targetRoomId, window.SyncEngine.isAdmin);
    if (elements.checkerConflictMsg) {
      if (conflict.hasConflict) {
        elements.checkerConflictMsg.style.display = "block";
        elements.checkerConflictMsg.className = "checker-conflict-message has-conflict";
        elements.checkerConflictMsg.innerHTML = conflict.isPast ? `⌛ <strong>Past Time:</strong> ${conflict.message}` : `⚠️ <strong>Time Conflict:</strong> ${conflict.message}`;
      } else {
        elements.checkerConflictMsg.style.display = "block";
        elements.checkerConflictMsg.className = "checker-conflict-message is-free";
        elements.checkerConflictMsg.innerHTML = `? <strong>Available:</strong> ${chosenStart} � ${chosenEnd} (${chosenDuration} mins) is currently free to book!`;
      }
    }
  }

  // Recurrence Modal Controller (Outlook Style)
  function openRecurrenceModal() {
    const startTime = elements.bookingStartTimeInput.value || "09:00";
    const duration = parseInt(elements.bookingDurationInput.value, 10) || 30;
    const startMins = window.SyncEngine.timeToMinutes(startTime);
    const endTime = window.SyncEngine.minutesToTime(startMins + duration);

    // Set time dropdowns
    elements.recTimeStart.value = startTime;
    elements.recTimeDuration.value = String(duration);
    elements.recTimeEnd.value = endTime;

    const startDate = elements.bookingDateInput.value || getISODate(state.selectedDate);
    elements.recRangeStart.value = startDate;

    // Set default end date to 12 weeks ahead
    const d = new Date(startDate);
    d.setDate(d.getDate() + 84);
    elements.recRangeEndDate.value = getISODate(d);

    // Pre-check weekday matching start date
    const startDay = new Date(startDate).getDay();
    document.querySelectorAll('input[name="rec-weekday"]').forEach(cb => {
      cb.checked = (parseInt(cb.value, 10) === startDay);
    });

    // If existing recurrence on state, populate from it
    if (state.recurrence && state.recurrence.pattern) {
      const p = state.recurrence.pattern;
      const r = state.recurrence.range;

      const radio = document.querySelector(`input[name="rec-pattern-type"][value="${p.type}"]`);
      if (radio) radio.checked = true;
      switchRecurrencePanel(p.type);

      if (p.type === "weekly") {
        if (elements.recWeeklyInterval) elements.recWeeklyInterval.value = p.interval || 1;
        if (Array.isArray(p.daysOfWeek)) {
          document.querySelectorAll('input[name="rec-weekday"]').forEach(cb => {
            cb.checked = p.daysOfWeek.includes(parseInt(cb.value, 10));
          });
        }
      } else if (p.type === "daily") {
        if (elements.recDailyInterval) elements.recDailyInterval.value = p.interval || 1;
        const dailyRadio = document.querySelector(`input[name="rec-daily-mode"][value="${p.weekdaysOnly ? 'weekdays' : 'interval'}"]`);
        if (dailyRadio) dailyRadio.checked = true;
      } else if (p.type === "monthly") {
        if (elements.recMonthlyDay) elements.recMonthlyDay.value = p.dayOfMonth || 1;
        if (elements.recMonthlyInterval) elements.recMonthlyInterval.value = p.interval || 1;
      }

      if (r) {
        const rangeRadio = document.querySelector(`input[name="rec-range-mode"][value="${r.type}"]`);
        if (rangeRadio) rangeRadio.checked = true;
        if (r.endDate && elements.recRangeEndDate) elements.recRangeEndDate.value = r.endDate;
        if (r.occurrences && elements.recRangeOccurrences) elements.recRangeOccurrences.value = r.occurrences;
      }
    } else {
      switchRecurrencePanel("weekly");
    }

    elements.recurrenceModal.classList.add("open");
  }

  function switchRecurrencePanel(patternType) {
    if (elements.recPanelWeekly) elements.recPanelWeekly.style.display = patternType === "weekly" ? "block" : "none";
    if (elements.recPanelDaily) elements.recPanelDaily.style.display = patternType === "daily" ? "block" : "none";
    if (elements.recPanelMonthly) elements.recPanelMonthly.style.display = patternType === "monthly" ? "block" : "none";
    if (elements.recPanelYearly) elements.recPanelYearly.style.display = patternType === "yearly" ? "block" : "none";
  }

  function handleRecurrenceOk() {
    const patternTypeRadio = document.querySelector('input[name="rec-pattern-type"]:checked');
    const rangeModeRadio = document.querySelector('input[name="rec-range-mode"]:checked');
    const patternType = patternTypeRadio ? patternTypeRadio.value : "weekly";
    const rangeMode = rangeModeRadio ? rangeModeRadio.value : "endBy";
    const startDate = elements.recRangeStart.value;

    const pattern = { type: patternType, interval: 1 };
    let summaryText = "";

    const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

    if (patternType === "weekly") {
      pattern.interval = parseInt(elements.recWeeklyInterval.value, 10) || 1;
      const checkedDays = Array.from(document.querySelectorAll('input[name="rec-weekday"]:checked'))
        .map(cb => parseInt(cb.value, 10));
      if (checkedDays.length === 0) {
        alert("Please select at least one day of the week for weekly recurrence.");
        return;
      }
      pattern.daysOfWeek = checkedDays;
      const daysStr = checkedDays.map(d => dayNames[d] + "s").join(", ");
      summaryText = pattern.interval === 1 ? `Weekly on ${daysStr}` : `Every ${pattern.interval} weeks on ${daysStr}`;
    } else if (patternType === "daily") {
      const dailyModeRadio = document.querySelector('input[name="rec-daily-mode"]:checked');
      const dailyMode = dailyModeRadio ? dailyModeRadio.value : "interval";
      if (dailyMode === "weekdays") {
        pattern.weekdaysOnly = true;
        summaryText = "Every weekday (Mon-Fri)";
      } else {
        pattern.interval = parseInt(elements.recDailyInterval.value, 10) || 1;
        summaryText = pattern.interval === 1 ? "Daily" : `Every ${pattern.interval} days`;
      }
    } else if (patternType === "monthly") {
      pattern.dayOfMonth = parseInt(elements.recMonthlyDay.value, 10) || 1;
      pattern.interval = parseInt(elements.recMonthlyInterval.value, 10) || 1;
      summaryText = `Monthly on day ${pattern.dayOfMonth}`;
    } else if (patternType === "yearly") {
      summaryText = "Yearly";
    }

    const range = {
      type: rangeMode,
      startDate: startDate,
      endDate: elements.recRangeEndDate.value,
      occurrences: parseInt(elements.recRangeOccurrences.value, 10) || 10
    };

    if (rangeMode === "endBy") {
      summaryText += ` until ${range.endDate}`;
    } else if (rangeMode === "endAfter") {
      summaryText += ` (${range.occurrences} times)`;
    } else {
      summaryText += " (for 6 months)";
    }

    // Sync back time to booking form
    elements.bookingStartTimeInput.value = elements.recTimeStart.value;
    setBookingDuration(parseInt(elements.recTimeDuration.value, 10));

    // Calculate dates
    const dates = window.SyncEngine.generateRecurrenceDates(startDate, pattern, range);

    state.recurrence = {
      isRecurring: true,
      pattern: pattern,
      range: range,
      summary: summaryText,
      totalCount: dates.length
    };

    if (elements.recurrenceSummaryText) {
      elements.recurrenceSummaryText.innerHTML = `?? <strong>${summaryText}</strong> (${dates.length} sessions planned)`;
    }
    elements.recurrenceModal.classList.remove("open");
    updateLiveAvailabilityPreview();
  }

  function handleRecurrenceRemove() {
    state.recurrence = null;
    if (elements.recurrenceSummaryText) {
      elements.recurrenceSummaryText.textContent = "One-off booking (Does not repeat)";
    }
    elements.recurrenceModal.classList.remove("open");
    updateLiveAvailabilityPreview();
  }

  function handleBookingFormSubmit(e) {
    e.preventDefault();

    const date = elements.bookingDateInput.value;
    const startTime = elements.bookingStartTimeInput.value;
    const duration = parseInt(elements.bookingDurationInput.value, 10);
    const startMins = window.SyncEngine.timeToMinutes(startTime);
    const endTime = window.SyncEngine.minutesToTime(startMins + duration);

    const name = elements.bookerNameInput.value.trim();
    const email = elements.bookerEmailInput.value.trim();
    const bleep = elements.bookerBleepInput.value.trim();
    const dept = elements.bookerDeptInput.value;
    const title = elements.bookingTitleInput.value.trim();
    const attendees = parseInt(elements.bookingAttendeesInput.value, 10) || 1;
    const pin = elements.bookerPinInput.value.trim();

    // Validations
    if (!name || !email || !bleep || !title) {
      alert("Please complete all required fields (Name, NHS Email, Bleep/Ext, Title).");
      return;
    }

    const weakPins = ["0000", "1234", "1111", "2222", "3333", "4444", "5555", "6666", "7777", "8888", "9999", "1212", "4321"];
    if (weakPins.includes(pin)) {
      alert("Security Warning: '" + pin + "' is too easily guessed.\n\nPlease choose a private 4-digit PIN to prevent other users from altering your booking.");
      if (elements.bookerPinInput) elements.bookerPinInput.focus();
      return;
    }

    if (!pin || pin.length !== 4 || !/^\d{4}$/.test(pin)) {
      alert("Please enter a 4-digit numeric PIN. You will need this PIN if you wish to edit or cancel your booking later.");
      return;
    }

    // NHS Email check recommendation
    if (!email.toLowerCase().includes("nhs")) {
      const proceed = confirm("Notice: You entered a non-NHS email address. Are you sure you wish to proceed with this email?");
      if (!proceed) return;
    }

    // Prevent booking past slots
    if (!window.SyncEngine.isAdmin && window.SyncEngine.isSlotInPast(date, startTime)) {
      alert("⚠️ Cannot Book Past Slot:\n\nThe selected date and time has already passed. Please select an upcoming date and time.");
      return;
    }

    // Equipment
    const selectedEquipment = [];
    document.querySelectorAll('input[name="equipment"]:checked').forEach(cb => {
      selectedEquipment.push(cb.value);
    });

    const targetRoomId = (elements.bookingRoom && elements.bookingRoom.value) || state.selectedRoomId || "qeh-seminar-room";

    const bookingPayload = {
      roomId: targetRoomId,
      date: date,
      startTime: startTime,
      endTime: endTime,
      durationMinutes: duration,
      bookerName: name,
      email: email,
      bleep: bleep,
      department: dept,
      title: title,
      attendees: attendees,
      equipment: selectedEquipment,
      pin: pin
    };

    try {
      if (state.editingBookingId) {
        window.SyncEngine.updateBooking(
          state.editingBookingId,
          bookingPayload,
          pin,
          window.SyncEngine.isAdmin
        );
        alert("Booking updated successfully!");
        closeAllModals();
        renderCurrentView();
        return;
      }

      // Handle Recurring Booking Series
      if (state.recurrence && state.recurrence.isRecurring) {
        const dates = window.SyncEngine.generateRecurrenceDates(date, state.recurrence.pattern, state.recurrence.range);
        if (dates.length === 0) {
          alert("No dates could be calculated for this recurrence pattern. Please check your range.");
          return;
        }

        const conflictCheck = window.SyncEngine.checkRecurrenceConflicts(dates, startTime, endTime, targetRoomId);
        let datesToBook = dates;

        if (conflictCheck.hasConflicts) {
          if (conflictCheck.freeDates.length === 0) {
            alert(`?? Cannot create recurring series:\nAll ${dates.length} requested sessions conflict with existing bookings.`);
            return;
          }

          const conflictLines = conflictCheck.conflicts.map(c => `� ${c.date}: ${c.message}`).slice(0, 5).join("\n");
          const moreText = conflictCheck.conflicts.length > 5 ? `\n...and ${conflictCheck.conflicts.length - 5} more` : "";
          const proceed = confirm(
            `?? Scheduling Conflict on ${conflictCheck.conflicts.length} of ${dates.length} sessions:\n\n` +
            `${conflictLines}${moreText}\n\n` +
            `Would you like to book the ${conflictCheck.freeDates.length} AVAILABLE sessions (skipping conflicting dates)?`
          );

          if (!proceed) return;
          datesToBook = conflictCheck.freeDates;
        }

        const result = window.SyncEngine.addRecurringBookings(bookingPayload, datesToBook, state.recurrence);
        closeAllModals();

        const downloadIcs = confirm(
          `? Recurring Booking Series Confirmed!\n\n` +
          `Series: ${state.recurrence.summary}\n` +
          `Total Sessions Booked: ${result.bookings.length}\n` +
          `Time: ${startTime} - ${endTime}\n\n` +
          `Would you like to download the Outlook Recurring Appointment invite (.ics) now?`
        );
        if (downloadIcs && result.bookings.length > 0) {
          window.SyncEngine.generateICS(result.bookings[0]);
        }
        renderCurrentView();
        return;
      }

      // Standard One-Off Booking
      const newBooking = window.SyncEngine.addBooking(bookingPayload, window.SyncEngine.isAdmin);
      closeAllModals();

      // Prompt Outlook .ics download immediately
      const downloadIcs = confirm(
        `? Seminar Room Booked Successfully!\n\n` +
        `Date: ${newBooking.date}\n` +
        `Time: ${newBooking.startTime} - ${newBooking.endTime}\n` +
        `Booking Ref: ${newBooking.id}\n\n` +
        `Would you like to download the NHS Outlook Calendar invite now?`
      );
      if (downloadIcs) {
        window.SyncEngine.generateICS(newBooking);
      }
      renderCurrentView();
    } catch (err) {
      alert("?? Booking Error: " + err.message);
    }
  }

  // Booking Details Modal
  let activeDetailsBooking = null;

  function openDetailsModal(booking) {
    activeDetailsBooking = booking;
    const equipList = Array.isArray(booking.equipment) && booking.equipment.length > 0
      ? booking.equipment.map(e => `<span class="tag-department">${e.replace(/_/g, ' ')}</span>`).join(" ")
      : "<span style='color:var(--text-light);'>None requested</span>";

    const recurrenceBadge = booking.isRecurring
      ? `<div style="background:#EFF6FF; border:1px solid #BFDBFE; color:#1E40AF; padding:0.4rem 0.6rem; border-radius:4px; font-size:0.8rem; margin-top:0.5rem;">
          ?? <strong>Recurring Series:</strong> ${booking.recurrenceSummary || 'Periodic Session'} (Session ${booking.occurrenceIndex || 1} of ${booking.totalOccurrences || 1})
         </div>`
      : "";

    elements.detailsBody.innerHTML = `
      <div style="background:var(--bg-surface-subtle); padding:1rem; border-radius:var(--radius-sm); border-left:4px solid var(--nhs-blue);">
        <h4 style="margin:0; font-size:1.1rem; color:var(--nhs-dark-blue);">${booking.title}</h4>
        <div style="font-size:0.85rem; color:var(--text-muted); margin-top:0.3rem;">
          ?? <strong>${booking.date}</strong> &nbsp;�&nbsp; ?? <strong>${booking.startTime} - ${booking.endTime}</strong> (${booking.durationMinutes} mins)
        </div>
        ${recurrenceBadge}
      </div>
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:0.75rem; font-size:0.85rem;">
        <div>
          <span style="color:var(--text-muted);">Booker:</span><br>
          <strong>${booking.bookerName}</strong>
        </div>
        <div>
          <span style="color:var(--text-muted);">Department / Specialty:</span><br>
          <strong>${booking.department}</strong>
        </div>
        <div>
          <span style="color:var(--text-muted);">NHS Email:</span><br>
          <a href="mailto:${booking.email}" style="color:var(--nhs-blue);">${booking.email}</a>
        </div>
        <div>
          <span style="color:var(--text-muted);">Contact Bleep / Ext:</span><br>
          <strong>${booking.bleep}</strong>
        </div>
        <div>
          <span style="color:var(--text-muted);">Expected Attendees:</span><br>
          <strong>${booking.attendees} people</strong>
        </div>
        <div>
          <span style="color:var(--text-muted);">Booking Reference:</span><br>
          <code>${booking.id}</code>
        </div>
      </div>
      <div>
        <span style="font-size:0.82rem; color:var(--text-muted);">Equipment Requested:</span>
        <div style="margin-top:0.35rem; display:flex; gap:0.4rem; flex-wrap:wrap;">
          ${equipList}
        </div>
      </div>
    `;

    elements.detailsModal.classList.add("open");
  }

  function escapeHtml(str) {
    if (!str) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  // Chief Pharmacist Guardrail Spoken Sound & Voice Selection (from POCS Replacement)
  function soundEnabled() {
    try { return localStorage.getItem("seminar_lockout_sound") !== "off"; } catch (e) { return true; }
  }

  function setSoundEnabled(on) {
    try { localStorage.setItem("seminar_lockout_sound", on ? "on" : "off"); } catch (e) {}
    if (!on) stopSpeech();
    refreshSoundButton();
  }

  function refreshSoundButton() {
    if (elements.lockoutSoundBtn) {
      elements.lockoutSoundBtn.textContent = soundEnabled() ? "?? Sound on" : "?? Sound off";
    }
  }

  /** Prefer the most natural-sounding British English voice available on this browser/PC (Edge/Chrome) */
  function pickVoice() {
    const voices = (window.speechSynthesis && speechSynthesis.getVoices()) || [];
    const en = voices.filter(v => /^en[-_]/i.test(v.lang));
    const score = (v) =>
      (/natural/i.test(v.name) ? 100 : 0) +
      (/online/i.test(v.name) ? 40 : 0) +
      (/en[-_]GB/i.test(v.lang) ? 20 : 0) +
      (/ryan|sonia|libby|thomas|hazel|george|susan/i.test(v.name) ? 10 : 0) +
      (/google/i.test(v.name) ? 5 : 0);
    return en.sort((a, b) => score(b) - score(a))[0] || null;
  }

  function speak(text) {
    try {
      if (!soundEnabled() || !window.speechSynthesis || !window.SpeechSynthesisUtterance) return;
      speechSynthesis.cancel();
      const u = new SpeechSynthesisUtterance(text);
      const v = pickVoice();
      if (v) { u.voice = v; u.lang = v.lang; } else { u.lang = "en-GB"; }
      u.rate = 0.92;
      u.pitch = 1;
      u.volume = 0.95;
      speechSynthesis.speak(u);
    } catch (e) {
      console.warn("Spoken lockout warning unavailable:", e);
    }
  }

  function stopSpeech() {
    try { if (window.speechSynthesis) speechSynthesis.cancel(); } catch (e) {}
  }

  if (typeof window !== "undefined" && window.speechSynthesis) {
    if (speechSynthesis.onvoiceschanged !== undefined) {
      speechSynthesis.onvoiceschanged = () => { pickVoice(); };
    }
  }

  // Synthetic Neuralyzer sound & screen flash
  function playNeuralyzerSound() {
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      const ctx = new AudioCtx();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();

      osc.type = "sawtooth";
      const now = ctx.currentTime;
      osc.frequency.setValueAtTime(3200, now);
      osc.frequency.exponentialRampToValueAtTime(120, now + 0.38);

      gain.gain.setValueAtTime(0.35, now);
      gain.gain.exponentialRampToValueAtTime(0.001, now + 0.4);

      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(now);
      osc.stop(now + 0.42);
    } catch (e) {
      // AudioContext policy fallback
    }
  }

  function triggerNeuralyzerFlash() {
    // Close existing open dialogs
    if (elements.cancelPinModal) elements.cancelPinModal.classList.remove("open");
    if (elements.detailsModal) elements.detailsModal.classList.remove("open");

    // Camera flash capacitor sweep sound
    playNeuralyzerSound();

    // Trigger full-screen blinding white flash
    if (elements.neuralyzerFlash) {
      elements.neuralyzerFlash.classList.add("flash-active");
      setTimeout(() => {
        elements.neuralyzerFlash.classList.remove("flash-active");
      }, 160);
    }

    // Open Men-In-Black pop-up with animated GIF, spoken voice & Neuralyzer
    if (elements.mibModal) {
      if (elements.mibAdminPassword) elements.mibAdminPassword.value = "";
      if (elements.mibAuthStatus) {
        elements.mibAuthStatus.style.display = "none";
        elements.mibAuthStatus.textContent = "";
      }

      // 1. Restart animated MIB Neuralyzer artwork from frame 1
      if (elements.mibGif) {
        elements.mibGif.src = "assets/mib_neuralyzer.svg?t=" + Date.now();
      }

      // 2. Spoken warning using natural British English voice
      refreshSoundButton();
      speak("Nothing to see here. That bright flash? Just swamp gas reflecting off Venus. You did not book this seminar room slot, so you cannot cancel it. You were actually on your way to get a cup of tea.");

      setTimeout(() => {
        elements.mibModal.classList.add("open");
        if (elements.mibAdminPassword) elements.mibAdminPassword.focus();
      }, 140);
    }
  }

  function executeBookingCancellation(booking, pin, isAdmin) {
    if (!booking) return;

    if (booking.isRecurring && booking.recurrenceGroupId) {
      const cancelAll = confirm(
        `This booking is part of a recurring series:\n"${booking.recurrenceSummary}"\n\n` +
        `� Click OK to cancel the ENTIRE RECURRING SERIES (${booking.totalOccurrences || ''} sessions).\n` +
        `� Click CANCEL to cancel ONLY THIS SINGLE SESSION on ${booking.date}.`
      );

      try {
        if (cancelAll) {
          const count = window.SyncEngine.deleteRecurringSeries(booking.recurrenceGroupId, pin, isAdmin);
          alert(`? Entire recurring series (${count} sessions) cancelled successfully.`);
        } else {
          window.SyncEngine.deleteBooking(booking.id, pin, isAdmin);
          alert(`? Single session on ${booking.date} cancelled successfully.`);
        }
        closeAllModals();
        renderCurrentView();
      } catch (err) {
        alert("?? Error: " + err.message);
      }
      return;
    }

    // Single booking cancel
    const confirmCancel = confirm(`Are you sure you want to cancel the booking for '${booking.title}' on ${booking.date}?`);
    if (!confirmCancel) return;

    try {
      window.SyncEngine.deleteBooking(booking.id, pin, isAdmin);
      alert("Booking successfully cancelled. The slot is now free for others to book.");
      closeAllModals();
      renderCurrentView();
    } catch (err) {
      alert("?? Error: " + err.message);
    }
  }

  function handleEditBooking() {
    if (!activeDetailsBooking) return;

    if (window.SyncEngine.isAdmin) {
      openEditBookingModal(activeDetailsBooking, "");
      return;
    }

    state.pendingPinAction = "edit";
    state.pinFailedAttempts = 0;

    if (elements.cancelPinTitle) {
      elements.cancelPinTitle.textContent = "Security Check: Edit Booking";
    }
    if (elements.cancelPinDesc) {
      elements.cancelPinDesc.innerHTML = "This slot was reserved by <strong>" + (activeDetailsBooking.bookerName || "Staff Member") + "</strong> (" + (activeDetailsBooking.department || "Pharmacy") + ").<br><br>To prevent unauthorized schedule alterations, please enter the <strong>4-digit PIN</strong> set by the booker. (Admins may enter the Directorate Admin PIN).";
    }
    if (elements.btnVerifyCancelPin) {
      elements.btnVerifyCancelPin.textContent = "Verify & Edit Booking";
      elements.btnVerifyCancelPin.style.background = "#005EB8";
      elements.btnVerifyCancelPin.style.borderColor = "#005EB8";
    }

    if (elements.cancelSlotPin) elements.cancelSlotPin.value = "";
    if (elements.cancelPinError) {
      elements.cancelPinError.style.display = "none";
      elements.cancelPinError.textContent = "";
    }
    if (elements.cancelPinModal) {
      elements.cancelPinModal.classList.add("open");
      if (elements.cancelSlotPin) {
        setTimeout(() => elements.cancelSlotPin.focus(), 100);
      }
    }
  }

  function openEditBookingModal(booking, verifiedPin = "") {
    closeAllModals();
    state.editingBookingId = booking.id;
    if (elements.bookingRoom) {
      populateRoomSelect(elements.bookingRoom, state.selectedSite || "QEH Heritage", booking.roomId || state.selectedRoomId || "qeh-seminar-room");
      elements.bookingRoom.value = booking.roomId || state.selectedRoomId || "qeh-seminar-room";
      updateBookingRoomCapacityHint();
    }
    elements.bookingModalTitle.textContent = "Edit Seminar Room Booking";
    elements.bookingDateInput.value = booking.date;
    elements.bookingStartTimeInput.value = booking.startTime;
    setBookingDuration(booking.durationMinutes || 30);
    elements.bookerNameInput.value = booking.bookerName;
    elements.bookerEmailInput.value = booking.email;
    elements.bookerBleepInput.value = booking.bleep;
    elements.bookerDeptInput.value = booking.department;
    elements.bookingTitleInput.value = booking.title;
    elements.bookingAttendeesInput.value = booking.attendees;
    elements.bookerPinInput.value = verifiedPin || booking.pin || "";
    elements.bookingModal.classList.add("open");
    updateLiveAvailabilityPreview();
  }

  function handleCancelBooking() {
    if (!activeDetailsBooking) return;

    if (window.SyncEngine.isAdmin) {
      executeBookingCancellation(activeDetailsBooking, "", true);
      return;
    }

    state.pendingPinAction = "cancel";
    state.pinFailedAttempts = 0;

    if (elements.cancelPinTitle) {
      elements.cancelPinTitle.textContent = "Cancel / Release Room Slot";
    }
    if (elements.cancelPinDesc) {
      elements.cancelPinDesc.innerHTML = "This slot was reserved by <strong>" + (activeDetailsBooking.bookerName || "Staff Member") + "</strong> (" + (activeDetailsBooking.department || "Pharmacy") + ").<br><br>To prevent unauthorized cancellations, please enter the <strong>4-digit PIN</strong> set by the booker. (Admins may enter the Directorate Admin PIN).";
    }
    if (elements.btnVerifyCancelPin) {
      elements.btnVerifyCancelPin.textContent = "Verify & Cancel Slot";
      elements.btnVerifyCancelPin.style.background = "#DA291C";
      elements.btnVerifyCancelPin.style.borderColor = "#DA291C";
    }

    // Open clean PIN verification modal
    if (elements.cancelSlotPin) elements.cancelSlotPin.value = "";
    if (elements.cancelPinError) {
      elements.cancelPinError.style.display = "none";
      elements.cancelPinError.textContent = "";
    }
    if (elements.cancelPinModal) {
      elements.cancelPinModal.classList.add("open");
      if (elements.cancelSlotPin) {
        setTimeout(() => elements.cancelSlotPin.focus(), 100);
      }
    }
  }

  function handleVerifyCancelPin() {
    if (!activeDetailsBooking) return;
    const pin = elements.cancelSlotPin ? elements.cancelSlotPin.value.trim() : "";

    // Check if booker PIN matches
    const expectedPin = activeDetailsBooking.pin ? String(activeDetailsBooking.pin).trim() : "";
    const isBookerPin = expectedPin ? (pin === expectedPin) : false;
    const isAdminPin = (((window.SyncEngine.config && pin === window.SyncEngine.config.adminPin) || pin === "7788"));

    if (isBookerPin || isAdminPin) {
      state.pinFailedAttempts = 0;
      if (isAdminPin) {
        window.SyncEngine.isAdmin = true;
      }
      if (elements.cancelPinModal) elements.cancelPinModal.classList.remove("open");

      if (state.pendingPinAction === "edit") {
        openEditBookingModal(activeDetailsBooking, pin);
      } else {
        executeBookingCancellation(activeDetailsBooking, pin, isAdminPin);
      }
    } else {
      state.pinFailedAttempts = (state.pinFailedAttempts || 0) + 1;
      if (state.pinFailedAttempts >= 3) {
        state.pinFailedAttempts = 0;
        triggerNeuralyzerFlash();
      } else {
        if (elements.cancelPinError) {
          elements.cancelPinError.style.display = "block";
          elements.cancelPinError.textContent = "? Incorrect PIN (Attempt " + state.pinFailedAttempts + " of 3). Only the booker or Pharmacy Directorate Admin can modify this booking.";
        }
        if (elements.cancelSlotPin) {
          elements.cancelSlotPin.value = "";
          elements.cancelSlotPin.focus();
        }
      }
    }
  }
    function handleMibAdminOverride() {
    if (!activeDetailsBooking) return;
    const pwd = elements.mibAdminPassword ? elements.mibAdminPassword.value.trim() : "";

    if (window.SyncEngine.verifyAdminPin(pwd)) {
      if (elements.mibAuthStatus) {
        elements.mibAuthStatus.style.display = "block";
        elements.mibAuthStatus.style.color = "#10B981";
        elements.mibAuthStatus.textContent = "? Directorate Admin Authorization Confirmed. Releasing slot...";
      }
      setTimeout(() => {
        if (elements.mibModal) elements.mibModal.classList.remove("open");
        executeBookingCancellation(activeDetailsBooking, "", true);
      }, 400);
    } else {
      // Shake MIB modal card and show access denied
      const card = elements.mibModal ? elements.mibModal.querySelector(".mib-card") : null;
      if (card) {
        card.classList.remove("shake");
        void card.offsetWidth; // retrigger reflow
        card.classList.add("shake");
      }
      if (elements.mibAuthStatus) {
        elements.mibAuthStatus.style.display = "block";
        elements.mibAuthStatus.style.color = "#FF4444";
        elements.mibAuthStatus.textContent = "? Invalid Directorate Admin Password. Access Denied.";
      }
      if (elements.mibAdminPassword) {
        elements.mibAdminPassword.value = "";
        elements.mibAdminPassword.focus();
      }
    }
  }

  // Door Sign & QR Generator Modal
  function renderDoorQrCode(targetUrl) {
    if (!elements.doorQrContainer) return;
    elements.doorQrContainer.innerHTML = "";

    try {
      if (window.QRCode) {
        window.QRCode(elements.doorQrContainer, {
          text: targetUrl,
          width: 220,
          height: 220,
          colorDark: "#003087",
          colorLight: "#FFFFFF",
          correctLevel: 1
        });
      }
    } catch (err) {
      console.warn("QR generation fallback:", err);
      elements.doorQrContainer.innerHTML = `<div style="padding:1rem; color:#005EB8; font-size:0.85rem;"><strong>${targetUrl}</strong></div>`;
    }
  }

  function openDoorSignModal() {
    let url = window.location.href;
    if (window.SyncEngine && window.SyncEngine.config && window.SyncEngine.config.cloudSyncUrl) {
      // If user has a cloud sync URL or mobile host
    }
    if (url.startsWith("file://")) {
      url = "http://localhost:8089";
    }

    if (elements.doorQrUrlInput) {
      elements.doorQrUrlInput.value = url;
    }

    renderDoorQrCode(url);
    elements.doorSignModal.classList.add("open");
  }

  // Admin Superuser Login
  function handleAdminLogin() {
    const pin = elements.adminPinInput.value.trim();
    if (window.SyncEngine.verifyAdminPin(pin)) {
      alert("? Admin Superuser Mode Activated!\nYou can now edit, cancel, or block any room slot.");
      closeAllModals();
      renderCurrentView();
    } else {
      alert("? Incorrect Admin PIN. Please check with the Pharmacy IT lead.");
    }
  }

  function handleAdminLogout() {
    window.SyncEngine.logoutAdmin();
    alert("Admin Superuser Mode deactivated.");
    renderCurrentView();
  }

  // Close All Modals
  function closeAllModals() {
    stopSpeech();
    document.querySelectorAll(".modal-overlay").forEach(m => m.classList.remove("open"));
  }

  // Event Listeners Setup
  
  // =========================================================================
  // Staff Autocomplete & User Directory Management
  // =========================================================================
  function updateStaffSuggestions() {
    if (!elements.staffSuggestions || !window.SyncEngine) return;
    const users = window.SyncEngine.getUsers ? window.SyncEngine.getUsers() : [];
    elements.staffSuggestions.innerHTML = "";
    users.forEach(u => {
      const opt = document.createElement("option");
      opt.value = u.name;
      opt.textContent = u.name + (u.department ? " � " + u.department : "") + (u.bleep ? " (" + u.bleep + ")" : "");
      elements.staffSuggestions.appendChild(opt);
    });
  }

  function handleBookerNameInput() {
    if (!elements.bookerNameInput || !window.SyncEngine) return;
    const val = elements.bookerNameInput.value.trim();
    if (!val) {
      if (elements.staffMatchBadge) elements.staffMatchBadge.style.display = "none";
      return;
    }
    const matched = window.SyncEngine.getUserByName(val);
    if (matched) {
      if (elements.bookerEmailInput && (!elements.bookerEmailInput.value || elements.bookerEmailInput.value.includes("@"))) {
        elements.bookerEmailInput.value = matched.email;
      }
      if (elements.bookerDeptSelect) {
        elements.bookerDeptSelect.value = matched.department;
      }
      if (elements.bookerBleepInput) {
        elements.bookerBleepInput.value = matched.bleep || "";
      }
      // Do not auto-fill PIN for security - booker must enter their own secret PIN
      if (elements.staffMatchBadge) {
        elements.staffMatchBadge.style.display = "inline";
        elements.staffMatchBadge.textContent = "? Recognized: " + matched.name + " (" + matched.department + ")";
      }
    } else {
      if (elements.staffMatchBadge) elements.staffMatchBadge.style.display = "none";
    }
  }

  function openManageUsersModal() {
    if (!elements.manageUsersModal) return;
    resetUserForm();
    renderAdminUsersList();
    elements.manageUsersModal.classList.add("open");
  }

  function resetUserForm() {
    if (!elements.adminAddUserForm) return;
    elements.adminAddUserForm.reset();
    if (elements.editUserId) elements.editUserId.value = "";
    if (elements.btnCancelEditUser) elements.btnCancelEditUser.style.display = "none";
    if (elements.btnSaveUserSubmit) elements.btnSaveUserSubmit.textContent = "? Save Staff Member";
    const title = document.getElementById("user-form-title");
    if (title) title.textContent = "? Add Staff Member";
  }

  function renderAdminUsersList(searchFilter = "") {
    if (!elements.adminUsersTableBody || !window.SyncEngine) return;
    const users = window.SyncEngine.getUsers ? window.SyncEngine.getUsers() : [];
    const filter = (searchFilter || "").toLowerCase().trim();
    const filtered = filter ? users.filter(u => 
      u.name.toLowerCase().includes(filter) || 
      u.email.toLowerCase().includes(filter) || 
      (u.department && u.department.toLowerCase().includes(filter)) ||
      (u.bleep && u.bleep.toLowerCase().includes(filter))
    ) : users;

    if (elements.adminTotalUsersCount) {
      elements.adminTotalUsersCount.textContent = users.length + " staff registered";
    }

    elements.adminUsersTableBody.innerHTML = "";
    if (filtered.length === 0) {
      elements.adminUsersTableBody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding: 1.5rem; color: var(--text-muted);">No staff members found matching search.</td></tr>';
      return;
    }

    filtered.forEach(u => {
      const tr = document.createElement("tr");
      tr.style.borderBottom = "1px solid var(--border-subtle)";
      
      const roleBadge = u.role === "admin" 
        ? '<span style="background: #fef2f2; color: #dc2626; padding: 2px 6px; border-radius: 4px; font-weight:700; font-size:0.75rem;">Admin</span>'
        : u.role === "lead" 
        ? '<span style="background: #eff6ff; color: #2563eb; padding: 2px 6px; border-radius: 4px; font-weight:700; font-size:0.75rem;">Lead</span>'
        : '<span style="background: #f0fdf4; color: #16a34a; padding: 2px 6px; border-radius: 4px; font-weight:600; font-size:0.75rem;">Staff</span>';

      tr.innerHTML = 
        '<td style="padding: 8px 10px; font-weight: 600;">' + (u.name || '') + '</td>' +
        '<td style="padding: 8px 10px; color: var(--text-muted); font-family: monospace;">' + (u.email || '') + '</td>' +
        '<td style="padding: 8px 10px;">' + (u.bleep || '�') + '</td>' +
        '<td style="padding: 8px 10px;">' + (u.department || 'Clinical Pharmacy') + '</td>' +
        '<td style="padding: 8px 10px;">' + roleBadge + '</td>' +
        '<td style="padding: 8px 10px; text-align: center;">' +
          '<button class="btn-secondary btn-sm btn-edit-user" data-id="' + u.id + '" style="padding: 2px 6px; font-size: 0.75rem; margin-right: 4px;">?? Edit</button>' +
          '<button class="btn-secondary btn-sm btn-delete-user" data-id="' + u.id + '" style="padding: 2px 6px; font-size: 0.75rem; color: #dc2626;">???</button>' +
        '</td>';
      elements.adminUsersTableBody.appendChild(tr);
    });

    elements.adminUsersTableBody.querySelectorAll(".btn-edit-user").forEach(btn => {
      btn.onclick = () => editStaffMember(btn.dataset.id);
    });
    elements.adminUsersTableBody.querySelectorAll(".btn-delete-user").forEach(btn => {
      btn.onclick = () => deleteStaffMember(btn.dataset.id);
    });
  }

  function editStaffMember(userId) {
    if (!window.SyncEngine) return;
    const user = (window.SyncEngine.getUsers() || []).find(u => u.id === userId);
    if (!user) return;

    if (elements.editUserId) elements.editUserId.value = user.id;
    if (elements.userFormName) elements.userFormName.value = user.name;
    if (elements.userFormEmail) elements.userFormEmail.value = user.email;
    if (elements.userFormBleep) elements.userFormBleep.value = user.bleep || "";
    if (elements.userFormDept) elements.userFormDept.value = user.department;
    if (elements.userFormRole) elements.userFormRole.value = user.role || "staff";
    if (elements.userFormPin) elements.userFormPin.value = user.pin || "1234";

    if (elements.btnCancelEditUser) elements.btnCancelEditUser.style.display = "inline-block";
    if (elements.btnSaveUserSubmit) elements.btnSaveUserSubmit.textContent = "?? Update Staff Member";
    const title = document.getElementById("user-form-title");
    if (title) title.textContent = "?? Edit Staff Member";
  }

  function deleteStaffMember(userId) {
    if (!window.SyncEngine) return;
    const user = (window.SyncEngine.getUsers() || []).find(u => u.id === userId);
    if (!user) return;
    if (confirm("Are you sure you want to remove '" + user.name + "' from the staff directory?")) {
      window.SyncEngine.deleteUser(userId);
      renderAdminUsersList(elements.userSearchInput ? elements.userSearchInput.value : "");
      updateStaffSuggestions();
    }
  }

  function initEvents() {
    // Nav buttons
    elements.btnPrevDate.onclick = () => {
      state.selectedDate.setDate(state.selectedDate.getDate() - 1);
      renderCurrentView();
    };

    elements.btnNextDate.onclick = () => {
      state.selectedDate.setDate(state.selectedDate.getDate() + 1);
      renderCurrentView();
    };

    elements.btnToday.onclick = () => {
      state.selectedDate = new Date();
      renderCurrentView();
    };

    if (elements.calendarDateInput) {
      elements.calendarDateInput.onchange = (e) => {
        if (!e.target.value) return;
        const parts = e.target.value.split("-");
        state.selectedDate = new Date(parts[0], parts[1] - 1, parts[2]);
        renderCurrentView();
      };
    }

    if (elements.hiddenDatePicker) {
      elements.hiddenDatePicker.onchange = (e) => {
        const parts = e.target.value.split("-");
        state.selectedDate = new Date(parts[0], parts[1] - 1, parts[2]);
        renderCurrentView();
      };
    }

    // View toggles
    elements.viewDayBtn.onclick = () => { state.viewMode = "day"; renderCurrentView(); };
    elements.viewWeekBtn.onclick = () => { state.viewMode = "week"; renderCurrentView(); };
    elements.viewAgendaBtn.onclick = () => { state.viewMode = "agenda"; renderCurrentView(); };

    // Department filter
    elements.departmentFilter.onchange = (e) => {
      state.filterDepartment = e.target.value;
      renderCurrentView();
    };

    // Quick Book & Future Session Buttons
    elements.quickBookBtn.onclick = () => {
      openBookingModal(getISODate(new Date()), state.activeSlotStart);
    };

    if (elements.btnBookSession) {
      if (elements.bookingRoom) {
      elements.bookingRoom.onchange = () => {
        updateBookingRoomCapacityHint();
        updateLiveAvailabilityPreview();
      };
    }

    if (elements.roomSelectDropdown) {
      elements.roomSelectDropdown.onchange = (e) => {
        selectRoom(e.target.value);
      };
    }

    elements.btnBookSession.onclick = () => {
        openBookingModal(getISODate(state.selectedDate), "09:00");
      };
    }

    // Modal Live Slot Availability inputs change
    if (elements.bookingDateInput) {
      elements.bookingDateInput.onchange = () => updateLiveAvailabilityPreview();
    }
    if (elements.bookingStartTimeInput) {
      elements.bookingStartTimeInput.onchange = () => updateLiveAvailabilityPreview();
    }

    // Recurrence Modal Trigger & Controls
    if (elements.btnOpenRecurrenceModal) {
      elements.btnOpenRecurrenceModal.onclick = openRecurrenceModal;
    }
    if (elements.btnRecOk) {
      elements.btnRecOk.onclick = handleRecurrenceOk;
    }
    if (elements.btnRecCancel) {
      elements.btnRecCancel.onclick = () => elements.recurrenceModal.classList.remove("open");
    }
    if (elements.btnRecRemove) {
      elements.btnRecRemove.onclick = handleRecurrenceRemove;
    }

    // Pattern Radio Switches in Recurrence Modal
    if (elements.recPatternRadios) {
      elements.recPatternRadios.forEach(radio => {
        radio.onchange = () => switchRecurrencePanel(radio.value);
      });
    }

    // Sync Recurrence Modal Times
    if (elements.recTimeStart && elements.recTimeDuration && elements.recTimeEnd) {
      const syncRecEndTime = () => {
        const startMins = window.SyncEngine.timeToMinutes(elements.recTimeStart.value);
        const duration = parseInt(elements.recTimeDuration.value, 10);
        elements.recTimeEnd.value = window.SyncEngine.minutesToTime(startMins + duration);
      };

      elements.recTimeStart.onchange = syncRecEndTime;
      elements.recTimeDuration.onchange = syncRecEndTime;

      elements.recTimeEnd.onchange = () => {
        const startMins = window.SyncEngine.timeToMinutes(elements.recTimeStart.value);
        const endMins = window.SyncEngine.timeToMinutes(elements.recTimeEnd.value);
        if (endMins > startMins) {
          const dur = endMins - startMins;
          elements.recTimeDuration.value = String(dur);
        }
      };
    }

    // Duration buttons in modal
    elements.durationButtons.forEach(btn => {
      btn.onclick = () => setBookingDuration(parseInt(btn.dataset.duration, 10));
    });

    // Form submit
    elements.bookingForm.onsubmit = handleBookingFormSubmit;

    // Details actions
    elements.btnOutlookInvite.onclick = () => {
      if (activeDetailsBooking) window.SyncEngine.generateICS(activeDetailsBooking);
    };
    elements.btnCancelBooking.onclick = handleCancelBooking;

    // Cancellation PIN Verification Modal
    if (elements.btnVerifyCancelPin) {
      elements.btnVerifyCancelPin.onclick = handleVerifyCancelPin;
    }
    if (elements.btnNotMyBooking) {
      elements.btnNotMyBooking.onclick = () => {
        triggerNeuralyzerFlash();
      };
    }
    if (elements.cancelSlotPin) {
      elements.cancelSlotPin.onkeydown = (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          handleVerifyCancelPin();
        }
      };
    }

    // MIB Neuralyzer Admin Override
    if (elements.btnMibOverrideCancel) {
      elements.btnMibOverrideCancel.onclick = handleMibAdminOverride;
    }
    if (elements.mibAdminPassword) {
      elements.mibAdminPassword.onkeydown = (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          handleMibAdminOverride();
        }
      };
    }
    if (elements.btnMibDismiss) {
      elements.btnMibDismiss.onclick = closeAllModals;
    }

    if (elements.lockoutSoundBtn) {
      elements.lockoutSoundBtn.onclick = () => {
        setSoundEnabled(!soundEnabled());
        if (soundEnabled()) {
          speak("Sound enabled.");
        }
      };
    }
    elements.btnEditBooking.onclick = handleEditBooking;

    // Door Sign & QR Code Management
    elements.printDoorSignBtn.onclick = openDoorSignModal;
    elements.printSignActionBtn.onclick = () => window.print();

    if (elements.btnCopyDoorUrl) {
      elements.btnCopyDoorUrl.onclick = () => {
        const url = elements.doorQrUrlInput ? elements.doorQrUrlInput.value : window.location.href;
        navigator.clipboard.writeText(url).then(() => {
          alert("?? Link copied to clipboard:\n" + url);
        }).catch(() => {
          prompt("Copy this URL for door sign / phone access:", url);
        });
      };
    }

    if (elements.btnRefreshDoorQr) {
      elements.btnRefreshDoorQr.onclick = () => {
        const url = elements.doorQrUrlInput ? elements.doorQrUrlInput.value.trim() : window.location.href;
        if (url) {
          renderDoorQrCode(url);
        }
      };
    }

    // Admin toggles
    elements.adminToggleBtn.onclick = () => {
      if (window.SyncEngine.isAdmin) {
        handleAdminLogout();
      } else {
        elements.adminPinInput.value = "";
        elements.adminLoginModal.classList.add("open");
      }
    };
    elements.adminLoginSubmit.onclick = handleAdminLogin;
    elements.adminLogoutBtn.onclick = handleAdminLogout;
    elements.adminExportBtn.onclick = () => window.SyncEngine.exportCSV();

    // Admin maintenance block
    elements.adminBlockBtn.onclick = () => {
      elements.adminBlockForm.reset();
      document.getElementById("block-date").value = getISODate(state.selectedDate);
      elements.adminBlockModal.classList.add("open");
    };

    elements.adminBlockForm.onsubmit = (e) => {
      e.preventDefault();
      const bDate = document.getElementById("block-date").value;
      const bStart = document.getElementById("block-start-time").value;
      const bEnd = document.getElementById("block-end-time").value;
      const bReason = document.getElementById("block-reason").value.trim();

      try {
        window.SyncEngine.addMaintenanceBlock(bDate, bStart, bEnd, bReason);
        alert("Maintenance block scheduled successfully.");
        closeAllModals();
        renderCurrentView();
      } catch (err) {
        alert("?? Error: " + err.message);
      }
    };

    // Cloud Settings Modal
    elements.syncStatusBadge.onclick = () => {
      if (elements.googleSheetUrlInput) elements.googleSheetUrlInput.value = window.SyncEngine.config.googleWebAppUrl || "";
      if (elements.googleSheetTestStatus) { elements.googleSheetTestStatus.textContent = ""; elements.googleSheetTestStatus.style.color = ""; }
      if (elements.supabaseUrlInput) elements.supabaseUrlInput.value = window.SyncEngine.config.supabaseUrl || "";
      if (elements.supabaseAnonKeyInput) elements.supabaseAnonKeyInput.value = window.SyncEngine.config.supabaseAnonKey || "";
      if (elements.cloudUrlInput) elements.cloudUrlInput.value = window.SyncEngine.config.cloudSyncUrl || "";
      if (elements.supabaseTestStatus) {
        elements.supabaseTestStatus.textContent = "";
        elements.supabaseTestStatus.style.color = "";
      }
      elements.cloudSettingsModal.classList.add("open");
    };

    if (elements.btnTestSupabase) {
      elements.btnTestSupabase.onclick = async () => {
        elements.supabaseTestStatus.textContent = "Testing connection...";
        elements.supabaseTestStatus.style.color = "var(--nhs-blue)";
        const url = elements.supabaseUrlInput.value.trim();
        const key = elements.supabaseAnonKeyInput.value.trim();
        const res = await window.SyncEngine.testSupabaseConnection(url, key);
        if (res.success) {
          elements.supabaseTestStatus.textContent = "?? Connected! Found " + res.roomsCount + " rooms, " + res.usersCount + " staff, " + res.bookingsCount + " bookings.";
          elements.supabaseTestStatus.style.color = "#007f3b";
        } else {
          elements.supabaseTestStatus.textContent = "?? " + res.error;
          elements.supabaseTestStatus.style.color = "#da291c";
        }
      };
    }

    elements.saveCloudSettingsBtn.onclick = () => {
      const supUrl = elements.supabaseUrlInput ? elements.supabaseUrlInput.value.trim() : "";
      const supKey = elements.supabaseAnonKeyInput ? elements.supabaseAnonKeyInput.value.trim() : "";
      const restUrl = elements.cloudUrlInput ? elements.cloudUrlInput.value.trim() : "";

      window.SyncEngine.saveConfig({
        supabaseUrl: supUrl,
        supabaseAnonKey: supKey,
        cloudSyncUrl: restUrl
      });

      alert("? Cloud database settings saved!\nReal-time synchronization is now active.");
      closeAllModals();
      updateSyncStatusBadge();
    };

    // Staff & Users Events
    if (elements.bookerNameInput) {
      elements.bookerNameInput.addEventListener("input", handleBookerNameInput);
      elements.bookerNameInput.addEventListener("change", handleBookerNameInput);
    }

    if (elements.adminManageUsersBtn) {
      elements.adminManageUsersBtn.onclick = openManageUsersModal;
    }

    if (elements.btnCancelEditUser) {
      elements.btnCancelEditUser.onclick = resetUserForm;
    }

    if (elements.userSearchInput) {
      elements.userSearchInput.oninput = () => {
        renderAdminUsersList(elements.userSearchInput.value);
      };
    }

    if (elements.btnExportUsersCsv) {
      elements.btnExportUsersCsv.onclick = () => {
        window.SyncEngine.exportUsersCsv();
      };
    }

    // Direct Excel (.xlsx / .xls) and CSV File Importer
    if (elements.btnImportUsersCsv && elements.userCsvFileInput) {
      elements.btnImportUsersCsv.onclick = () => {
        elements.userCsvFileInput.click();
      };
      elements.userCsvFileInput.onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const fileName = file.name.toLowerCase();
        const isExcel = fileName.endsWith(".xlsx") || fileName.endsWith(".xls");

        if (isExcel) {
          const reader = new FileReader();
          reader.onload = (evt) => {
            try {
              if (typeof XLSX === "undefined") {
                throw new Error("Excel parsing engine (xlsx.full.min.js) is not loaded.");
              }
              const data = new Uint8Array(evt.target.result);
              const workbook = XLSX.read(data, { type: "array" });
              
              // If 'Staff Directory' sheet exists, use that; otherwise first sheet
              const sheetName = workbook.SheetNames.find(s => s.toLowerCase().includes("staff")) || workbook.SheetNames[0];
              const worksheet = workbook.Sheets[sheetName];
              const csvText = XLSX.utils.sheet_to_csv(worksheet);

              const count = window.SyncEngine.importUsersFromCsv(csvText);
              alert("? Successfully imported " + count + " staff members directly from '" + file.name + "' (Sheet: " + sheetName + ")!");
              renderAdminUsersList();
              updateStaffSuggestions();
            } catch (err) {
              alert("?? Excel Import Error: " + err.message);
            }
            elements.userCsvFileInput.value = "";
          };
          reader.readAsArrayBuffer(file);
        } else {
          // Standard CSV
          const reader = new FileReader();
          reader.onload = (evt) => {
            try {
              const count = window.SyncEngine.importUsersFromCsv(evt.target.result);
              alert("? Successfully imported " + count + " staff members from '" + file.name + "'!");
              renderAdminUsersList();
              updateStaffSuggestions();
            } catch (err) {
              alert("?? CSV Import Error: " + err.message);
            }
            elements.userCsvFileInput.value = "";
          };
          reader.readAsText(file);
        }
      };
    }

    if (elements.adminAddUserForm) {
      elements.adminAddUserForm.onsubmit = (e) => {
        e.preventDefault();
        const editId = elements.editUserId.value;
        const userData = {
          name: elements.userFormName.value.trim(),
          email: elements.userFormEmail.value.trim(),
          bleep: elements.userFormBleep.value.trim(),
          department: elements.userFormDept.value,
          role: elements.userFormRole.value,
          pin: elements.userFormPin.value.trim() || "1234"
        };

        try {
          if (editId) {
            window.SyncEngine.updateUser(editId, userData);
            alert("Staff member updated successfully.");
          } else {
            window.SyncEngine.addUser(userData);
            alert("Staff member added successfully.");
          }
          resetUserForm();
          renderAdminUsersList();
          updateStaffSuggestions();
        } catch (err) {
          alert("?? Error: " + err.message);
        }
      };
    }

    // Theme toggle
        if (elements.btnInstallPwa) {
      elements.btnInstallPwa.onclick = async () => {
        if (deferredInstallPrompt) {
          deferredInstallPrompt.prompt();
          const choice = await deferredInstallPrompt.userChoice;
          if (choice.outcome === "accepted") {
            elements.btnInstallPwa.style.display = "none";
          }
          deferredInstallPrompt = null;
        } else {
          alert("To install this app on your phone:\n� iPhone (Safari): Tap the Share button, then 'Add to Home Screen'\n� Android (Chrome): Tap the menu (?), then 'Install app' or 'Add to Home screen'");
        }
      };
    }

    elements.themeToggleBtn.onclick = () => {
      state.isDarkMode = !state.isDarkMode;
      localStorage.setItem("qe_theme", state.isDarkMode ? "dark" : "light");
      applyTheme();
    };

    // Modal Close Buttons
    document.querySelectorAll(".modal-close, .btn-modal-cancel").forEach(btn => {
      btn.onclick = closeAllModals;
    });

    // Close on overlay backdrop click
    document.querySelectorAll(".modal-overlay").forEach(overlay => {
      overlay.onclick = (e) => {
        if (e.target === overlay) closeAllModals();
      };
    });

    // Sync engine listener
    window.SyncEngine.onUpdate((eventType) => {
      renderCurrentView();
      updateSyncStatusBadge();
    });

    // Periodic live countdown timer (every 10 seconds)
    setInterval(updateRoomStatus, 10000);
  }

  function updateSyncStatusBadge() {
    const hasGoogle = Boolean(window.SyncEngine.isGoogleSheetsConfigured && window.SyncEngine.isGoogleSheetsConfigured());
    const hasCloud = Boolean(window.SyncEngine.config.cloudSyncUrl || window.SyncEngine.config.supabaseUrl);
    if (hasGoogle || hasCloud) {
      elements.syncDot.className = "sync-dot";
      elements.syncText.textContent = hasGoogle ? "Google Sheets Live" : "Live Cloud Sync";
    } else {
      elements.syncDot.className = "sync-dot offline";
      elements.syncText.textContent = "Local & Network Mode";
    }
  }

  // Initialization
  // =========================================================================
  // Progressive Web App (PWA) Service Worker & Installation
  // =========================================================================
  if ("serviceWorker" in navigator && (window.location.protocol === "https:" || window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1")) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("./sw.js").then((reg) => {
        console.log("PWA Service Worker registered:", reg.scope);
      }).catch((err) => {
        console.warn("PWA Service Worker registration skipped:", err);
      });
    });
  }

  let deferredInstallPrompt = null;
  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    deferredInstallPrompt = e;
    if (elements.btnInstallPwa) {
      elements.btnInstallPwa.style.display = "inline-flex";
    }
  });

  window.addEventListener("appinstalled", () => {
    if (elements.btnInstallPwa) elements.btnInstallPwa.style.display = "none";
    deferredInstallPrompt = null;
    console.log("QE Booking PWA installed successfully");
  });

  window.addEventListener("DOMContentLoaded", () => {
    initElements();
    applyTheme();
    initEvents();
    renderCurrentView();
    updateSyncStatusBadge();
  });

})();


