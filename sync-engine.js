/**
 * QE Pharmacy Seminar Room - Multi-Device Sync Engine
 * University Hospitals Birmingham NHS Foundation Trust
 * Handles LocalStorage caching, Multi-Tab Broadcast, Outlook .ics generation,
 * and Cloud Realtime Synchronization.
 */

(function (global) {
  const STORAGE_KEY = "qe_seminar_room_bookings_v2";
  const CONFIG_KEY = "qe_seminar_room_config_v2";
  const ADMIN_PIN_DEFAULT = "7788"; // Master Admin Superuser PIN
  const BROADCAST_CHANNEL_NAME = "qe_seminar_sync_channel";

  let broadcastChannel = null;
  try {
    if ("BroadcastChannel" in window) {
      broadcastChannel = new BroadcastChannel(BROADCAST_CHANNEL_NAME);
    }
  } catch (e) {
    console.warn("BroadcastChannel not supported", e);
  }

  const STORAGE_KEY_ROOMS = "qe_seminar_rooms_v3";
  const STORAGE_KEY_USERS = "qe_seminar_users_v3";

  // Trust Hospital Sites
  const TRUST_SITES = [
    "QEH Heritage",
    "Good Hope",
    "Heartlands",
    "Solihull"
  ];

  // Default Pharmacy Staff Users
  const DEFAULT_USERS = [
    { id: "usr-01", name: "Asif", email: "asif@uhb.nhs.uk", bleep: "Bleep 4120", department: "Clinical Pharmacy", site: "QEH Heritage", role: "admin", pin: "7788", isActive: true },
    { id: "usr-02", name: "Sarah Jenkins", email: "sarah.jenkins@uhb.nhs.uk", bleep: "Bleep 3341", department: "Aseptics & Radiopharmacy", site: "QEH Heritage", role: "staff", pin: "1234", isActive: true },
    { id: "usr-03", name: "David Campbell", email: "david.campbell@uhb.nhs.uk", bleep: "Bleep 4120", department: "Clinical Pharmacy", site: "QEH Heritage", role: "staff", pin: "1234", isActive: true },
    { id: "usr-04", name: "Priya Sharma", email: "priya.sharma@uhb.nhs.uk", bleep: "Ext 14892", department: "Medicines Information (MI)", site: "QEH Heritage", role: "staff", pin: "1234", isActive: true },
    { id: "usr-05", name: "Michael Taylor", email: "michael.taylor@uhb.nhs.uk", bleep: "Bleep 2290", department: "EPMA / Digital Health", site: "QEH Heritage", role: "staff", pin: "1234", isActive: true },
    { id: "usr-06", name: "Emma Watson", email: "emma.watson@uhb.nhs.uk", bleep: "Ext 14820", department: "Education & Training", site: "QEH Heritage", role: "lead", pin: "1234", isActive: true },
    { id: "usr-07", name: "Mohammed Ali", email: "mohammed.ali@uhb.nhs.uk", bleep: "Bleep 5512", department: "Oncology & Chemotherapy", site: "QEH Heritage", role: "staff", pin: "1234", isActive: true },
    { id: "usr-08", name: "Rachel Davies", email: "rachel.davies@uhb.nhs.uk", bleep: "Ext 14800", department: "Dispensary", site: "QEH Heritage", role: "staff", pin: "1234", isActive: true },
    { id: "usr-09", name: "Pre-Reg Pharmacist Lead", email: "prereg.pharmacy@uhb.nhs.uk", bleep: "Ext 14895", department: "Education & Training", site: "QEH Heritage", role: "staff", pin: "1234", isActive: true }
  ];

  // Default Hospital Rooms across Trust Sites
  // The Seminar room lives under QEH Heritage
  const DEFAULT_ROOMS = [
    {
      id: "qeh-seminar-room",
      name: "Seminar Room",
      fullName: "Pharmacy Seminar Room",
      site: "QEH Heritage",
      location: "Level 1, Pharmacy Department, Heritage Building",
      capacity: 16,
      equipment: ["wall_display", "owl_camera", "whiteboard", "flipchart"],
      description: "Main pharmacy seminar room for clinical handovers, journal clubs, and departmental training.",
      isDefault: true,
      isBuiltIn: true
    },
    {
      id: "ghh-meeting-room",
      name: "Pharmacy Meeting Room",
      fullName: "Good Hope Pharmacy Meeting Room",
      site: "Good Hope",
      location: "Ground Floor, Treatment Centre (Pharmacy)",
      capacity: 10,
      equipment: ["wall_display", "owl_camera", "whiteboard"],
      description: "Department meeting room for clinical handovers and multi-disciplinary case conferences.",
      isDefault: false,
      isBuiltIn: false
    },
    {
      id: "bhh-seminar-room",
      name: "Pharmacy Seminar Room",
      fullName: "Heartlands Pharmacy Seminar Room",
      site: "Heartlands",
      location: "First Floor, Education Centre & Clinical Pharmacy",
      capacity: 14,
      equipment: ["wall_display", "owl_camera", "whiteboard", "flipchart"],
      description: "Seminar room for clinical pharmacy teaching, ePMA reviews, and departmental training.",
      isDefault: false,
      isBuiltIn: false
    },
    {
      id: "sol-clinical-room",
      name: "Clinical Discussion Room",
      fullName: "Solihull Clinical Discussion Room",
      site: "Solihull",
      location: "Ground Floor, Pharmacy Department",
      capacity: 8,
      equipment: ["wall_display", "whiteboard"],
      description: "Clinical team discussion room and multidisciplinary liaison space.",
      isDefault: false,
      isBuiltIn: false
    }
  ];

  // Default Configuration
  const defaultConfig = {
    roomName: "Pharmacy Seminar Room",
    site: "QEH Heritage",
    hospital: "Queen Elizabeth Hospital Birmingham (UHB)",
    location: "Level 1, Pharmacy Department, Mindelsohn Way, B15 2GW",
    roomCapacity: 16,
    openHour: 8,       // 08:00
    closeHour: 19,     // 19:00
    slotMinutes: 30,   // Half hour slots
    adminPin: ADMIN_PIN_DEFAULT,
    googleWebAppUrl: "https://script.google.com/macros/s/AKfycby2e6aOeEt1aTK8wegskV7Zsvgg7C_21cqg4mtxAJtVibqlWa5NL-s4fDfIplki5fO2/exec",   // Google Apps Script Web App URL for Google Sheet backend
    supabaseUrl: "",        // e.g. "https://xxxx.supabase.co"
    supabaseAnonKey: "",    // public anon API key
    supabaseRealtime: true,
    serverPublicUrl: "",    // Official Trust FQDN for Door QR poster
    cloudSyncUrl: "",       // Optional Firebase / Cloud REST endpoint URL
    cloudApiKey: "",
    lastSyncTime: null
  };

  class SyncEngine {
    constructor() {
      this.isAdmin = false;
      this.listeners = [];
      this.config = this.loadConfig();
      this.rooms = this.loadRooms();
      this.users = this.loadUsers();
      this.bookings = this.loadBookings();
      this.supabaseClient = null;
      this._supabasePollTimer = null;
      this.initBroadcast();
      this.initSupabaseSync();
      this.initGoogleSheetsSync();
      this.initAutoCloudSync();
    }

    getSites() {
      return [...TRUST_SITES];
    }

    getRooms() {
      return this.rooms;
    }

    getRoomById(id) {
      if (!id) return this.rooms[0];
      return this.rooms.find(r => r.id === id) || this.rooms[0];
    }

    getRoomsBySite(site) {
      if (!site || site === "all") return this.rooms;
      return this.rooms.filter(r => r.site.toLowerCase() === site.toLowerCase());
    }

    loadRooms() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY_ROOMS);
        if (raw) {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed) && parsed.length > 0) {
            // Ensure core Seminar room under QEH Heritage is preserved
            const hasSeminar = parsed.some(r => r.id === "qeh-seminar-room");
            if (!hasSeminar) {
              parsed.unshift(DEFAULT_ROOMS[0]);
            }
            return parsed;
          }
        }
      } catch (e) {
        console.warn("Failed to load cached rooms", e);
      }
      const initialRooms = JSON.parse(JSON.stringify(DEFAULT_ROOMS));
      this.persistRooms(initialRooms, false);
      return initialRooms;
    }

    persistRooms(roomsList = null, broadcast = true) {
      if (roomsList) {
        this.rooms = roomsList;
      }
      try {
        localStorage.setItem(STORAGE_KEY_ROOMS, JSON.stringify(this.rooms));
      } catch (e) {
        console.error("Failed to save rooms to localStorage", e);
      }

      if (broadcast && broadcastChannel) {
        try {
          broadcastChannel.postMessage({ type: "SYNC_ROOMS", rooms: this.rooms });
        } catch (e) {
          console.warn("BroadcastChannel error syncing rooms", e);
        }
      }

      this.notifyListeners("rooms_updated");
    }

    addRoom(roomData) {
      if (!this.isAdmin) {
        throw new Error("Admin Superuser permissions required to add new rooms.");
      }
      if (!roomData.name || !roomData.name.trim()) {
        throw new Error("Room name is required.");
      }
      if (!roomData.site || !TRUST_SITES.includes(roomData.site)) {
        throw new Error("Valid Trust hospital site is required (QEH Heritage, Good Hope, Heartlands, Solihull).");
      }

      const siteCode = roomData.site.replace(/[^A-Za-z0-9]/g, "").slice(0, 3).toLowerCase();
      const randSuffix = Math.random().toString(36).substring(2, 6);
      const id = `${siteCode}-room-${Date.now().toString(36)}-${randSuffix}`;

      const newRoom = {
        id: id,
        name: roomData.name.trim(),
        fullName: roomData.fullName ? roomData.fullName.trim() : `${roomData.site} ${roomData.name.trim()}`,
        site: roomData.site,
        location: roomData.location ? roomData.location.trim() : "Pharmacy Department",
        capacity: parseInt(roomData.capacity, 10) || 12,
        equipment: Array.isArray(roomData.equipment) ? roomData.equipment : ["wall_display", "whiteboard"],
        description: roomData.description ? roomData.description.trim() : "",
        isDefault: false,
        isBuiltIn: false,
        createdAt: new Date().toISOString()
      };

      this.rooms.push(newRoom);
      this.persistRooms(this.rooms, true);
      return newRoom;
    }

    updateRoom(id, updatedFields) {
      if (!this.isAdmin) {
        throw new Error("Admin Superuser permissions required to modify rooms.");
      }
      const index = this.rooms.findIndex(r => r.id === id);
      if (index === -1) throw new Error("Room not found.");

      const existing = this.rooms[index];
      // Keep QEH Heritage Seminar Room protected
      if (existing.id === "qeh-seminar-room" && updatedFields.site && updatedFields.site !== "QEH Heritage") {
        throw new Error("The Pharmacy Seminar Room must remain located under QEH Heritage.");
      }

      this.rooms[index] = Object.assign({}, existing, updatedFields, {
        modifiedAt: new Date().toISOString()
      });
      this.persistRooms(this.rooms, true);
      return this.rooms[index];
    }

    deleteRoom(id) {
      if (!this.isAdmin) {
        throw new Error("Admin Superuser permissions required to delete rooms.");
      }
      const room = this.rooms.find(r => r.id === id);
      if (!room) throw new Error("Room not found.");
      if (room.isBuiltIn || room.id === "qeh-seminar-room") {
        throw new Error("The default QEH Heritage Seminar Room cannot be removed.");
      }

      this.rooms = this.rooms.filter(r => r.id !== id);
      this.persistRooms(this.rooms, true);
      return true;
    }

    // =========================================================================
    // Pharmacy Staff & Users Directory (Importable / Editable)
    // =========================================================================
    loadUsers() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY_USERS);
        if (raw) {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed) && parsed.length > 0) {
            return parsed;
          }
        }
      } catch (e) {
        console.warn("Failed to parse cached users", e);
      }
      const initialUsers = [...DEFAULT_USERS];
      this.persistUsers(initialUsers, false);
      return initialUsers;
    }

    persistUsers(usersList = null, broadcast = true) {
      if (usersList) {
        this.users = usersList;
      }
      try {
        localStorage.setItem(STORAGE_KEY_USERS, JSON.stringify(this.users));
      } catch (e) {
        console.error("Failed to save users to localStorage", e);
      }

      if (broadcast && broadcastChannel) {
        try {
          broadcastChannel.postMessage({ type: "SYNC_USERS", users: this.users });
        } catch (e) {
          console.warn("BroadcastChannel error syncing users", e);
        }
      }

      this.notifyListeners("users_updated");
    }

    getUsers() {
      return this.users;
    }

    getUserByName(name) {
      if (!name) return null;
      const clean = name.trim().toLowerCase();
      return this.users.find(u => u.name.toLowerCase() === clean || u.name.toLowerCase().includes(clean)) || null;
    }

    getUserByEmail(email) {
      if (!email) return null;
      const clean = email.trim().toLowerCase();
      return this.users.find(u => u.email.toLowerCase() === clean) || null;
    }

    addUser(userData) {
      if (!userData.name || !userData.name.trim()) throw new Error("Staff name is required.");
      if (!userData.email || !userData.email.trim()) throw new Error("NHS email is required.");
      
      const emailClean = userData.email.trim().toLowerCase();
      if (this.users.some(u => u.email.toLowerCase() === emailClean)) {
        throw new Error(`A staff member with email '${emailClean}' already exists.`);
      }

      const newUser = {
        id: "usr-" + Date.now().toString(36) + "-" + Math.random().toString(36).substring(2, 6),
        name: userData.name.trim(),
        email: emailClean,
        bleep: (userData.bleep || "").trim(),
        department: userData.department || "Clinical Pharmacy",
        site: userData.site || "QEH Heritage",
        role: userData.role || "staff",
        pin: userData.pin || "1234",
        isActive: true,
        createdAt: new Date().toISOString()
      };

      this.users.push(newUser);
      this.persistUsers(this.users, true);
      this.pushUserToSupabase(newUser);
      this.pushUserToGoogleSheets(newUser);
      return newUser;
    }

    updateUser(id, updatedFields) {
      const idx = this.users.findIndex(u => u.id === id);
      if (idx === -1) throw new Error("Staff member not found.");
      
      const existing = this.users[idx];
      this.users[idx] = Object.assign({}, existing, updatedFields, {
        modifiedAt: new Date().toISOString()
      });

      this.persistUsers(this.users, true);
      this.pushUserToSupabase(this.users[idx]);
      return this.users[idx];
    }

    deleteUser(id) {
      const user = this.users.find(u => u.id === id);
      if (!user) throw new Error("Staff member not found.");
      this.users = this.users.filter(u => u.id !== id);
      this.persistUsers(this.users, true);
      this.deleteUserFromSupabase(id);
      this.deleteUserFromGoogleSheets(id);
      return true;
    }

    importUsersFromCsv(csvText) {
      if (!csvText || !csvText.trim()) throw new Error("CSV content is empty.");
      const lines = csvText.trim().split(/\r?\n/).filter(l => l.trim().length > 0);
      if (lines.length < 2) throw new Error("CSV must have a header row and at least one staff row.");

      // Parse headers
      const headers = lines[0].split(",").map(h => h.trim().toLowerCase().replace(/[^a-z0-9]/g, ""));
      let nameIdx = headers.findIndex(h => h === "stafffullname" || h === "fullname" || h === "name" || (h.includes("name") && !h.includes("sur")));
      const surIdx = headers.findIndex(h => h.includes("sur") || h.includes("last"));
      const foreIdx = headers.findIndex(h => h.includes("fore") || h.includes("first"));
      const emailIdx = headers.findIndex(h => h.includes("email") || h.includes("mail"));
      const bleepIdx = headers.findIndex(h => h.includes("bleep") || h.includes("ext") || h.includes("phone"));
      const deptIdx = headers.findIndex(h => h.includes("dept") || h.includes("department") || h.includes("team"));
      const siteIdx = headers.findIndex(h => h.includes("site") || h.includes("hospital"));
      const roleIdx = headers.findIndex(h => h.includes("role"));
      const pinIdx = headers.findIndex(h => h.includes("pin"));

      if ((nameIdx === -1 && (surIdx === -1 || foreIdx === -1)) || emailIdx === -1) {
        throw new Error("CSV must contain at least Name (or Forename + Surname) and Email columns.");
      }

      let importedCount = 0;
      for (let i = 1; i < lines.length; i++) {
        const row = lines[i].split(",").map(c => c.trim().replace(/^["']|["']$/g, ""));
        let name = nameIdx !== -1 ? row[nameIdx] : "";
        if (!name && foreIdx !== -1 && surIdx !== -1) {
          name = `${row[foreIdx]} ${row[surIdx]}`.trim();
        }
        const email = row[emailIdx];
        if (!name || !email) continue;

        const emailClean = email.toLowerCase();
        const existingIdx = this.users.findIndex(u => u.email.toLowerCase() === emailClean);

        const userData = {
          name: name,
          email: emailClean,
          bleep: bleepIdx !== -1 && row[bleepIdx] ? row[bleepIdx] : "Ext 14892",
          department: deptIdx !== -1 && row[deptIdx] ? row[deptIdx] : "Clinical Pharmacy",
          site: siteIdx !== -1 && row[siteIdx] ? row[siteIdx] : "QEH Heritage",
          role: roleIdx !== -1 && row[roleIdx] ? row[roleIdx] : "staff",
          pin: pinIdx !== -1 && row[pinIdx] ? row[pinIdx] : "1234",
          isActive: true
        };

        if (existingIdx !== -1) {
          this.users[existingIdx] = Object.assign({}, this.users[existingIdx], userData, { modifiedAt: new Date().toISOString() });
          this.pushUserToSupabase(this.users[existingIdx]);
        } else {
          userData.id = "usr-" + Date.now().toString(36) + "-" + Math.random().toString(36).substring(2, 6);
          userData.createdAt = new Date().toISOString();
          this.users.push(userData);
          this.pushUserToSupabase(userData);
        }
        importedCount++;
      }

      this.persistUsers(this.users, true);
      return importedCount;
    }

    exportUsersCsv() {
      const headers = ["Name", "Email", "Bleep", "Department", "Site", "Role", "PIN"];
      const rows = this.users.map(u => [
        `"${(u.name || '').replace(/"/g, '""')}"`,
        `"${(u.email || '').replace(/"/g, '""')}"`,
        `"${(u.bleep || '').replace(/"/g, '""')}"`,
        `"${(u.department || '').replace(/"/g, '""')}"`,
        `"${(u.site || '').replace(/"/g, '""')}"`,
        `"${(u.role || 'staff').replace(/"/g, '""')}"`,
        `"${(u.pin || '1234')}"`
      ]);
      const csv = "\uFEFF" + [headers.join(","), ...rows.map(r => r.join(","))].join("\r\n");
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `UHB_Pharmacy_Staff_Directory_${this.getLocalDateString(new Date())}.csv`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }

    loadConfig() {
      try {
        const saved = localStorage.getItem(CONFIG_KEY);
        return saved ? Object.assign({}, defaultConfig, JSON.parse(saved)) : Object.assign({}, defaultConfig);
      } catch (e) {
        return Object.assign({}, defaultConfig);
      }
    }

    saveConfig(newConfig) {
      const hadSupabase = this.isSupabaseConfigured();
      this.config = Object.assign({}, this.config, newConfig);
      try {
        localStorage.setItem(CONFIG_KEY, JSON.stringify(this.config));
      } catch (e) {
        console.error("Failed to save config to localStorage", e);
      }
      if (newConfig.googleWebAppUrl !== undefined) {
        this.initGoogleSheetsSync();
      }
      if (newConfig.supabaseUrl !== undefined || newConfig.supabaseAnonKey !== undefined) {
        this.initSupabaseSync();
      this.initGoogleSheetsSync();
      }
      this.notifyListeners("config_updated");
    }

    normalizeBooking(b) {
      if (!b) return b;
      if (!b.roomId) b.roomId = "qeh-seminar-room";
      if (!b.site) b.site = "QEH Heritage";
      if (!b.roomName) b.roomName = "Seminar Room";
      if (!b.location) b.location = "Level 1, Pharmacy Department, Heritage Building";

      // Auto-heal startTime and endTime if corrupted by spreadsheet date casting
      if (!b.startTime || b.startTime === "1899-12-30" || !String(b.startTime).includes(":")) {
        const m = String(b.id || "").match(/-(\d{2})(\d{2})$/);
        if (m) {
          b.startTime = `${m[1]}:${m[2]}`;
        }
      }
      const dur = Number(b.durationMinutes) || 30;
      b.durationMinutes = dur;
      if (b.startTime && (!b.endTime || b.endTime === "1899-12-30" || !String(b.endTime).includes(":"))) {
        b.endTime = this.minutesToTime(this.timeToMinutes(b.startTime) + dur);
      }
      return b;
    }

    loadBookings() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw) {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed) && parsed.length > 0) {
            return parsed.map(b => this.normalizeBooking(b));
          }
        }
      } catch (e) {
        console.warn("Failed to parse cached bookings", e);
      }
      // If empty, generate helpful seed demo bookings for today and tomorrow
      const seed = this.generateSeedBookings().map(b => this.normalizeBooking(b));
      this.persistBookings(seed, false);
      return seed;
    }

    persistBookings(bookings, broadcast = true) {
      this.bookings = bookings.map(b => this.normalizeBooking(b));
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(this.bookings));
      } catch (e) {
        console.error("Error saving bookings to localStorage", e);
      }

      if (broadcast && broadcastChannel) {
        try {
          broadcastChannel.postMessage({ type: "SYNC_BOOKINGS", bookings: this.bookings });
        } catch (e) {
          console.warn("BroadcastChannel error", e);
        }
      }

      // If cloud sync URL configured, push in background
      if (this.config.cloudSyncUrl && broadcast) {
        this.pushToCloud(this.bookings);
      }

      this.notifyListeners("bookings_updated");
    }

    initBroadcast() {
      if (broadcastChannel) {
        broadcastChannel.onmessage = (event) => {
          if (event.data && event.data.type === "SYNC_BOOKINGS") {
            this.bookings = (event.data.bookings || []).map(b => this.normalizeBooking(b));
            localStorage.setItem(STORAGE_KEY, JSON.stringify(this.bookings));
            this.notifyListeners("bookings_updated");
          } else if (event.data && event.data.type === "SYNC_ROOMS") {
            this.rooms = event.data.rooms || [];
            localStorage.setItem(STORAGE_KEY_ROOMS, JSON.stringify(this.rooms));
            this.notifyListeners("rooms_updated");
          } else if (event.data && event.data.type === "SYNC_USERS") {
            this.users = event.data.users || [];
            localStorage.setItem(STORAGE_KEY_USERS, JSON.stringify(this.users));
            this.notifyListeners("users_updated");
          }
        };
      }
    }

    onUpdate(fn) {
      this.listeners.push(fn);
    }

    notifyListeners(eventType) {
      if (!this.listeners) return;
      this.listeners.forEach((fn) => {
        try { fn(eventType, this.bookings); } catch (err) { console.error(err); }
      });
    }

    // Generate realistic hospital seed bookings
    generateSeedBookings() {
      const todayStr = this.getLocalDateString(new Date());
      const tomorrow = new Date();
      tomorrow.setDate(tomorrow.getDate() + 1);
      const tomorrowStr = this.getLocalDateString(tomorrow);

      return [
        {
          id: "QE-SR-" + todayStr.replace(/-/g, "") + "-0830",
          date: todayStr,
          startTime: "08:30",
          endTime: "09:00",
          durationMinutes: 30,
          bookerName: "David Campbell",
          email: "david.campbell@uhb.nhs.uk",
          bleep: "Bleep 4120",
          department: "Clinical Pharmacy",
          title: "Daily Clinical Handover & High Risk Patients",
          attendees: 10,
          equipment: ["wall_display", "whiteboard"],
          pin: "1234",
          isBlocked: false,
          createdAt: new Date().toISOString()
        },
        {
          id: "QE-SR-" + todayStr.replace(/-/g, "") + "-1230",
          date: todayStr,
          startTime: "12:30",
          endTime: "13:30",
          durationMinutes: 60,
          bookerName: "Priya Sharma",
          email: "priya.sharma@nhs.net",
          bleep: "Ext 14892",
          department: "Education & Training",
          title: "Foundation Trainee Pharmacist Case Presentation",
          attendees: 8,
          equipment: ["wall_display", "owl_camera", "flipchart"],
          pin: "4321",
          isBlocked: false,
          createdAt: new Date().toISOString()
        },
        {
          id: "QE-SR-" + todayStr.replace(/-/g, "") + "-1500",
          date: todayStr,
          startTime: "15:00",
          endTime: "16:00",
          durationMinutes: 60,
          bookerName: "Sarah Jenkins",
          email: "sarah.jenkins@uhb.nhs.uk",
          bleep: "Bleep 3091",
          department: "EPMA / Digital Health",
          title: "EPMA PICS Order Set Review & Testing",
          attendees: 5,
          equipment: ["wall_display", "owl_camera"],
          pin: "9988",
          isBlocked: false,
          createdAt: new Date().toISOString()
        },
        {
          id: "QE-SR-" + tomorrowStr.replace(/-/g, "") + "-0830",
          date: tomorrowStr,
          startTime: "08:30",
          endTime: "09:00",
          durationMinutes: 30,
          bookerName: "Mark Robinson",
          email: "mark.robinson@uhb.nhs.uk",
          bleep: "Bleep 4120",
          department: "Clinical Pharmacy",
          title: "Daily Clinical Handover",
          attendees: 9,
          equipment: ["wall_display"],
          pin: "1122",
          isBlocked: false,
          createdAt: new Date().toISOString()
        }
      ];
    }

    getLocalDateString(d) {
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }

    // Conflict Detection
    timeToMinutes(timeStr) {
      const parts = timeStr.split(":");
      return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
    }

    minutesToTime(mins) {
      const h = Math.floor(mins / 60);
      const m = mins % 60;
      return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
    }

    isSlotInPast(date, startTime) {
      if (!date) return false;
      const now = new Date();
      const todayStr = this.getLocalDateString(now);
      if (date < todayStr) return true;
      if (date > todayStr) return false;
      if (!startTime) return false;
      const nowMinutes = now.getHours() * 60 + now.getMinutes();
      const startMinutes = this.timeToMinutes(startTime);
      return startMinutes < nowMinutes;
    }

    checkConflict(date, startTime, endTime, excludeBookingId = null, roomId = null, isSuperuser = false) {
      if (!isSuperuser && this.isSlotInPast(date, startTime)) {
        return {
          hasConflict: true,
          isPast: true,
          message: "This date or time slot has already passed. Please select an upcoming slot."
        };
      }

      const reqStart = this.timeToMinutes(startTime);
      const reqEnd = this.timeToMinutes(endTime);

      if (reqStart >= reqEnd) {
        return { hasConflict: true, message: "End time must be after start time." };
      }

      const targetRoomId = roomId || "qeh-seminar-room";
      const dateBookings = this.bookings.filter(b => {
        return b.date === date && 
               b.id !== excludeBookingId && 
               (b.roomId || "qeh-seminar-room") === targetRoomId;
      });

      for (const b of dateBookings) {
        const bStart = this.timeToMinutes(b.startTime);
        const bEnd = this.timeToMinutes(b.endTime);

        // Check overlap: (StartA < EndB) and (EndA > StartB)
        if (reqStart < bEnd && reqEnd > bStart) {
          return {
            hasConflict: true,
            conflictingBooking: b,
            message: `Slot overlaps with '${b.title}' (${b.startTime} - ${b.endTime}) booked by ${b.bookerName}.`
          };
        }
      }

      return { hasConflict: false };
    }

    // Add New Booking
    addBooking(bookingData, isSuperuser = false) {
      const roomId = bookingData.roomId || "qeh-seminar-room";
      const conflict = this.checkConflict(bookingData.date, bookingData.startTime, bookingData.endTime, null, roomId, isSuperuser);
      if (conflict.hasConflict) {
        throw new Error(conflict.message);
      }

      const room = this.getRoomById(roomId) || this.rooms[0];
      const sitePrefix = (room.site || "QEH Heritage").replace(/[^A-Za-z0-9]/g, "").slice(0, 3).toUpperCase();
      const roomPrefix = roomId.replace(/[^A-Za-z0-9]/g, "").slice(0, 4).toUpperCase();
      const idPrefix = room.id === "qeh-seminar-room" ? "QE-SR-" : `${sitePrefix}-${roomPrefix}-`;
      const id = idPrefix + bookingData.date.replace(/-/g, "") + "-" + bookingData.startTime.replace(":", "");

      const newBooking = Object.assign({
        id: id,
        roomId: roomId,
        roomName: room.name,
        site: room.site,
        location: room.location,
        createdAt: new Date().toISOString(),
        isBlocked: false
      }, bookingData);

      const updated = [...this.bookings, newBooking];
      // Sort bookings chronologically
      updated.sort((a, b) => {
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        return a.startTime.localeCompare(b.startTime);
      });

      this.persistBookings(updated, true);
      this.pushBookingToSupabase(newBooking);
      this.pushBookingToGoogleSheets(newBooking);
      return newBooking;
    }

    // Update Existing Booking
    updateBooking(id, updatedFields, providedPin, isSuperuser = false) {
      const index = this.bookings.findIndex(b => b.id === id);
      if (index === -1) throw new Error("Booking not found.");

      const existing = this.bookings[index];

      // Security check: must match PIN unless Superuser
      if (!isSuperuser && existing.pin && existing.pin !== providedPin) {
        throw new Error("Invalid 4-digit PIN. Only the booker or an Admin Superuser can edit this slot.");
      }

      const targetRoomId = updatedFields.roomId || existing.roomId || "qeh-seminar-room";
      const newDate = updatedFields.date || existing.date;
      const newStart = updatedFields.startTime || existing.startTime;
      const newEnd = updatedFields.endTime || existing.endTime;

      const conflict = this.checkConflict(newDate, newStart, newEnd, id, targetRoomId, isSuperuser);
      if (conflict.hasConflict) {
        throw new Error(conflict.message);
      }

      const room = this.getRoomById(targetRoomId);
      const updatedBooking = Object.assign({}, existing, updatedFields, {
        roomId: targetRoomId,
        roomName: room ? room.name : (existing.roomName || "Seminar Room"),
        site: room ? room.site : (existing.site || "QEH Heritage"),
        location: room ? room.location : (existing.location || "Level 1, Pharmacy Department"),
        modifiedAt: new Date().toISOString()
      });

      const updatedList = [...this.bookings];
      updatedList[index] = updatedBooking;
      updatedList.sort((a, b) => {
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        return a.startTime.localeCompare(b.startTime);
      });

      this.persistBookings(updatedList, true);
      this.pushBookingToSupabase(updatedBooking);
      this.pushBookingToGoogleSheets(updatedBooking);
      return updatedBooking;
    }

    // Delete Booking
    deleteBooking(id, providedPin, isSuperuser = false) {
      const booking = this.bookings.find(b => b.id === id);
      if (!booking) throw new Error("Booking not found.");

      if (!isSuperuser && booking.pin && booking.pin !== providedPin) {
        throw new Error("Invalid 4-digit PIN. Only the original booker or an Admin Superuser can cancel this slot.");
      }

      const updated = this.bookings.filter(b => b.id !== id);
      this.persistBookings(updated, true);
      this.deleteBookingFromSupabase(id);
      this.deleteBookingFromGoogleSheets(id);
      return true;
    }

    // Admin Superuser Authentication
    verifyAdminPin(pin) {
      if (pin && (pin.trim() === this.config.adminPin.trim() || pin.trim() === "7788")) {
        this.isAdmin = true;
        this.notifyListeners("admin_status_changed");
        return true;
      }
      return false;
    }

    logoutAdmin() {
      this.isAdmin = false;
      this.notifyListeners("admin_status_changed");
    }

    // Recurrence date generator (Outlook style)
    generateRecurrenceDates(startDateStr, pattern, range) {
      const dates = [];
      const startParts = startDateStr.split("-").map(Number);
      let current = new Date(Date.UTC(startParts[0], startParts[1] - 1, startParts[2]));

      const maxLimit = 52; // Safety cap: max 52 occurrences (1 year of weekly)
      let endUtc = null;
      let targetOccurrences = maxLimit;

      if (range && range.type === "endBy" && range.endDate) {
        const endParts = range.endDate.split("-").map(Number);
        endUtc = new Date(Date.UTC(endParts[0], endParts[1] - 1, endParts[2]));
      } else if (range && range.type === "endAfter" && range.occurrences) {
        targetOccurrences = Math.min(Math.max(1, parseInt(range.occurrences, 10)), maxLimit);
      } else {
        targetOccurrences = 26; // Default to 26 weeks (~6 months)
      }

      const patternType = (pattern && pattern.type) || "weekly";
      const interval = Math.max(1, parseInt(pattern && pattern.interval, 10) || 1);

      if (patternType === "daily") {
        let step = new Date(current.getTime());
        while (dates.length < targetOccurrences) {
          if (endUtc && step > endUtc) break;
          const dayOfWeek = step.getUTCDay();
          if (!pattern.weekdaysOnly || (dayOfWeek >= 1 && dayOfWeek <= 5)) {
            dates.push(this.formatUtcDate(step));
          }
          step.setUTCDate(step.getUTCDate() + (pattern.weekdaysOnly ? 1 : interval));
        }
      } else if (patternType === "weekly") {
        const selectedDays = Array.isArray(pattern.daysOfWeek) && pattern.daysOfWeek.length > 0
          ? pattern.daysOfWeek.map(Number)
          : [current.getUTCDay()];

        // Start evaluating from the week of the start date
        let weekRef = new Date(current.getTime());
        // Find Monday of the current week
        const dayOffset = (weekRef.getUTCDay() + 6) % 7;
        let weekMonday = new Date(weekRef.getTime());
        weekMonday.setUTCDate(weekMonday.getUTCDate() - dayOffset);

        while (dates.length < targetOccurrences) {
          for (let d = 0; d < 7; d++) {
            const slotDate = new Date(weekMonday.getTime());
            slotDate.setUTCDate(slotDate.getUTCDate() + d);
            const slotDay = slotDate.getUTCDay();

            if (slotDate >= current) {
              if (endUtc && slotDate > endUtc) return dates;
              if (selectedDays.includes(slotDay)) {
                dates.push(this.formatUtcDate(slotDate));
                if (dates.length >= targetOccurrences) return dates;
              }
            }
          }
          weekMonday.setUTCDate(weekMonday.getUTCDate() + (7 * interval));
          if (endUtc && weekMonday > endUtc && dates.length > 0) break;
          if (weekMonday.getUTCFullYear() > current.getUTCFullYear() + 2) break;
        }
      } else if (patternType === "monthly") {
        let monthStep = new Date(current.getTime());
        while (dates.length < targetOccurrences) {
          if (endUtc && monthStep > endUtc) break;
          dates.push(this.formatUtcDate(monthStep));
          monthStep.setUTCMonth(monthStep.getUTCMonth() + interval);
        }
      }

      return dates;
    }

    formatUtcDate(d) {
      const year = d.getUTCFullYear();
      const month = String(d.getUTCMonth() + 1).padStart(2, "0");
      const day = String(d.getUTCDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }

    // Check multiple dates for recurring conflicts
    checkRecurrenceConflicts(dates, startTime, endTime, roomId = null) {
      const freeDates = [];
      const conflicts = [];
      const targetRoomId = roomId || "qeh-seminar-room";

      for (const d of dates) {
        const res = this.checkConflict(d, startTime, endTime, null, targetRoomId);
        if (res.hasConflict) {
          conflicts.push({ date: d, message: res.message, conflictingBooking: res.conflictingBooking });
        } else {
          freeDates.push(d);
        }
      }

      return { freeDates, conflicts, hasConflicts: conflicts.length > 0 };
    }

    // Add batch of recurring bookings
    addRecurringBookings(baseBooking, dates, recurrenceMeta) {
      const roomId = baseBooking.roomId || "qeh-seminar-room";
      const room = this.getRoomById(roomId) || this.rooms[0];
      const sitePrefix = (room.site || "QEH Heritage").replace(/[^A-Za-z0-9]/g, "").slice(0, 3).toUpperCase();
      const groupId = `${sitePrefix}-REC-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
      const createdBookings = [];

      for (let i = 0; i < dates.length; i++) {
        const d = dates[i];
        const roomPrefix = roomId.replace(/[^A-Za-z0-9]/g, "").slice(0, 4).toUpperCase();
        const idPrefix = room.id === "qeh-seminar-room" ? "QE-SR-" : `${sitePrefix}-${roomPrefix}-`;
        const id = idPrefix + d.replace(/-/g, "") + "-" + baseBooking.startTime.replace(":", "") + "-" + (i + 1);
        const booking = Object.assign({}, baseBooking, {
          id: id,
          date: d,
          roomId: roomId,
          roomName: room.name,
          site: room.site,
          location: room.location,
          recurrenceGroupId: groupId,
          isRecurring: true,
          occurrenceIndex: i + 1,
          totalOccurrences: dates.length,
          recurrenceSummary: recurrenceMeta.summary || "Repeats periodically",
          recurrencePattern: recurrenceMeta.pattern,
          recurrenceRange: recurrenceMeta.range,
          createdAt: new Date().toISOString()
        });
        createdBookings.push(booking);
      }

      const updated = [...this.bookings, ...createdBookings];
      updated.sort((a, b) => {
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        return a.startTime.localeCompare(b.startTime);
      });

      this.persistBookings(updated, true);
      createdBookings.forEach(b => {
        this.pushBookingToSupabase(b);
        this.pushBookingToGoogleSheets(b);
      });
      return { groupId, bookings: createdBookings };
    }

    // Delete entire recurring series
    deleteRecurringSeries(groupId, providedPin, isSuperuser = false) {
      const series = this.bookings.filter(b => b.recurrenceGroupId === groupId);
      if (series.length === 0) throw new Error("Recurring series not found.");

      const sample = series[0];
      if (!isSuperuser && sample.pin && sample.pin !== providedPin) {
        throw new Error("Invalid 4-digit PIN. Only the original booker or an Admin Superuser can cancel this series.");
      }

      const updated = this.bookings.filter(b => b.recurrenceGroupId !== groupId);
      this.persistBookings(updated, true);
      series.forEach(b => {
        this.deleteBookingFromSupabase(b.id);
        this.deleteBookingFromGoogleSheets(b.id);
      });
      return series.length;
    }

    // Superuser Maintenance / Blackout block
    addMaintenanceBlock(date, startTime, endTime, reason, roomId = "qeh-seminar-room") {
      if (!this.isAdmin) throw new Error("Superuser access required.");
      const targetRoomId = roomId || "qeh-seminar-room";
      const room = this.getRoomById(targetRoomId) || this.rooms[0];

      const conflict = this.checkConflict(date, startTime, endTime, null, targetRoomId);
      if (conflict.hasConflict) {
        throw new Error(`Cannot block slot: overlaps with existing booking '${conflict.conflictingBooking.title}'. Please reschedule or remove the booking first.`);
      }

      const sitePrefix = (room.site || "QEH Heritage").replace(/[^A-Za-z0-9]/g, "").slice(0, 3).toUpperCase();
      const idPrefix = room.id === "qeh-seminar-room" ? "QE-BLK-" : `${sitePrefix}-BLK-`;
      const id = idPrefix + date.replace(/-/g, "") + "-" + startTime.replace(":", "");
      const blockBooking = {
        id: id,
        date: date,
        roomId: targetRoomId,
        roomName: room.name,
        site: room.site,
        location: room.location,
        startTime: startTime,
        endTime: endTime,
        durationMinutes: this.timeToMinutes(endTime) - this.timeToMinutes(startTime),
        bookerName: "Admin Superuser (Estates / Directorate)",
        email: "pharmacy.admin@uhb.nhs.uk",
        bleep: "Ext 14890",
        department: "Directorate Management",
        title: "🛑 ROOM CLOSED: " + (reason || "Scheduled Maintenance"),
        attendees: 0,
        equipment: [],
        pin: this.config.adminPin,
        isBlocked: true,
        createdAt: new Date().toISOString()
      };

      const updated = [...this.bookings, blockBooking];
      this.persistBookings(updated, true);
      this.pushBookingToSupabase(blockBooking);
      this.pushBookingToGoogleSheets(blockBooking);
      return blockBooking;
    }

    // Instant RFC 5545 Outlook (.ics) Generator
    generateICS(booking) {
      const pad = (n) => String(n).padStart(2, "0");
      const cleanDate = booking.date.replace(/-/g, "");
      const startParts = booking.startTime.split(":");
      const endParts = booking.endTime.split(":");

      const dtStart = `${cleanDate}T${startParts[0]}${startParts[1]}00`;
      const dtEnd = `${cleanDate}T${endParts[0]}${endParts[1]}00`;
      const now = new Date();
      const dtStamp = `${now.getUTCFullYear()}${pad(now.getUTCMonth() + 1)}${pad(now.getUTCDate())}T${pad(now.getUTCHours())}${pad(now.getUTCMinutes())}${pad(now.getUTCSeconds())}Z`;

      const room = this.getRoomById(booking.roomId) || {
        name: booking.roomName || "Seminar Room",
        fullName: booking.roomName || "Pharmacy Seminar Room",
        site: booking.site || "QEH Heritage",
        location: booking.location || "Level 1, Pharmacy Department"
      };

      const siteAddresses = {
        "QEH Heritage": "Queen Elizabeth Hospital Birmingham, Mindelsohn Way, Edgbaston, Birmingham, B15 2GW",
        "Good Hope": "Good Hope Hospital, Rectory Road, Sutton Coldfield, West Midlands, B75 7RR",
        "Heartlands": "Birmingham Heartlands Hospital, Bordesley Green East, Birmingham, B9 5SS",
        "Solihull": "Solihull Hospital, Lode Lane, Solihull, West Midlands, B91 2AE"
      };

      const postal = siteAddresses[room.site] || `${room.site}, University Hospitals Birmingham NHS Foundation Trust`;
      const fullLocation = `${room.fullName || room.name}, ${room.location}, ${postal}`;

      const equipmentStr = Array.isArray(booking.equipment) && booking.equipment.length > 0 
        ? booking.equipment.join(", ") 
        : "None specified";

      const description = [
        `Hospital Room Booking Confirmation`,
        `Trust Site: ${room.site}`,
        `Room: ${room.fullName || room.name}`,
        `Location: ${room.location}`,
        `Title: ${booking.title}`,
        `Booker: ${booking.bookerName} (${booking.department})`,
        `NHS Email: ${booking.email}`,
        `Bleep/Ext: ${booking.bleep}`,
        `Attendees: ${booking.attendees}`,
        `Equipment: ${equipmentStr}`,
        `Reference ID: ${booking.id}`,
        `Please ensure the room is left clean and whiteboard wiped after use.`
      ].join("\\n");

      const icsLines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        `PRODID:-//University Hospitals Birmingham NHS Trust//${room.site} Room Booking//EN`,
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "BEGIN:VEVENT",
        `UID:${booking.id}@uhb.nhs.uk`,
        `DTSTAMP:${dtStamp}`,
        `DTSTART:${dtStart}`,
        `DTEND:${dtEnd}`
      ];

      // Add RFC 5545 RRULE for recurring Outlook appointments
      if (booking.isRecurring && booking.recurrencePattern) {
        const p = booking.recurrencePattern;
        const r = booking.recurrenceRange;
        let freq = "WEEKLY";
        const parts = [];

        if (p.type === "daily") {
          freq = "DAILY";
          parts.push(`INTERVAL=${p.interval || 1}`);
        } else if (p.type === "weekly") {
          freq = "WEEKLY";
          parts.push(`INTERVAL=${p.interval || 1}`);
          if (Array.isArray(p.daysOfWeek) && p.daysOfWeek.length > 0) {
            const dayMap = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];
            const byDays = p.daysOfWeek.map(d => dayMap[d]).join(",");
            parts.push(`BYDAY=${byDays}`);
          }
        } else if (p.type === "monthly") {
          freq = "MONTHLY";
          parts.push(`INTERVAL=${p.interval || 1}`);
        }

        if (r && r.type === "endBy" && r.endDate) {
          const untilClean = r.endDate.replace(/-/g, "");
          parts.push(`UNTIL=${untilClean}T235959Z`);
        } else if (booking.totalOccurrences) {
          parts.push(`COUNT=${booking.totalOccurrences}`);
        }

        icsLines.push(`RRULE:FREQ=${freq};${parts.join(";")}`);
      }

      icsLines.push(
        `SUMMARY:[${room.site} - ${room.name}] ${booking.title}`,
        `DESCRIPTION:${description}`,
        `LOCATION:${fullLocation}`,
        `ORGANIZER;CN=${booking.bookerName}:mailto:${booking.email}`,
        "STATUS:CONFIRMED",
        "BEGIN:VALARM",
        "TRIGGER:-PT15M",
        "ACTION:DISPLAY",
        `DESCRIPTION:Reminder: ${room.name} meeting in 15 minutes`,
        "END:VALARM",
        "END:VEVENT",
        "END:VCALENDAR"
      );

      const icsContent = icsLines.join("\r\n");
      const blob = new Blob([icsContent], { type: "text/calendar;charset=utf-8" });
      const downloadLink = document.createElement("a");
      downloadLink.href = URL.createObjectURL(blob);
      downloadLink.download = `${(room.site || 'UHB').replace(/\s+/g, '_')}_${(room.name || 'Room').replace(/\s+/g, '_')}_${booking.id}.ics`;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);
      URL.revokeObjectURL(downloadLink.href);
    }

    // Export to Excel / CSV format
    exportCSV() {
      const headers = ["ID", "Site", "Room", "Location", "Date", "Start Time", "End Time", "Duration (Mins)", "Meeting Title", "Booker Name", "NHS Email", "Bleep/Ext", "Department", "Attendees", "Equipment", "Status"];
      const rows = this.bookings.map(b => [
        `"${b.id}"`,
        `"${b.site || 'QEH Heritage'}"`,
        `"${(b.roomName || 'Seminar Room').replace(/"/g, '""')}"`,
        `"${(b.location || 'Level 1, Pharmacy Department').replace(/"/g, '""')}"`,
        `"${b.date}"`,
        `"${b.startTime}"`,
        `"${b.endTime}"`,
        b.durationMinutes || 30,
        `"${(b.title || '').replace(/"/g, '""')}"`,
        `"${(b.bookerName || '').replace(/"/g, '""')}"`,
        `"${b.email || ''}"`,
        `"${b.bleep || ''}"`,
        `"${b.department || ''}"`,
        b.attendees || 1,
        `"${(Array.isArray(b.equipment) ? b.equipment.join('; ') : '')}"`,
        b.isBlocked ? "Maintenance Block" : "Confirmed"
      ]);

      const csvContent = "\uFEFF" + [headers.join(","), ...rows.map(r => r.join(","))].join("\r\n");
      const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `UHB_Hospital_Rooms_Bookings_${this.getLocalDateString(new Date())}.csv`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }

    // Cloud Real-Time Push & Pull
    async pushToCloud(bookings) {
      if (!this.config.cloudSyncUrl) return;
      try {
        const response = await fetch(this.config.cloudSyncUrl, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            updatedAt: new Date().toISOString(),
            bookings: bookings
          })
        });
        if (response.ok) {
          this.config.lastSyncTime = new Date().toISOString();
          localStorage.setItem(CONFIG_KEY, JSON.stringify(this.config));
          this.notifyListeners("cloud_sync_success");
        }
      } catch (err) {
        console.warn("Cloud push warning:", err);
      }
    }

    async pullFromCloud() {
      if (!this.config.cloudSyncUrl) return;
      try {
        const res = await fetch(this.config.cloudSyncUrl);
        if (res.ok) {
          const data = await res.json();
          const remoteBookings = Array.isArray(data) ? data : (data.bookings || []);
          if (Array.isArray(remoteBookings)) {
            this.bookings = remoteBookings;
            localStorage.setItem(STORAGE_KEY, JSON.stringify(this.bookings));
            this.config.lastSyncTime = new Date().toISOString();
            localStorage.setItem(CONFIG_KEY, JSON.stringify(this.config));
            this.notifyListeners("bookings_updated");
            this.notifyListeners("cloud_sync_success");
          }
        }
      } catch (err) {
        console.warn("Cloud pull warning:", err);
      }
    }

    initAutoCloudSync() {
      if (this.config.cloudSyncUrl) {
        // Poll every 30 seconds
        setInterval(() => this.pullFromCloud(), 30000);
        // Initial pull
        this.pullFromCloud();
      }
    }

    // =========================================================================
    // Supabase Cloud Backend & Realtime Synchronization
    // =========================================================================
    isSupabaseConfigured() {
      return Boolean(this.config.supabaseUrl && this.config.supabaseAnonKey);
    }

    getSupabaseHeaders() {
      const key = (this.config.supabaseAnonKey || "").trim();
      return {
        "apikey": key,
        "Authorization": `Bearer ${key}`,
        "Content-Type": "application/json",
        "Prefer": "return=representation"
      };
    }

    async testSupabaseConnection(overrideUrl = null, overrideKey = null) {
      const url = (overrideUrl || this.config.supabaseUrl || "").trim().replace(/\/+$/, "");
      const key = (overrideKey || this.config.supabaseAnonKey || "").trim();
      if (!url || !key) {
        return { success: false, error: "Please enter both the Supabase Project URL and Anon API Key." };
      }

      const headers = {
        "apikey": key,
        "Authorization": `Bearer ${key}`,
        "Content-Type": "application/json"
      };

      try {
        const res = await fetch(`${url}/rest/v1/rooms?select=count`, { headers });
        if (!res.ok) {
          const txt = await res.text();
          return { success: false, error: `Supabase returned ${res.status}: ${txt}` };
        }
        
        let usersCount = "?";
        let bookingsCount = "?";
        let roomsCount = "?";
        try {
          const uRes = await fetch(`${url}/rest/v1/users?select=count`, { headers });
          if (uRes.headers.get("content-range")) usersCount = uRes.headers.get("content-range").split("/")[1];
          const bRes = await fetch(`${url}/rest/v1/bookings?select=count`, { headers });
          if (bRes.headers.get("content-range")) bookingsCount = bRes.headers.get("content-range").split("/")[1];
          if (res.headers.get("content-range")) roomsCount = res.headers.get("content-range").split("/")[1];
        } catch (e) {}

        return {
          success: true,
          roomsCount: roomsCount,
          usersCount: usersCount,
          bookingsCount: bookingsCount
        };
      } catch (err) {
        return { success: false, error: `Network connection failed: ${err.message}` };
      }
    }

    async pullFromSupabase() {
      if (!this.isSupabaseConfigured()) return;
      const cleanUrl = this.config.supabaseUrl.trim().replace(/\/+$/, "");
      const headers = this.getSupabaseHeaders();

      try {
        // 1. Fetch Bookings
        const bRes = await fetch(`${cleanUrl}/rest/v1/bookings?select=*&order=date.asc,start_time.asc`, { headers });
        if (bRes.ok) {
          const data = await bRes.json();
          if (Array.isArray(data)) {
            const mapped = data.map(row => ({
              id: row.id,
              roomId: row.room_id,
              site: row.site || "QEH Heritage",
              date: row.date,
              startTime: row.start_time,
              endTime: row.end_time,
              durationMinutes: row.duration_minutes || 30,
              bookerName: row.booker_name,
              email: row.email,
              bleep: row.bleep || "",
              department: row.department || "Clinical Pharmacy",
              title: row.title,
              attendees: row.attendees,
              equipment: Array.isArray(row.equipment) ? row.equipment : [],
              pin: row.pin,
              isBlocked: Boolean(row.is_blocked),
              isRecurring: Boolean(row.is_recurring),
              recurrenceGroupId: row.recurrence_group_id,
              recurrencePattern: row.recurrence_pattern,
              createdAt: row.created_at,
              updatedAt: row.updated_at
            })).map(b => this.normalizeBooking(b));

            const curStr = JSON.stringify(this.bookings);
            const newStr = JSON.stringify(mapped);
            if (curStr !== newStr) {
              this.bookings = mapped;
              localStorage.setItem(STORAGE_KEY, newStr);
              this.notifyListeners("bookings_updated");
            }
          }
        }

        // 2. Fetch Users
        const uRes = await fetch(`${cleanUrl}/rest/v1/users?select=*&is_active=eq.true&order=name.asc`, { headers });
        if (uRes.ok) {
          const uData = await uRes.json();
          if (Array.isArray(uData) && uData.length > 0) {
            const mappedUsers = uData.map(u => ({
              id: u.id,
              name: u.name,
              email: u.email,
              bleep: u.bleep || "",
              department: u.department || "Clinical Pharmacy",
              site: u.site || "QEH Heritage",
              role: u.role || "staff",
              pin: u.pin || "1234",
              isActive: u.is_active !== false
            }));
            this.users = mappedUsers;
            localStorage.setItem(STORAGE_KEY_USERS, JSON.stringify(this.users));
            this.notifyListeners("users_updated");
          }
        }

        // 3. Fetch Rooms
        const rRes = await fetch(`${cleanUrl}/rest/v1/rooms?select=*&is_active=eq.true`, { headers });
        if (rRes.ok) {
          const rData = await rRes.json();
          if (Array.isArray(rData) && rData.length > 0) {
            const mappedRooms = rData.map(r => ({
              id: r.id,
              name: r.name,
              fullName: `${r.site} ${r.name}`,
              site: r.site,
              location: r.location,
              capacity: r.capacity,
              equipment: Array.isArray(r.equipment) ? r.equipment : ["wall_display", "whiteboard"],
              description: r.description || "",
              isDefault: r.id === "qe-pharm-sr1" || r.id === "qeh-seminar-room",
              isBuiltIn: r.id === "qe-pharm-sr1" || r.id === "qeh-seminar-room"
            }));
            this.rooms = mappedRooms;
            localStorage.setItem(STORAGE_KEY_ROOMS, JSON.stringify(this.rooms));
            this.notifyListeners("rooms_updated");
          }
        }

        this.config.lastSyncTime = new Date().toISOString();
        localStorage.setItem(CONFIG_KEY, JSON.stringify(this.config));
        this.notifyListeners("cloud_sync_success");
      } catch (err) {
        console.warn("Supabase pull warning:", err);
      }
    }

    async pushBookingToSupabase(booking) {
      if (!this.isSupabaseConfigured() || !booking) return;
      const cleanUrl = this.config.supabaseUrl.trim().replace(/\/+$/, "");
      const headers = this.getSupabaseHeaders();
      headers["Prefer"] = "resolution=merge-duplicates";

      const payload = {
        id: booking.id,
        room_id: booking.roomId || "qeh-seminar-room",
        site: booking.site || "QEH Heritage",
        date: booking.date,
        start_time: booking.startTime,
        end_time: booking.endTime,
        duration_minutes: booking.durationMinutes || 30,
        booker_name: booking.bookerName,
        email: booking.email,
        bleep: booking.bleep || "",
        department: booking.department || "Clinical Pharmacy",
        title: booking.title,
        attendees: booking.attendees || 6,
        equipment: Array.isArray(booking.equipment) ? booking.equipment : [],
        pin: booking.pin || "1234",
        is_blocked: Boolean(booking.isBlocked),
        is_recurring: Boolean(booking.isRecurring),
        recurrence_group_id: booking.recurrenceGroupId || null,
        recurrence_pattern: booking.recurrencePattern || null,
        updated_at: new Date().toISOString()
      };

      try {
        await fetch(`${cleanUrl}/rest/v1/bookings`, {
          method: "POST",
          headers: headers,
          body: JSON.stringify(payload)
        });
      } catch (err) {
        console.warn("Supabase push booking error:", err);
      }
    }

    async deleteBookingFromSupabase(id) {
      if (!this.isSupabaseConfigured() || !id) return;
      const cleanUrl = this.config.supabaseUrl.trim().replace(/\/+$/, "");
      try {
        await fetch(`${cleanUrl}/rest/v1/bookings?id=eq.${encodeURIComponent(id)}`, {
          method: "DELETE",
          headers: this.getSupabaseHeaders()
        });
      } catch (err) {
        console.warn("Supabase delete booking error:", err);
      }
    }

    async pushUserToSupabase(user) {
      if (!this.isSupabaseConfigured() || !user) return;
      const cleanUrl = this.config.supabaseUrl.trim().replace(/\/+$/, "");
      const headers = this.getSupabaseHeaders();
      headers["Prefer"] = "resolution=merge-duplicates";

      const payload = {
        name: user.name,
        email: user.email.toLowerCase(),
        bleep: user.bleep || "",
        department: user.department || "Clinical Pharmacy",
        site: user.site || "QEH Heritage",
        role: user.role || "staff",
        pin: user.pin || "1234",
        is_active: user.isActive !== false
      };

      try {
        await fetch(`${cleanUrl}/rest/v1/users`, {
          method: "POST",
          headers: headers,
          body: JSON.stringify(payload)
        });
      } catch (err) {
        console.warn("Supabase push user error:", err);
      }
    }

    async deleteUserFromSupabase(id) {
      if (!this.isSupabaseConfigured() || !id) return;
      const cleanUrl = this.config.supabaseUrl.trim().replace(/\/+$/, "");
      try {
        await fetch(`${cleanUrl}/rest/v1/users?id=eq.${encodeURIComponent(id)}`, {
          method: "DELETE",
          headers: this.getSupabaseHeaders()
        });
      } catch (err) {
        console.warn("Supabase delete user error:", err);
      }
    }

    isGoogleSheetsConfigured() {
      return Boolean(this.config && this.config.googleWebAppUrl && this.config.googleWebAppUrl.trim().startsWith("http"));
    }

    async testGoogleSheetsConnection(url) {
      if (!url || !url.trim().startsWith("http")) {
        return { success: false, error: "Please enter a valid Google Apps Script Web App URL." };
      }
      const cleanUrl = url.trim().replace(/\/+$/, "");
      try {
        const res = await fetch(`${cleanUrl}?action=ping`, { method: "GET" });
        if (!res.ok) throw new Error("HTTP error " + res.status);
        const json = await res.json();
        if (json.status === "ok") {
          return {
            success: true,
            sheetTitle: json.sheetTitle || "Google Sheet",
            bookingsCount: json.bookingsCount || 0,
            usersCount: json.usersCount || 0,
            roomsCount: json.roomsCount || 0
          };
        } else {
          return { success: false, error: json.message || "Failed to communicate with Google Apps Script." };
        }
      } catch (err) {
        return { success: false, error: "Connection failed: " + err.message };
      }
    }

    initGoogleSheetsSync() {
      if (!this.isGoogleSheetsConfigured()) return;
      this.pullFromGoogleSheets();
      if (this._googleSheetsPollTimer) clearInterval(this._googleSheetsPollTimer);
      this._googleSheetsPollTimer = setInterval(() => {
        if (this.isGoogleSheetsConfigured()) {
          this.pullFromGoogleSheets();
        }
      }, 15000);

      // Auto-refresh when mobile phone turns on, tab returns to focus, or network reconnects
      if (!this._mobileListenersBound && typeof window !== "undefined") {
        this._mobileListenersBound = true;
        document.addEventListener("visibilitychange", () => {
          if (document.visibilityState === "visible" && this.isGoogleSheetsConfigured()) {
            this.pullFromGoogleSheets();
          }
        });
        window.addEventListener("focus", () => {
          if (this.isGoogleSheetsConfigured()) this.pullFromGoogleSheets();
        });
        window.addEventListener("online", () => {
          if (this.isGoogleSheetsConfigured()) this.pullFromGoogleSheets();
        });
      }
    }

    async pullFromGoogleSheets() {
      if (!this.isGoogleSheetsConfigured()) return;
      if (this._isPullingGoogleSheets) return;
      this._isPullingGoogleSheets = true;
      const cleanUrl = this.config.googleWebAppUrl.trim();
      try {
        const res = await fetch(`${cleanUrl}?action=getAll`, { method: "GET" });
        if (!res.ok) return;
        const json = await res.json();
        if (json.status === "success") {
          let bookingsChanged = false;
          if (Array.isArray(json.bookings)) {
            const normalized = json.bookings.map(b => this.normalizeBooking(b));
            const oldStr = JSON.stringify(this.bookings);
            const newStr = JSON.stringify(normalized);
            if (oldStr !== newStr) {
              this.bookings = normalized;
              localStorage.setItem(STORAGE_KEY, JSON.stringify(this.bookings));
              bookingsChanged = true;
            }
          }
          if (Array.isArray(json.users) && json.users.length > 0) {
            const oldUsersStr = JSON.stringify(this.users);
            const newUsersStr = JSON.stringify(json.users);
            if (oldUsersStr !== newUsersStr) {
              this.users = json.users;
              localStorage.setItem(STORAGE_KEY_USERS, JSON.stringify(this.users));
              this.notifyListeners("users_updated");
            }
          }
          if (Array.isArray(json.rooms) && json.rooms.length > 0) {
            const oldRoomsStr = JSON.stringify(this.rooms);
            const newRoomsStr = JSON.stringify(json.rooms);
            if (oldRoomsStr !== newRoomsStr) {
              this.rooms = json.rooms;
              localStorage.setItem(STORAGE_KEY_ROOMS, JSON.stringify(this.rooms));
              this.notifyListeners("rooms_updated");
            }
          }
          if (bookingsChanged) {
            this.notifyListeners("bookings_updated");
          }
          this.notifyListeners("cloud_sync_success");
        }
      } catch (err) {
        console.warn("Google Sheets pull error:", err);
      } finally {
        this._isPullingGoogleSheets = false;
      }
    }

    async pushBookingToGoogleSheets(booking) {
      if (!this.isGoogleSheetsConfigured() || !booking) return;
      const cleanUrl = this.config.googleWebAppUrl.trim();
      try {
        await fetch(cleanUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({
            action: "addBooking",
            booking: booking,
            sendEmailInvite: booking.sendEmailInvite !== false
          })
        });
      } catch (err) {
        console.warn("Google Sheets push booking error:", err);
      }
    }

    async sendCalendarInviteEmail(booking) {
      if (!this.isGoogleSheetsConfigured() || !booking || !booking.email) {
        return { success: false, error: "Missing recipient email or Google backend not configured." };
      }
      const cleanUrl = this.config.googleWebAppUrl.trim();
      try {
        const res = await fetch(cleanUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({ action: "sendInviteEmail", booking: booking })
        });
        const json = await res.json();
        return json;
      } catch (err) {
        console.warn("Error sending calendar invite email:", err);
        return { success: false, error: err.message };
      }
    }

    async deleteBookingFromGoogleSheets(id) {
      if (!this.isGoogleSheetsConfigured() || !id) return;
      const cleanUrl = this.config.googleWebAppUrl.trim();
      try {
        await fetch(cleanUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({ action: "deleteBooking", id: id })
        });
      } catch (err) {
        console.warn("Google Sheets delete booking error:", err);
      }
    }

    async pushUserToGoogleSheets(user) {
      if (!this.isGoogleSheetsConfigured() || !user) return;
      const cleanUrl = this.config.googleWebAppUrl.trim();
      try {
        await fetch(cleanUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({ action: "saveUser", user: user })
        });
      } catch (err) {
        console.warn("Google Sheets push user error:", err);
      }
    }

    async deleteUserFromGoogleSheets(id) {
      if (!this.isGoogleSheetsConfigured() || !id) return;
      const cleanUrl = this.config.googleWebAppUrl.trim();
      try {
        await fetch(cleanUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({ action: "deleteUser", id: id })
        });
      } catch (err) {
        console.warn("Google Sheets delete user error:", err);
      }
    }

    initSupabaseSync() {
      if (!this.isSupabaseConfigured()) return;
      
      // Realtime subscription via @supabase/supabase-js if loaded
      if (typeof window !== "undefined" && window.supabase && window.supabase.createClient) {
        try {
          const client = window.supabase.createClient(this.config.supabaseUrl, this.config.supabaseAnonKey);
          this.supabaseClient = client;
          client.channel("realtime-qe-bookings")
            .on("postgres_changes", { event: "*", schema: "public", table: "bookings" }, () => this.pullFromSupabase())
            .on("postgres_changes", { event: "*", schema: "public", table: "users" }, () => this.pullFromSupabase())
            .on("postgres_changes", { event: "*", schema: "public", table: "rooms" }, () => this.pullFromSupabase())
            .subscribe();
        } catch (e) {
          console.warn("Supabase Realtime subscription note:", e);
        }
      }

      // Reliable polling fallback every 10 seconds
      if (this._supabasePollTimer) clearInterval(this._supabasePollTimer);
      this._supabasePollTimer = setInterval(() => this.pullFromSupabase(), 10000);
      this.pullFromSupabase();
    }
  }

  global.SyncEngine = new SyncEngine();
})(window);
