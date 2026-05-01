const STORAGE_KEY = "hotel-dashboard-state";
const DEFAULT_SETTINGS_PASSWORD = "inexus12345";
const SITE_ACCESS_PASSWORD = "inexus12345";
const SITE_UNLOCK_KEY = "morishita-site-unlocked";

const state = {
  settings: {
    totalRooms: 6,
  },
  inventory: [],
  bookings: [],
  notes: [],
};

const roomCapacities = {
  "101": 3,
  "102": 4,
  "201": 3,
  "202": 4,
  "301": 3,
  "302": 4,
};

const cannedReplies = [
  {
    question: "Airbnbの掲載ページはありますか？",
    answer: "はい、Airbnbの掲載ページはこちらです。 https://www.airbnb.jp/rooms/1634547511797444361?unique_share_id=5d4f20cd-9444-4730-af29-d416ade8544b&viralityEntryPoint=1&s=76",
  },
  {
    question: "エレベーターはありますか？",
    answer: "エレベーターはありませんが、空き状況に応じて1階や2階のお部屋を優先してご案内できます。",
  },
  {
    question: "駐車場はありますか？",
    answer: "駐車場のご利用可否は日によって変わります。必要な場合は事前にご相談ください。",
  },
  {
    question: "チェックインは何時からですか？",
    answer: "チェックイン可能時間はご案内時にお伝えしています。早めの到着予定がある場合は事前にご連絡ください。",
  },
  {
    question: "遅い時間のチェックインはできますか？",
    answer: "到着予定時刻が分かれば、できる範囲で調整します。まずはご相談ください。",
  },
  {
    question: "連泊中の清掃はありますか？",
    answer: "清掃のタイミングは当日の運用状況に合わせてご案内しています。必要があれば事前にお知らせください。",
  },
];

const inventoryInput = document.getElementById("inventoryInput");
const loadInventoryBtn = document.getElementById("loadInventoryBtn");
const downloadInventoryBtn = document.getElementById("downloadInventoryBtn");
const downloadBookingsBtn = document.getElementById("downloadBookingsBtn");
const settingsForm = document.getElementById("settingsForm");
const totalRoomsInput = document.getElementById("totalRooms");
const settingsPassword = document.getElementById("settingsPassword");
const unlockSettingsBtn = document.getElementById("unlockSettingsBtn");
const settingsFeedback = document.getElementById("settingsFeedback");
const availabilityForm = document.getElementById("availabilityForm");
const bookingForm = document.getElementById("bookingForm");
const availabilityResult = document.getElementById("availabilityResult");
const bookingFeedback = document.getElementById("bookingFeedback");
const guestCount = document.getElementById("guestCount");
const bookRoomType = document.getElementById("bookRoomType");
const bookGuestCount = document.getElementById("bookGuestCount");
const bookingSource = document.getElementById("bookingSource");
const bookingAmount = document.getElementById("bookingAmount");
const paymentStatus = document.getElementById("paymentStatus");
const idVerified = document.getElementById("idVerified");
const checkInGuideSent = document.getElementById("checkInGuideSent");
const reviewRequestSent = document.getElementById("reviewRequestSent");
const bookingMemo = document.getElementById("bookingMemo");
const calendarGrid = document.getElementById("calendarGrid");
const calendarNoteForm = document.getElementById("calendarNoteForm");
const calendarNoteFeedback = document.getElementById("calendarNoteFeedback");
const noteDate = document.getElementById("noteDate");
const noteRoom = document.getElementById("noteRoom");
const noteText = document.getElementById("noteText");
const downloadCalendarPdfBtn = document.getElementById("downloadCalendarPdfBtn");
const bookingTableBody = document.getElementById("bookingTableBody");
const bookingSearchInput = document.getElementById("bookingSearchInput");
const bookingSourceFilter = document.getElementById("bookingSourceFilter");
const bookingRoomFilter = document.getElementById("bookingRoomFilter");
const bookingMonthFilter = document.getElementById("bookingMonthFilter");
const bookingPaymentFilter = document.getElementById("bookingPaymentFilter");
const bookingCleaningFilter = document.getElementById("bookingCleaningFilter");
const bookingSortSelect = document.getElementById("bookingSortSelect");
const bookingListSummary = document.getElementById("bookingListSummary");
const bookingToolbar = document.getElementById("bookingToolbar");
const todaySummary = document.getElementById("todaySummary");
const faqTemplates = document.getElementById("faqTemplates");
const siteNav = document.getElementById("siteNav");
const siteNavMenu = document.getElementById("siteNavMenu");
const mobileNavToggle = document.getElementById("mobileNavToggle");
const siteLockOverlay = document.getElementById("siteLockOverlay");
const siteLockForm = document.getElementById("siteLockForm");
const sitePasswordInput = document.getElementById("sitePasswordInput");
const siteLockFeedback = document.getElementById("siteLockFeedback");
const calendarYearSelect = document.getElementById("calendarYearSelect");
const calendarMonthSelect = document.getElementById("calendarMonthSelect");
const tabButtons = [...document.querySelectorAll(".tab-button")];
const tabPanels = [...document.querySelectorAll(".tab-panel")];
const todayStayingCount = document.getElementById("todayStayingCount");
const todayCheckInCount = document.getElementById("todayCheckInCount");
const todayCheckOutCount = document.getElementById("todayCheckOutCount");
const todayCleaningCount = document.getElementById("todayCleaningCount");
const todayStayingList = document.getElementById("todayStayingList");
const todayCheckInList = document.getElementById("todayCheckInList");
const todayCheckOutList = document.getElementById("todayCheckOutList");
const todayCleaningList = document.getElementById("todayCleaningList");
const topCheckInCount = document.getElementById("topCheckInCount");
const topCheckOutCount = document.getElementById("topCheckOutCount");
const topCleaningCount = document.getElementById("topCleaningCount");
const unpaidAlertCount = document.getElementById("unpaidAlertCount");
const uncleanedAlertCount = document.getElementById("uncleanedAlertCount");
const idPendingAlertCount = document.getElementById("idPendingAlertCount");
const guidePendingAlertCount = document.getElementById("guidePendingAlertCount");
const monthSelect = document.getElementById("monthSelect");
const monthOccupancy = document.getElementById("monthOccupancy");
const monthBookings = document.getElementById("monthBookings");
const monthRevenue = document.getElementById("monthRevenue");
const monthVacancyAverage = document.getElementById("monthVacancyAverage");
const monthCleaningCount = document.getElementById("monthCleaningCount");
const monthCheckInCount = document.getElementById("monthCheckInCount");
const monthCheckOutCount = document.getElementById("monthCheckOutCount");
const monthNarrative = document.getElementById("monthNarrative");
const compareMonthA = document.getElementById("compareMonthA");
const compareMonthB = document.getElementById("compareMonthB");
const compareHeaderA = document.getElementById("compareHeaderA");
const compareHeaderB = document.getElementById("compareHeaderB");
const comparisonNarrative = document.getElementById("comparisonNarrative");
const comparisonTableBody = document.getElementById("comparisonTableBody");
const backupJsonBtn = document.getElementById("backupJsonBtn");
const restoreJsonBtn = document.getElementById("restoreJsonBtn");
const restoreJsonInput = document.getElementById("restoreJsonInput");

let settingsUnlocked = false;

function safeLocalStorageGet() {
  try {
    return localStorage.getItem(STORAGE_KEY);
  } catch (error) {
    return null;
  }
}

function safeLocalStorageSet(payload) {
  try {
    localStorage.setItem(STORAGE_KEY, payload);
  } catch (error) {
    console.warn("localStorage unavailable", error);
  }
}

function parseInventory(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line, index) => {
      const parts = line.split(/\t|,/).map((part) => part.trim());
      if (parts.length < 3) {
        throw new Error(`${index + 1}行目の形式が正しくありません。`);
      }
      const [date, type, stockRaw] = parts;
      const stock = Number(stockRaw);
      if (!date || !type || Number.isNaN(stock)) {
        throw new Error(`${index + 1}行目の値を確認してください。`);
      }
      return { date, type, stock };
    });
}

function toInventoryTsv(items) {
  return items.map((item) => `${item.date}\t${item.type}\t${item.stock}`).join("\n");
}

function saveState() {
  safeLocalStorageSet(JSON.stringify(state));
}

function formatCurrency(amount) {
  const value = Number(amount) || 0;
  return `¥${value.toLocaleString("ja-JP")}`;
}

function normalizePaymentStatus(status) {
  const value = String(status || "").trim();
  if (!value || value === "一部入金") {
    return "未払い";
  }
  if (["未払い", "支払い済み", "現地払い", "返金済み"].includes(value)) {
    return value;
  }
  return "未払い";
}

function createBookingId() {
  return `booking-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
}

function ensureBookingIds() {
  let changed = false;
  state.bookings = state.bookings.map((booking) => {
    const nextBooking = { ...booking };

    if (!nextBooking.id) {
      nextBooking.id = createBookingId();
      changed = true;
    }

    if (typeof nextBooking.amount !== "number") {
      nextBooking.amount = Number(nextBooking.amount) || 0;
      changed = true;
    }

    const normalizedPaymentStatus = normalizePaymentStatus(nextBooking.paymentStatus);
    if (nextBooking.paymentStatus !== normalizedPaymentStatus) {
      nextBooking.paymentStatus = normalizedPaymentStatus;
      changed = true;
    }

    if (!nextBooking.cleaningStatus) {
      nextBooking.cleaningStatus = "未清掃";
      changed = true;
    }

    if (typeof nextBooking.cleaningMemo !== "string") {
      nextBooking.cleaningMemo = String(nextBooking.cleaningMemo || "");
      changed = true;
    }

    if (typeof nextBooking.idVerified !== "boolean") {
      nextBooking.idVerified = false;
      changed = true;
    }

    if (typeof nextBooking.checkInGuideSent !== "boolean") {
      nextBooking.checkInGuideSent = false;
      changed = true;
    }

    if (typeof nextBooking.reviewRequestSent !== "boolean") {
      nextBooking.reviewRequestSent = false;
      changed = true;
    }

    if (!nextBooking.source) {
      nextBooking.source = "直接予約";
      changed = true;
    }

    return nextBooking;
  });

  if (changed) {
    saveState();
  }
}

function getPreloadedInventory() {
  const inventory = window.MORISHITA_PRELOADED_DATA?.inventory;
  return Array.isArray(inventory) ? inventory : [];
}

function loadState() {
  const saved = safeLocalStorageGet();
  const preloaded = getPreloadedInventory();

  if (!saved) {
    state.inventory = preloaded.length ? preloaded : parseInventory(inventoryInput.value);
    state.bookings = [];
    inventoryInput.value = toInventoryTsv(state.inventory);
    saveState();
    return;
  }

  try {
    const parsed = JSON.parse(saved);
    state.settings = {
      totalRooms: parsed.settings?.totalRooms || 6,
    };
    state.inventory = preloaded.length ? preloaded : (parsed.inventory || []);
    state.bookings = Array.isArray(parsed.bookings) ? parsed.bookings : [];
    state.notes = Array.isArray(parsed.notes) ? parsed.notes : [];
  } catch (error) {
    state.inventory = preloaded.length ? preloaded : parseInventory(inventoryInput.value);
    state.bookings = [];
    state.notes = [];
  }

  totalRoomsInput.value = state.settings.totalRooms;
  inventoryInput.value = toInventoryTsv(state.inventory);
}

function getRoomTypes() {
  return [...new Set(state.inventory.map((item) => String(item.type)))].sort((a, b) => Number(a) - Number(b));
}

function getRoomCapacity(room) {
  return roomCapacities[String(room)] || 0;
}

function getNoteKey(date, room) {
  return `${date}|${room}`;
}

function getNoteMap() {
  return Object.fromEntries(state.notes.map((note) => [getNoteKey(note.date, note.room), note.text]));
}

function getMonthKeys() {
  return [...new Set(state.inventory.map((item) => item.date.slice(0, 7)))].sort();
}

function nightsBetween(startDate, endDate) {
  const nights = [];
  const cursor = new Date(startDate);
  const end = new Date(endDate);

  while (cursor < end) {
    nights.push(cursor.toISOString().slice(0, 10));
    cursor.setDate(cursor.getDate() + 1);
  }

  return nights;
}

function shiftDate(dateText, amount) {
  const date = new Date(dateText);
  date.setDate(date.getDate() + amount);
  return date.toISOString().slice(0, 10);
}

function formatJaDate(dateText) {
  const [year, month, day] = dateText.split("-");
  return `${Number(month)}/${Number(day)}`;
}

function formatMonthLabel(monthKey) {
  const [year, month] = monthKey.split("-");
  return `${year}年${Number(month)}月`;
}

function summarizeDateRange(dates) {
  if (!dates.length) {
    return "";
  }

  const sorted = [...dates].sort();
  const ranges = [];
  let start = sorted[0];
  let previous = sorted[0];

  for (let index = 1; index < sorted.length; index += 1) {
    const current = sorted[index];
    if (shiftDate(previous, 1) !== current) {
      ranges.push([start, previous]);
      start = current;
    }
    previous = current;
  }

  ranges.push([start, previous]);
  return ranges
    .map(([rangeStart, rangeEnd]) => rangeStart === rangeEnd
      ? formatJaDate(rangeStart)
      : `${formatJaDate(rangeStart)}〜${formatJaDate(rangeEnd)}`)
    .join(", ");
}

function availableStock(date, room) {
  const inventoryRow = state.inventory.find((item) => item.date === date && String(item.type) === String(room));
  if (!inventoryRow) {
    return null;
  }

  const bookedCount = state.bookings.filter((booking) => {
    if (String(booking.type) !== String(room)) {
      return false;
    }
    return nightsBetween(booking.checkIn, booking.checkOut).includes(date);
  }).length;

  return inventoryRow.stock - bookedCount;
}

function getTodayDate() {
  const actualToday = new Date().toISOString().slice(0, 10);
  if (state.inventory.some((item) => item.date === actualToday)) {
    return actualToday;
  }
  const dates = [...new Set(state.inventory.map((item) => item.date))].sort();
  return dates[0] || actualToday;
}

function getPreferredMonthKey() {
  return getTodayDate().slice(0, 7);
}

function getRoomSnapshot(date) {
  return getRoomTypes().map((room) => ({
    room,
    occupied: (availableStock(date, room) ?? 0) <= 0,
  }));
}

function inferMonthlyFlowStats(monthKey) {
  const monthDates = [...new Set(
    state.inventory
      .filter((item) => item.date.startsWith(monthKey))
      .map((item) => item.date)
  )].sort();

  let checkInCount = 0;
  let checkOutCount = 0;

  getRoomTypes().forEach((room) => {
    monthDates.forEach((date) => {
      const occupiedToday = (availableStock(date, room) ?? 0) <= 0;
      const occupiedYesterday = (availableStock(shiftDate(date, -1), room) ?? 0) <= 0;

      if (occupiedToday && !occupiedYesterday) {
        checkInCount += 1;
      }

      if (!occupiedToday && occupiedYesterday) {
        checkOutCount += 1;
      }
    });
  });

  return {
    checkInCount,
    checkOutCount,
  };
}

function getAvailableRooms(checkIn, checkOut, guestTotal) {
  const nights = nightsBetween(checkIn, checkOut);

  return getRoomTypes()
    .filter((room) => getRoomCapacity(room) >= guestTotal)
    .map((room) => {
      const unavailableDates = nights.filter((date) => {
        const stock = availableStock(date, room);
        return stock === null || stock <= 0;
      });

      return {
        room,
        available: unavailableDates.length === 0,
        unavailableDates,
      };
    })
    .sort((a, b) => Number(a.room) - Number(b.room));
}

function normalizeBookingSource(source) {
  const value = String(source || "").trim().toLowerCase();
  if (!value) {
    return "direct";
  }
  if (value.includes("airbnb") || value.includes("air bnb")) {
    return "airbnb";
  }
  if (value.includes("booking")) {
    return "booking";
  }
  if (value.includes("expedia")) {
    return "expedia";
  }
  if (value.includes("携程") || value.includes("ctrip") || value.includes("trip.com") || value.includes("tripcom")) {
    return "ctrip";
  }
  if (value.includes("agoda")) {
    return "agoda";
  }
  if (value.includes("rakuten") || value.includes("楽天")) {
    return "rakuten";
  }
  return "other";
}

function getBookingSourceMeta(source) {
  const normalized = normalizeBookingSource(source);
  const sourceMap = {
    airbnb: { label: "Airbnb", className: "source-airbnb" },
    booking: { label: "Booking.com", className: "source-booking" },
    expedia: { label: "Expedia", className: "source-expedia" },
    ctrip: { label: "携程", className: "source-ctrip" },
    agoda: { label: "Agoda", className: "source-agoda" },
    rakuten: { label: "楽天", className: "source-rakuten" },
    direct: { label: "直接", className: "source-direct" },
    other: { label: source || "その他", className: "source-other" },
  };
  return sourceMap[normalized] || sourceMap.other;
}

function findBookingByDateAndRoom(date, room) {
  return state.bookings.find((booking) => {
    if (String(booking.type) !== String(room)) {
      return false;
    }
    return nightsBetween(booking.checkIn, booking.checkOut).includes(date);
  }) || null;
}

function getBookingsForDate(date) {
  return state.bookings.filter((booking) => nightsBetween(booking.checkIn, booking.checkOut).includes(date));
}

function getBookingsCheckingInOn(date) {
  return state.bookings.filter((booking) => booking.checkIn === date);
}

function getBookingsCheckingOutOn(date) {
  return state.bookings.filter((booking) => booking.checkOut === date);
}

function getCleaningBookingsForDate(date) {
  return state.bookings.filter((booking) => booking.checkOut === date && booking.cleaningStatus !== "清掃済み");
}

function bookingsOverlap(left, right) {
  return left.checkIn < right.checkOut && left.checkOut > right.checkIn;
}

function findBookingConflicts(candidateBooking, ignoreBookingId = "") {
  return state.bookings.filter((booking) => {
    if (ignoreBookingId && booking.id === ignoreBookingId) {
      return false;
    }
    if (String(booking.type) !== String(candidateBooking.type)) {
      return false;
    }
    return bookingsOverlap(candidateBooking, booking);
  });
}

function getBookingTaskSummary(booking) {
  const pending = [];
  if (booking.paymentStatus !== "支払い済み") {
    pending.push("未払い");
  }
  if (booking.cleaningStatus !== "清掃済み") {
    pending.push("未清掃");
  }
  if (!booking.idVerified) {
    pending.push("身分証未確認");
  }
  if (!booking.checkInGuideSent) {
    pending.push("案内未送信");
  }
  if (!booking.reviewRequestSent) {
    pending.push("レビュー依頼未送信");
  }
  return pending.length ? pending.join(" / ") : "対応済み";
}

function getBookingFilterState() {
  return {
    query: (bookingSearchInput?.value || "").trim().toLowerCase(),
    source: bookingSourceFilter?.value || "",
    room: bookingRoomFilter?.value || "",
    month: bookingMonthFilter?.value || "",
    paymentStatus: bookingPaymentFilter?.value || "",
    cleaningStatus: bookingCleaningFilter?.value || "",
    sortKey: bookingSortSelect?.value || "checkin-asc",
  };
}

function sortBookings(bookings, sortKey) {
  const sorted = [...bookings];
  sorted.sort((left, right) => {
    switch (sortKey) {
      case "checkin-desc":
        return right.checkIn.localeCompare(left.checkIn);
      case "checkout-asc":
        return left.checkOut.localeCompare(right.checkOut);
      case "guest-asc":
        return String(left.guest || "").localeCompare(String(right.guest || ""), "ja");
      case "amount-desc":
        return (Number(right.amount) || 0) - (Number(left.amount) || 0);
      case "amount-asc":
        return (Number(left.amount) || 0) - (Number(right.amount) || 0);
      case "checkin-asc":
      default:
        return left.checkIn.localeCompare(right.checkIn);
    }
  });
  return sorted;
}

function getVisibleBookings() {
  const filters = getBookingFilterState();
  const filtered = state.bookings.filter((booking) => {
    const matchesQuery = !filters.query || [
      booking.guest,
      booking.source,
      booking.type,
      booking.memo,
      booking.paymentStatus,
      booking.cleaningStatus,
    ]
      .map((value) => String(value || "").toLowerCase())
      .some((value) => value.includes(filters.query));

    const matchesSource = !filters.source || normalizeBookingSource(booking.source) === normalizeBookingSource(filters.source);
    const matchesRoom = !filters.room || String(booking.type) === String(filters.room);
    const matchesMonth = !filters.month || booking.checkIn.startsWith(filters.month) || booking.checkOut.startsWith(filters.month);
    const matchesPayment = !filters.paymentStatus || booking.paymentStatus === filters.paymentStatus;
    const matchesCleaning = !filters.cleaningStatus || booking.cleaningStatus === filters.cleaningStatus;
    return matchesQuery && matchesSource && matchesRoom && matchesMonth && matchesPayment && matchesCleaning;
  });

  return sortBookings(filtered, filters.sortKey);
}

function checkAvailability(checkIn, checkOut, room, guestTotal) {
  const nights = nightsBetween(checkIn, checkOut);
  if (!nights.length) {
    return { ok: false, message: "チェックアウトはチェックインより後の日付にしてください。" };
  }

  const capacity = getRoomCapacity(room);
  if (guestTotal > capacity) {
    const alternatives = getAvailableRooms(checkIn, checkOut, guestTotal).filter((item) => item.available);
    const optionText = alternatives.length
      ? `候補は ${alternatives.map((item) => `${item.room}号室`).join("、")} です。`
      : "条件に合う他のお部屋はありません。";
    return {
      ok: false,
      message: `${room}号室は${guestTotal}人では利用できません。定員は${capacity}人です。${optionText}`,
    };
  }

  const unavailableDates = nights.filter((date) => {
    const stock = availableStock(date, room);
    return stock === null || stock <= 0;
  });

  if (unavailableDates.length) {
    const alternatives = getAvailableRooms(checkIn, checkOut, guestTotal)
      .filter((item) => String(item.room) !== String(room) && item.available);
    const optionText = alternatives.length
      ? `代わりに ${alternatives.map((item) => `${item.room}号室`).join("、")} が連続で空いています。`
      : "連続で空いている他のお部屋はありません。";
    return {
      ok: false,
      message: `${room}号室は ${summarizeDateRange(unavailableDates)} が埋まっています。${optionText}`,
    };
  }

  return {
    ok: true,
    message: `${room}号室は ${nights.length}泊分空いています。`,
  };
}

function findRoomCandidates(checkIn, checkOut, guestTotal) {
  const nights = nightsBetween(checkIn, checkOut);
  if (!nights.length) {
    return { ok: false, message: "チェックアウトはチェックインより後の日付にしてください。", rooms: [] };
  }

  const candidates = getAvailableRooms(checkIn, checkOut, guestTotal).filter((item) => item.available);
  if (!candidates.length) {
    return {
      ok: false,
      message: `${formatJaDate(checkIn)}〜${formatJaDate(shiftDate(checkOut, -1))} で連続して空いているお部屋がありません。`,
      rooms: [],
    };
  }

  return {
    ok: true,
    message: `候補は ${candidates.map((item) => `${item.room}号室`).join("、")} です。`,
    rooms: candidates,
  };
}

function download(filename, text) {
  const blob = new Blob([text], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function downloadJson(filename, payload) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function setResult(node, message, kind) {
  node.textContent = message;
  node.className = `result-card ${kind}`;
}

function setOptions(selectNode) {
  const rooms = getRoomTypes();
  selectNode.innerHTML = rooms
    .map((room) => `<option value="${room}">${room}号室（${getRoomCapacity(room)}人）</option>`)
    .join("");
}

function renderSelectors() {
  const previousRoomFilter = bookingRoomFilter?.value || "";
  const previousMonthFilter = bookingMonthFilter?.value || "";
  setOptions(bookRoomType);
  if (noteRoom) {
    setOptions(noteRoom);
  }
  if (bookingRoomFilter) {
    bookingRoomFilter.innerHTML = `<option value="">すべて</option>${getRoomTypes().map((room) => `<option value="${room}">${room}号室</option>`).join("")}`;
    if ([...bookingRoomFilter.options].some((option) => option.value === previousRoomFilter)) {
      bookingRoomFilter.value = previousRoomFilter;
    }
  }
  if (bookingMonthFilter) {
    const bookingMonths = [...new Set(state.bookings.flatMap((booking) => [booking.checkIn.slice(0, 7), booking.checkOut.slice(0, 7)]).filter(Boolean))].sort();
    bookingMonthFilter.innerHTML = `<option value="">すべて</option>${bookingMonths.map((monthKey) => `<option value="${monthKey}">${formatMonthLabel(monthKey)}</option>`).join("")}`;
    if ([...bookingMonthFilter.options].some((option) => option.value === previousMonthFilter)) {
      bookingMonthFilter.value = previousMonthFilter;
    }
  }
}

function getCalendarEventLabel(date, room) {
  const occupiedToday = (availableStock(date, room) ?? 0) <= 0;
  const occupiedYesterday = (availableStock(shiftDate(date, -1), room) ?? 0) <= 0;
  const occupiedTomorrow = (availableStock(shiftDate(date, 1), room) ?? 0) <= 0;

  if (!occupiedToday && occupiedYesterday) {
    return { text: "退去", className: "check-out" };
  }
  if (occupiedToday && !occupiedYesterday) {
    return { text: "入居", className: "check-in" };
  }
  if (occupiedToday && occupiedTomorrow) {
    return null;
  }
  return null;
}

function renderCalendarSelectors() {
  const years = [...new Set(state.inventory.map((item) => item.date.slice(0, 4)))].sort();
  const preferredMonthKey = getPreferredMonthKey();
  const preferredYear = preferredMonthKey.slice(0, 4);
  const selectedYear = years.includes(calendarYearSelect.value)
    ? calendarYearSelect.value
    : (years.includes(preferredYear) ? preferredYear : years[years.length - 1]);
  calendarYearSelect.innerHTML = years.map((year) => `<option value="${year}">${year}年</option>`).join("");
  if (selectedYear) {
    calendarYearSelect.value = selectedYear;
  }

  const months = [...new Set(
    state.inventory
      .map((item) => item.date)
      .filter((date) => date.startsWith(selectedYear))
      .map((date) => date.slice(5, 7))
  )].sort();

  const preferredMonth = preferredMonthKey.slice(5, 7);
  const selectedMonth = months.includes(calendarMonthSelect.value)
    ? calendarMonthSelect.value
    : (months.includes(preferredMonth) ? preferredMonth : months[months.length - 1]);
  calendarMonthSelect.innerHTML = months.map((month) => `<option value="${month}">${Number(month)}月</option>`).join("");
  if (selectedMonth) {
    calendarMonthSelect.value = selectedMonth;
  }
}

function renderCalendar() {
  const selectedYear = calendarYearSelect.value;
  const selectedMonth = calendarMonthSelect.value;
  const dates = [...new Set(
    state.inventory
      .map((item) => item.date)
      .filter((date) => {
        const [year, month] = date.split("-");
        return (!selectedYear || year === selectedYear) && (!selectedMonth || month === selectedMonth);
      })
  )].sort();

  if (!dates.length) {
    calendarGrid.innerHTML = `<div class="calendar-card">該当する運用状況がありません。</div>`;
    return;
  }

  const rooms = getRoomTypes();
  const today = getTodayDate();
  const noteMap = getNoteMap();
  const headerCells = dates.map((date) => `
    <th class="${date === today ? "is-today" : ""}">
      <span>${formatJaDate(date)}</span>
    </th>
  `).join("");

  const bodyRows = rooms.map((room) => {
    const cells = dates.map((date) => {
      const remaining = availableStock(date, room);
      const soldout = remaining !== null && remaining <= 0;
      const label = remaining === null ? "-" : soldout ? "埋" : "空";
      const className = remaining === null ? "unknown" : soldout ? "soldout" : "available";
      const event = getCalendarEventLabel(date, room);
      const note = noteMap[getNoteKey(date, room)];
      const booking = findBookingByDateAndRoom(date, room);
      const sourceMeta = booking ? getBookingSourceMeta(booking.source) : null;
      return `
        <td class="calendar-status ${className} ${date === today ? "is-today" : ""}">
          <div class="calendar-cell-body">
            <strong>${label}</strong>
            ${sourceMeta ? `<span class="source-badge ${sourceMeta.className}">${sourceMeta.label}</span>` : ""}
            ${event ? `<span class="calendar-event ${event.className}">${event.text}</span>` : ""}
            ${note ? `<span class="calendar-note">${note}</span>` : ""}
          </div>
        </td>
      `;
    }).join("");

    return `
      <tr>
        <th scope="row" class="room-label">
          <div>${room}号室</div>
          <small>${getRoomCapacity(room)}人部屋</small>
        </th>
        ${cells}
      </tr>
    `;
  }).join("");

  calendarGrid.innerHTML = `
    <article class="calendar-card calendar-overview">
      <div class="calendar-legend">
        <span class="legend-chip available">空室</span>
        <span class="legend-chip soldout">埋まり</span>
        <span class="legend-chip unknown">データなし</span>
        <span class="legend-chip source-airbnb">Airbnb</span>
        <span class="legend-chip source-booking">Booking.com</span>
        <span class="legend-chip source-expedia">Expedia</span>
        <span class="legend-chip source-ctrip">携程</span>
      </div>
      <div class="calendar-matrix-wrap">
        <table class="calendar-matrix">
          <thead>
            <tr>
              <th>部屋</th>
              ${headerCells}
            </tr>
          </thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </div>
    </article>
  `;
}

function renderBookings() {
  ensureBookingIds();

  if (bookingToolbar) {
    bookingToolbar.hidden = state.bookings.length === 0;
  }
  if (bookingListSummary) {
    bookingListSummary.hidden = state.bookings.length === 0;
    bookingListSummary.textContent = "";
  }

  if (!state.bookings.length) {
    bookingTableBody.innerHTML = `<tr><td colspan="13">まだ予約は登録されていません。</td></tr>`;
    return;
  }

  const visibleBookings = getVisibleBookings();

  if (bookingListSummary) {
    bookingListSummary.textContent = `${visibleBookings.length}件表示 / 全${state.bookings.length}件`;
  }

  if (!visibleBookings.length) {
    bookingTableBody.innerHTML = `<tr><td colspan="13">条件に合う予約はありません。</td></tr>`;
    return;
  }

  bookingTableBody.innerHTML = visibleBookings
    .map((booking) => {
      const sourceMeta = getBookingSourceMeta(booking.source);
      return `
      <tr>
        <td>${booking.guest}</td>
        <td>
          <div class="booking-source-stack">
            <span class="source-badge ${sourceMeta.className}">${sourceMeta.label}</span>
          <input
            type="text"
            class="booking-inline-input ${sourceMeta.className}"
            data-booking-field="source"
            data-booking-id="${booking.id}"
            value="${booking.source || ""}"
            list="bookingSourceOptions"
            placeholder="予約元"
          >
          </div>
        </td>
        <td>${booking.type}号室</td>
        <td>${booking.checkIn}</td>
        <td>${booking.checkOut}</td>
        <td>${booking.guestTotal}人 / ${nightsBetween(booking.checkIn, booking.checkOut).length}泊</td>
        <td>
          <input
            type="number"
            class="booking-inline-input booking-amount-input"
            data-booking-field="amount"
            data-booking-id="${booking.id}"
            value="${Number(booking.amount) || 0}"
            min="0"
            step="1"
          >
        </td>
        <td>
          <select
            class="booking-inline-input booking-select-input payment-status"
            data-booking-field="paymentStatus"
            data-booking-id="${booking.id}"
          >
            <option value="未払い" ${booking.paymentStatus === "未払い" ? "selected" : ""}>未払い</option>
            <option value="支払い済み" ${booking.paymentStatus === "支払い済み" ? "selected" : ""}>支払い済み</option>
            <option value="現地払い" ${booking.paymentStatus === "現地払い" ? "selected" : ""}>現地払い</option>
            <option value="返金済み" ${booking.paymentStatus === "返金済み" ? "selected" : ""}>返金済み</option>
          </select>
        </td>
        <td>
          <select
            class="booking-inline-input booking-select-input cleaning-status"
            data-booking-field="cleaningStatus"
            data-booking-id="${booking.id}"
          >
            <option value="未清掃" ${booking.cleaningStatus === "未清掃" ? "selected" : ""}>未清掃</option>
            <option value="清掃予定" ${booking.cleaningStatus === "清掃予定" ? "selected" : ""}>清掃予定</option>
            <option value="清掃済み" ${booking.cleaningStatus === "清掃済み" ? "selected" : ""}>清掃済み</option>
          </select>
        </td>
        <td>
          <input
            type="text"
            class="booking-inline-input"
            data-booking-field="cleaningMemo"
            data-booking-id="${booking.id}"
            value="${booking.cleaningMemo || ""}"
            placeholder="清掃メモ"
          >
        </td>
        <td>
          <div class="status-mini-grid">
            <label class="status-mini-item">
              <input type="checkbox" data-booking-field="idVerified" data-booking-id="${booking.id}" ${booking.idVerified ? "checked" : ""}>
              <span>身分証</span>
            </label>
            <label class="status-mini-item">
              <input type="checkbox" data-booking-field="checkInGuideSent" data-booking-id="${booking.id}" ${booking.checkInGuideSent ? "checked" : ""}>
              <span>案内</span>
            </label>
            <label class="status-mini-item">
              <input type="checkbox" data-booking-field="reviewRequestSent" data-booking-id="${booking.id}" ${booking.reviewRequestSent ? "checked" : ""}>
              <span>レビュー</span>
            </label>
          </div>
          <div class="booking-status-summary">${getBookingTaskSummary(booking)}</div>
        </td>
        <td>
          <input
            type="text"
            class="booking-inline-input"
            data-booking-field="memo"
            data-booking-id="${booking.id}"
            value="${booking.memo || ""}"
            placeholder="あとからメモ追加"
          >
        </td>
        <td class="booking-action-cell">
          <button type="button" class="button secondary" onclick="window.saveBookingMeta('${booking.id}')">保存</button>
          <button type="button" class="button secondary" onclick="window.cancelBookingById('${booking.id}')">取り消し</button>
        </td>
      </tr>
    `;
    })
    .join("");
}

function renderSummary() {
  const targetDate = getTodayDate();
  const remaining = getRoomTypes().reduce((sum, room) => sum + Math.max(availableStock(targetDate, room) ?? 0, 0), 0);
  todaySummary.innerHTML = `${targetDate}<br>残り ${remaining} / ${state.settings.totalRooms} 室`;
}

function renderFaqTemplates() {
  faqTemplates.innerHTML = cannedReplies.map((item) => `
    <article class="faq-card">
      <p class="faq-question">${item.question}</p>
      <p class="faq-answer">${item.answer}</p>
    </article>
  `).join("");
}

function renderStatusList(node, items, emptyText) {
  if (!items.length) {
    node.innerHTML = `<p class="status-empty">${emptyText}</p>`;
    return;
  }

  node.innerHTML = items.map((item) => `
    <div class="status-item">
      <strong>${item.room}号室</strong>
      <span>${item.label}</span>
    </div>
  `).join("");
}

function renderTodayOperations() {
  const today = getTodayDate();
  const staying = getBookingsForDate(today)
    .sort((a, b) => Number(a.type) - Number(b.type))
    .map((booking) => ({
      room: booking.type,
      label: `${booking.guest} / ${booking.guestTotal}人 / ${getBookingSourceMeta(booking.source).label}`,
    }));

  const checkIns = getBookingsCheckingInOn(today)
    .sort((a, b) => Number(a.type) - Number(b.type))
    .map((booking) => ({
      room: booking.type,
      label: `${booking.guest} / ${booking.guestTotal}人 / ${getBookingSourceMeta(booking.source).label}`,
    }));

  const checkOuts = getBookingsCheckingOutOn(today)
    .sort((a, b) => Number(a.type) - Number(b.type))
    .map((booking) => ({
      room: booking.type,
      label: `${booking.guest} / ${booking.paymentStatus || "未払い"}`,
    }));

  const cleaning = getCleaningBookingsForDate(today)
    .sort((a, b) => Number(a.type) - Number(b.type))
    .map((booking) => ({
      room: booking.type,
      label: `${booking.guest} / ${booking.cleaningStatus || "未清掃"}`,
    }));

  todayStayingCount.textContent = String(staying.length);
  todayCheckInCount.textContent = String(checkIns.length);
  todayCheckOutCount.textContent = String(checkOuts.length);
  todayCleaningCount.textContent = String(cleaning.length);
  if (topCheckInCount) {
    topCheckInCount.textContent = String(checkIns.length);
  }
  if (topCheckOutCount) {
    topCheckOutCount.textContent = String(checkOuts.length);
  }
  if (topCleaningCount) {
    topCleaningCount.textContent = String(cleaning.length);
  }

  renderStatusList(todayStayingList, staying, "本日宿泊中のお部屋はありません。");
  renderStatusList(todayCheckInList, checkIns, "本日のチェックインはありません。");
  renderStatusList(todayCheckOutList, checkOuts, "本日のチェックアウトはありません。");
  renderStatusList(todayCleaningList, cleaning, "本日の清掃対象はありません。");
}

function renderAlerts() {
  const unpaid = state.bookings.filter((booking) => booking.paymentStatus !== "支払い済み");
  const uncleaned = state.bookings.filter((booking) => booking.cleaningStatus !== "清掃済み");
  const idPending = state.bookings.filter((booking) => !booking.idVerified);
  const guidePending = state.bookings.filter((booking) => !booking.checkInGuideSent);

  if (unpaidAlertCount) {
    unpaidAlertCount.textContent = String(unpaid.length);
  }
  if (uncleanedAlertCount) {
    uncleanedAlertCount.textContent = String(uncleaned.length);
  }
  if (idPendingAlertCount) {
    idPendingAlertCount.textContent = String(idPending.length);
  }
  if (guidePendingAlertCount) {
    guidePendingAlertCount.textContent = String(guidePending.length);
  }
}

function buildMonthStats(monthKey) {
  const monthRows = state.inventory.filter((item) => item.date.startsWith(monthKey));
  const dates = [...new Set(monthRows.map((item) => item.date))].sort();
  const totalSlots = monthRows.length;
  const occupiedSlots = monthRows.filter((item) => (availableStock(item.date, item.type) ?? 0) <= 0).length;
  const vacancyPerDay = dates.map((date) => getRoomTypes().reduce((sum, room) => sum + Math.max(availableStock(date, room) ?? 0, 0), 0));
  const averageVacancy = dates.length ? vacancyPerDay.reduce((sum, value) => sum + value, 0) / dates.length : 0;
  const flowStats = inferMonthlyFlowStats(monthKey);
  const monthlyBookings = state.bookings.filter((booking) => booking.checkIn.startsWith(monthKey));
  const bookingCount = monthlyBookings.length || flowStats.checkInCount;
  const revenue = monthlyBookings.reduce((sum, booking) => sum + (Number(booking.amount) || 0), 0);
  const cleaningValue = flowStats.checkOutCount;

  return {
    occupancyRate: totalSlots ? (occupiedSlots / totalSlots) * 100 : 0,
    bookingCount,
    revenue,
    averageVacancy,
    cleaningCount: cleaningValue,
    checkInCount: flowStats.checkInCount,
    checkOutCount: flowStats.checkOutCount,
    occupiedSlots,
    totalSlots,
  };
}

function renderMonthSelectors() {
  const monthKeys = getMonthKeys();
  const options = monthKeys.map((monthKey) => `<option value="${monthKey}">${formatMonthLabel(monthKey)}</option>`).join("");
  monthSelect.innerHTML = options;
  compareMonthA.innerHTML = options;
  compareMonthB.innerHTML = options;
  const preferredMonthKey = getPreferredMonthKey();

  if (monthKeys.length) {
    if (!monthKeys.includes(monthSelect.value)) {
      monthSelect.value = monthKeys.includes(preferredMonthKey) ? preferredMonthKey : monthKeys[monthKeys.length - 1];
    }
    if (!monthKeys.includes(compareMonthA.value)) {
      compareMonthA.value = monthKeys.includes(preferredMonthKey)
        ? preferredMonthKey
        : monthKeys[Math.max(monthKeys.length - 2, 0)];
    }
    if (!monthKeys.includes(compareMonthB.value)) {
      compareMonthB.value = monthKeys.includes(preferredMonthKey) ? preferredMonthKey : monthKeys[monthKeys.length - 1];
    }
  }
}

function renderMonthlySummary() {
  const monthKey = monthSelect.value || getMonthKeys()[0];
  if (!monthKey) {
    monthNarrative.textContent = "月次データがありません。";
    return;
  }

  const stats = buildMonthStats(monthKey);
  monthOccupancy.textContent = `${stats.occupancyRate.toFixed(1)}%`;
  monthBookings.textContent = String(stats.bookingCount);
  monthRevenue.textContent = formatCurrency(stats.revenue);
  monthVacancyAverage.textContent = `${stats.averageVacancy.toFixed(1)}室`;
  monthCleaningCount.textContent = String(stats.cleaningCount);
  monthCheckInCount.textContent = String(stats.checkInCount);
  monthCheckOutCount.textContent = String(stats.checkOutCount);
  monthNarrative.textContent = `${formatMonthLabel(monthKey)} は売上 ${formatCurrency(stats.revenue)}、予約 ${stats.bookingCount} 件、延べ ${stats.totalSlots} 室日のうち ${stats.occupiedSlots} 室日が埋まっており、入居率は ${stats.occupancyRate.toFixed(1)}% です。チェックイン ${stats.checkInCount} 件、チェックアウト ${stats.checkOutCount} 件です。`;
}

function renderComparison() {
  const monthA = compareMonthA.value;
  const monthB = compareMonthB.value;
  if (!monthA || !monthB) {
    comparisonNarrative.textContent = "比較できる月がありません。";
    comparisonTableBody.innerHTML = "";
    return;
  }

  const statsA = buildMonthStats(monthA);
  const statsB = buildMonthStats(monthB);
  const diff = statsB.occupancyRate - statsA.occupancyRate;

  compareHeaderA.textContent = formatMonthLabel(monthA);
  compareHeaderB.textContent = formatMonthLabel(monthB);
  comparisonNarrative.textContent = `${formatMonthLabel(monthB)} の入居率は ${statsB.occupancyRate.toFixed(1)}% で、${formatMonthLabel(monthA)} と比べて ${diff >= 0 ? "+" : ""}${diff.toFixed(1)}pt です。`;

  comparisonTableBody.innerHTML = `
    <tr><td>入居率</td><td>${statsA.occupancyRate.toFixed(1)}%</td><td>${statsB.occupancyRate.toFixed(1)}%</td></tr>
    <tr><td>予約件数</td><td>${statsA.bookingCount}</td><td>${statsB.bookingCount}</td></tr>
    <tr><td>売上</td><td>${formatCurrency(statsA.revenue)}</td><td>${formatCurrency(statsB.revenue)}</td></tr>
    <tr><td>チェックイン数</td><td>${statsA.checkInCount}</td><td>${statsB.checkInCount}</td></tr>
    <tr><td>チェックアウト数</td><td>${statsA.checkOutCount}</td><td>${statsB.checkOutCount}</td></tr>
    <tr><td>平均空室数/日</td><td>${statsA.averageVacancy.toFixed(1)}室</td><td>${statsB.averageVacancy.toFixed(1)}室</td></tr>
    <tr><td>清掃回数</td><td>${statsA.cleaningCount}</td><td>${statsB.cleaningCount}</td></tr>
  `;
}

function renderAll() {
  renderSelectors();
  renderCalendarSelectors();
  renderCalendar();
  renderBookings();
  renderSummary();
  renderAlerts();
  renderFaqTemplates();
  renderTodayOperations();
  renderMonthSelectors();
  renderMonthlySummary();
  renderComparison();
}

function downloadCalendarPdf() {
  document.body.classList.add("print-calendar");
  activateTab("reservations");
  window.print();
  window.setTimeout(() => {
    document.body.classList.remove("print-calendar");
  }, 300);
}

function isSiteUnlocked() {
  try {
    return localStorage.getItem(SITE_UNLOCK_KEY) === "true";
  } catch (error) {
    return false;
  }
}

function setSiteUnlocked(unlocked) {
  try {
    localStorage.setItem(SITE_UNLOCK_KEY, unlocked ? "true" : "false");
  } catch (error) {
    console.warn("site unlock persistence unavailable", error);
  }
}

function updateSiteLockUi() {
  const unlocked = isSiteUnlocked();
  document.body.classList.toggle("site-locked", !unlocked);
  if (siteLockOverlay) {
    siteLockOverlay.classList.toggle("hidden", unlocked);
  }
}

function updateNavScrollUi() {
  document.body.classList.toggle("nav-scrolled", window.scrollY > 12);
}

function closeMobileNav() {
  if (!siteNavMenu || !mobileNavToggle) {
    return;
  }
  siteNavMenu.classList.remove("open");
  mobileNavToggle.setAttribute("aria-expanded", "false");
}

function updateSettingsLockUi() {
  totalRoomsInput.disabled = !settingsUnlocked;
  unlockSettingsBtn.textContent = settingsUnlocked ? "解除中" : "ロック解除";
}

function activateTab(tabName) {
  tabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === tabName);
  });
  tabPanels.forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.panel === tabName);
  });
}

function focusTabPanel(tabName) {
  const targetPanel = tabPanels.find((panel) => panel.dataset.panel === tabName);
  if (!targetPanel) {
    return;
  }

  window.requestAnimationFrame(() => {
    const navOffset = siteNav ? siteNav.getBoundingClientRect().height + 18 : 12;
    const targetTop = targetPanel.getBoundingClientRect().top + window.scrollY - navOffset;
    window.scrollTo({
      top: Math.max(targetTop, 0),
      behavior: "smooth",
    });
  });
}

function validateInventoryAgainstCapacity(items) {
  const totalsByDate = items.reduce((map, item) => {
    map[item.date] = (map[item.date] || 0) + item.stock;
    return map;
  }, {});

  Object.entries(totalsByDate).forEach(([date, total]) => {
    if (total > state.settings.totalRooms) {
      throw new Error(`${date} の在庫合計 ${total} 室が総部屋数 ${state.settings.totalRooms} 室を超えています。`);
    }
  });
}

loadInventoryBtn.addEventListener("click", () => {
  try {
    const parsed = parseInventory(inventoryInput.value);
    validateInventoryAgainstCapacity(parsed);
    state.inventory = parsed;
    saveState();
    renderAll();
    setResult(availabilityResult, "在庫データを更新しました。", "ok");
  } catch (error) {
    setResult(availabilityResult, error.message, "error");
  }
});

settingsForm.addEventListener("submit", (event) => {
  event.preventDefault();
  if (!settingsUnlocked) {
    settingsFeedback.textContent = "総部屋数を変更するには、先にパスワードでロック解除してください。";
    return;
  }

  const totalRooms = Number(totalRoomsInput.value);
  if (!Number.isInteger(totalRooms) || totalRooms <= 0) {
    settingsFeedback.textContent = "総部屋数は1以上で入力してください。";
    return;
  }

  const previous = state.settings.totalRooms;
  state.settings.totalRooms = totalRooms;
  try {
    validateInventoryAgainstCapacity(state.inventory);
    saveState();
    settingsFeedback.textContent = `総部屋数を ${totalRooms} 室に更新しました。`;
    settingsUnlocked = false;
    settingsPassword.value = "";
    updateSettingsLockUi();
    renderSummary();
  } catch (error) {
    state.settings.totalRooms = previous;
    totalRoomsInput.value = previous;
    settingsFeedback.textContent = error.message;
  }
});

unlockSettingsBtn.addEventListener("click", () => {
  settingsUnlocked = settingsPassword.value === DEFAULT_SETTINGS_PASSWORD;
  settingsFeedback.textContent = settingsUnlocked
    ? "設定を編集できます。保存すると再度ロックされます。"
    : "管理用パスワードが違います。";
  updateSettingsLockUi();
});

availabilityForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const checkIn = document.getElementById("checkIn").value;
  const checkOut = document.getElementById("checkOut").value;
  const guests = Number(guestCount.value);
  const result = findRoomCandidates(checkIn, checkOut, guests);
  setResult(availabilityResult, result.message, result.ok ? "ok" : "error");

  if (result.ok && result.rooms.length) {
    document.getElementById("bookCheckIn").value = checkIn;
    document.getElementById("bookCheckOut").value = checkOut;
    bookGuestCount.value = String(guests);
    bookRoomType.value = result.rooms[0].room;
  }
});

bookingForm.addEventListener("submit", (event) => {
  event.preventDefault();

  const booking = {
    id: createBookingId(),
    guest: document.getElementById("guestName").value.trim(),
    checkIn: document.getElementById("bookCheckIn").value,
    checkOut: document.getElementById("bookCheckOut").value,
    type: bookRoomType.value,
    guestTotal: Number(bookGuestCount.value),
    source: bookingSource.value.trim(),
    amount: Number(bookingAmount.value) || 0,
    paymentStatus: normalizePaymentStatus(paymentStatus.value),
    cleaningStatus: "未清掃",
    cleaningMemo: "",
    idVerified: Boolean(idVerified?.checked),
    checkInGuideSent: Boolean(checkInGuideSent?.checked),
    reviewRequestSent: Boolean(reviewRequestSent?.checked),
    memo: bookingMemo.value.trim(),
  };

  if (!booking.guest || !booking.source) {
    setResult(bookingFeedback, "予約者名と予約サイトを入力してください。", "error");
    return;
  }

  const result = checkAvailability(booking.checkIn, booking.checkOut, booking.type, booking.guestTotal);
  if (!result.ok) {
    setResult(bookingFeedback, result.message, "error");
    return;
  }

  const conflicts = findBookingConflicts(booking);
  if (conflicts.length) {
    const conflictText = conflicts.map((item) => `${item.type}号室 ${item.checkIn} - ${item.checkOut} (${item.guest})`).join(" / ");
    setResult(bookingFeedback, `ダブルブッキングの可能性があります。${conflictText}`, "error");
    return;
  }

  state.bookings.push(booking);
  saveState();
  renderAll();
  bookingForm.reset();
  setResult(bookingFeedback, `${booking.guest}様の予約を登録しました。`, "ok");
});

downloadInventoryBtn.addEventListener("click", () => {
  const csv = `date,room,stock\n${state.inventory.map((item) => [item.date, item.type, item.stock].join(",")).join("\n")}`;
  download("inventory.csv", csv);
});

downloadBookingsBtn.addEventListener("click", () => {
  const csv = `guest,source,room,check_in,check_out,nights,guest_total,amount,payment_status,cleaning_status,cleaning_memo,id_verified,guide_sent,review_request_sent,memo\n${state.bookings.map((booking) => [
    booking.guest,
    booking.source || "",
    booking.type,
    booking.checkIn,
    booking.checkOut,
    nightsBetween(booking.checkIn, booking.checkOut).length,
    booking.guestTotal,
    Number(booking.amount) || 0,
    booking.paymentStatus || "",
    booking.cleaningStatus || "",
    booking.cleaningMemo || "",
    booking.idVerified ? "1" : "0",
    booking.checkInGuideSent ? "1" : "0",
    booking.reviewRequestSent ? "1" : "0",
    booking.memo || "",
  ].join(",")).join("\n")}`;
  download("bookings.csv", csv);
});

if (backupJsonBtn) {
  backupJsonBtn.addEventListener("click", () => {
    downloadJson("morishita-backup.json", {
      exportedAt: new Date().toISOString(),
      settings: state.settings,
      inventory: state.inventory,
      bookings: state.bookings,
      notes: state.notes,
    });
    setResult(bookingFeedback, "JSONバックアップを保存しました。", "ok");
  });
}

if (restoreJsonBtn && restoreJsonInput) {
  restoreJsonBtn.addEventListener("click", () => restoreJsonInput.click());
  restoreJsonInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      state.settings = {
        totalRooms: Number(parsed.settings?.totalRooms) || state.settings.totalRooms,
      };
      state.inventory = Array.isArray(parsed.inventory) && parsed.inventory.length ? parsed.inventory : state.inventory;
      state.bookings = Array.isArray(parsed.bookings) ? parsed.bookings : [];
      state.notes = Array.isArray(parsed.notes) ? parsed.notes : [];
      ensureBookingIds();
      totalRoomsInput.value = state.settings.totalRooms;
      inventoryInput.value = toInventoryTsv(state.inventory);
      saveState();
      renderAll();
      setResult(bookingFeedback, "JSONバックアップから復元しました。", "ok");
    } catch (error) {
      setResult(bookingFeedback, `JSON復元に失敗しました: ${error.message}`, "error");
    } finally {
      restoreJsonInput.value = "";
    }
  });
}

calendarNoteForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const payload = {
    date: noteDate.value,
    room: noteRoom.value,
    text: noteText.value.trim(),
  };

  if (!payload.date || !payload.room || !payload.text) {
    setResult(calendarNoteFeedback, "日付・部屋・注釈メモを入力してください。", "error");
    return;
  }

  state.notes = state.notes.filter((note) => getNoteKey(note.date, note.room) !== getNoteKey(payload.date, payload.room));
  state.notes.push(payload);
  saveState();
  renderCalendar();
  calendarNoteForm.reset();
  setResult(calendarNoteFeedback, `${payload.date} の ${payload.room}号室に注釈を保存しました。`, "ok");
});

downloadCalendarPdfBtn.addEventListener("click", downloadCalendarPdf);

siteLockForm.addEventListener("submit", (event) => {
  event.preventDefault();
  if (sitePasswordInput.value === SITE_ACCESS_PASSWORD) {
    setSiteUnlocked(true);
    updateSiteLockUi();
    siteLockFeedback.textContent = "";
    sitePasswordInput.value = "";
    return;
  }

  siteLockFeedback.textContent = "サイトパスワードが違います。";
});

function cancelBookingById(bookingId) {
  const targetBooking = state.bookings.find((booking) => booking.id === bookingId);
  if (!targetBooking) {
    return;
  }

  const confirmed = window.confirm(`${targetBooking.guest}様の予約を取り消しますか？`);
  if (!confirmed) {
    return;
  }

  state.bookings = state.bookings.filter((booking) => booking.id !== bookingId);
  saveState();
  renderAll();
  setResult(bookingFeedback, `${targetBooking.guest}様の予約を取り消しました。`, "ok");
}

function saveBookingMeta(bookingId) {
  const sourceInput = document.querySelector(`[data-booking-field="source"][data-booking-id="${bookingId}"]`);
  const amountInput = document.querySelector(`[data-booking-field="amount"][data-booking-id="${bookingId}"]`);
  const paymentStatusInput = document.querySelector(`[data-booking-field="paymentStatus"][data-booking-id="${bookingId}"]`);
  const cleaningStatusInput = document.querySelector(`[data-booking-field="cleaningStatus"][data-booking-id="${bookingId}"]`);
  const cleaningMemoInput = document.querySelector(`[data-booking-field="cleaningMemo"][data-booking-id="${bookingId}"]`);
  const idVerifiedInput = document.querySelector(`[data-booking-field="idVerified"][data-booking-id="${bookingId}"]`);
  const checkInGuideSentInput = document.querySelector(`[data-booking-field="checkInGuideSent"][data-booking-id="${bookingId}"]`);
  const reviewRequestSentInput = document.querySelector(`[data-booking-field="reviewRequestSent"][data-booking-id="${bookingId}"]`);
  const memoInput = document.querySelector(`[data-booking-field="memo"][data-booking-id="${bookingId}"]`);
  const booking = state.bookings.find((item) => item.id === bookingId);

  if (!booking || !sourceInput || !amountInput || !paymentStatusInput || !cleaningStatusInput || !cleaningMemoInput || !idVerifiedInput || !checkInGuideSentInput || !reviewRequestSentInput || !memoInput) {
    return;
  }

  booking.source = sourceInput.value.trim();
  booking.amount = Number(amountInput.value) || 0;
  booking.paymentStatus = normalizePaymentStatus(paymentStatusInput.value);
  booking.cleaningStatus = cleaningStatusInput.value;
  booking.cleaningMemo = cleaningMemoInput.value.trim();
  booking.idVerified = idVerifiedInput.checked;
  booking.checkInGuideSent = checkInGuideSentInput.checked;
  booking.reviewRequestSent = reviewRequestSentInput.checked;
  booking.memo = memoInput.value.trim();
  saveState();
  renderAll();
  setResult(bookingFeedback, `${booking.guest}様の予約情報を更新しました。`, "ok");
}

window.cancelBookingById = cancelBookingById;
window.saveBookingMeta = saveBookingMeta;

[
  [bookingSearchInput, "input"],
  [bookingSourceFilter, "change"],
  [bookingRoomFilter, "change"],
  [bookingMonthFilter, "change"],
  [bookingPaymentFilter, "change"],
  [bookingCleaningFilter, "change"],
  [bookingSortSelect, "change"],
].forEach(([node, eventName]) => {
  if (!node) {
    return;
  }
  node.addEventListener(eventName, renderBookings);
});

tabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    activateTab(button.dataset.tab);
    focusTabPanel(button.dataset.tab);
    closeMobileNav();
  });
});

if (mobileNavToggle && siteNavMenu) {
  mobileNavToggle.addEventListener("click", () => {
    const isOpen = siteNavMenu.classList.toggle("open");
    mobileNavToggle.setAttribute("aria-expanded", isOpen ? "true" : "false");
  });
}

document.addEventListener("click", (event) => {
  if (!siteNavMenu || !mobileNavToggle || !siteNav) {
    return;
  }
  if (!siteNavMenu.classList.contains("open")) {
    return;
  }
  if (siteNav.contains(event.target)) {
    return;
  }
  closeMobileNav();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeMobileNav();
  }
});

monthSelect.addEventListener("change", renderMonthlySummary);
compareMonthA.addEventListener("change", renderComparison);
compareMonthB.addEventListener("change", renderComparison);
calendarYearSelect.addEventListener("change", () => {
  renderCalendarSelectors();
  renderCalendar();
});
calendarMonthSelect.addEventListener("change", renderCalendar);

function initialize() {
  loadState();
  ensureBookingIds();
  updateSiteLockUi();
  updateNavScrollUi();
  activateTab("reservations");
  updateSettingsLockUi();
  renderAll();
}

window.addEventListener("scroll", updateNavScrollUi, { passive: true });

try {
  initialize();
} catch (error) {
  console.error(error);
  todaySummary.textContent = `表示エラー: ${error.message}`;
  calendarGrid.innerHTML = `<div class="calendar-card">表示エラー: ${error.message}</div>`;
}
