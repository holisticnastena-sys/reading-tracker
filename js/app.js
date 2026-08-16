const state = {
  books: [],
  activeStatus: "Все",
  visibleCount: 24,
  coverCache: JSON.parse(localStorage.getItem("readingTrackerCoverCache") || "{}")
};

const FALLBACK = "assets/no-cover.svg";

document.addEventListener("DOMContentLoaded", () => {
  bindNavigation();
  bindFilters();
  document.getElementById("reloadBtn").addEventListener("click", loadBooks);
  loadBooks();
});

function bindNavigation() {
  document.querySelectorAll(".nav-item").forEach(btn => {
    btn.addEventListener("click", () => showPage(btn.dataset.page));
  });
  document.querySelectorAll("[data-go]").forEach(btn => {
    btn.addEventListener("click", () => showPage(btn.dataset.go));
  });
}

function showPage(id) {
  document.querySelectorAll(".page").forEach(p => p.classList.toggle("active", p.id === id));
  document.querySelectorAll(".nav-item").forEach(b => b.classList.toggle("active", b.dataset.page === id));
  window.scrollTo({top: 0, behavior: "smooth"});
}

function bindFilters() {
  document.querySelectorAll("#statusFilters .chip").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll("#statusFilters .chip").forEach(x => x.classList.remove("active"));
      btn.classList.add("active");
      state.activeStatus = btn.dataset.status;
      state.visibleCount = 24;
      renderLibrary();
    });
  });

  ["searchInput", "formatFilter", "seriesFilter", "ratingFilter"].forEach(id => {
    document.getElementById(id).addEventListener(id === "searchInput" ? "input" : "change", () => {
      state.visibleCount = 24;
      renderLibrary();
    });
  });

  ["statsYear", "statsMonth", "statsFormat", "statsStatus", "statsSeries", "statsRating"].forEach(id => {
    document.getElementById(id).addEventListener("change", renderStatsPage);
  });

  document.getElementById("resetStatsFilters").addEventListener("click", () => {
    ["statsYear", "statsMonth", "statsFormat", "statsStatus", "statsSeries", "statsRating"]
      .forEach(id => document.getElementById(id).value = "");
    renderStatsPage();
  });

  document.getElementById("loadMoreBtn").addEventListener("click", () => {
    state.visibleCount += 24;
    renderLibrary();
  });
}

async function loadBooks() {
  const errorBox = document.getElementById("errorBox");
  errorBox.classList.add("hidden");

  try {
    const response = await fetch(`books.xlsx?v=${Date.now()}`);
    if (!response.ok) throw new Error(`Не удалось открыть books.xlsx (${response.status})`);

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });

    state.books = rows
      .filter(row => String(row["Название"] || "").trim())
      .map(normalizeBook);

    populateFilters();
    renderAll();
  } catch (err) {
    console.error(err);
    errorBox.textContent =
      "Не удалось прочитать books.xlsx. Если ты открыла index.html двойным щелчком, браузер блокирует чтение локального файла. На GitHub Pages всё будет работать; для локальной проверки запусти сайт через Live Server.";
    errorBox.classList.remove("hidden");
  }
}

function normalizeBook(row, index) {
  return {
    id: index + 1,
    title: text(row["Название"]),
    author: text(row["Автор"]),
    series: text(row["Серия"]),
    seriesNumber: text(row["Книга в серии"]),
    status: text(row["Статус"]),
    format: text(row["Формат"]),
    pages: number(row["Страниц"]),
    readPages: number(row["Прочитано"]),
    start: toDate(row["Начало"]),
    end: toDate(row["Конец"]),
    duration: number(row["Общее время прочтения"]),
    rating: number(row["Оценка"]),
    cover: text(row["Обложка"])
  };
}

function text(v) {
  return v == null ? "" : String(v).trim();
}

function number(v) {
  if (v === "" || v == null) return null;
  const n = Number(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number") {
    const parsed = XLSX.SSF.parse_date_code(v);
    return parsed ? new Date(parsed.y, parsed.m - 1, parsed.d) : null;
  }
  const parts = String(v).match(/(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})/);
  if (parts) return new Date(+parts[3], +parts[2] - 1, +parts[1]);
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function renderAll() {
  renderStats();
  renderHome();
  renderLibrary();
  renderSeries();
  renderStatsPage();
}

function renderStats() {
  const total = state.books.length;
  const finished = byStatus("Я все прочитал!").length;
  const reading = byStatus("Чтение").length;
  const paused = byStatus("Пауза").length;
  const dropped = byStatus("Бросил читать").length;

  const rated = state.books.filter(b => b.rating != null && b.rating > 0);
  const avg = rated.length ? rated.reduce((s,b) => s + b.rating, 0) / rated.length : null;

  setText("totalCount", total);
  setText("finishedCount", finished);
  setText("readingCount", reading);
  setText("pausedCount", paused);
  setText("droppedCount", dropped);
  setText("avgRating", avg ? avg.toFixed(1) : "—");
  setText("libraryCount", total);
}

function renderHome() {
  const reading = byStatus("Чтение").slice(0, 3);
  const readingNow = document.getElementById("readingNow");
  readingNow.innerHTML = reading.length ? "" : `<div class="muted">Сейчас нет книг со статусом «Чтение».</div>`;
  reading.forEach(book => readingNow.appendChild(makeReadingCard(book)));

  const finished = byStatus("Я все прочитал!")
    .filter(b => b.end)
    .sort((a,b) => b.end - a.end)
    .slice(0, 6);

  const recent = document.getElementById("recentFinished");
  recent.innerHTML = "";
  finished.forEach(book => recent.appendChild(makeCompactCard(book)));
}

function makeReadingCard(book) {
  const el = document.createElement("article");
  el.className = "reading-card";
  const progress = getProgress(book);
  el.innerHTML = `
    <img src="${FALLBACK}" alt="">
    <div>
      <h3>${escapeHtml(book.title)}</h3>
      <p>${escapeHtml(book.author)}</p>
      ${book.series ? `<p>${escapeHtml(book.series)}${book.seriesNumber ? ` · книга ${escapeHtml(book.seriesNumber)}` : ""}</p>` : ""}
      ${book.format ? `<span class="tag">${escapeHtml(book.format)}</span>` : ""}
      ${progress != null ? `
        <div class="progress-line"><span style="width:${progress}%"></span></div>
        <div class="progress-text"><span>${book.readPages ?? 0} / ${book.pages ?? "—"} стр.</span><b>${progress}%</b></div>
      ` : ""}
      ${book.start ? `<div class="progress-text"><span>Начато ${formatDate(book.start)}</span></div>` : ""}
    </div>
  `;
  hydrateCover(el.querySelector("img"), book);
  return el;
}

function makeCompactCard(book) {
  const el = document.createElement("article");
  el.className = "compact-card";
  el.innerHTML = `
    <img src="${FALLBACK}" alt="">
    <h3>${escapeHtml(book.title)}</h3>
    <p>${escapeHtml(book.author)}</p>
    <div class="rating">${book.rating ? `★ ${book.rating.toFixed(1)}` : "Без оценки"}</div>
  `;
  hydrateCover(el.querySelector("img"), book);
  return el;
}

function renderLibrary() {
  const query = document.getElementById("searchInput").value.trim().toLowerCase();
  const format = document.getElementById("formatFilter").value;
  const series = document.getElementById("seriesFilter").value;
  const minRating = Number(document.getElementById("ratingFilter").value || 0);

  let books = state.books.filter(b => {
    if (state.activeStatus !== "Все" && b.status !== state.activeStatus) return false;
    if (format && b.format !== format) return false;
    if (series && b.series !== series) return false;
    if (minRating && (!b.rating || b.rating < minRating)) return false;
    if (query && !`${b.title} ${b.author} ${b.series}`.toLowerCase().includes(query)) return false;
    return true;
  });

  const grid = document.getElementById("libraryGrid");
  grid.innerHTML = "";

  books.slice(0, state.visibleCount).forEach(book => {
    const card = document.getElementById("bookCardTemplate").content.firstElementChild.cloneNode(true);
    const img = card.querySelector(".cover");
    img.src = FALLBACK;
    img.alt = `Обложка: ${book.title}`;

    card.querySelector(".status-badge").textContent = shortStatus(book.status);
    card.querySelector(".book-title").textContent = book.title;
    card.querySelector(".book-author").textContent = book.author;

    const meta = [book.format, book.series ? `${book.series}${book.seriesNumber ? ` #${book.seriesNumber}` : ""}` : ""].filter(Boolean);
    card.querySelector(".book-meta").textContent = meta.join(" · ");
    card.querySelector(".rating").textContent = book.rating ? `★ ${book.rating.toFixed(1)}` : "Без оценки";

    grid.appendChild(card);
    hydrateCover(img, book);
  });

  document.getElementById("loadMoreBtn").classList.toggle("hidden", books.length <= state.visibleCount);
}

function populateFilters() {
  const formats = unique(state.books.map(b => b.format));
  const series = unique(state.books.map(b => b.series));

  fillSelect("formatFilter", formats, "Все форматы");
  fillSelect("seriesFilter", series, "Все серии");
  fillSelect("statsFormat", formats, "Все форматы");
  fillSelect("statsSeries", series, "Все серии");

  const years = unique(
    state.books
      .map(b => (b.end || b.start))
      .filter(Boolean)
      .map(d => String(d.getFullYear()))
  ).sort((a,b) => Number(b) - Number(a));
  fillSelect("statsYear", years, "Всё время");
}

function fillSelect(id, values, label) {
  const select = document.getElementById(id);
  select.innerHTML = `<option value="">${label}</option>`;
  values.forEach(v => {
    const o = document.createElement("option");
    o.value = v; o.textContent = v;
    select.appendChild(o);
  });
}

function renderSeries() {
  const groups = {};
  state.books.filter(b => b.series).forEach(b => {
    (groups[b.series] ||= []).push(b);
  });

  const root = document.getElementById("seriesGrid");
  root.innerHTML = "";

  Object.entries(groups)
    .sort((a,b) => a[0].localeCompare(b[0], "ru"))
    .forEach(([name, books]) => {
      books.sort((a,b) => numericSeries(a.seriesNumber) - numericSeries(b.seriesNumber));
      const finished = books.filter(b => b.status === "Я все прочитал!").length;
      const el = document.createElement("article");
      el.className = "series-card";
      el.innerHTML = `
        <h3>${escapeHtml(name)}</h3>
        <p>${books.length} ${declension(books.length, ["книга","книги","книг"])} · завершено ${finished}/${books.length}</p>
        ${books.map(b => `
          <div class="series-book">
            <span>${b.seriesNumber ? `${escapeHtml(b.seriesNumber)}. ` : ""}${escapeHtml(b.title)}</span>
            <em>${shortStatus(b.status)}${b.rating ? ` · ★ ${b.rating.toFixed(1)}` : ""}</em>
          </div>
        `).join("")}
      `;
      root.appendChild(el);
    });
}

function renderStatsPage() {
  const year = document.getElementById("statsYear").value;
  const month = document.getElementById("statsMonth").value;
  const format = document.getElementById("statsFormat").value;
  const status = document.getElementById("statsStatus").value;
  const series = document.getElementById("statsSeries").value;
  const minRating = Number(document.getElementById("statsRating").value || 0);

  const filtered = state.books.filter(book => {
    const activityDate = book.end || book.start;

    if (year && (!activityDate || String(activityDate.getFullYear()) !== year)) return false;
    if (month !== "" && (!activityDate || String(activityDate.getMonth()) !== month)) return false;
    if (format && book.format !== format) return false;
    if (status && book.status !== status) return false;
    if (series && book.series !== series) return false;
    if (minRating && (!book.rating || book.rating < minRating)) return false;
    return true;
  });

  const finished = filtered.filter(b => b.status === "Я все прочитал!");
  const dropped = filtered.filter(b => b.status === "Бросил читать");
  const reading = filtered.filter(b => b.status === "Чтение");
  const paused = filtered.filter(b => b.status === "Пауза");
  const rated = filtered.filter(b => b.rating && b.rating > 0);
  const avg = rated.length ? rated.reduce((s,b) => s+b.rating,0)/rated.length : null;
  const totalDays = filtered.reduce((s,b) => s + (b.duration || 0), 0);

  document.getElementById("statsPage").innerHTML = `
    <article class="big-stat"><strong>${filtered.length}</strong><span>в выборке</span></article>
    <article class="big-stat"><strong>${finished.length}</strong><span>прочитано</span></article>
    <article class="big-stat"><strong>${dropped.length}</strong><span>брошено</span></article>
    <article class="big-stat"><strong>${reading.length}</strong><span>читаю</span></article>
    <article class="big-stat"><strong>${avg ? avg.toFixed(1) : "—"}</strong><span>средняя оценка</span></article>
    <article class="big-stat"><strong>${totalDays}</strong><span>дней чтения суммарно</span></article>
  `;

  renderMonthChart(filtered);
  renderStatusChart(filtered);
  renderRatingChart(filtered);
}

function renderMonthChart(books) {
  const months = ["Янв","Фев","Мар","Апр","Май","Июн","Июл","Авг","Сен","Окт","Ноя","Дек"];
  const counts = Array(12).fill(0);

  books.forEach(book => {
    const d = book.end || book.start;
    if (d) counts[d.getMonth()]++;
  });

  renderBars("monthChart", months.map((label, i) => ({ label, value: counts[i] })));
}

function renderStatusChart(books) {
  const defs = [
    ["Прочитано", "Я все прочитал!"],
    ["Брошено", "Бросил читать"],
    ["Читаю", "Чтение"],
    ["На паузе", "Пауза"],
    ["К прочтению", "К прочтению"]
  ];

  renderBars("statusChart", defs.map(([label, value]) => ({
    label,
    value: books.filter(b => b.status === value).length
  })));
}

function renderRatingChart(books) {
  const ratings = {};
  books
    .filter(b => b.rating && b.rating > 0)
    .forEach(b => {
      const key = b.rating.toFixed(1);
      ratings[key] = (ratings[key] || 0) + 1;
    });

  const data = Object.entries(ratings)
    .sort((a,b) => Number(b[0]) - Number(a[0]))
    .map(([label, value]) => ({ label: `★ ${label}`, value }));

  renderBars("ratingChart", data);
}

function renderBars(id, data) {
  const root = document.getElementById(id);
  if (!data.length || data.every(d => d.value === 0)) {
    root.innerHTML = `<div class="empty-chart">Нет данных для выбранных фильтров</div>`;
    return;
  }

  const max = Math.max(...data.map(d => d.value), 1);
  root.innerHTML = data.map(d => `
    <div class="bar-row">
      <span class="bar-label">${escapeHtml(d.label)}</span>
      <div class="bar-track"><span style="width:${(d.value/max)*100}%"></span></div>
      <strong>${d.value}</strong>
    </div>
  `).join("");
}

function groupSeries() {
  return state.books.filter(b=>b.series).reduce((a,b)=>((a[b.series] ||= []).push(b),a),{});
}

function byStatus(status) {
  return state.books.filter(b => b.status === status);
}

function getProgress(book) {
  if (!book.pages || book.readPages == null) return null;
  return Math.max(0, Math.min(100, Math.round(book.readPages / book.pages * 100)));
}

async function hydrateCover(img, book) {
  img.loading = "lazy";

  if (book.cover) {
    img.src = book.cover;
    img.onerror = () => {
      img.onerror = null;
      img.src = FALLBACK;
      hydrateAutoCover(img, book);
    };
    return;
  }

  hydrateAutoCover(img, book);
}

function hydrateAutoCover(img, book) {
  const key = `${book.title}__${book.author}`.toLowerCase();

  if (state.coverCache[key]) {
    img.src = state.coverCache[key];
    return;
  }

  if (!isNearViewport(img)) {
    const observer = new IntersectionObserver(entries => {
      if (entries.some(e => e.isIntersecting)) {
        observer.disconnect();
        findCover(img, book, key);
      }
    }, { rootMargin: "500px" });
    observer.observe(img);
  } else {
    findCover(img, book, key);
  }
}

async function findCover(img, book, key) {
  try {
    const q = `intitle:${book.title} inauthor:${book.author}`;
    const url = `https://www.googleapis.com/books/v1/volumes?q=${encodeURIComponent(q)}&maxResults=5`;
    const response = await fetch(url);
    if (!response.ok) return;
    const data = await response.json();
    const items = data.items || [];

    let image = null;
    for (const item of items) {
      const links = item.volumeInfo && item.volumeInfo.imageLinks;
      image = links && (links.thumbnail || links.smallThumbnail);
      if (image) break;
    }

    if (image) {
      image = image.replace("http://", "https://").replace("&edge=curl", "");
      state.coverCache[key] = image;
      localStorage.setItem("readingTrackerCoverCache", JSON.stringify(state.coverCache));
      img.src = image;
    }
  } catch (e) {
    console.warn("Обложка не найдена:", book.title);
  }
}

function isNearViewport(el) {
  const r = el.getBoundingClientRect();
  return r.top < innerHeight + 500;
}

function shortStatus(status) {
  const map = {
    "Я все прочитал!": "Прочитано",
    "Бросил читать": "Брошено",
    "К прочтению": "К прочтению",
    "Чтение": "Читаю",
    "Пауза": "На паузе"
  };
  return map[status] || status || "Без статуса";
}

function formatDate(d) {
  return new Intl.DateTimeFormat("ru-RU", {day:"2-digit", month:"2-digit", year:"numeric"}).format(d);
}

function setText(id, value) {
  document.getElementById(id).textContent = value;
}

function unique(arr) {
  return [...new Set(arr.filter(Boolean))].sort((a,b)=>a.localeCompare(b,"ru"));
}

function numericSeries(v) {
  if (!v) return 9999;
  const n = parseFloat(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : 9999;
}

function declension(n, forms) {
  n = Math.abs(n) % 100;
  const n1 = n % 10;
  if (n > 10 && n < 20) return forms[2];
  if (n1 > 1 && n1 < 5) return forms[1];
  if (n1 === 1) return forms[0];
  return forms[2];
}

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, ch => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#039;"
  }[ch]));
}
