(() => {
  "use strict";

  const STORAGE_KEY = "growthquest_v3";

  /** @typedef {{id:string,title:string,description:string,targetDate:string|null,targetAmount:number|null,createdAt:string}} Goal */
  /** @typedef {{id:string,title:string,goalId:string|null,createdAt:string,completedDates:string[]}} Habit */
  /** @typedef {{id:string,date:string,amount:number,source:string}} IncomeEntry */
  /** @typedef {{id:string,date:string,amount:number,category:string}} ExpenseEntry */
  /** @typedef {{id:string,text:string,date:string}} QuickTask */
  /** @typedef {{id:string,text:string,date:string,createdAt:string,syncedToSheets:boolean}} EmotionLog */

  function loadState() {
    const empty = {
      goals: [], habits: [], incomes: [], expenses: [],
      todayTasks: [], emotionLogs: [],
      lastSeenReportMonth: null, lastRecurringExpenseMonth: null,
      googleClientId: null, googleSpreadsheetId: null, googleFeelingsSheetReady: false, lastSyncedMonth: null,
      googleSheetStyled: false, googleSummarySheetReady: false, googleSheetMigratedJa: false,
      googleTxSplitByMonth: false,
    };
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return empty;
      const parsed = JSON.parse(raw);
      return {
        goals: parsed.goals || [],
        habits: parsed.habits || [],
        incomes: parsed.incomes || [],
        expenses: parsed.expenses || [],
        todayTasks: parsed.todayTasks || [],
        emotionLogs: parsed.emotionLogs || [],
        lastSeenReportMonth: parsed.lastSeenReportMonth ?? null,
        lastRecurringExpenseMonth: parsed.lastRecurringExpenseMonth ?? null,
        googleClientId: parsed.googleClientId ?? null,
        googleSpreadsheetId: parsed.googleSpreadsheetId ?? null,
        googleFeelingsSheetReady: parsed.googleFeelingsSheetReady ?? false,
        lastSyncedMonth: parsed.lastSyncedMonth ?? null,
        googleSheetStyled: parsed.googleSheetStyled ?? false,
        googleSummarySheetReady: parsed.googleSummarySheetReady ?? false,
        googleSheetMigratedJa: parsed.googleSheetMigratedJa ?? false,
        googleTxSplitByMonth: parsed.googleTxSplitByMonth ?? false,
      };
    } catch (e) {
      console.error("state load failed", e);
      return empty;
    }
  }

  function saveState() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }

  function uid() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  function pad2(n) { return String(n).padStart(2, "0"); }
  function dateKey(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }
  function dateOnly(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
  /** "YYYY-MM-DD" 文字列をタイムゾーンの影響を受けずにローカル日付として解釈する */
  function parseDateKey(s) {
    const [y, m, d] = s.split("-").map(Number);
    return new Date(y, m - 1, d);
  }
  function monthKeyNum(year, month) { return year * 12 + month; }
  function formatYen(n) { return `¥${Math.round(n).toLocaleString("ja-JP")}`; }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  let state = loadState();
  // 月が変わったら自動で追加する固定費(毎月1日付けの支出として計上)
  const RECURRING_MONTHLY_EXPENSES = [
    { category: "家賃", amount: 35000 },
    { category: "携帯代+アプリ", amount: 15000 },
    { category: "イラストレーター", amount: 9000 },
    { category: "ジム", amount: 3300 },
    { category: "クラウド", amount: 3300 },
  ];

  const WEEKDAY_LABELS = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];
  const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  function formatMonthYear(year, month) { return `${MONTH_NAMES[month]} ${year}`; }
  let today = new Date();

  // ---------------------------------------------------------------------
  // Tab navigation
  // ---------------------------------------------------------------------
  const views = {
    home: document.getElementById("view-home"),
    habits: document.getElementById("view-habits"),
    income: document.getElementById("view-income"),
    expense: document.getElementById("view-expense"),
    goals: document.getElementById("view-goals"),
  };
  const topbarTitle = document.getElementById("topbar-title");
  const titles = { home: "Growth Quest", habits: "Habits", income: "Income", expense: "Expense", goals: "Goals" };

  function renderView(name) {
    if (name === "home") renderHome();
    if (name === "habits") renderHabitsView(true);
    if (name === "income") renderIncomeView();
    if (name === "expense") renderExpenseView();
    if (name === "goals") renderGoals();
  }

  function activeViewName() {
    return Object.entries(views).find(([, el]) => el.classList.contains("active"))[0];
  }

  function showView(name) {
    Object.entries(views).forEach(([key, el]) => el.classList.toggle("active", key === name));
    document.querySelectorAll(".tab-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.view === name));
    topbarTitle.textContent = titles[name];
    fab.style.display = (name === "home" || name === "income" || name === "expense") ? "none" : "flex";
    renderView(name);
  }

  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => showView(btn.dataset.view));
  });

  // ---------------------------------------------------------------------
  // FAB (add button) — behavior depends on active tab
  // ---------------------------------------------------------------------
  const fab = document.getElementById("fab");
  fab.addEventListener("click", () => {
    const active = activeViewName();
    if (active === "habits") openHabitSheet();
    if (active === "goals") openGoalSheet();
  });

  // ---------------------------------------------------------------------
  // Sheet helpers
  // ---------------------------------------------------------------------
  function openSheet(id) { document.getElementById(id).classList.add("open"); }
  function closeSheet(id) { document.getElementById(id).classList.remove("open"); }

  document.querySelectorAll("[data-close-sheet]").forEach(btn => {
    btn.addEventListener("click", () => closeSheet(btn.dataset.closeSheet));
  });
  document.querySelectorAll(".sheet-overlay").forEach(overlay => {
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) overlay.classList.remove("open");
    });
  });

  document.getElementById("settings-btn").addEventListener("click", () => {
    document.getElementById("google-client-id-input").value = state.googleClientId || "";
    updateGoogleStatusUI();
    openSheet("sheet-settings");
  });

  // ---------------------------------------------------------------------
  // Backup / transfer (export all data as a JSON file, import it back)
  // ---------------------------------------------------------------------
  function showBackupFeedback(text) {
    const el = document.getElementById("backup-feedback");
    el.textContent = text;
  }

  document.getElementById("export-backup-btn").addEventListener("click", () => {
    const json = JSON.stringify(state, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const stamp = dateKey(new Date());
    a.href = url;
    a.download = `growth-quest-backup-${stamp}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    showBackupFeedback("Backup file saved. Check your Files app / Downloads.");
  });

  document.getElementById("import-backup-btn").addEventListener("click", () => {
    document.getElementById("import-backup-input").click();
  });

  document.getElementById("import-backup-input").addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result));
        if (typeof parsed !== "object" || parsed === null) throw new Error("invalid file");
        state = {
          goals: parsed.goals || [],
          habits: parsed.habits || [],
          incomes: parsed.incomes || [],
          expenses: parsed.expenses || [],
          todayTasks: parsed.todayTasks || [],
          emotionLogs: parsed.emotionLogs || [],
          lastSeenReportMonth: parsed.lastSeenReportMonth ?? null,
          lastRecurringExpenseMonth: parsed.lastRecurringExpenseMonth ?? null,
          googleClientId: parsed.googleClientId ?? null,
          googleSpreadsheetId: parsed.googleSpreadsheetId ?? null,
          googleFeelingsSheetReady: parsed.googleFeelingsSheetReady ?? false,
          lastSyncedMonth: parsed.lastSyncedMonth ?? null,
          googleSheetStyled: parsed.googleSheetStyled ?? false,
          googleSummarySheetReady: parsed.googleSummarySheetReady ?? false,
          googleSheetMigratedJa: parsed.googleSheetMigratedJa ?? false,
          googleTxSplitByMonth: parsed.googleTxSplitByMonth ?? false,
        };
        saveState();
        showBackupFeedback("Backup restored. Reloading…");
        setTimeout(() => location.reload(), 800);
      } catch (err) {
        console.error("import failed", err);
        showBackupFeedback("Couldn't read that file — make sure it's a Growth Quest backup JSON.");
      }
      e.target.value = "";
    };
    reader.readAsText(file);
  });

  // ---------------------------------------------------------------------
  // GOALS
  // ---------------------------------------------------------------------
  const goalTitleInput = document.getElementById("goal-title-input");
  const goalDescInput = document.getElementById("goal-desc-input");
  const goalHasDateInput = document.getElementById("goal-has-date-input");
  const goalDateInput = document.getElementById("goal-date-input");
  const goalHasAmountInput = document.getElementById("goal-has-amount-input");
  const goalAmountInput = document.getElementById("goal-amount-input");
  const goalDeleteBtn = document.getElementById("goal-delete-btn");
  const goalSheetTitle = document.getElementById("goal-sheet-title");
  let editingGoalId = null;

  goalHasDateInput.addEventListener("change", () => {
    goalDateInput.classList.toggle("hidden-field", !goalHasDateInput.checked);
  });
  goalHasAmountInput.addEventListener("change", () => {
    goalAmountInput.classList.toggle("hidden-field", !goalHasAmountInput.checked);
  });

  function openGoalSheet(goal) {
    editingGoalId = goal ? goal.id : null;
    goalSheetTitle.textContent = goal ? "Edit Goal" : "Add Goal";
    goalTitleInput.value = goal ? goal.title : "";
    goalDescInput.value = goal ? goal.description : "";
    goalHasDateInput.checked = !!(goal && goal.targetDate);
    goalDateInput.value = (goal && goal.targetDate) || "";
    goalDateInput.classList.toggle("hidden-field", !goalHasDateInput.checked);
    goalHasAmountInput.checked = !!(goal && goal.targetAmount);
    goalAmountInput.value = (goal && goal.targetAmount) || "";
    goalAmountInput.classList.toggle("hidden-field", !goalHasAmountInput.checked);
    goalDeleteBtn.classList.toggle("hidden-field", !goal);
    openSheet("sheet-goal");
  }

  document.getElementById("goal-save-btn").addEventListener("click", () => {
    const title = goalTitleInput.value.trim();
    if (!title) return;
    const targetDate = goalHasDateInput.checked ? goalDateInput.value || null : null;
    const targetAmount = goalHasAmountInput.checked ? (Number(goalAmountInput.value) || null) : null;

    if (editingGoalId) {
      const goal = state.goals.find(g => g.id === editingGoalId);
      goal.title = title;
      goal.description = goalDescInput.value.trim();
      goal.targetDate = targetDate;
      goal.targetAmount = targetAmount;
    } else {
      state.goals.push({
        id: uid(),
        title,
        description: goalDescInput.value.trim(),
        targetDate,
        targetAmount,
        createdAt: new Date().toISOString(),
      });
    }
    saveState();
    closeSheet("sheet-goal");
    renderGoals();
    renderHabitGoalOptions();
    renderHome();
  });

  goalDeleteBtn.addEventListener("click", () => {
    if (!editingGoalId) return;
    state.goals = state.goals.filter(g => g.id !== editingGoalId);
    state.habits.forEach(h => { if (h.goalId === editingGoalId) h.goalId = null; });
    saveState();
    closeSheet("sheet-goal");
    renderGoals();
    renderHabitGoalOptions();
    renderHome();
  });

  /** 目標に紐づく習慣の「継続率(チェック日数 / 習慣作成からの経過日数)」の平均 */
  function goalCompletionRate(goalId) {
    const habits = state.habits.filter(h => h.goalId === goalId);
    if (habits.length === 0) return 0;
    const todayD = dateOnly(new Date());
    const rates = habits.map(h => {
      const created = dateOnly(new Date(h.createdAt));
      const days = Math.max(1, Math.round((todayD - created) / 86400000) + 1);
      return h.completedDates.length / days;
    });
    return rates.reduce((a, b) => a + b, 0) / rates.length;
  }

  /** 目標作成日以降に記録された収入の合計から、金額目標までの距離を算出 */
  function goalMoneyProgress(goal) {
    if (!goal.targetAmount) return null;
    const created = dateOnly(new Date(goal.createdAt));
    const total = state.incomes
      .filter(inc => dateOnly(parseDateKey(inc.date)) >= created)
      .reduce((sum, inc) => sum + inc.amount, 0);
    return {
      total,
      remaining: Math.max(0, goal.targetAmount - total),
      pct: Math.min(1, total / goal.targetAmount),
    };
  }

  /** 期限日までの残り日数(過ぎていれば負の値) */
  function goalDaysRemaining(goal) {
    if (!goal.targetDate) return null;
    const todayD = dateOnly(new Date());
    const target = dateOnly(new Date(goal.targetDate));
    return Math.round((target - todayD) / 86400000);
  }

  function goalDistanceHtml(goal) {
    const parts = [];
    const money = goalMoneyProgress(goal);
    if (money) {
      parts.push(`<div class="money-line">${formatYen(money.total)} / ${formatYen(goal.targetAmount)} (${formatYen(money.remaining)} to go)</div>`);
    }
    const days = goalDaysRemaining(goal);
    if (days !== null) {
      if (days < 0) parts.push(`<div class="days-line overdue">${Math.abs(days)} days overdue</div>`);
      else if (days === 0) parts.push(`<div class="days-line">Due today</div>`);
      else parts.push(`<div class="days-line">${days} days left</div>`);
    }
    return parts.length ? `<div class="goal-distance">${parts.join("")}</div>` : "";
  }

  /** 目標の「現在地」を0〜1で算出。金額目標 > 紐づく習慣の継続率 > 期限までの経過度合い、の優先順で採用 */
  function goalProgressRatio(goal) {
    const money = goalMoneyProgress(goal);
    if (money) return { ratio: money.pct, source: "Funding progress" };

    const linkedHabits = state.habits.filter(h => h.goalId === goal.id);
    if (linkedHabits.length > 0) return { ratio: Math.min(1, goalCompletionRate(goal.id)), source: "Habit consistency" };

    if (goal.targetDate) {
      const created = dateOnly(new Date(goal.createdAt)).getTime();
      const target = dateOnly(new Date(goal.targetDate)).getTime();
      const now = dateOnly(new Date()).getTime();
      if (target <= created) return null;
      return { ratio: Math.max(0, Math.min(1, (now - created) / (target - created))), source: "Time elapsed" };
    }
    return null;
  }

  /** ゴールへの道のりをRPG風のクエストマップ(SVG)として描画 */
  function questMapSvg(ratio, sourceLabel) {
    const r = Math.max(0, Math.min(1, ratio));
    const P0 = { x: 22, y: 110 }, P1 = { x: 108, y: 134 }, P2 = { x: 206, y: 12 }, P3 = { x: 294, y: 24 };
    const bez = (t) => {
      const mt = 1 - t;
      return {
        x: mt * mt * mt * P0.x + 3 * mt * mt * t * P1.x + 3 * mt * t * t * P2.x + t * t * t * P3.x,
        y: mt * mt * mt * P0.y + 3 * mt * mt * t * P1.y + 3 * mt * t * t * P2.y + t * t * t * P3.y,
      };
    };
    const pathD = `M ${P0.x} ${P0.y} C ${P1.x} ${P1.y}, ${P2.x} ${P2.y}, ${P3.x} ${P3.y}`;
    const checkpoints = [0.2, 0.4, 0.6, 0.8].map(t => ({ t, ...bez(t) }));
    const start = bez(0);
    const end = bez(1);
    const cur = bez(r);

    const checkpointDots = checkpoints.map(cp => {
      const done = r >= cp.t - 0.001;
      return `<circle cx="${cp.x}" cy="${cp.y}" r="6" style="fill:${done ? "var(--success)" : "var(--surface)"};stroke:${done ? "var(--success)" : "var(--border)"};stroke-width:2" />`;
    }).join("");

    return `
      <div class="quest-map">
        <svg viewBox="0 0 320 150" class="quest-svg" preserveAspectRatio="xMidYMid meet">
          <path d="${pathD}" style="fill:none;stroke:var(--border);stroke-width:3;stroke-dasharray:2 7;stroke-linecap:round" />
          ${checkpointDots}
          <text x="${start.x}" y="${start.y + 22}" font-size="16" text-anchor="middle">🚩</text>
          <text x="${end.x}" y="${end.y - 10}" font-size="20" text-anchor="middle">🏁</text>
          <circle cx="${cur.x}" cy="${cur.y}" r="9" style="fill:var(--accent);stroke:var(--surface);stroke-width:2.5" />
          <text x="${cur.x}" y="${cur.y - 16}" font-size="12" font-weight="700" text-anchor="middle" style="fill:var(--accent)">${Math.round(r * 100)}%</text>
        </svg>
        <div class="quest-map-caption">Position based on ${sourceLabel}</div>
      </div>
    `;
  }

  function renderGoals() {
    const list = document.getElementById("goal-list");
    list.innerHTML = "";
    document.getElementById("goal-empty").classList.toggle("show", state.goals.length === 0);

    state.goals.forEach(goal => {
      const li = document.createElement("li");
      li.className = "goal-row";
      const progress = goalProgressRatio(goal);
      const progressHtml = progress
        ? questMapSvg(progress.ratio, progress.source)
        : `<div class="bar-track"><div class="bar-fill" style="width:0%"></div></div>`;
      li.innerHTML = `
        <div class="goal-title"></div>
        <div class="goal-desc"></div>
        ${goalDistanceHtml(goal)}
        ${progressHtml}
      `;
      li.querySelector(".goal-title").textContent = goal.title;
      const descEl = li.querySelector(".goal-desc");
      if (goal.description) {
        descEl.textContent = goal.description;
      } else {
        descEl.remove();
      }
      li.addEventListener("click", () => openGoalSheet(goal));
      list.appendChild(li);
    });
  }

  // ---------------------------------------------------------------------
  // HABITS (spreadsheet-style daily checklist)
  // ---------------------------------------------------------------------
  const habitTitleInput = document.getElementById("habit-title-input");
  const habitGoalSelect = document.getElementById("habit-goal-select");
  const habitDeleteBtn = document.getElementById("habit-delete-btn");
  const habitSheetTitle = document.getElementById("habit-sheet-title");
  let editingHabitId = null;

  let viewMonth = { year: today.getFullYear(), month: today.getMonth() };

  function renderHabitGoalOptions() {
    habitGoalSelect.innerHTML = '<option value="">None</option>';
    state.goals.forEach(g => {
      const opt = document.createElement("option");
      opt.value = g.id;
      opt.textContent = g.title;
      habitGoalSelect.appendChild(opt);
    });
  }

  function openHabitSheet(habit) {
    editingHabitId = habit ? habit.id : null;
    habitSheetTitle.textContent = habit ? "Edit Habit" : "Add Habit";
    habitTitleInput.value = habit ? habit.title : "";
    habitGoalSelect.value = (habit && habit.goalId) || "";
    habitDeleteBtn.classList.toggle("hidden-field", !habit);
    openSheet("sheet-habit");
  }

  document.getElementById("habit-save-btn").addEventListener("click", () => {
    const title = habitTitleInput.value.trim();
    if (!title) return;
    const goalId = habitGoalSelect.value || null;

    if (editingHabitId) {
      const habit = state.habits.find(h => h.id === editingHabitId);
      habit.title = title;
      habit.goalId = goalId;
    } else {
      state.habits.push({
        id: uid(),
        title,
        goalId,
        createdAt: new Date().toISOString(),
        completedDates: [],
      });
    }
    saveState();
    closeSheet("sheet-habit");
    renderHabitsView();
    renderHome();
  });

  habitDeleteBtn.addEventListener("click", () => {
    if (!editingHabitId) return;
    state.habits = state.habits.filter(h => h.id !== editingHabitId);
    saveState();
    closeSheet("sheet-habit");
    renderHabitsView();
    renderHome();
  });

  function toggleHabitDate(habit, key) {
    const idx = habit.completedDates.indexOf(key);
    if (idx >= 0) habit.completedDates.splice(idx, 1); else habit.completedDates.push(key);
    saveState();
  }

  function daysInMonth(year, month) {
    const count = new Date(year, month + 1, 0).getDate();
    return Array.from({ length: count }, (_, i) => new Date(year, month, i + 1));
  }

  /** その日時点で存在していた習慣のうち、完了した数/全体数 */
  function dayStats(date, habits) {
    const key = dateKey(date);
    const existing = habits.filter(h => dateOnly(new Date(h.createdAt)) <= date);
    const done = existing.filter(h => h.completedDates.includes(key)).length;
    const total = existing.length;
    return { done, total, notDone: total - done, pct: total === 0 ? 0 : Math.round((done / total) * 100) };
  }

  /** 指定した月の日次データから習慣の月間平均達成率(%)を算出。当月の場合は今日までで計算 */
  function monthlyHabitRate(year, month) {
    const days = daysInMonth(year, month);
    const todayOnly = dateOnly(new Date());
    const relevantDays = days.filter(d => d <= todayOnly);
    let doneSum = 0, totalSum = 0;
    relevantDays.forEach(d => {
      const s = dayStats(dateOnly(d), state.habits);
      doneSum += s.done; totalSum += s.total;
    });
    return totalSum === 0 ? 0 : Math.round((doneSum / totalSum) * 100);
  }

  const habitGrid = document.getElementById("habit-grid-table");
  const habitGridScroll = document.getElementById("habit-grid-scroll");

  function buildHabitGrid() {
    const days = daysInMonth(viewMonth.year, viewMonth.month);
    const todayKeyStr = dateKey(new Date());
    const todayOnly = dateOnly(new Date());

    let head = '<thead><tr><th class="col-name">Habit</th>';
    days.forEach(d => {
      const isWeekend = d.getDay() === 0 || d.getDay() === 6;
      const isToday = dateKey(d) === todayKeyStr;
      head += `<th class="day-head${isWeekend ? " weekend" : ""}${isToday ? " today" : ""}">
        <span class="wd">${WEEKDAY_LABELS[d.getDay()]}</span><span class="dn">${d.getDate()}</span>
      </th>`;
    });
    head += "</tr></thead>";

    let body = "<tbody>";
    state.habits.forEach(habit => {
      body += `<tr><td class="col-name" data-habit-edit="${habit.id}"><span class="habit-name-text">${escapeHtml(habit.title)}</span></td>`;
      const createdOnly = dateOnly(new Date(habit.createdAt));
      days.forEach(d => {
        const key = dateKey(d);
        const done = habit.completedDates.includes(key);
        const isFuture = d > todayOnly;
        const beforeStart = d < createdOnly;
        const isToday = key === todayKeyStr;
        const classes = ["habit-cell-btn"];
        if (done) classes.push("done");
        if (isFuture || beforeStart) classes.push("future");
        if (isToday) classes.push("today-col");
        body += `<td><button type="button" class="${classes.join(" ")}" data-habit="${habit.id}" data-date="${key}"><span class="box">✓</span></button></td>`;
      });
      body += "</tr>";
    });
    body += "</tbody>";

    let foot = "<tfoot>";
    foot += '<tr><td class="col-name">Rate</td>' + days.map(d => `<td>${dayStats(dateOnly(d), state.habits).pct}%</td>`).join("") + "</tr>";
    foot += '<tr><td class="col-name">Done</td>' + days.map(d => `<td>${dayStats(dateOnly(d), state.habits).done}</td>`).join("") + "</tr>";
    foot += '<tr><td class="col-name">Left</td>' + days.map(d => `<td>${dayStats(dateOnly(d), state.habits).notDone}</td>`).join("") + "</tr>";
    foot += "</tfoot>";

    habitGrid.innerHTML = head + body + foot;
  }

  habitGrid.addEventListener("click", (e) => {
    const cellBtn = e.target.closest(".habit-cell-btn");
    if (cellBtn && !cellBtn.classList.contains("future")) {
      const habit = state.habits.find(h => h.id === cellBtn.dataset.habit);
      if (habit) {
        toggleHabitDate(habit, cellBtn.dataset.date);
        renderHabitsView();
        renderHome();
      }
      return;
    }
    const nameCell = e.target.closest(".col-name[data-habit-edit]");
    if (nameCell) {
      const habit = state.habits.find(h => h.id === nameCell.dataset.habitEdit);
      if (habit) openHabitSheet(habit);
    }
  });

  document.getElementById("month-prev").addEventListener("click", () => {
    viewMonth.month -= 1;
    if (viewMonth.month < 0) { viewMonth.month = 11; viewMonth.year -= 1; }
    renderHabitsView(true);
  });
  document.getElementById("month-next").addEventListener("click", () => {
    viewMonth.month += 1;
    if (viewMonth.month > 11) { viewMonth.month = 0; viewMonth.year += 1; }
    renderHabitsView(true);
  });

  function renderHabitsView(jumpToToday = false) {
    document.getElementById("month-label").textContent = formatMonthYear(viewMonth.year, viewMonth.month);
    document.getElementById("habit-count-stat").textContent = String(state.habits.length);

    const hasHabits = state.habits.length > 0;
    document.getElementById("habit-empty").classList.toggle("show", !hasHabits);
    document.querySelector(".grid-card").style.display = hasHabits ? "" : "none";
    document.getElementById("habit-chart-card").style.display = hasHabits ? "" : "none";
    if (!hasHabits) return;

    const prevScrollLeft = habitGridScroll.scrollLeft;
    buildHabitGrid();
    if (jumpToToday) {
      const todayHead = habitGrid.querySelector(".day-head.today");
      if (todayHead) todayHead.scrollIntoView({ inline: "center", block: "nearest" });
      else habitGridScroll.scrollLeft = 0;
    } else {
      habitGridScroll.scrollLeft = prevScrollLeft;
    }

    document.getElementById("habit-rate-stat").textContent = `${monthlyHabitRate(viewMonth.year, viewMonth.month)}%`;

    const days = daysInMonth(viewMonth.year, viewMonth.month);
    const todayOnly = dateOnly(new Date());
    const relevantDays = days.filter(d => d <= todayOnly);
    const points = relevantDays.map(d => {
      const s = dayStats(dateOnly(d), state.habits);
      return { date: d, rate: s.total === 0 ? 0 : s.done / s.total };
    });
    drawAreaChart(document.getElementById("habit-chart"), points, "--success");
  }

  window.addEventListener("resize", () => {
    if (views.habits.classList.contains("active")) renderHabitsView();
  });

  // ---------------------------------------------------------------------
  // INCOME (daily entries + monthly/yearly view)
  // ---------------------------------------------------------------------
  let incomeView = { year: today.getFullYear(), month: today.getMonth() };
  const incomeDateInput = document.getElementById("income-date-input");
  const incomeAmountInput = document.getElementById("income-amount-input");
  const incomeSourceInput = document.getElementById("income-source-input");
  const incomeDeleteBtn = document.getElementById("income-delete-btn");
  let editingIncomeId = null;

  function updateIncomeSourceOptions() {
    const datalist = document.getElementById("income-source-list");
    const unique = [...new Set(state.incomes.map(i => i.source).filter(Boolean))];
    datalist.innerHTML = unique.map(s => `<option value="${escapeHtml(s)}"></option>`).join("");
  }

  function openIncomeSheet(entry) {
    editingIncomeId = entry.id;
    incomeDateInput.value = entry.date;
    incomeAmountInput.value = entry.amount;
    incomeSourceInput.value = entry.source;
    updateIncomeSourceOptions();
    openSheet("sheet-income");
  }

  document.getElementById("income-save-btn").addEventListener("click", () => {
    if (!editingIncomeId) return;
    const entry = state.incomes.find(i => i.id === editingIncomeId);
    entry.date = incomeDateInput.value || entry.date;
    entry.amount = Number(incomeAmountInput.value) || 0;
    entry.source = incomeSourceInput.value.trim();
    saveState();
    closeSheet("sheet-income");
    renderIncomeView();
    renderGoals();
    renderHome();
  });

  // 常時表示のクイック入力フォーム
  const incomeQuickDate = document.getElementById("income-quick-date");
  const incomeQuickAmount = document.getElementById("income-quick-amount");
  const incomeQuickSource = document.getElementById("income-quick-source");
  incomeQuickDate.value = dateKey(new Date());

  function addIncomeQuick() {
    const amount = Number(incomeQuickAmount.value) || 0;
    if (amount <= 0) { incomeQuickAmount.focus(); return; }
    const date = incomeQuickDate.value || dateKey(new Date());
    const source = incomeQuickSource.value.trim();
    state.incomes.push({ id: uid(), date, amount, source });
    saveState();
    incomeQuickAmount.value = "";
    incomeQuickSource.value = "";
    incomeQuickAmount.focus();
    renderIncomeView();
    renderGoals();
    renderHome();
  }
  document.getElementById("income-quick-add-btn").addEventListener("click", addIncomeQuick);
  incomeQuickSource.addEventListener("keydown", (e) => { if (e.key === "Enter") addIncomeQuick(); });

  incomeDeleteBtn.addEventListener("click", () => {
    if (!editingIncomeId) return;
    state.incomes = state.incomes.filter(i => i.id !== editingIncomeId);
    saveState();
    closeSheet("sheet-income");
    renderIncomeView();
    renderGoals();
    renderHome();
  });

  document.getElementById("income-month-prev").addEventListener("click", () => {
    incomeView.month -= 1;
    if (incomeView.month < 0) { incomeView.month = 11; incomeView.year -= 1; }
    renderIncomeView();
  });
  document.getElementById("income-month-next").addEventListener("click", () => {
    incomeView.month += 1;
    if (incomeView.month > 11) { incomeView.month = 0; incomeView.year += 1; }
    renderIncomeView();
  });

  function renderLedgerList(listEl, emptyEl, entries, kind) {
    listEl.innerHTML = "";
    emptyEl.classList.toggle("show", entries.length === 0);
    entries.slice().sort((a, b) => b.date.localeCompare(a.date)).forEach(entry => {
      const li = document.createElement("li");
      li.className = "ledger-row";
      const d = parseDateKey(entry.date);
      const label = kind === "income" ? entry.source : entry.category;
      li.innerHTML = `
        <div class="ledger-date">${d.getMonth() + 1}/${d.getDate()}(${WEEKDAY_LABELS[d.getDay()]})</div>
        <div class="ledger-source"></div>
        <div class="ledger-amount ${kind}-amount">${kind === "income" ? "+" : "-"}${formatYen(entry.amount)}</div>
      `;
      li.querySelector(".ledger-source").textContent = label || (kind === "income" ? "Income" : "Expense");
      li.addEventListener("click", () => {
        if (kind === "income") openIncomeSheet(entry); else openExpenseSheet(entry);
      });
      listEl.appendChild(li);
    });
  }

  function renderIncomeView() {
    updateIncomeSourceOptions();
    document.getElementById("income-month-label").textContent = formatMonthYear(incomeView.year, incomeView.month);

    const monthEntries = state.incomes.filter(i => {
      const d = parseDateKey(i.date);
      return d.getFullYear() === incomeView.year && d.getMonth() === incomeView.month;
    });
    renderLedgerList(document.getElementById("income-list"), document.getElementById("income-empty"), monthEntries, "income");

    const monthTotal = monthEntries.reduce((s, i) => s + i.amount, 0);
    document.getElementById("income-month-total-stat").textContent = formatYen(monthTotal);
    document.getElementById("income-list-total").textContent = formatYen(monthTotal);

    const yearEntries = state.incomes.filter(i => parseDateKey(i.date).getFullYear() === incomeView.year);
    const yearTotal = yearEntries.reduce((s, i) => s + i.amount, 0);
    document.getElementById("income-year-total-label").textContent = `${incomeView.year} Total`;
    document.getElementById("income-year-total-stat").textContent = formatYen(yearTotal);

    const dailyAmounts = dailyAmountsForMonth(monthEntries, incomeView.year, incomeView.month);
    drawBarChart(document.getElementById("income-chart"), dailyAmounts, todayHighlightIndex(incomeView.year, incomeView.month), "--success", (i) => `${i + 1}`);
  }

  /** 指定月の日数分の配列を作り、entries を日ごとに合計する */
  function dailyAmountsForMonth(entries, year, month) {
    const count = new Date(year, month + 1, 0).getDate();
    return Array.from({ length: count }, (_, i) =>
      entries.filter(e => parseDateKey(e.date).getDate() === i + 1).reduce((s, e) => s + e.amount, 0)
    );
  }

  /** 表示中の月が今日を含む場合、今日の日(0始まり)のインデックスを返す。それ以外は -1 */
  function todayHighlightIndex(year, month) {
    const now = new Date();
    if (year !== now.getFullYear() || month !== now.getMonth()) return -1;
    return now.getDate() - 1;
  }

  window.addEventListener("resize", () => {
    if (views.income.classList.contains("active")) renderIncomeView();
  });

  // ---------------------------------------------------------------------
  // EXPENSE (daily entries + monthly/yearly view)
  // ---------------------------------------------------------------------
  let expenseView = { year: today.getFullYear(), month: today.getMonth() };
  const expenseDateInput = document.getElementById("expense-date-input");
  const expenseAmountInput = document.getElementById("expense-amount-input");
  const expenseCategoryInput = document.getElementById("expense-category-input");
  const expenseDeleteBtn = document.getElementById("expense-delete-btn");
  let editingExpenseId = null;

  const DEFAULT_EXPENSE_CATEGORIES = ["Groceries", "Rent", "Transport", "Utilities", "Entertainment", "Health", "Other"];

  function updateExpenseCategoryOptions() {
    const datalist = document.getElementById("expense-category-list");
    const used = state.expenses.map(e => e.category).filter(Boolean);
    const unique = [...new Set([...DEFAULT_EXPENSE_CATEGORIES, ...used])];
    datalist.innerHTML = unique.map(c => `<option value="${escapeHtml(c)}"></option>`).join("");
  }

  function openExpenseSheet(entry) {
    editingExpenseId = entry.id;
    expenseDateInput.value = entry.date;
    expenseAmountInput.value = entry.amount;
    expenseCategoryInput.value = entry.category;
    updateExpenseCategoryOptions();
    openSheet("sheet-expense");
  }

  document.getElementById("expense-save-btn").addEventListener("click", () => {
    if (!editingExpenseId) return;
    const entry = state.expenses.find(e => e.id === editingExpenseId);
    entry.date = expenseDateInput.value || entry.date;
    entry.amount = Number(expenseAmountInput.value) || 0;
    entry.category = expenseCategoryInput.value.trim();
    saveState();
    closeSheet("sheet-expense");
    renderExpenseView();
    renderHome();
  });

  // 常時表示のクイック入力フォーム
  const expenseQuickDate = document.getElementById("expense-quick-date");
  const expenseQuickAmount = document.getElementById("expense-quick-amount");
  const expenseQuickCategory = document.getElementById("expense-quick-category");
  expenseQuickDate.value = dateKey(new Date());

  function addExpenseQuick() {
    const amount = Number(expenseQuickAmount.value) || 0;
    if (amount <= 0) { expenseQuickAmount.focus(); return; }
    const date = expenseQuickDate.value || dateKey(new Date());
    const category = expenseQuickCategory.value.trim();
    state.expenses.push({ id: uid(), date, amount, category });
    saveState();
    expenseQuickAmount.value = "";
    expenseQuickCategory.value = "";
    expenseQuickAmount.focus();
    renderExpenseView();
    renderHome();
  }
  document.getElementById("expense-quick-add-btn").addEventListener("click", addExpenseQuick);
  expenseQuickCategory.addEventListener("keydown", (e) => { if (e.key === "Enter") addExpenseQuick(); });

  expenseDeleteBtn.addEventListener("click", () => {
    if (!editingExpenseId) return;
    state.expenses = state.expenses.filter(e => e.id !== editingExpenseId);
    saveState();
    closeSheet("sheet-expense");
    renderExpenseView();
    renderHome();
  });

  document.getElementById("expense-month-prev").addEventListener("click", () => {
    expenseView.month -= 1;
    if (expenseView.month < 0) { expenseView.month = 11; expenseView.year -= 1; }
    renderExpenseView();
  });
  document.getElementById("expense-month-next").addEventListener("click", () => {
    expenseView.month += 1;
    if (expenseView.month > 11) { expenseView.month = 0; expenseView.year += 1; }
    renderExpenseView();
  });

  function renderExpenseView() {
    updateExpenseCategoryOptions();
    document.getElementById("expense-month-label").textContent = formatMonthYear(expenseView.year, expenseView.month);

    const monthEntries = state.expenses.filter(e => {
      const d = parseDateKey(e.date);
      return d.getFullYear() === expenseView.year && d.getMonth() === expenseView.month;
    });
    renderLedgerList(document.getElementById("expense-list"), document.getElementById("expense-empty"), monthEntries, "expense");

    const monthTotal = monthEntries.reduce((s, e) => s + e.amount, 0);
    document.getElementById("expense-month-total-stat").textContent = formatYen(monthTotal);
    document.getElementById("expense-list-total").textContent = formatYen(monthTotal);

    const yearEntries = state.expenses.filter(e => parseDateKey(e.date).getFullYear() === expenseView.year);
    const yearTotal = yearEntries.reduce((s, e) => s + e.amount, 0);
    document.getElementById("expense-year-total-label").textContent = `${expenseView.year} Total`;
    document.getElementById("expense-year-total-stat").textContent = formatYen(yearTotal);

    const dailyAmounts = dailyAmountsForMonth(monthEntries, expenseView.year, expenseView.month);
    drawBarChart(document.getElementById("expense-chart"), dailyAmounts, todayHighlightIndex(expenseView.year, expenseView.month), "--danger", (i) => `${i + 1}`);
  }

  window.addEventListener("resize", () => {
    if (views.expense.classList.contains("active")) renderExpenseView();
  });

  function drawBarChart(canvas, values, highlightIndex, highlightColorVar, labelFn) {
    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    const w = rect.width || 320;
    const h = 180;
    canvas.width = w * dpr;
    canvas.height = h * dpr;
    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, w, h);

    const padding = { top: 12, right: 8, bottom: 18, left: 44 };
    const plotW = w - padding.left - padding.right;
    const plotH = h - padding.top - padding.bottom;
    const maxVal = Math.max(1, ...values);

    const styles = getComputedStyle(document.documentElement);
    const highlight = styles.getPropertyValue(highlightColorVar).trim() || "#6c4fd6";
    const gridColor = styles.getPropertyValue("--border").trim() || "#e7e3f5";
    const labelColor = styles.getPropertyValue("--text-secondary").trim() || "#6b6580";

    ctx.strokeStyle = gridColor;
    ctx.lineWidth = 1;
    ctx.font = "10px -apple-system, sans-serif";
    ctx.fillStyle = labelColor;
    ctx.textBaseline = "middle";
    [0, 0.5, 1].forEach(f => {
      const y = padding.top + plotH * (1 - f);
      ctx.beginPath();
      ctx.setLineDash([3, 3]);
      ctx.moveTo(padding.left, y);
      ctx.lineTo(w - padding.right, y);
      ctx.stroke();
      ctx.setLineDash([]);
      const label = f === 0 ? "¥0" : `¥${Math.round((maxVal * f) / 1000).toLocaleString("ja-JP")}k`;
      ctx.fillText(label, 2, y);
    });

    const count = values.length;
    const barSlot = plotW / count;
    const barWidth = Math.max(2, barSlot * (count > 20 ? 0.65 : 0.55));
    const labelStep = Math.max(1, Math.ceil(count / 8));
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    values.forEach((v, i) => {
      const barH = maxVal === 0 ? 0 : (v / maxVal) * plotH;
      const x = padding.left + barSlot * i + (barSlot - barWidth) / 2;
      const y = padding.top + plotH - barH;
      ctx.fillStyle = i === highlightIndex ? highlight : hexToRgba(highlight, 0.4);
      ctx.beginPath();
      const r = Math.min(4, barWidth / 2);
      ctx.moveTo(x, y + barH);
      ctx.lineTo(x, y + r);
      ctx.arcTo(x, y, x + r, y, r);
      ctx.lineTo(x + barWidth - r, y);
      ctx.arcTo(x + barWidth, y, x + barWidth, y + r, r);
      ctx.lineTo(x + barWidth, y + barH);
      ctx.closePath();
      if (i === highlightIndex) {
        ctx.shadowColor = hexToRgba(highlight, 0.7);
        ctx.shadowBlur = 10;
      }
      ctx.fill();
      ctx.shadowBlur = 0;

      if (i % labelStep === 0 || i === count - 1) {
        ctx.fillStyle = labelColor;
        ctx.font = "9.5px -apple-system, sans-serif";
        ctx.fillText(labelFn(i), x + barWidth / 2, padding.top + plotH + 4);
      }
    });
  }

  // ---------------------------------------------------------------------
  // HOME (growth ring + all-time chart + 収支 + per-goal progress)
  // ---------------------------------------------------------------------
  function todayRate() {
    const todayOnly = dateOnly(new Date());
    const s = dayStats(todayOnly, state.habits);
    return s.total === 0 ? 0 : s.done / s.total;
  }

  function renderRing() {
    const rate = todayRate();
    const circumference = 2 * Math.PI * 84;
    const ringFg = document.getElementById("ring-fg");
    ringFg.style.strokeDashoffset = String(circumference * (1 - rate));
    document.getElementById("ring-percent").textContent = `${Math.round(rate * 100)}%`;
  }

  function drawAreaChart(canvas, points, colorVarName) {
    if (points.length < 2) { canvas.parentElement.style.display = "none"; return; }
    canvas.parentElement.style.display = "block";

    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    const w = rect.width || 320;
    const h = 180;
    canvas.width = w * dpr;
    canvas.height = h * dpr;
    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, w, h);

    const padding = { top: 12, right: 10, bottom: 8, left: 34 };
    const plotW = w - padding.left - padding.right;
    const plotH = h - padding.top - padding.bottom;

    const styles = getComputedStyle(document.documentElement);
    const color = styles.getPropertyValue(colorVarName).trim() || "#2ea36b";
    const gridColor = styles.getPropertyValue("--border").trim() || "#e7e3f5";
    const labelColor = styles.getPropertyValue("--text-secondary").trim() || "#6b6580";

    ctx.strokeStyle = gridColor;
    ctx.lineWidth = 1;
    ctx.font = "10px -apple-system, sans-serif";
    ctx.fillStyle = labelColor;
    ctx.textBaseline = "middle";
    [0.25, 0.5, 0.75].forEach(f => {
      const y = padding.top + plotH * (1 - f);
      ctx.beginPath();
      ctx.setLineDash([3, 3]);
      ctx.moveTo(padding.left, y);
      ctx.lineTo(w - padding.right, y);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillText(`${Math.round(f * 100)}%`, 2, y);
    });

    const xFor = (i) => padding.left + (plotW * i) / (points.length - 1);
    const yFor = (rate) => padding.top + plotH * (1 - rate);

    ctx.beginPath();
    ctx.moveTo(xFor(0), yFor(points[0].rate));
    points.forEach((p, i) => ctx.lineTo(xFor(i), yFor(p.rate)));
    ctx.lineTo(xFor(points.length - 1), padding.top + plotH);
    ctx.lineTo(xFor(0), padding.top + plotH);
    ctx.closePath();
    const gradient = ctx.createLinearGradient(0, padding.top, 0, padding.top + plotH);
    gradient.addColorStop(0, hexToRgba(color, 0.35));
    gradient.addColorStop(1, hexToRgba(color, 0));
    ctx.fillStyle = gradient;
    ctx.fill();

    ctx.beginPath();
    points.forEach((p, i) => {
      const x = xFor(i), y = yFor(p.rate);
      if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
    });
    ctx.strokeStyle = color;
    ctx.lineWidth = 2.5;
    ctx.lineJoin = "round";
    ctx.lineCap = "round";
    ctx.shadowColor = hexToRgba(color, 0.85);
    ctx.shadowBlur = 8;
    ctx.stroke();
    ctx.shadowBlur = 0;

    const lastX = xFor(points.length - 1), lastY = yFor(points[points.length - 1].rate);
    ctx.beginPath();
    ctx.arc(lastX, lastY, 4, 0, Math.PI * 2);
    ctx.fillStyle = color;
    ctx.shadowColor = hexToRgba(color, 0.9);
    ctx.shadowBlur = 10;
    ctx.fill();
    ctx.shadowBlur = 0;
  }

  function hexToRgba(hex, alpha) {
    const h = hex.replace("#", "");
    const bigint = parseInt(h.length === 3 ? h.split("").map(c => c + c).join("") : h, 16);
    const r = (bigint >> 16) & 255, g = (bigint >> 8) & 255, b = bigint & 255;
    return `rgba(${r},${g},${b},${alpha})`;
  }

  /** 指定年月の収入合計・支出合計 */
  function monthlyBalance(year, month) {
    const incomeTotal = state.incomes
      .filter(i => { const d = parseDateKey(i.date); return d.getFullYear() === year && d.getMonth() === month; })
      .reduce((s, i) => s + i.amount, 0);
    const expenseTotal = state.expenses
      .filter(e => { const d = parseDateKey(e.date); return d.getFullYear() === year && d.getMonth() === month; })
      .reduce((s, e) => s + e.amount, 0);
    return { incomeTotal, expenseTotal, net: incomeTotal - expenseTotal };
  }

  document.getElementById("view-report-btn").addEventListener("click", () => {
    const now = new Date();
    showMonthlyReport(now.getFullYear(), now.getMonth(), false);
    closeSheet("sheet-settings");
  });

  // --- Daily Quests(チェックすると消える。未達成のものは日をまたいでも残り続ける) ---
  function getTodayTasks() {
    return state.todayTasks;
  }

  function renderTodayTasks() {
    const tasks = getTodayTasks();
    const list = document.getElementById("today-task-list");
    list.innerHTML = "";
    document.getElementById("today-task-empty").classList.toggle("show", tasks.length === 0);
    tasks.forEach(task => {
      const li = document.createElement("li");
      li.className = "today-task-row";
      li.innerHTML = `<button type="button" class="today-task-check"><span class="box">✓</span></button><span class="today-task-text"></span>`;
      li.querySelector(".today-task-text").textContent = task.text;
      li.querySelector(".today-task-check").addEventListener("click", () => {
        state.todayTasks = state.todayTasks.filter(t => t.id !== task.id);
        saveState();
        renderTodayTasks();
      });
      list.appendChild(li);
    });
  }

  function addTodayTaskFromInput() {
    const input = document.getElementById("task-quick-input");
    const text = input.value.trim();
    if (!text) return;
    state.todayTasks.push({ id: uid(), text, date: dateKey(new Date()) });
    saveState();
    input.value = "";
    renderTodayTasks();
  }
  document.getElementById("task-quick-add-btn").addEventListener("click", addTodayTaskFromInput);
  document.getElementById("task-quick-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter") addTodayTaskFromInput();
  });

  // --- 今日の気持ち(消さずに記録として貯める。連携中ならGoogle Sheetsにも即時保存) ---
  function renderEmotionLog() {
    const list = document.getElementById("emotion-log-list");
    list.innerHTML = "";
    const entries = state.emotionLogs.slice().sort((a, b) => b.createdAt.localeCompare(a.createdAt)).slice(0, 30);
    document.getElementById("emotion-log-empty").classList.toggle("show", entries.length === 0);
    entries.forEach(entry => {
      const row = document.createElement("div");
      row.className = "emotion-log-row";
      const d = parseDateKey(entry.date);
      row.innerHTML = `<div class="emotion-log-date"></div><div class="emotion-log-text"></div>`;
      row.querySelector(".emotion-log-date").textContent = `${d.getMonth() + 1}/${d.getDate()}(${WEEKDAY_LABELS[d.getDay()]})`;
      row.querySelector(".emotion-log-text").textContent = entry.text;
      if (state.googleClientId) {
        const badge = document.createElement("div");
        badge.className = "emotion-log-sync" + (entry.syncedToSheets ? " synced" : "");
        badge.textContent = entry.syncedToSheets ? "☁️ Saved to Sheets" : "Not synced";
        row.appendChild(badge);
      }
      list.appendChild(row);
    });
  }

  function showEmotionSyncFeedback(text) {
    const el = document.getElementById("emotion-sync-feedback");
    el.textContent = text;
    el.classList.toggle("show", !!text);
  }

  document.getElementById("emotion-save-btn").addEventListener("click", async () => {
    const input = document.getElementById("emotion-input");
    const text = input.value.trim();
    if (!text) return;
    const entry = { id: uid(), text, date: dateKey(new Date()), createdAt: new Date().toISOString(), syncedToSheets: false };
    state.emotionLogs.push(entry);
    saveState();
    input.value = "";
    renderEmotionLog();

    if (!state.googleClientId) return;

    const btn = document.getElementById("emotion-save-btn");
    btn.disabled = true;
    showEmotionSyncFeedback("Saving to Google Sheets…");
    const result = await syncEmotionEntryToSheets(entry);
    btn.disabled = false;
    if (result.ok) {
      entry.syncedToSheets = true;
      saveState();
      renderEmotionLog();
      showEmotionSyncFeedback("Saved to Google Sheets");
    } else {
      showEmotionSyncFeedback("Couldn't save to Google Sheets. You can retry later from Settings.");
    }
  });

  function renderHome() {
    renderRing();
    renderTodayTasks();
    renderEmotionLog();
  }

  // ---------------------------------------------------------------------
  // 月次レポート
  // ---------------------------------------------------------------------
  function groupSum(entries, keyFn) {
    const map = new Map();
    entries.forEach(e => {
      const key = keyFn(e) || "Other";
      map.set(key, (map.get(key) || 0) + e.amount);
    });
    return [...map.entries()].sort((a, b) => b[1] - a[1]);
  }

  function showMonthlyReport(year, month, isAutoPopup) {
    const { incomeTotal, expenseTotal, net } = monthlyBalance(year, month);
    const monthEntries = {
      incomes: state.incomes.filter(i => { const d = parseDateKey(i.date); return d.getFullYear() === year && d.getMonth() === month; }),
      expenses: state.expenses.filter(e => { const d = parseDateKey(e.date); return d.getFullYear() === year && d.getMonth() === month; }),
    };
    const habitRate = monthlyHabitRate(year, month);
    const bySource = groupSum(monthEntries.incomes, i => i.source);
    const byCategory = groupSum(monthEntries.expenses, e => e.category);

    const isCurrentMonth = year === today.getFullYear() && month === today.getMonth();
    document.getElementById("report-title").textContent = `${formatMonthYear(year, month)} Report${isCurrentMonth ? " (in progress)" : ""}`;

    let html = `
      <div class="report-summary">
        <div class="balance-item"><div class="balance-label">Income</div><div class="balance-value income">${formatYen(incomeTotal)}</div></div>
        <div class="balance-item"><div class="balance-label">Expense</div><div class="balance-value expense">${formatYen(expenseTotal)}</div></div>
        <div class="balance-item"><div class="balance-label">Net</div><div class="balance-value net ${net >= 0 ? "positive" : "negative"}">${formatYen(net)}</div></div>
      </div>
      <div class="report-section-title">Habit average this month</div>
      <div class="report-list-row"><span>${habitRate}%</span></div>
    `;

    html += `<div class="report-section-title">Income breakdown</div>`;
    html += bySource.length
      ? bySource.map(([name, total]) => `<div class="report-list-row"><span>${escapeHtml(name)}</span><span class="report-row-value">${formatYen(total)}</span></div>`).join("")
      : `<p class="empty-note show">No income this month</p>`;

    html += `<div class="report-section-title">Expense breakdown</div>`;
    html += byCategory.length
      ? byCategory.map(([name, total]) => `<div class="report-list-row"><span>${escapeHtml(name)}</span><span class="report-row-value">${formatYen(total)}</span></div>`).join("")
      : `<p class="empty-note show">No expenses this month</p>`;

    if (isAutoPopup) {
      html += `<p class="empty-note show" style="margin-top:16px">Last month just wrapped up, so here's its report automatically.</p>`;
    }

    document.getElementById("report-body").innerHTML = html;
    openSheet("sheet-report");
  }

  /** 前回アプリを開いてから月が変わっていたら、固定費(家賃・携帯代など)を新しい月の1日付けで自動追加する。
   *  初回起動時は過去分を遡って追加しないよう、今月を基準として記録するだけにする。 */
  function maybeAddRecurringExpenses() {
    const curKey = monthKeyNum(today.getFullYear(), today.getMonth());
    if (state.lastRecurringExpenseMonth === null) {
      state.lastRecurringExpenseMonth = curKey;
      saveState();
      return;
    }
    if (state.lastRecurringExpenseMonth === curKey) return;

    const dateStr = `${today.getFullYear()}-${pad2(today.getMonth() + 1)}-01`;
    RECURRING_MONTHLY_EXPENSES.forEach(item => {
      state.expenses.push({ id: uid(), date: dateStr, amount: item.amount, category: item.category });
    });
    state.lastRecurringExpenseMonth = curKey;
    saveState();
  }

  /** 前回アプリを開いてから月が変わっていたら、直前の月のレポートを自動表示する。
   *  ただしこれは「アプリを開いたタイミング」での判定であり、月末ちょうどに通知されるわけではない。 */
  function maybeShowMonthlyReport() {
    const curKey = monthKeyNum(today.getFullYear(), today.getMonth());
    if (state.lastSeenReportMonth === null) {
      state.lastSeenReportMonth = curKey;
      saveState();
      return;
    }
    if (state.lastSeenReportMonth === curKey) return;

    const prevKey = state.lastSeenReportMonth;
    const prevYear = Math.floor(prevKey / 12);
    const prevMonth = prevKey % 12;
    const hasData =
      state.incomes.some(i => { const d = parseDateKey(i.date); return d.getFullYear() === prevYear && d.getMonth() === prevMonth; }) ||
      state.expenses.some(e => { const d = parseDateKey(e.date); return d.getFullYear() === prevYear && d.getMonth() === prevMonth; }) ||
      state.habits.some(h => h.completedDates.some(k => { const d = parseDateKey(k); return d.getFullYear() === prevYear && d.getMonth() === prevMonth; }));

    state.lastSeenReportMonth = curKey;
    saveState();

    if (hasData) showMonthlyReport(prevYear, prevMonth, true);

    if (state.googleClientId && monthKeyNum(prevYear, prevMonth) > (state.lastSyncedMonth ?? -1)) {
      syncMonthToGoogleSheets(prevYear, prevMonth, { interactive: false });
    }
  }

  // ---------------------------------------------------------------------
  // Google Sheets 連携(収入・支出の取引明細を月ごとに自動送信)
  // ---------------------------------------------------------------------
  const GOOGLE_SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
  // Google Sheets側の表記は日本語で統一する(アプリ本体のUIは英語のまま)
  const SPREADSHEET_TITLE = "成長クエスト家計簿";
  const SHEET_TX = "取引";
  const SHEET_MOODS = "気持ち";
  const SHEET_SUMMARY = "サマリー";
  const TYPE_INCOME = "収入";
  const TYPE_EXPENSE = "支出";
  let googleTokenClient = null;
  let googleAccessToken = null;
  let googleTokenExpiresAt = 0;

  function isGisReady() {
    return typeof google !== "undefined" && google.accounts && google.accounts.oauth2;
  }

  function ensureTokenClient() {
    if (!isGisReady() || !state.googleClientId) return null;
    if (!googleTokenClient || googleTokenClient.__clientId !== state.googleClientId) {
      googleTokenClient = google.accounts.oauth2.initTokenClient({
        client_id: state.googleClientId,
        scope: GOOGLE_SHEETS_SCOPE,
        callback: () => {},
      });
      googleTokenClient.__clientId = state.googleClientId;
    }
    return googleTokenClient;
  }

  /** アクセストークンを取得する。interactive=false なら、既存のGoogleセッションがある場合のみ
   *  ポップアップなしで取得を試み、失敗したら null を返す(ブラウザのポップアップブロックを避けるため)。 */
  function requestGoogleToken(interactive) {
    return new Promise((resolve) => {
      const client = ensureTokenClient();
      if (!client) { resolve(null); return; }
      if (googleAccessToken && Date.now() < googleTokenExpiresAt) { resolve(googleAccessToken); return; }

      let settled = false;
      const finish = (value) => {
        if (settled) return;
        settled = true;
        clearTimeout(timeoutId);
        resolve(value);
      };
      // ポップアップがブロックされた場合など、コールバックが一切呼ばれないケースがあるため
      // 一定時間で必ず抜けられるようにする
      const timeoutId = setTimeout(() => finish(null), 20000);

      client.callback = (resp) => {
        if (!resp || resp.error) { finish(null); return; }
        googleAccessToken = resp.access_token;
        googleTokenExpiresAt = Date.now() + (Number(resp.expires_in || 3600) * 1000 - 60000);
        finish(googleAccessToken);
      };
      try {
        client.requestAccessToken({ prompt: interactive ? "consent" : "" });
      } catch (e) {
        finish(null);
      }
    });
  }

  async function ensureSpreadsheet(token) {
    if (state.googleSpreadsheetId) return state.googleSpreadsheetId;
    const res = await fetch("https://sheets.googleapis.com/v4/spreadsheets", {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        properties: { title: SPREADSHEET_TITLE },
        sheets: [
          { properties: { title: SHEET_SUMMARY } },
          { properties: { title: SHEET_MOODS } },
        ],
      }),
    });
    if (!res.ok) throw new Error(`create spreadsheet failed: ${res.status}`);
    const data = await res.json();
    state.googleSpreadsheetId = data.spreadsheetId;
    state.googleFeelingsSheetReady = true;
    state.googleSummarySheetReady = true;
    state.googleSheetMigratedJa = true; // 最初から日本語表記で作成しているので移行不要
    state.googleTxSplitByMonth = true; // 最初から月ごとのタブで作成しているので移行不要
    saveState();
    await appendRowsToRange(token, `${SHEET_MOODS}!A:B`, [["日時", "内容"]]);
    try { const now = new Date(); await syncMonthTransactionsSheet(token, now.getFullYear(), now.getMonth()); } catch (e) { console.error("initial transactions failed", e); }
    try { await syncSummarySheet(token); } catch (e) { console.error("initial summary failed", e); }
    try { await applySheetFormatting(token); } catch (e) { console.error("initial sheet styling failed", e); }
    return state.googleSpreadsheetId;
  }

  /** 既存のスプレッドシートに「気持ち」シートがまだ無ければ追加する(スプレッドシート作成時にも
   *  一緒に作られるので、通常は新規作成時以外は再チェックのみで済む) */
  async function ensureFeelingsSheetTab(token) {
    if (state.googleFeelingsSheetReady) return;
    const metaRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}?fields=sheets.properties.title`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!metaRes.ok) throw new Error(`get spreadsheet meta failed: ${metaRes.status}`);
    const meta = await metaRes.json();
    const hasFeelingsSheet = (meta.sheets || []).some(s => s.properties.title === SHEET_MOODS || s.properties.title === "Moods");
    if (!hasFeelingsSheet) {
      const batchRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}:batchUpdate`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ requests: [{ addSheet: { properties: { title: SHEET_MOODS } } }] }),
      });
      if (!batchRes.ok) throw new Error(`add feelings sheet failed: ${batchRes.status}`);
      await appendRowsToRange(token, `${SHEET_MOODS}!A:B`, [["日時", "内容"]]);
    }
    state.googleFeelingsSheetReady = true;
    saveState();
  }

  /** サマリータブがまだ無ければ、一番左のタブとして追加する */
  async function ensureSummarySheetTab(token) {
    if (state.googleSummarySheetReady) return;
    const metaRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}?fields=sheets.properties.title`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!metaRes.ok) throw new Error(`get spreadsheet meta failed: ${metaRes.status}`);
    const meta = await metaRes.json();
    const hasSummary = (meta.sheets || []).some(s => s.properties.title === SHEET_SUMMARY);
    if (!hasSummary) {
      const batchRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}:batchUpdate`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ requests: [{ addSheet: { properties: { title: SHEET_SUMMARY, index: 0 } } }] }),
      });
      if (!batchRes.ok) throw new Error(`add summary sheet failed: ${batchRes.status}`);
    }
    state.googleSummarySheetReady = true;
    saveState();
  }

  /** 古い英語表記(Transactions/Moods/Growth Quest Finances)で作られた既存のスプレッドシートを、
   *  タブ名・ヘッダー・種類の値(Income/Expense)ごと日本語に置き換える。一度だけ実行すればよい移行処理。 */
  async function migrateSheetNamesIfNeeded(token) {
    if (state.googleSheetMigratedJa) return;
    const metaRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}?fields=properties.title,sheets.properties(sheetId,title)`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!metaRes.ok) throw new Error(`get spreadsheet meta failed: ${metaRes.status}`);
    const meta = await metaRes.json();
    const sheetsByTitle = {};
    (meta.sheets || []).forEach(s => { sheetsByTitle[s.properties.title] = s.properties; });

    const renameRequests = [];
    if (sheetsByTitle["Transactions"]) {
      renameRequests.push({ updateSheetProperties: { properties: { sheetId: sheetsByTitle["Transactions"].sheetId, title: SHEET_TX }, fields: "title" } });
    }
    if (sheetsByTitle["Moods"]) {
      renameRequests.push({ updateSheetProperties: { properties: { sheetId: sheetsByTitle["Moods"].sheetId, title: SHEET_MOODS }, fields: "title" } });
    }
    if (meta.properties && meta.properties.title !== SPREADSHEET_TITLE) {
      renameRequests.push({ updateSpreadsheetProperties: { properties: { title: SPREADSHEET_TITLE }, fields: "title" } });
    }
    if (renameRequests.length > 0) {
      const renameRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}:batchUpdate`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ requests: renameRequests }),
      });
      if (!renameRes.ok) throw new Error(`rename sheets failed: ${renameRes.status}`);
    }

    // 旧「取引」タブ自体は migrateToMonthlyTransactionTabs が月ごとのタブへの移行時に削除するので、
    // ここではリネームだけ行えば十分。
    if (sheetsByTitle["Moods"]) {
      await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}/values/${encodeURIComponent(SHEET_MOODS + "!A1:B1")}?valueInputOption=USER_ENTERED`,
        { method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, body: JSON.stringify({ values: [["日時", "内容"]] }) }
      );
    }

    state.googleSheetMigratedJa = true;
    saveState();
  }

  /** 収入・支出を月別/年別/全期間で集計し、サマリータブの中身をまるごと書き直す */
  function buildSummaryRows() {
    const monthMap = new Map();
    const yearMap = new Map();
    const addTo = (map, key, income, expense, extra) => {
      const cur = map.get(key) || { income: 0, expense: 0, ...extra };
      cur.income += income;
      cur.expense += expense;
      map.set(key, cur);
    };
    state.incomes.forEach(i => {
      const d = parseDateKey(i.date);
      const y = d.getFullYear(), m = d.getMonth();
      addTo(monthMap, `${y}-${pad2(m + 1)}`, i.amount, 0, { year: y, month: m });
      addTo(yearMap, y, i.amount, 0, {});
    });
    state.expenses.forEach(e => {
      const d = parseDateKey(e.date);
      const y = d.getFullYear(), m = d.getMonth();
      addTo(monthMap, `${y}-${pad2(m + 1)}`, 0, e.amount, { year: y, month: m });
      addTo(yearMap, y, 0, e.amount, {});
    });

    const monthRows = [...monthMap.keys()].sort().map(k => {
      const v = monthMap.get(k);
      return [`${v.year}年${v.month + 1}月`, v.income, v.expense, v.income - v.expense];
    });
    const yearRows = [...yearMap.keys()].sort((a, b) => a - b).map(y => {
      const v = yearMap.get(y);
      return [`${y}年 合計`, v.income, v.expense, v.income - v.expense];
    });
    const totalIncome = state.incomes.reduce((s, i) => s + i.amount, 0);
    const totalExpense = state.expenses.reduce((s, e) => s + e.amount, 0);

    const rows = [["年月", "収入", "支出", "差額"]];
    rows.push(...monthRows);
    rows.push(["", "", "", ""]);
    rows.push(["年別合計", "", "", ""]);
    rows.push(...yearRows);
    rows.push(["", "", "", ""]);
    rows.push(["全期間合計", totalIncome, totalExpense, totalIncome - totalExpense]);
    return rows;
  }

  /** サマリータブを、今わかっている収入・支出の全データから毎回まるごと再生成する */
  async function syncSummarySheet(token) {
    const rows = buildSummaryRows();
    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}/values/${encodeURIComponent(SHEET_SUMMARY + "!A1:D1000")}:clear`,
      { method: "POST", headers: { Authorization: `Bearer ${token}` } }
    );
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}/values/${encodeURIComponent(SHEET_SUMMARY + "!A1")}?valueInputOption=USER_ENTERED`,
      { method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, body: JSON.stringify({ values: rows }) }
    );
    if (!res.ok) throw new Error(`summary update failed: ${res.status}`);
  }

  /** 収入(A〜C列)と支出(E〜G列)を左右に分けた2次元配列を組み立てる。
   *  1行目=区分見出し、2行目=列見出し、以降データ、最終行=合計・差額。 */
  function buildIncomeExpenseGrid(incomes, expenses) {
    const incomeRows = incomes.slice().sort((a, b) => a.date.localeCompare(b.date))
      .map(i => [i.date, i.source || "", i.amount]);
    const expenseRows = expenses.slice().sort((a, b) => a.date.localeCompare(b.date))
      .map(e => [e.date, e.category || "", e.amount]);
    const rowCount = Math.max(incomeRows.length, expenseRows.length);

    const grid = [
      [TYPE_INCOME, "", "", "", TYPE_EXPENSE, "", ""],
      ["日付", "内容", "金額", "", "日付", "内容", "金額"],
    ];
    for (let i = 0; i < rowCount; i++) {
      const inc = incomeRows[i] || ["", "", ""];
      const exp = expenseRows[i] || ["", "", ""];
      grid.push([...inc, "", ...exp]);
    }
    const totalIncome = incomes.reduce((s, i) => s + i.amount, 0);
    const totalExpense = expenses.reduce((s, e) => s + e.amount, 0);
    grid.push(["合計", "", totalIncome, "", "合計", "", totalExpense]);
    grid.push(["差額(収入-支出)", "", totalIncome - totalExpense, "", "", "", ""]);
    return grid;
  }

  /** "2026年8月" のようなタブ名を作る */
  function monthSheetTitle(year, month) { return `${year}年${month + 1}月`; }

  /** 指定の年月の収入・支出だけを抜き出したグリッドを組み立てる */
  function buildMonthGrid(year, month) {
    const incomes = state.incomes.filter(i => { const d = parseDateKey(i.date); return d.getFullYear() === year && d.getMonth() === month; });
    const expenses = state.expenses.filter(e => { const d = parseDateKey(e.date); return d.getFullYear() === year && d.getMonth() === month; });
    return buildIncomeExpenseGrid(incomes, expenses);
  }

  /** 指定タブの中身(A1起点、G列まで)をまるごとクリアしてから書き直す */
  async function writeGridToSheet(token, sheetTitle, grid) {
    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}/values/${encodeURIComponent(sheetTitle + "!A1:G3000")}:clear`,
      { method: "POST", headers: { Authorization: `Bearer ${token}` } }
    );
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}/values/${encodeURIComponent(sheetTitle + "!A1")}?valueInputOption=USER_ENTERED`,
      { method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, body: JSON.stringify({ values: grid }) }
    );
    if (!res.ok) throw new Error(`sheet "${sheetTitle}" update failed: ${res.status}`);
  }

  /** 指定タブがまだ無ければ追加する(月ごとの取引タブ用) */
  async function ensureNamedSheetTab(token, title) {
    const metaRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}?fields=sheets.properties.title`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!metaRes.ok) throw new Error(`get spreadsheet meta failed: ${metaRes.status}`);
    const meta = await metaRes.json();
    const exists = (meta.sheets || []).some(s => s.properties.title === title);
    if (!exists) {
      const batchRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}:batchUpdate`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ requests: [{ addSheet: { properties: { title } } }] }),
      });
      if (!batchRes.ok) throw new Error(`add sheet "${title}" failed: ${batchRes.status}`);
    }
  }

  /** 月ごとの取引タブに、ヘッダー色付け・見出し固定・合計行の強調を適用する */
  async function styleNamedTransactionsSheet(token, title, grid) {
    const props = await getSheetPropertiesMap(token);
    const sheet = props[title];
    if (!sheet) return;
    const requests = transactionsHeaderStyleRequests(sheet.sheetId, grid.length - 2);
    const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}:batchUpdate`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ requests }),
    });
    if (!res.ok) throw new Error(`style sheet "${title}" failed: ${res.status}`);
  }

  /** 指定の年月の取引タブを、用意→書き込み→色付けまでまとめて行う */
  async function syncMonthTransactionsSheet(token, year, month) {
    const title = monthSheetTitle(year, month);
    await ensureNamedSheetTab(token, title);
    const grid = buildMonthGrid(year, month);
    await writeGridToSheet(token, title, grid);
    await styleNamedTransactionsSheet(token, title, grid);
    return title;
  }

  /** 旧レイアウト(収入・支出が1本の「取引」タブに混在)を使っていた場合、
   *  既存データがある月ごとにタブを作り直し、古い「取引」タブは削除する。一度だけ実行すればよい移行処理。 */
  async function migrateToMonthlyTransactionTabs(token) {
    if (state.googleTxSplitByMonth) return;
    const monthKeys = new Set();
    state.incomes.forEach(i => { const d = parseDateKey(i.date); monthKeys.add(`${d.getFullYear()}-${d.getMonth()}`); });
    state.expenses.forEach(e => { const d = parseDateKey(e.date); monthKeys.add(`${d.getFullYear()}-${d.getMonth()}`); });
    for (const key of monthKeys) {
      const [y, m] = key.split("-").map(Number);
      await syncMonthTransactionsSheet(token, y, m);
    }
    const props = await getSheetPropertiesMap(token);
    if (props[SHEET_TX]) {
      const delRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}:batchUpdate`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ requests: [{ deleteSheet: { sheetId: props[SHEET_TX].sheetId } }] }),
      });
      if (!delRes.ok) throw new Error(`delete old transactions sheet failed: ${delRes.status}`);
    }
    state.googleTxSplitByMonth = true;
    saveState();
  }

  /** サインイン後や月次同期のたびに呼ぶ: 旧英語表記の移行・気持ち/サマリータブの用意までまとめて行う */
  async function ensureModernSheetLayout(token) {
    await migrateSheetNamesIfNeeded(token);
    await ensureFeelingsSheetTab(token);
    await ensureSummarySheetTab(token);
    await migrateToMonthlyTransactionTabs(token);
  }

  async function appendRowsToRange(token, sheetRange, rows) {
    const range = encodeURIComponent(sheetRange);
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}/values/${range}:append?valueInputOption=USER_ENTERED`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ values: rows }),
      }
    );
    if (!res.ok) throw new Error(`append failed: ${res.status}`);
    return res.json();
  }

  /** シートのタイトル→properties(sheetId, gridProperties 等)のマップを取得する */
  async function getSheetPropertiesMap(token) {
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}?fields=sheets.properties(sheetId,title,gridProperties)`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) throw new Error(`get spreadsheet meta failed: ${res.status}`);
    const meta = await res.json();
    const map = {};
    (meta.sheets || []).forEach(s => { map[s.properties.title] = s.properties; });
    return map;
  }

  /** ヘッダー行の色付け・固定・列幅を整える(何度呼んでも安全)リクエストを作る */
  function headerStyleRequests(sheetId, columnCount) {
    return [
      {
        repeatCell: {
          range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: columnCount },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 1, green: 0.427, blue: 0.161 },
              textFormat: { bold: true, fontSize: 11, foregroundColor: { red: 1, green: 1, blue: 1 } },
              horizontalAlignment: "CENTER",
              verticalAlignment: "MIDDLE",
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
        },
      },
      {
        updateSheetProperties: {
          properties: { sheetId, gridProperties: { frozenRowCount: 1 } },
          fields: "gridProperties.frozenRowCount",
        },
      },
      {
        updateDimensionProperties: {
          range: { sheetId, dimension: "COLUMNS", startIndex: 0, endIndex: columnCount },
          properties: { pixelSize: 150 },
          fields: "pixelSize",
        },
      },
    ];
  }

  /** 取引タブ専用: 収入(A〜C列・緑)と支出(E〜G列・赤)を左右に分けた2段見出しの色付けリクエストを作る */
  function transactionsHeaderStyleRequests(sheetId, totalsStartRowIndex) {
    const bandCell = (bg, fg) => ({
      userEnteredFormat: {
        backgroundColor: bg,
        textFormat: { bold: true, fontSize: 11, foregroundColor: fg },
        horizontalAlignment: "CENTER",
        verticalAlignment: "MIDDLE",
      },
    });
    const fields = "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)";
    const white = { red: 1, green: 1, blue: 1 };
    const requests = [
      { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 3 }, cell: bandCell({ red: 0.24, green: 0.62, blue: 0.38 }, white), fields } },
      { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 4, endColumnIndex: 7 }, cell: bandCell({ red: 0.85, green: 0.29, blue: 0.22 }, white), fields } },
      { repeatCell: { range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 3 }, cell: bandCell({ red: 0.86, green: 0.96, blue: 0.89 }, { red: 0.11, green: 0.35, blue: 0.19 }), fields } },
      { repeatCell: { range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 4, endColumnIndex: 7 }, cell: bandCell({ red: 1, green: 0.9, blue: 0.88 }, { red: 0.5, green: 0.1, blue: 0.06 }), fields } },
      { updateSheetProperties: { properties: { sheetId, gridProperties: { frozenRowCount: 2 } }, fields: "gridProperties.frozenRowCount" } },
      { updateDimensionProperties: { range: { sheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 3 }, properties: { pixelSize: 130 }, fields: "pixelSize" } },
      { updateDimensionProperties: { range: { sheetId, dimension: "COLUMNS", startIndex: 3, endIndex: 4 }, properties: { pixelSize: 24 }, fields: "pixelSize" } },
      { updateDimensionProperties: { range: { sheetId, dimension: "COLUMNS", startIndex: 4, endIndex: 7 }, properties: { pixelSize: 130 }, fields: "pixelSize" } },
    ];
    if (typeof totalsStartRowIndex === "number") {
      // 「合計」「差額」の2行を太字+薄い色の帯にして目立たせる
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: totalsStartRowIndex, endRowIndex: totalsStartRowIndex + 2, startColumnIndex: 0, endColumnIndex: 7 },
          cell: { userEnteredFormat: { backgroundColor: { red: 0.94, green: 0.94, blue: 0.94 }, textFormat: { bold: true } } },
          fields: "userEnteredFormat(backgroundColor,textFormat)",
        },
      });
    }
    return requests;
  }

  /** 見やすいデザインに整える: ヘッダーの色付け・見出し固定・列幅に加えて、
   *  Transactionsシートは収入=緑/支出=赤の行に、Moodsシートは1行おきの帯色をつける。
   *  帯・条件付き書式は重複登録を避けるため googleSheetStyled フラグで一度だけ実行する。 */
  async function applySheetFormatting(token) {
    const props = await getSheetPropertiesMap(token);
    const moods = props[SHEET_MOODS];
    const summary = props[SHEET_SUMMARY];
    const requests = [];

    // 月ごとの取引タブは syncMonthTransactionsSheet 内で個別に色付けされるので、ここでは対象外
    if (moods) requests.push(...headerStyleRequests(moods.sheetId, 2));
    if (summary) requests.push(...headerStyleRequests(summary.sheetId, 4));

    if (!state.googleSheetStyled) {
      if (moods) {
        requests.push({
          addBanding: {
            bandedRange: {
              range: { sheetId: moods.sheetId, startRowIndex: 1, startColumnIndex: 0, endColumnIndex: 2 },
              rowProperties: {
                firstBandColor: { red: 1, green: 1, blue: 1 },
                secondBandColor: { red: 1, green: 0.96, blue: 0.92 },
              },
            },
          },
        });
      }
    }

    if (requests.length === 0) return;
    const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${state.googleSpreadsheetId}:batchUpdate`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ requests }),
    });
    if (!res.ok) throw new Error(`format sheet failed: ${res.status}`);
    state.googleSheetStyled = true;
    saveState();
  }

  /** 「今日の気持ち」を記録した瞬間にGoogle Sheetsへ即時保存する */
  async function syncEmotionEntryToSheets(entry, { interactive } = {}) {
    if (!state.googleClientId) return { ok: false, reason: "no-client-id" };
    let token = await requestGoogleToken(false);
    if (!token && interactive !== false) token = await requestGoogleToken(true);
    if (!token) return { ok: false, reason: "auth-required" };

    try {
      await ensureSpreadsheet(token);
      await ensureModernSheetLayout(token);
      const d = new Date(entry.createdAt);
      const timestamp = `${entry.date} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
      await appendRowsToRange(token, `${SHEET_MOODS}!A:B`, [[timestamp, entry.text]]);
      return { ok: true };
    } catch (e) {
      console.error("emotion sync failed", e);
      return { ok: false, reason: "api-error" };
    }
  }

  /** 未送信のまま残っている気持ちの記録を、まとめて再送信する */
  async function syncPendingEmotionLogs({ interactive } = {}) {
    const pending = state.emotionLogs.filter(e => !e.syncedToSheets);
    let successCount = 0;
    for (const entry of pending) {
      const result = await syncEmotionEntryToSheets(entry, { interactive });
      if (!result.ok) break; // 認証切れなどの場合は以降も失敗するため打ち切る
      entry.syncedToSheets = true;
      successCount++;
    }
    if (successCount > 0) saveState();
    return { successCount, remaining: pending.length - successCount };
  }

  async function syncMonthToGoogleSheets(year, month, { interactive } = {}) {
    if (!state.googleClientId) return { ok: false, reason: "no-client-id" };
    const token = await requestGoogleToken(!!interactive);
    if (!token) return { ok: false, reason: "auth-required" };

    try {
      await ensureSpreadsheet(token);
      await ensureModernSheetLayout(token);
      await syncMonthTransactionsSheet(token, year, month);
      await syncSummarySheet(token);
      state.lastSyncedMonth = monthKeyNum(year, month);
      saveState();
      return { ok: true, rowCount: state.incomes.length + state.expenses.length };
    } catch (e) {
      console.error("Google Sheets sync failed", e);
      return { ok: false, reason: "api-error" };
    }
  }

  function updateGoogleStatusUI() {
    const box = document.getElementById("google-status-value");
    if (!state.googleClientId) {
      box.textContent = "Not connected (no client ID)";
      box.className = "settings-status-value";
    } else if (!state.googleSpreadsheetId) {
      box.textContent = "Client ID saved (not signed in)";
      box.className = "settings-status-value";
    } else {
      box.textContent = "Connected";
      box.className = "settings-status-value connected";
    }
  }

  function showGoogleFeedback(text) {
    const el = document.getElementById("google-feedback");
    el.textContent = text;
    el.classList.add("show");
  }

  document.getElementById("google-client-id-save-btn").addEventListener("click", () => {
    const value = document.getElementById("google-client-id-input").value.trim();
    state.googleClientId = value || null;
    googleTokenClient = null;
    googleAccessToken = null;
    saveState();
    updateGoogleStatusUI();
    showGoogleFeedback(value ? "Client ID saved" : "Client ID cleared");
  });

  document.getElementById("google-signin-btn").addEventListener("click", async () => {
    if (!state.googleClientId) { showGoogleFeedback("Save an OAuth Client ID first"); return; }
    if (!isGisReady()) { showGoogleFeedback("Couldn't load Google sign-in. Check your connection."); return; }
    showGoogleFeedback("Signing in…");
    const token = await requestGoogleToken(true);
    if (!token) { showGoogleFeedback("Google sign-in failed. Check for a blocked popup or an incorrect Client ID."); return; }
    try {
      await ensureSpreadsheet(token);
      await ensureModernSheetLayout(token);
      const now = new Date();
      await syncMonthTransactionsSheet(token, now.getFullYear(), now.getMonth());
      await syncSummarySheet(token);
      await applySheetFormatting(token);
      updateGoogleStatusUI();
      showGoogleFeedback("Signed in — your spreadsheet is ready");
    } catch (e) {
      showGoogleFeedback("Couldn't set up the spreadsheet");
    }
  });

  document.getElementById("google-sync-now-btn").addEventListener("click", async () => {
    const now = new Date();
    const btn = document.getElementById("google-sync-now-btn");
    btn.disabled = true;
    showGoogleFeedback("Syncing…");
    const monthResult = await syncMonthToGoogleSheets(now.getFullYear(), now.getMonth(), { interactive: true });
    const emotionResult = await syncPendingEmotionLogs({ interactive: true });
    btn.disabled = false;
    updateGoogleStatusUI();
    renderEmotionLog();
    if (monthResult.ok) {
      showGoogleFeedback(`Sheet updated (${monthResult.rowCount} transaction(s) total) and ${emotionResult.successCount} mood entries sent`);
    } else {
      showGoogleFeedback("Sync failed. Check your Client ID and sign-in status.");
    }
  });

  document.getElementById("google-format-btn").addEventListener("click", async () => {
    if (!state.googleSpreadsheetId) { showGoogleFeedback("Sign in first so a spreadsheet exists"); return; }
    const btn = document.getElementById("google-format-btn");
    btn.disabled = true;
    showGoogleFeedback("Styling sheet…");
    const token = await requestGoogleToken(true);
    if (!token) { btn.disabled = false; showGoogleFeedback("Google sign-in failed"); return; }
    try {
      await ensureModernSheetLayout(token);
      const now = new Date();
      await syncMonthTransactionsSheet(token, now.getFullYear(), now.getMonth());
      await syncSummarySheet(token);
      await applySheetFormatting(token);
      showGoogleFeedback("Sheet updated — Japanese labels, summary tab, and colors applied");
    } catch (e) {
      console.error("format sheet failed", e);
      showGoogleFeedback("Couldn't style the sheet. Try again in a moment.");
    }
    btn.disabled = false;
  });

  document.getElementById("google-disconnect-btn").addEventListener("click", () => {
    state.googleClientId = null;
    state.googleSpreadsheetId = null;
    state.googleFeelingsSheetReady = false;
    state.lastSyncedMonth = null;
    state.googleSheetStyled = false;
    state.googleSummarySheetReady = false;
    state.googleSheetMigratedJa = false;
    state.googleTxSplitByMonth = false;
    googleTokenClient = null;
    googleAccessToken = null;
    saveState();
    document.getElementById("google-client-id-input").value = "";
    updateGoogleStatusUI();
    renderEmotionLog();
    showGoogleFeedback("Disconnected");
  });

  // ---------------------------------------------------------------------
  // 日付が変わったら(アプリを開きっぱなしにしていても)「今日」を追従させる
  // ---------------------------------------------------------------------
  function checkDateRollover() {
    const now = new Date();
    if (dateKey(now) === dateKey(today)) return;

    const oldTodayKey = dateKey(today);
    today = now;

    [
      [incomeQuickDate, "value"],
      [expenseQuickDate, "value"],
    ].forEach(([input]) => {
      if (input.value === oldTodayKey) input.value = dateKey(now);
    });

    maybeAddRecurringExpenses();
    renderView(activeViewName());
    maybeShowMonthlyReport();
  }
  setInterval(checkDateRollover, 30000);

  // ---------------------------------------------------------------------
  // Service worker (offline support / installability)
  // ---------------------------------------------------------------------
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("service-worker.js").catch(() => {});
    });
  }

  // ---------------------------------------------------------------------
  // Init
  // ---------------------------------------------------------------------
  maybeAddRecurringExpenses();
  renderHabitGoalOptions();
  renderGoals();
  renderHabitsView();
  renderIncomeView();
  renderExpenseView();
  renderHome();
  showView("home");
  maybeShowMonthlyReport();
})();
