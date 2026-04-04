// --- Data Store ---
const STORAGE_KEYS = { tasks: "tm_tasks", members: "tm_members" };

function loadData(key) {
  try {
    return JSON.parse(localStorage.getItem(key)) || [];
  } catch {
    return [];
  }
}

function saveData(key, data) {
  localStorage.setItem(key, JSON.stringify(data));
}

let tasks = loadData(STORAGE_KEYS.tasks);
let members = loadData(STORAGE_KEYS.members);

// --- DOM References ---
const $ = (sel) => document.querySelector(sel);
const taskList = $("#task-list");
const emptyMessage = $("#empty-message");
const taskModal = $("#task-modal");
const memberModal = $("#member-modal");
const exportModal = $("#export-modal");
const taskForm = $("#task-form");

// --- Toast Notification ---
function showToast(message) {
  const existing = document.querySelector(".toast");
  if (existing) existing.remove();
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 2500);
}

// --- Members ---
function renderMemberOptions() {
  const selects = [$("#task-assignee"), $("#filter-assignee")];
  selects.forEach((sel) => {
    const currentVal = sel.value;
    const isFilter = sel.id === "filter-assignee";
    sel.innerHTML = isFilter
      ? '<option value="all">すべて</option>'
      : '<option value="">未割当</option>';
    members.forEach((m) => {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = m;
      sel.appendChild(opt);
    });
    sel.value = currentVal;
  });
}

function renderMemberList() {
  const list = $("#member-list");
  list.innerHTML = "";
  if (members.length === 0) {
    list.innerHTML =
      '<li style="color:var(--text-muted)">メンバーが登録されていません</li>';
    return;
  }
  members.forEach((m) => {
    const li = document.createElement("li");
    li.innerHTML = `<span>${escapeHtml(m)}</span>
      <button class="btn btn-danger btn-sm" data-member="${escapeHtml(m)}">削除</button>`;
    li.querySelector("button").addEventListener("click", () => {
      members = members.filter((x) => x !== m);
      saveData(STORAGE_KEYS.members, members);
      renderMemberList();
      renderMemberOptions();
      showToast(`${m} を削除しました`);
    });
    list.appendChild(li);
  });
}

// --- Task Rendering ---
function getPriorityOrder(p) {
  return { 高: 0, 中: 1, 低: 2 }[p] ?? 1;
}

function getStatusBadgeClass(status) {
  if (status === "完了") return "status-done";
  if (status === "進行中") return "status-wip";
  if (status === "レビュー待ち") return "status-review";
  return "";
}

function getPriorityBadgeClass(p) {
  if (p === "高") return "";
  if (p === "中") return "p-medium";
  return "p-low";
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

function isOverdue(deadline) {
  if (!deadline) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return new Date(deadline) < today;
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
}

function getFilteredAndSortedTasks() {
  const statusFilter = $("#filter-status").value;
  const assigneeFilter = $("#filter-assignee").value;
  const priorityFilter = $("#filter-priority").value;
  const sortBy = $("#sort-by").value;

  let filtered = tasks.filter((t) => {
    if (statusFilter !== "all" && t.status !== statusFilter) return false;
    if (assigneeFilter !== "all" && t.assignee !== assigneeFilter) return false;
    if (priorityFilter !== "all" && t.priority !== priorityFilter) return false;
    return true;
  });

  filtered.sort((a, b) => {
    switch (sortBy) {
      case "created-desc":
        return b.createdAt - a.createdAt;
      case "created-asc":
        return a.createdAt - b.createdAt;
      case "priority":
        return getPriorityOrder(a.priority) - getPriorityOrder(b.priority);
      case "deadline":
        if (!a.deadline) return 1;
        if (!b.deadline) return -1;
        return new Date(a.deadline) - new Date(b.deadline);
      case "estimate":
        return (b.estimate || 0) - (a.estimate || 0);
      default:
        return 0;
    }
  });

  return filtered;
}

function renderTasks() {
  const filtered = getFilteredAndSortedTasks();

  // Clear task cards but keep empty message
  taskList
    .querySelectorAll(".task-card")
    .forEach((card) => card.remove());

  if (filtered.length === 0) {
    emptyMessage.style.display = "block";
    emptyMessage.textContent =
      tasks.length === 0
        ? 'タスクがありません。「+ タスク追加」から始めましょう。'
        : "条件に一致するタスクがありません。";
  } else {
    emptyMessage.style.display = "none";
    filtered.forEach((task) => {
      taskList.appendChild(createTaskCard(task));
    });
  }

  updateSummary();
}

function createTaskCard(task) {
  const card = document.createElement("div");
  card.className = `task-card priority-${
    task.priority === "高" ? "high" : task.priority === "低" ? "low" : "medium"
  }${task.status === "完了" ? " status-done" : ""}`;

  const overdue = isOverdue(task.deadline) && task.status !== "完了";

  card.innerHTML = `
    <div class="task-header">
      <span class="task-title">${escapeHtml(task.title)}</span>
      <div class="task-actions">
        <button class="btn btn-secondary btn-sm btn-edit" data-id="${task.id}">編集</button>
        <button class="btn btn-danger btn-sm btn-delete" data-id="${task.id}">削除</button>
      </div>
    </div>
    ${task.action ? `<div class="task-description">${escapeHtml(task.action)}</div>` : ""}
    <div class="task-meta">
      <span class="task-badge badge-status ${getStatusBadgeClass(task.status)}">${escapeHtml(task.status)}</span>
      <span class="task-badge badge-priority ${getPriorityBadgeClass(task.priority)}">優先: ${escapeHtml(task.priority)}</span>
      ${task.category ? `<span class="task-badge badge-category">${escapeHtml(task.category)}</span>` : ""}
      ${task.assignee ? `<span class="task-assignee">担当: ${escapeHtml(task.assignee)}</span>` : ""}
      ${task.estimate ? `<span class="task-estimate">見積: ${task.estimate}h</span>` : ""}
      ${task.deadline ? `<span class="task-deadline ${overdue ? "overdue" : ""}">期限: ${formatDate(task.deadline)}${overdue ? " (期限超過)" : ""}</span>` : ""}
    </div>
  `;

  card.querySelector(".btn-edit").addEventListener("click", () => openEditTask(task.id));
  card.querySelector(".btn-delete").addEventListener("click", () => deleteTask(task.id));

  return card;
}

function updateSummary() {
  $("#total-count").textContent = tasks.length;
  $("#todo-count").textContent = tasks.filter((t) => t.status === "未着手").length;
  $("#wip-count").textContent = tasks.filter(
    (t) => t.status === "進行中" || t.status === "レビュー待ち"
  ).length;
  $("#done-count").textContent = tasks.filter((t) => t.status === "完了").length;

  const totalHours = tasks.reduce((sum, t) => sum + (t.estimate || 0), 0);
  $("#total-estimate").textContent = `${totalHours}h`;
}

// --- CRUD ---
function openAddTask() {
  taskForm.reset();
  $("#task-id").value = "";
  $("#modal-title").textContent = "タスク追加";
  taskModal.classList.add("active");
  $("#task-title").focus();
}

function openEditTask(id) {
  const task = tasks.find((t) => t.id === id);
  if (!task) return;
  $("#task-id").value = task.id;
  $("#task-title").value = task.title;
  $("#task-action").value = task.action || "";
  $("#task-assignee").value = task.assignee || "";
  $("#task-status").value = task.status;
  $("#task-priority").value = task.priority;
  $("#task-estimate").value = task.estimate || "";
  $("#task-deadline").value = task.deadline || "";
  $("#task-category").value = task.category || "";
  $("#modal-title").textContent = "タスク編集";
  taskModal.classList.add("active");
}

function saveTask(e) {
  e.preventDefault();
  const id = $("#task-id").value;
  const data = {
    title: $("#task-title").value.trim(),
    action: $("#task-action").value.trim(),
    assignee: $("#task-assignee").value,
    status: $("#task-status").value,
    priority: $("#task-priority").value,
    estimate: parseFloat($("#task-estimate").value) || 0,
    deadline: $("#task-deadline").value,
    category: $("#task-category").value,
  };

  if (!data.title) return;

  if (id) {
    const idx = tasks.findIndex((t) => t.id === id);
    if (idx !== -1) {
      tasks[idx] = { ...tasks[idx], ...data, updatedAt: Date.now() };
    }
    showToast("タスクを更新しました");
  } else {
    tasks.push({
      id: crypto.randomUUID(),
      ...data,
      createdAt: Date.now(),
      updatedAt: Date.now(),
    });
    showToast("タスクを追加しました");
  }

  saveData(STORAGE_KEYS.tasks, tasks);
  taskModal.classList.remove("active");
  renderTasks();
}

function deleteTask(id) {
  const task = tasks.find((t) => t.id === id);
  if (!task) return;
  if (!confirm(`「${task.title}」を削除しますか？`)) return;
  tasks = tasks.filter((t) => t.id !== id);
  saveData(STORAGE_KEYS.tasks, tasks);
  renderTasks();
  showToast("タスクを削除しました");
}

// --- Export / Import ---
function exportJSON() {
  const data = JSON.stringify({ tasks, members }, null, 2);
  navigator.clipboard.writeText(data).then(() => {
    showToast("JSONをクリップボードにコピーしました");
  });
}

function exportCSV() {
  const headers = [
    "タスク名",
    "アクション",
    "担当者",
    "ステータス",
    "優先度",
    "見積もり(h)",
    "期限",
    "カテゴリ",
  ];
  const rows = tasks.map((t) =>
    [
      t.title,
      t.action,
      t.assignee,
      t.status,
      t.priority,
      t.estimate,
      t.deadline,
      t.category,
    ]
      .map((v) => `"${String(v || "").replace(/"/g, '""')}"`)
      .join(",")
  );
  return "\uFEFF" + headers.join(",") + "\n" + rows.join("\n");
}

function copyCSV() {
  navigator.clipboard.writeText(exportCSV()).then(() => {
    showToast("CSVをクリップボードにコピーしました");
  });
}

function downloadCSV() {
  const blob = new Blob([exportCSV()], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `tasks_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast("CSVをダウンロードしました");
}

function importData() {
  const raw = $("#import-data").value.trim();
  if (!raw) return;
  try {
    const data = JSON.parse(raw);
    if (data.tasks && Array.isArray(data.tasks)) {
      const newCount = data.tasks.filter(
        (nt) => !tasks.some((t) => t.id === nt.id)
      ).length;
      data.tasks.forEach((nt) => {
        if (!tasks.some((t) => t.id === nt.id)) {
          tasks.push(nt);
        }
      });
      saveData(STORAGE_KEYS.tasks, tasks);
    }
    if (data.members && Array.isArray(data.members)) {
      data.members.forEach((m) => {
        if (!members.includes(m)) members.push(m);
      });
      saveData(STORAGE_KEYS.members, members);
      renderMemberOptions();
    }
    renderTasks();
    $("#import-data").value = "";
    exportModal.classList.remove("active");
    showToast("データをインポートしました");
  } catch {
    showToast("JSONの形式が正しくありません");
  }
}

// --- Event Listeners ---
// Task modal
$("#btn-add-task").addEventListener("click", openAddTask);
$("#modal-close").addEventListener("click", () =>
  taskModal.classList.remove("active")
);
$("#btn-cancel").addEventListener("click", () =>
  taskModal.classList.remove("active")
);
taskForm.addEventListener("submit", saveTask);

// Member modal
$("#btn-manage-members").addEventListener("click", () => {
  renderMemberList();
  memberModal.classList.add("active");
});
$("#member-modal-close").addEventListener("click", () =>
  memberModal.classList.remove("active")
);
$("#btn-add-member").addEventListener("click", () => {
  const name = $("#member-name").value.trim();
  if (!name) return;
  if (members.includes(name)) {
    showToast("同名のメンバーが既に存在します");
    return;
  }
  members.push(name);
  saveData(STORAGE_KEYS.members, members);
  $("#member-name").value = "";
  renderMemberList();
  renderMemberOptions();
  showToast(`${name} を追加しました`);
});

// Export modal
$("#btn-export").addEventListener("click", () =>
  exportModal.classList.add("active")
);
$("#export-modal-close").addEventListener("click", () =>
  exportModal.classList.remove("active")
);
$("#btn-copy-json").addEventListener("click", exportJSON);
$("#btn-copy-csv").addEventListener("click", copyCSV);
$("#btn-download-csv").addEventListener("click", downloadCSV);
$("#btn-import").addEventListener("click", importData);

// Filters
["filter-status", "filter-assignee", "filter-priority", "sort-by"].forEach(
  (id) => {
    $(`#${id}`).addEventListener("change", renderTasks);
  }
);

// Close modals on overlay click
[taskModal, memberModal, exportModal].forEach((modal) => {
  modal.addEventListener("click", (e) => {
    if (e.target === modal) modal.classList.remove("active");
  });
});

// Keyboard: Escape to close modals
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    [taskModal, memberModal, exportModal].forEach((m) =>
      m.classList.remove("active")
    );
  }
});

// --- Init ---
renderMemberOptions();
renderTasks();
