const apiUrl = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:3000/api";

export type User = {
  _id?: string;
  id?: string;
  firstName?: string;
  lastName?: string;
  email?: string;
  mobile?: string;
  roles?: string;
  workField?: WorkField;
  isActive?: boolean;
  score?: number;
};

export type WorkField = "it" | "human_resources" | "finance" | "sales" | "operations";
export type FixedTaskRecurrence = "daily" | "weekly" | "monthly";
export type DeadlineStatus = "overdue" | "within_deadline";

export type Task = {
  _id?: string;
  id?: string;
  title: string;
  description?: string;
  createdBy?: string | User;
  assignedTo?: Array<string | User>;
  projectId?: string | Project;
  status?: string;
  taskComment?: string;
  isPublic?: boolean;
  startDate?: string;
  dueDate?: string;
  createdAt?: string;
  updatedAt?: string;
};

export type Project = {
  _id?: string;
  id?: string;
  title: string;
  description?: string;
  owner?: string | User;
  members?: Array<string | User>;
  tasks?: Array<string | Task>;
  status?: string;
  workField?: WorkField;
  supervisorId?: string | User;
  assigneeId?: string | User;
  isArchived?: boolean;
  isActive?: boolean;
  isPublic?: boolean;
};

export type Notification = {
  _id?: string;
  id?: string;
  title: string;
  message: string;
  type?: string;
  isRead?: boolean;
  link?: string;
  createdAt?: string;
};

export type LeaveRequest = {
  _id?: string;
  id?: string;
  user?: string | User;
  startDate: string;
  endDate: string;
  reason?: string;
  status?: string;
  approvedBy?: string | User;
};

export type ManagerStats = {
  activeProjects: number;
  openTasks: number;
  activeUsers: number;
};

export type ProjectProgress = {
  projectId: string;
  projectName: string;
  totalTasks: number;
  completedTasks: number;
  inProgressTasks: number;
  pendingTasks: number;
  progressPercentage: number;
};

export type ProjectProgressItem = {
  projectId?: string;
  projectName?: string;
  title?: string;
  progressPercentage?: number;
  totalTasks?: number;
  completedTasks?: number;
  inProgressTasks?: number;
  pendingTasks?: number;
};

export type SupervisorStats = {
  supervisedProjects?: number;
  supervisedProjectsInProgressTasks?: number;
  participatingProjectsDoneTasks?: number;
  supervisorDoneTasks?: number;
  inProgressTasks?: number;
  successfulTasksInProjects?: number;
  successfulTasksAssignedToSupervisor?: number;
};

export type MemberPerformance = {
  userId?: string;
  _id?: string;
  id?: string;
  userName?: string;
  firstName?: string;
  lastName?: string;
  fullName?: string;
  taskCount?: number;
  totalTasks?: number;
  completedCount?: number;
  completedTasks?: number;
  todoTasks?: number;
  inProgressTasks?: number;
  doneTasks?: number;
  pendingTasks?: number;
  overdueTasks?: number;
  projectsCount?: number;
  projects?: string[];
  email?: string;
  mobile?: string;
  isActive?: boolean;
  completionRate?: number;
  score?: number;
};

export type ProjectReport = {
  totalTasks?: number;
  todoTasks?: number;
  inProgressTasks?: number;
  doneTasks?: number;
  overdueTasks?: number;
  overdueCount?: number;
  memberCount?: number;
  assigneeCount?: number;
  completionRate?: number;
};

export type TaskStatusOverview = {
  totalTasks?: number;
  todoTasks?: number;
  inProgressTasks?: number;
  doneTasks?: number;
  todo?: number;
  inProgress?: number;
  in_progress?: number;
  done?: number;
  total?: number;
};

export type UserTaskCount = {
  userId?: string;
  _id?: string;
  id?: string;
  userName?: string;
  firstName?: string;
  lastName?: string;
  fullName?: string;
  email?: string;
  totalTasks?: number;
  todoTasks?: number;
  inProgressTasks?: number;
  doneTasks?: number;
  todo?: number;
  inProgress?: number;
  in_progress?: number;
  done?: number;
  total?: number;
};

export type MonthlyPerformance = {
  userId?: string;
  _id?: string;
  id?: string;
  userName?: string;
  firstName?: string;
  lastName?: string;
  fullName?: string;
  month?: number;
  year?: number;
  email?: string;
  totalTasks?: number;
  completedTasks?: number;
  inProgressTasks?: number;
  pendingTasks?: number;
  completionRate?: number;
  score?: number;
};

export type MonthlyPerformanceResponse = {
  month?: number;
  year?: number;
  projectId?: string;
  users?: MonthlyPerformance[];
};

export type AuthResponse = { user: User; accessToken: string };

export type ProjectMember = {
  _id?: string;
  id?: string;
  project?: string | Project;
  user?: string | User;
  role?: string;
  isActive?: boolean;
};

export type ExcelFile = {
  _id?: string;
  id?: string;
  fileName?: string;
  originalName?: string;
  type?: string;
  status?: string;
  totalRows?: number;
  successRows?: number;
  errorRows?: number;
  createdAt?: string;
};

export type ExcelStatistics = {
  totalImports: number;
  totalExports: number;
  completedImports: number;
  failedImports: number;
  totalFiles: number;
};

export type FixedTask = {
  _id?: string;
  id?: string;
  title: string;
  assignedTo?: string | User;
  createdBy?: string | User;
  projectId?: string | Project;
  recurrence: FixedTaskRecurrence;
  description?: string;
  isActive?: boolean;
  lastGeneratedAt?: string;
  nextRunAt?: string;
  sourceExcel?: string;
  sourceSheet?: string;
  sourceRow?: number;
  createdAt?: string;
  updatedAt?: string;
};

export type IncompleteFixedTask = FixedTask & {
  deadline?: string;
  periodStart?: string;
  periodEnd?: string;
  deadlineStatus?: DeadlineStatus;
};

export function getId(item?: { _id?: string; id?: string } | string) {
  if (!item) return "";
  return typeof item === "string" ? item : item._id ?? item.id ?? "";
}

export function normalizeList<T>(v: T[] | { data?: T[]; items?: T[]; docs?: T[] }) {
  if (Array.isArray(v)) return v;
  return v.data ?? v.items ?? v.docs ?? [];
}

export async function request<T>(path: string, opts: RequestInit = {}, token?: string): Promise<T> {
  let res: Response;
  try {
    res = await fetch(`${apiUrl}${path}`, {
      ...opts,
      headers: {
        ...(opts.body instanceof FormData ? {} : { "Content-Type": "application/json" }),
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        ...opts.headers,
      },
    });
  } catch {
    throw new Error("اتصال به بک‌اند برقرار نشد.");
  }
  const body = await res.text();
  const data = body ? JSON.parse(body) : null;
  if (!res.ok) {
    const msg = data?.message;
    throw new Error(Array.isArray(msg) ? msg.join("، ") : msg ?? "درخواست ناموفق بود");
  }
  return data as T;
}

export async function downloadBlob(path: string, filename: string, token?: string, opts: RequestInit = {}) {
  let res: Response;
  try {
    res = await fetch(`${apiUrl}${path}`, {
      ...opts,
      headers: {
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        ...opts.headers,
      },
    });
  } catch {
    throw new Error("اتصال به بک‌اند برقرار نشد.");
  }
  if (!res.ok) {
    const body = await res.text();
    let msg = "دانلود ناموفق بود";
    try {
      const data = body ? JSON.parse(body) : null;
      msg = data?.message ?? msg;
    } catch { /* ignore */ }
    throw new Error(Array.isArray(msg) ? msg.join("، ") : msg);
  }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function qs(params?: Record<string, string | number | boolean | undefined>) {
  if (!params) return "";
  const q = Object.entries(params)
    .filter(([, v]) => v !== undefined && v !== "")
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`)
    .join("&");
  return q ? `?${q}` : "";
}

// ─── Auth ────────────────────────────────────────────────────────────────────
export const authApi = {
  login: (body: { mobile: string; password: string }) =>
    request<AuthResponse>("/auth/login", { method: "POST", body: JSON.stringify(body) }),
  register: (body: Record<string, string>) =>
    request<AuthResponse>("/auth/register", { method: "POST", body: JSON.stringify(body) }),
};

// ─── Users ───────────────────────────────────────────────────────────────────
export const userApi = {
  list: (token: string, page = 1, limit = 50) =>
    request<User[] | { data?: User[] }>(`/users${qs({ page, limit })}`, {}, token),
  get: (token: string, id: string) => request<User>(`/users/${id}`, {}, token),
  create: (token: string, body: Record<string, unknown>) =>
    request<User>("/users", { method: "POST", body: JSON.stringify(body) }, token),
  update: (token: string, id: string, body: Record<string, unknown>) =>
    request<User>(`/users/${id}`, { method: "PATCH", body: JSON.stringify(body) }, token),
  delete: (token: string, id: string) =>
    request(`/users/${id}`, { method: "DELETE" }, token),
  approve: (token: string, id: string) =>
    request(`/users/${id}/approve`, { method: "PATCH" }, token),
  increaseScore: (token: string, body: { userId: string; score: number }) =>
    request("/users/increase-score", { method: "POST", body: JSON.stringify(body) }, token),
};

// ─── Tasks ───────────────────────────────────────────────────────────────────
export const taskApi = {
  list: (token: string, params?: Record<string, string | number | undefined>) =>
    request<Task[] | { data?: Task[] }>(`/tasks${qs({ page: 1, limit: 50, ...params })}`, {}, token),
  get: (token: string, id: string) => request<Task>(`/tasks/${id}`, {}, token),
  create: (token: string, body: Record<string, unknown>) =>
    request<Task>("/tasks", { method: "POST", body: JSON.stringify(body) }, token),
  update: (token: string, id: string, body: Record<string, unknown>) =>
    request<Task>(`/tasks/${id}`, { method: "PATCH", body: JSON.stringify(body) }, token),
  delete: (token: string, id: string) =>
    request(`/tasks/${id}`, { method: "DELETE" }, token),
  updateStatus: (token: string, id: string, status: string) =>
    request(`/tasks/${id}/status`, { method: "PATCH", body: JSON.stringify({ status }) }, token),
  completionStats: (token: string, body: Record<string, unknown>) =>
    request("/tasks/completion-stats", { method: "POST", body: JSON.stringify(body) }, token),
  dateCount: (token: string, body: Record<string, unknown>) =>
    request("/tasks/date-count", { method: "POST", body: JSON.stringify(body) }, token),
};

// ─── Projects ────────────────────────────────────────────────────────────────
export const projectApi = {
  list: (token: string, params?: Record<string, string | number | undefined>) =>
    request<Project[] | { data?: Project[] }>(`/projects${qs(params)}`, {}, token),
  get: (token: string, id: string) => request<Project>(`/projects/${id}`, {}, token),
  byOwner: (token: string, ownerId: string) =>
    request<Project[]>(`/projects/owner/${ownerId}`, {}, token),
  create: (token: string, body: Record<string, unknown>) =>
    request<Project>("/projects", { method: "POST", body: JSON.stringify(body) }, token),
  update: (token: string, id: string, body: Record<string, unknown>) =>
    request<Project>(`/projects/${id}`, { method: "PATCH", body: JSON.stringify(body) }, token),
  delete: (token: string, id: string) =>
    request(`/projects/${id}`, { method: "DELETE" }, token),
  addTask: (token: string, projectId: string, taskId: string) =>
    request(`/projects/${projectId}/tasks/${taskId}`, { method: "PATCH" }, token),
  removeTask: (token: string, projectId: string, taskId: string) =>
    request(`/projects/${projectId}/tasks/${taskId}`, { method: "DELETE" }, token),
  progress: (token: string, id: string) =>
    request<ProjectProgress>(`/projects/${id}/progress`, {}, token),
};

// ─── Manager ─────────────────────────────────────────────────────────────────
export const managerApi = {
  statistics: (token: string) => request<ManagerStats>("/manager/statistics", {}, token),
  users: (token: string, params?: Record<string, string | number | boolean | undefined>) =>
    request<User[] | { data?: User[] }>(`/manager/users${qs(params)}`, {}, token),
  updateUserRole: (token: string, userId: string, role: string) =>
    request(`/manager/users/${userId}/role`, { method: "PATCH", body: JSON.stringify({ role }) }, token),
  setProjectActivation: (token: string, projectId: string, isActive: boolean) =>
    request(`/manager/projects/${projectId}/activation`, { method: "PATCH", body: JSON.stringify({ isActive }) }, token),
  projectAssignee: (token: string, projectId: string) =>
    request<User | null>(`/manager/projects/${projectId}/assignee`, {}, token),
  projectProgress: (token: string, projectId: string) =>
    request<ProjectProgress>(`/manager/projects/${projectId}/progress`, {}, token),
  projectsProgress: (token: string, params?: Record<string, string | number | undefined>) =>
    request<ProjectProgressItem[] | { data?: ProjectProgressItem[] }>(`/manager/projects/progress${qs(params)}`, {}, token),
  taskStatusOverview: (token: string, projectId?: string) =>
    request<TaskStatusOverview>(`/manager/tasks/status${qs(projectId ? { projectId } : {})}`, {}, token),
  taskCountsByUsers: (token: string, projectId?: string) =>
    request<UserTaskCount[] | { data?: UserTaskCount[] }>(`/manager/tasks/users/counts${qs(projectId ? { projectId } : {})}`, {}, token),
  monthlyPerformance: (token: string, params?: Record<string, string | number | undefined>) =>
    request<MonthlyPerformance[] | { data?: MonthlyPerformance[] } | MonthlyPerformanceResponse>(`/manager/users/monthly-performance${qs(params)}`, {}, token),
};

// ─── Supervisor ───────────────────────────────────────────────────────────────
export const supervisorApi = {
  statistics: (token: string) =>
    request<SupervisorStats>("/supervisor/statistics", {}, token),
  projects: (token: string, params?: Record<string, string | number | undefined>) =>
    request<Project[] | { data?: Project[] }>(`/supervisor/projects${qs(params)}`, {}, token),
  assignProjectSpecialist: (token: string, projectId: string, assigneeId: string) =>
    request(`/supervisor/projects/${projectId}/assignee`, { method: "PATCH", body: JSON.stringify({ assigneeId }) }, token),
  projectAssignee: (token: string, projectId: string) =>
    request<User | null>(`/supervisor/projects/${projectId}/assignee`, {}, token),
  projectAssigneePerformance: (token: string, projectId: string) =>
    request<MemberPerformance | MemberPerformance[] | { data?: MemberPerformance[] } | { members?: MemberPerformance[] }>(`/supervisor/projects/${projectId}/assignee/performance`, {}, token),
  updateTaskStatus: (token: string, projectId: string, taskId: string, status: string) =>
    request(`/supervisor/projects/${projectId}/tasks/${taskId}/status`, { method: "PATCH", body: JSON.stringify({ status }) }, token),
  overdueTasks: (token: string, params?: Record<string, string | number | undefined>) =>
    request<Task[] | { data?: Task[] }>(`/supervisor/tasks/overdue${qs(params)}`, {}, token),
  projectReport: (token: string, projectId: string) =>
    request<ProjectReport>(`/supervisor/projects/${projectId}/report`, {}, token),
  teamPerformance: (token: string) =>
    request<MemberPerformance[] | { data?: MemberPerformance[] } | { members?: MemberPerformance[] }>("/supervisor/team/performance", {}, token),
};

// ─── Notifications ───────────────────────────────────────────────────────────
export const notificationApi = {
  list: (token: string, params?: Record<string, string | number | boolean | undefined>) =>
    request<Notification[] | { data?: Notification[] }>(`/notifications/me${qs(params)}`, {}, token),
  unreadCount: (token: string) =>
    request<{ unreadCount: number }>("/notifications/me/unread-count", {}, token),
  get: (token: string, id: string) => request<Notification>(`/notifications/${id}`, {}, token),
  markRead: (token: string, id: string) =>
    request(`/notifications/${id}`, { method: "PATCH", body: JSON.stringify({ isRead: true }) }, token),
  markAllRead: (token: string) =>
    request("/notifications/me/read-all", { method: "PATCH" }, token),
  delete: (token: string, id: string) =>
    request(`/notifications/${id}`, { method: "DELETE" }, token),
  deleteRead: (token: string) =>
    request("/notifications/me/read", { method: "DELETE" }, token),
};

// ─── Fixed Tasks ─────────────────────────────────────────────────────────────
export const fixedTaskApi = {
  list: (token: string, params?: Record<string, string | number | boolean | undefined>) =>
    request<FixedTask[] | { data?: FixedTask[] }>(`/fixed-tasks${qs(params)}`, {}, token),
  get: (token: string, id: string) => request<FixedTask>(`/fixed-tasks/${id}`, {}, token),
  create: (token: string, body: Record<string, unknown>) =>
    request<FixedTask>("/fixed-tasks", { method: "POST", body: JSON.stringify(body) }, token),
  update: (token: string, id: string, body: Record<string, unknown>) =>
    request<FixedTask>(`/fixed-tasks/${id}`, { method: "PATCH", body: JSON.stringify(body) }, token),
  delete: (token: string, id: string) =>
    request(`/fixed-tasks/${id}`, { method: "DELETE" }, token),
  incompleteReport: (token: string, params?: Record<string, string | number | boolean | undefined>) =>
    request<IncompleteFixedTask[] | { data?: IncompleteFixedTask[] }>(`/fixed-tasks/reports/incomplete${qs(params)}`, {}, token),
};

// ─── Leave Requests ──────────────────────────────────────────────────────────
export const leaveApi = {
  list: (token: string, params?: Record<string, string | number | undefined>) =>
    request<LeaveRequest[] | { data?: LeaveRequest[] }>(`/leave-requests${qs(params)}`, {}, token),
  byUser: (token: string, userId: string) =>
    request<LeaveRequest[]>(`/leave-requests/user/${userId}`, {}, token),
  get: (token: string, id: string) => request<LeaveRequest>(`/leave-requests/${id}`, {}, token),
  create: (token: string, body: Record<string, unknown>) =>
    request<LeaveRequest>("/leave-requests", { method: "POST", body: JSON.stringify(body) }, token),
  update: (token: string, id: string, body: Record<string, unknown>) =>
    request<LeaveRequest>(`/leave-requests/${id}`, { method: "PATCH", body: JSON.stringify(body) }, token),
  delete: (token: string, id: string) =>
    request(`/leave-requests/${id}`, { method: "DELETE" }, token),
  approve: (token: string, id: string, approvedBy: string) =>
    request(`/leave-requests/${id}/approve`, { method: "POST", body: JSON.stringify({ approvedBy }) }, token),
  reject: (token: string, id: string, approvedBy: string, rejectionReason?: string) =>
    request(`/leave-requests/${id}/reject`, { method: "POST", body: JSON.stringify({ approvedBy, rejectionReason }) }, token),
};

// ─── Excel ───────────────────────────────────────────────────────────────────
export const excelApi = {
  list: (token: string, params?: Record<string, string | number | undefined>) =>
    request<ExcelFile[] | { data?: ExcelFile[] }>(`/excel${qs(params)}`, {}, token),
  get: (token: string, id: string) => request<ExcelFile>(`/excel/${id}`, {}, token),
  create: (token: string, body: Record<string, unknown>) =>
    request<ExcelFile>("/excel", { method: "POST", body: JSON.stringify(body) }, token),
  update: (token: string, id: string, body: Record<string, unknown>) =>
    request<ExcelFile>(`/excel/${id}`, { method: "PATCH", body: JSON.stringify(body) }, token),
  delete: (token: string, id: string) =>
    request(`/excel/${id}`, { method: "DELETE" }, token),
  upload: async (token: string, file: File, createdBy: string, type: "import" | "export" = "import") => {
    const form = new FormData();
    form.append("file", file);
    let res: Response;
    try {
      res = await fetch(`${apiUrl}/excel/upload?createdBy=${encodeURIComponent(createdBy)}&type=${type}`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}` },
        body: form,
      });
    } catch {
      throw new Error("اتصال به بک‌اند برقرار نشد.");
    }
    const body = await res.text();
    const data = body ? JSON.parse(body) : null;
    if (!res.ok) {
      const msg = data?.message;
      throw new Error(Array.isArray(msg) ? msg.join("، ") : msg ?? "آپلود ناموفق بود");
    }
    return data as ExcelFile;
  },
  download: (token: string, id: string, filename: string) =>
    downloadBlob(`/excel/${id}/download`, filename, token),
  process: (token: string, id: string) =>
    request(`/excel/${id}/process`, { method: "POST" }, token),
  statistics: (token: string, userId: string) =>
    request<ExcelStatistics>(`/excel/statistics/${userId}`, {}, token),
  generateExport: (token: string, body: Record<string, unknown>, filename = "export.xlsx") =>
    downloadBlob("/excel/export/generate", filename, token, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }),
};
