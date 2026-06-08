"use client";

import { DragDropContext, Draggable, Droppable, DropResult } from "@hello-pangea/dnd";
import {
  Activity,
  AlertCircle,
  AlertTriangle,
  Award,
  BarChart2,
  Bell,
  CalendarDays,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  CircleDashed,
  ClipboardList,
  Download,
  FileSpreadsheet,
  Flag,
  FolderKanban,
  GripVertical,
  LayoutDashboard,
  ListFilter,
  Loader2,
  LogIn,
  LogOut,
  Moon,
  MoreHorizontal,
  Plus,
  RefreshCw,
  Search,
  Settings,
  ShieldCheck,
  Sparkles,
  Sun,
  Target,
  Trash2,
  TrendingUp,
  Upload,
  UserPlus,
  UsersRound,
  X,
  Zap,
} from "lucide-react";
import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from "react";
import {
  authApi,
  excelApi,
  fixedTaskApi,
  getId,
  leaveApi,
  managerApi,
  normalizeList,
  notificationApi,
  projectApi,
  supervisorApi,
  taskApi,
  type FixedTask,
  type IncompleteFixedTask,
  userApi,
  type ExcelFile,
  type ExcelStatistics,
  type LeaveRequest,
  type ManagerStats,
  type MemberPerformance,
  type MonthlyPerformance,
  type Notification,
  type Project,
  type ProjectProgress,
  type ProjectProgressItem,
  type ProjectReport,
  type SupervisorStats,
  type Task,
  type TaskStatusOverview,
  type User,
  type UserTaskCount,
  type WorkField,
} from "@/lib/api";

// ─── Types ────────────────────────────────────────────────────────────────────
type Priority = "urgent" | "high" | "normal" | "low";
type TaskPeriod = "daily" | "weekly" | "monthly";
type View = "dashboard" | "projects" | "tasks" | "team" | "excel" | "settings" | "analytics" | "fixed-reports" | "supervisor-projects" | "supervisor-team";

const TASK_PERIODS: Array<[TaskPeriod, string]> = [
  ["daily", "روزانه"],
  ["weekly", "هفتگی"],
  ["monthly", "ماهانه"],
];

const WORK_FIELDS: Array<[WorkField, string]> = [
  ["operations", "عملیات"],
  ["it", "فناوری اطلاعات"],
  ["human_resources", "منابع انسانی"],
  ["finance", "مالی"],
  ["sales", "فروش"],
];

const COLUMNS = [
  {
    status: "todo", title: "در انتظار",
    dot: "bg-slate-400",
    colBg: "dark:from-slate-800/60 dark:to-slate-900/40 bg-gradient-to-b from-slate-50 to-white",
    border: "border-slate-200 dark:border-slate-700",
    headerGrad: "from-slate-100 to-slate-50 dark:from-slate-800 dark:to-slate-800/60",
    headerText: "text-slate-600 dark:text-slate-300",
    badge: "bg-slate-200 text-slate-600 dark:bg-slate-700 dark:text-slate-300",
    emptyBorder: "border-slate-200 dark:border-slate-700",
    cardBorder: "border-t-slate-300 dark:border-t-slate-500",
    btnHover: "hover:bg-slate-100 hover:text-slate-700 dark:hover:bg-slate-700 dark:hover:text-slate-200",
  },
  {
    status: "in_progress", title: "در حال انجام",
    dot: "bg-[#1f7a8c]",
    colBg: "dark:from-[#0f3040] dark:to-slate-900/40 bg-gradient-to-b from-[#e8f6f9] to-[#f5fbfd]",
    border: "border-[#b8dfe8] dark:border-[#1f5060]",
    headerGrad: "from-[#d0eef5] to-[#e8f6f9] dark:from-[#0f3040] dark:to-[#0f3040]/60",
    headerText: "text-[#1f7a8c] dark:text-[#4fc3d5]",
    badge: "bg-[#1f7a8c]/15 text-[#1f7a8c] dark:bg-[#1f7a8c]/20 dark:text-[#4fc3d5]",
    emptyBorder: "border-[#9dd4dc] dark:border-[#1f5060]",
    cardBorder: "border-t-[#1f7a8c] dark:border-t-[#2a9db2]",
    btnHover: "hover:bg-[#e0f4f8] hover:text-[#1f7a8c] dark:hover:bg-[#0f3040] dark:hover:text-[#4fc3d5]",
  },
  {
    status: "done", title: "تکمیل شده",
    dot: "bg-emerald-500",
    colBg: "dark:from-emerald-950/40 dark:to-slate-900/40 bg-gradient-to-b from-emerald-50 to-white",
    border: "border-emerald-200 dark:border-emerald-900",
    headerGrad: "from-emerald-100 to-emerald-50 dark:from-emerald-950/60 dark:to-emerald-950/30",
    headerText: "text-emerald-700 dark:text-emerald-400",
    badge: "bg-emerald-200/80 text-emerald-700 dark:bg-emerald-900/50 dark:text-emerald-400",
    emptyBorder: "border-emerald-200 dark:border-emerald-900",
    cardBorder: "border-t-emerald-400 dark:border-t-emerald-600",
    btnHover: "hover:bg-emerald-50 hover:text-emerald-700 dark:hover:bg-emerald-950/40 dark:hover:text-emerald-400",
  },
];

const PRIORITY: Record<Priority, { label: string; color: string; icon: string }> = {
  urgent: { label: "فوری", color: "text-red-600 bg-red-50 border border-red-200 dark:text-red-400 dark:bg-red-950/40 dark:border-red-900", icon: "🔴" },
  high:   { label: "بالا", color: "text-orange-600 bg-orange-50 border border-orange-200 dark:text-orange-400 dark:bg-orange-950/40 dark:border-orange-900", icon: "🟠" },
  normal: { label: "متوسط", color: "text-blue-600 bg-blue-50 border border-blue-200 dark:text-blue-400 dark:bg-blue-950/40 dark:border-blue-900", icon: "🔵" },
  low:    { label: "پایین", color: "text-slate-500 bg-slate-50 border border-slate-200 dark:text-slate-400 dark:bg-slate-800 dark:border-slate-700", icon: "⚪" },
};

const PROJECT_COVERS = [
  "from-[#1f7a8c] to-[#2aa3bd]",
  "from-violet-500 to-violet-400",
  "from-amber-500 to-orange-400",
  "from-emerald-500 to-teal-400",
  "from-rose-500 to-pink-400",
  "from-indigo-500 to-blue-400",
];

// ─── Helpers ──────────────────────────────────────────────────────────────────
function userName(user?: User | string) {
  if (!user) return "نامشخص";
  if (typeof user === "string") return user;
  return [user.firstName, user.lastName].filter(Boolean).join(" ") || user.mobile || user.email || getId(user);
}

function initials(user?: User | string) {
  const n = userName(user);
  if (!n || n === "نامشخص") return "؟";
  return n.split(" ").filter(Boolean).slice(0, 2).map((p) => p[0]).join("").toUpperCase();
}

function statusLabel(s?: string) {
  const m: Record<string, string> = { todo: "در انتظار", in_progress: "در حال انجام", done: "تکمیل شده", pending: "در انتظار", completed: "تکمیل شده" };
  return m[s ?? ""] ?? s ?? "نامشخص";
}

function workFieldLabel(field?: string) {
  return WORK_FIELDS.find(([value]) => value === field)?.[1] ?? field ?? "عملیات";
}

function nextStatus(s?: string) {
  if (s === "todo") return "in_progress";
  if (s === "in_progress") return "done";
  return "todo";
}

function formatDate(d?: string) {
  if (!d) return "—";
  try {
    return new Intl.DateTimeFormat("fa-IR", { month: "short", day: "numeric" }).format(new Date(d));
  } catch {
    return d;
  }
}

function taskDate(task: Task) {
  const rawDate = task.dueDate || task.createdAt;
  if (!rawDate) return null;
  const date = new Date(rawDate);
  return Number.isNaN(date.getTime()) ? null : date;
}

function isTaskInPeriod(task: Task, period: TaskPeriod) {
  const date = taskDate(task);
  if (!date) return false;

  const now = new Date();
  const start = new Date(now);
  start.setHours(0, 0, 0, 0);

  if (period === "weekly") {
    const day = start.getDay();
    const daysFromSaturday = (day + 1) % 7;
    start.setDate(start.getDate() - daysFromSaturday);
  }

  if (period === "monthly") {
    start.setDate(1);
  }

  const end = new Date(start);
  if (period === "daily") end.setDate(start.getDate() + 1);
  if (period === "weekly") end.setDate(start.getDate() + 7);
  if (period === "monthly") end.setMonth(start.getMonth() + 1);

  return date >= start && date < end;
}

function isUnassignedTask(task: Task) {
  return (task.assignedTo ?? []).length === 0;
}

// ─── Main Component ───────────────────────────────────────────────────────────
export default function Home() {
  const [token, setToken] = useState("");
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [authMode, setAuthMode] = useState<"login" | "register">("login");
  const [authLoading, setAuthLoading] = useState(false);
  const [loadingData, setLoadingData] = useState(false);
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");
  const [users, setUsers] = useState<User[]>([]);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [projects, setProjects] = useState<Project[]>([]);
  const [taskTitle, setTaskTitle] = useState("");
  const [taskAssignee, setTaskAssignee] = useState("");
  const [taskProjectId, setTaskProjectId] = useState("");
  const [taskQuery, setTaskQuery] = useState("");
  const [projectTitle, setProjectTitle] = useState("");
  const [projectDescription, setProjectDescription] = useState("");
  const [projectSupervisorId, setProjectSupervisorId] = useState("");
  const [projectAssigneeId, setProjectAssigneeId] = useState("");
  const [projectWorkField, setProjectWorkField] = useState<WorkField>("operations");
  const [showProjectForm, setShowProjectForm] = useState(false);

  // API data
  const [managerStats, setManagerStats] = useState<ManagerStats | null>(null);
  const [notifications, setNotifications] = useState<Notification[]>([]);
  const [unreadCount, setUnreadCount] = useState(0);
  const [leaveRequests, setLeaveRequests] = useState<LeaveRequest[]>([]);
  const [projectProgress, setProjectProgress] = useState<Record<string, ProjectProgress>>({});
  const [excelFiles, setExcelFiles] = useState<ExcelFile[]>([]);
  const [excelStats, setExcelStats] = useState<ExcelStatistics | null>(null);
  const [excelUploading, setExcelUploading] = useState(false);
  const [membersProject, setMembersProject] = useState<Project | null>(null);
  const [projectMembers, setProjectMembers] = useState<Array<{ user?: string | User; role?: string; id?: string }>>([]);
  const [fixedTasks, setFixedTasks] = useState<FixedTask[]>([]);
  const [incompleteFixedTasks, setIncompleteFixedTasks] = useState<IncompleteFixedTask[]>([]);
  const [fixedReportsTab, setFixedReportsTab] = useState<"templates" | "incomplete">("templates");
  const [showFixedTaskForm, setShowFixedTaskForm] = useState(false);
  const [editingFixedTask, setEditingFixedTask] = useState<FixedTask | null>(null);
  const [ftTitle, setFtTitle] = useState("");
  const [ftAssignee, setFtAssignee] = useState("");
  const [ftRecurrence, setFtRecurrence] = useState<"daily" | "weekly" | "monthly">("daily");
  const [ftDescription, setFtDescription] = useState("");
  const [ftProjectId, setFtProjectId] = useState("");

  // Manager analytics state
  const [managerTaskStatus, setManagerTaskStatus] = useState<TaskStatusOverview | null>(null);
  const [managerUserCounts, setManagerUserCounts] = useState<UserTaskCount[]>([]);
  const [managerMonthlyPerf, setManagerMonthlyPerf] = useState<MonthlyPerformance[]>([]);
  const [managerProjectsProgress, setManagerProjectsProgress] = useState<ProjectProgressItem[]>([]);

  // Supervisor state
  const [supervisorStats, setSupervisorStats] = useState<SupervisorStats | null>(null);
  const [supervisorProjects, setSupervisorProjects] = useState<any[]>([]);
  const [overdueTasks, setOverdueTasks] = useState<Task[]>([]);
  const [teamPerformance, setTeamPerformance] = useState<any>(null);
  const [selectedSupervisorProjectId, setSelectedSupervisorProjectId] = useState<string>("");
  const [supervisorProjectReport, setSupervisorProjectReport] = useState<ProjectReport | null>(null);
  const [supervisorProjectMembersPerf, setSupervisorProjectMembersPerf] = useState<MemberPerformance[]>([]);
  const [loadingSupervisorProject, setLoadingSupervisorProject] = useState(false);

  // UI state
  const [activeView, setActiveView] = useState<View>("dashboard");
  const [selectedProjectFilter, setSelectedProjectFilter] = useState("");
  const [selectedStatusFilter, setSelectedStatusFilter] = useState("");
  const [selectedAssigneeFilter, setSelectedAssigneeFilter] = useState("");
  const [selectedPriorityFilter, setSelectedPriorityFilter] = useState<Priority | "">("");
  const [selectedPeriodFilter, setSelectedPeriodFilter] = useState<TaskPeriod | "">("");
  const [showNotifications, setShowNotifications] = useState(false);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [darkMode, setDarkMode] = useState(false);
  const [selectedTask, setSelectedTask] = useState<Task | null>(null);
  const [selectedProject, setSelectedProject] = useState<Project | null>(null);
  const [taskPriorities, setTaskPriorities] = useState<Record<string, Priority>>({});
  const [leaveStart, setLeaveStart] = useState("");
  const [leaveEnd, setLeaveEnd] = useState("");
  const [leaveReason, setLeaveReason] = useState("");
  const [settingsFirstName, setSettingsFirstName] = useState("");
  const [settingsLastName, setSettingsLastName] = useState("");
  const [settingsEmail, setSettingsEmail] = useState("");

  useEffect(() => {
    document.documentElement.classList.toggle("dark", darkMode);
  }, [darkMode]);

  useEffect(() => {
    queueMicrotask(() => {
      const t = localStorage.getItem("taskino-token") ?? "";
      const u = localStorage.getItem("taskino-user");
      if (t) setToken(t);
      if (u) setCurrentUser(JSON.parse(u) as User);
    });
  }, []);

  const myId = getId(currentUser ?? undefined);
  const isManager = currentUser?.roles === "manager";
  const isSupervisor = currentUser?.roles === "supervisor";
  const isSpecialist = currentUser?.roles === "specialist";
  const doneTasks = tasks.filter((t) => t.status === "done").length;
  const activeTasks = managerStats?.openTasks ?? tasks.filter((t) => t.status !== "done").length;
  const inProgressTasks = tasks.filter((t) => t.status === "in_progress").length;
  const todoCount = tasks.filter((t) => (t.status ?? "todo") === "todo").length;
  const progress = tasks.length ? Math.round((doneTasks / tasks.length) * 100) : 0;
  const statsProjects = managerStats?.activeProjects ?? projects.length;
  const statsUsers = managerStats?.activeUsers ?? users.length;
  const supervisorInProgressReports = supervisorStats?.supervisedProjectsInProgressTasks ?? supervisorStats?.inProgressTasks ?? 0;
  const supervisorProjectDoneReports = supervisorStats?.participatingProjectsDoneTasks ?? supervisorStats?.successfulTasksInProjects ?? 0;
  const supervisorOwnDoneReports = supervisorStats?.supervisorDoneTasks ?? supervisorStats?.successfulTasksAssignedToSupervisor ?? 0;
  const teamAssignees = Array.isArray(teamPerformance?.assignees)
    ? teamPerformance.assignees
    : Array.isArray(teamPerformance?.members)
    ? teamPerformance.members
    : [];
  const teamAssigneeCount = teamPerformance?.assigneeCount ?? teamPerformance?.membersCount ?? teamAssignees.length;

  const filteredTasks = useMemo(() => {
    let list = tasks;
    if (selectedProjectFilter) {
      list = list.filter((t) => getId(t.projectId) === selectedProjectFilter);
    }
    if (selectedStatusFilter) {
      list = list.filter((t) => (t.status ?? "todo") === selectedStatusFilter);
    }
    if (selectedAssigneeFilter) {
      list = list.filter((t) => (t.assignedTo ?? []).some((u) => getId(u) === selectedAssigneeFilter));
    }
    if (selectedPriorityFilter) {
      list = list.filter((t) => taskPriorities[getId(t)] === selectedPriorityFilter);
    }
    if (selectedPeriodFilter) {
      list = list.filter((t) => isTaskInPeriod(t, selectedPeriodFilter));
    }
    const q = taskQuery.trim().toLowerCase();
    if (!q) return list;
    return list.filter((t) => `${t.title} ${(t.assignedTo ?? []).map(userName).join(" ")} ${statusLabel(t.status)}`.toLowerCase().includes(q));
  }, [taskPriorities, taskQuery, tasks, selectedAssigneeFilter, selectedPeriodFilter, selectedPriorityFilter, selectedProjectFilter, selectedStatusFilter]);

  const hasTaskFilters = Boolean(taskQuery || selectedProjectFilter || selectedStatusFilter || selectedAssigneeFilter || selectedPriorityFilter || selectedPeriodFilter);

  useEffect(() => {
    if (!error && !message) return;
    const id = window.setTimeout(() => { setError(""); setMessage(""); }, 3500);
    return () => window.clearTimeout(id);
  }, [error, message]);

  async function loadData(authToken = token) {
    if (!authToken) return;
    setLoadingData(true);
    setError("");
    try {
      const storedUser = JSON.parse(localStorage.getItem("taskino-user") ?? "{}") as User;
      const uid = myId || getId(storedUser);
      const currentRole = (currentUser ?? storedUser)?.roles;
      const userIsManager = currentRole === "manager";
      const userIsSupervisor = currentRole === "supervisor";
      const reportParams = !userIsManager && !userIsSupervisor && uid ? { assignedTo: uid } : undefined;
      const [u, t, p, leaves, statsRes, unreadRes, notifRes] = await Promise.all([
        userIsManager ? managerApi.users(authToken, { limit: 100 }).catch(() => userApi.list(authToken).catch(() => [])) : userApi.list(authToken).catch(() => []),
        taskApi.list(authToken, reportParams),
        projectApi.list(authToken),
        uid && !userIsManager
          ? leaveApi.list(authToken, { limit: 50, user: uid })
          : leaveApi.list(authToken, { limit: 50 }),
        managerApi.statistics(authToken).catch(() => null),
        notificationApi.unreadCount(authToken).catch(() => ({ unreadCount: 0 })),
        notificationApi.list(authToken, { isRead: false, limit: 20 }).catch(() => []),
      ]);
      const projectList = normalizeList(p);
      const taskList = normalizeList(t);
      setUsers(normalizeList(u));
      setTasks(taskList);
      const priorities: Record<string, Priority> = {};
      taskList.forEach((task) => {
        const match = task.taskComment?.match(/^priority:(urgent|high|normal|low)$/);
        if (match) priorities[getId(task)] = match[1] as Priority;
      });
      setTaskPriorities(priorities);
      setProjects(projectList);
      setLeaveRequests(normalizeList(leaves));
      setManagerStats(statsRes);
      setUnreadCount(unreadRes.unreadCount);
      setNotifications(normalizeList(notifRes as Notification[] | { data?: Notification[] }));
      if (projectList.length) {
        const progresses = await Promise.all(
          projectList.map((proj) =>
            projectApi.progress(authToken, getId(proj)).catch(() => null),
          ),
        );
        const map: Record<string, ProjectProgress> = {};
        progresses.forEach((prog) => { if (prog) map[prog.projectId] = prog; });
        setProjectProgress(map);
      }
      const role = (currentUser ?? storedUser)?.roles;
      if (role === "manager") void loadManagerAnalytics(authToken);
      if (role === "supervisor") void loadSupervisorData(authToken);
    } catch (err) {
      setError(err instanceof Error ? err.message : "دریافت اطلاعات ناموفق بود");
    } finally {
      setLoadingData(false);
    }
  }

  async function loadManagerAnalytics(authToken = token) {
    if (!authToken) return;
    try {
      const [taskStatus, userCounts, monthlyPerf, projectsProgress, recurring, incomplete] = await Promise.all([
        managerApi.taskStatusOverview(authToken).catch(() => null),
        managerApi.taskCountsByUsers(authToken).catch(() => []),
        managerApi.monthlyPerformance(authToken).catch(() => []),
        managerApi.projectsProgress(authToken, { limit: 50 }).catch(() => []),
        fixedTaskApi.list(authToken, { limit: 50 }).catch(() => []),
        fixedTaskApi.incompleteReport(authToken, { limit: 50 }).catch(() => []),
      ]);
      setManagerTaskStatus(taskStatus);
      setManagerUserCounts(normalizeList(userCounts as UserTaskCount[] | { data?: UserTaskCount[] }));
      const monthlyRaw = monthlyPerf as any;
      setManagerMonthlyPerf(monthlyRaw?.users ? monthlyRaw.users : normalizeList(monthlyRaw as MonthlyPerformance[] | { data?: MonthlyPerformance[] }));
      setManagerProjectsProgress(normalizeList(projectsProgress as ProjectProgressItem[] | { data?: ProjectProgressItem[] }));
      setFixedTasks(normalizeList(recurring as FixedTask[] | { data?: FixedTask[] }));
      setIncompleteFixedTasks(normalizeList(incomplete as IncompleteFixedTask[] | { data?: IncompleteFixedTask[] }));
    } catch {}
  }

  async function loadSupervisorData(authToken = token) {
    if (!authToken) return;
    try {
      const [stats, projects, overdue, team] = await Promise.all([
        supervisorApi.statistics(authToken).catch(() => null),
        supervisorApi.projects(authToken, { limit: 50 }).catch(() => []),
        supervisorApi.overdueTasks(authToken, { limit: 20 }).catch(() => []),
        supervisorApi.teamPerformance(authToken).catch(() => null),
      ]);
      setSupervisorStats(stats);
      setSupervisorProjects(normalizeList(projects as any[] | { data?: any[] }));
      setOverdueTasks(normalizeList(overdue as Task[] | { data?: Task[] }));
      setTeamPerformance(team);
    } catch {}
  }

  async function loadSupervisorProject(projectId: string, authToken = token) {
    if (!authToken || !projectId) return;
    setLoadingSupervisorProject(true);
    try {
      const [report, membersPerf] = await Promise.all([
        supervisorApi.projectReport(authToken, projectId).catch(() => null),
        supervisorApi.projectAssigneePerformance(authToken, projectId).catch(() => null),
      ]);
      setSupervisorProjectReport(report);
      const raw = membersPerf as any;
      setSupervisorProjectMembersPerf(
        raw?.members ? raw.members : Array.isArray(raw) || raw?.data ? normalizeList(raw as MemberPerformance[] | { data?: MemberPerformance[] }) : raw ? [raw] : [],
      );
    } catch {} finally {
      setLoadingSupervisorProject(false);
    }
  }

  useEffect(() => {
    queueMicrotask(() => void loadData());
  }, [token]);

  async function handleAuth(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    setAuthLoading(true); setError(""); setMessage("");
    const form = new FormData(e.currentTarget);
    const payload = authMode === "login"
      ? { mobile: String(form.get("mobile") ?? ""), password: String(form.get("password") ?? "") }
      : { firstName: String(form.get("firstName") ?? ""), lastName: String(form.get("lastName") ?? ""), email: String(form.get("email") ?? ""), mobile: String(form.get("mobile") ?? ""), password: String(form.get("password") ?? ""), workField: String(form.get("workField") ?? "operations") };
    try {
      const data = authMode === "login"
        ? await authApi.login(payload as { mobile: string; password: string })
        : await authApi.register(payload as Record<string, string>);
      setToken(data.accessToken); setCurrentUser(data.user);
      localStorage.setItem("taskino-token", data.accessToken);
      localStorage.setItem("taskino-user", JSON.stringify(data.user));
      setMessage(authMode === "login" ? "با موفقیت وارد شدی." : "ثبت‌نام انجام شد.");
      await loadData(data.accessToken);
    } catch (err) { setError(err instanceof Error ? err.message : "ورود ناموفق بود"); }
    finally { setAuthLoading(false); }
  }

  async function createTask(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    if (!myId || !taskTitle.trim()) return;
    try {
      const created = await taskApi.create(token, {
        title: taskTitle.trim(),
        assignedTo: taskAssignee ? [taskAssignee] : [],
        projectId: taskProjectId || undefined,
        status: "todo",
      });
      if (taskProjectId) {
        await projectApi.addTask(token, taskProjectId, getId(created)).catch(() => null);
      }
      setTaskTitle(""); setTaskAssignee(""); setTaskProjectId("");
      setMessage("گزارش جدید ساخته شد."); await loadData();
    } catch (err) { setError(err instanceof Error ? err.message : "ساخت گزارش ناموفق بود"); }
  }

  async function claimTask(taskId: string) {
    if (!myId) return;
    try {
      const updated = await taskApi.update(token, taskId, { assignedTo: [myId] });
      setTasks((prev) => prev.map((t) => getId(t) === taskId ? updated : t));
      if (selectedTask && getId(selectedTask) === taskId) setSelectedTask(updated);
      setMessage("گزارش برای شما اساین شد.");
    } catch (err) {
      setError(err instanceof Error ? err.message : "اساین گزارش ناموفق بود");
    }
  }

  async function createProject(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    if (!myId || !projectTitle.trim() || !projectSupervisorId) return;
    try {
      await projectApi.create(token, {
        title: projectTitle.trim(),
        description: projectDescription.trim() || undefined,
        owner: myId,
        supervisorId: projectSupervisorId,
        assigneeId: projectAssigneeId || undefined,
        workField: projectWorkField,
        status: "pending",
      });
      setProjectTitle(""); setProjectDescription(""); setProjectSupervisorId(""); setProjectAssigneeId(""); setShowProjectForm(false);
      setMessage("پروژه جدید ساخته شد."); await loadData();
    } catch (err) { setError(err instanceof Error ? err.message : "ساخت پروژه ناموفق بود"); }
  }

  async function moveTask(taskId: string, newStatus: string) {
    try {
      const task = tasks.find((t) => getId(t) === taskId);
      const projectId = getId(task?.projectId);
      if (isSupervisor && projectId) {
        await supervisorApi.updateTaskStatus(token, projectId, taskId, newStatus);
      } else {
        await taskApi.updateStatus(token, taskId, newStatus);
      }
      setTasks((prev) => prev.map((t) => getId(t) === taskId ? { ...t, status: newStatus } : t));
      if (selectedTask && getId(selectedTask) === taskId) setSelectedTask((prev) => prev ? { ...prev, status: newStatus } : prev);
    } catch (err) { setError(err instanceof Error ? err.message : "تغییر وضعیت ناموفق بود"); }
  }

  async function updateTask(taskId: string, body: Record<string, unknown>) {
    try {
      const updated = await taskApi.update(token, taskId, body);
      setTasks((prev) => prev.map((t) => getId(t) === taskId ? { ...t, ...updated } : t));
      if (selectedTask && getId(selectedTask) === taskId) setSelectedTask((prev) => prev ? { ...prev, ...updated } : prev);
      setMessage("گزارش بروزرسانی شد.");
    } catch (err) { setError(err instanceof Error ? err.message : "بروزرسانی گزارش ناموفق بود"); }
  }

  async function deleteTask(taskId: string) {
    try {
      await taskApi.delete(token, taskId);
      setTasks((prev) => prev.filter((t) => getId(t) !== taskId));
      setSelectedTask(null);
      setMessage("گزارش حذف شد.");
    } catch (err) { setError(err instanceof Error ? err.message : "حذف گزارش ناموفق بود"); }
  }

  async function deleteProject(projectId: string) {
    try {
      await projectApi.delete(token, projectId);
      setMessage("پروژه حذف شد.");
      await loadData();
    } catch (err) { setError(err instanceof Error ? err.message : "حذف پروژه ناموفق بود"); }
  }

  async function markNotificationRead(id: string) {
    try {
      await notificationApi.markRead(token, id);
      setNotifications((prev) => prev.filter((n) => getId(n) !== id));
      setUnreadCount((c) => Math.max(0, c - 1));
    } catch (err) { setError(err instanceof Error ? err.message : "خطا در خواندن اعلان"); }
  }

  async function markAllNotificationsRead() {
    try {
      await notificationApi.markAllRead(token);
      setNotifications([]);
      setUnreadCount(0);
      setMessage("همه اعلان‌ها خوانده شد.");
    } catch (err) { setError(err instanceof Error ? err.message : "خطا در خواندن اعلان‌ها"); }
  }

  async function createLeaveRequest(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    if (!myId || !leaveStart || !leaveEnd) return;
    try {
      await leaveApi.create(token, { user: myId, startDate: leaveStart, endDate: leaveEnd, reason: leaveReason, approvedBy: myId });
      setLeaveStart(""); setLeaveEnd(""); setLeaveReason("");
      setMessage("درخواست مرخصی ثبت شد.");
      await loadData();
    } catch (err) { setError(err instanceof Error ? err.message : "ثبت مرخصی ناموفق بود"); }
  }

  async function handleLeaveAction(id: string, action: "approve" | "reject") {
    if (!myId) return;
    try {
      if (action === "approve") await leaveApi.approve(token, id, myId);
      else await leaveApi.reject(token, id, myId);
      setMessage(action === "approve" ? "مرخصی تأیید شد." : "مرخصی رد شد.");
      await loadData();
    } catch (err) { setError(err instanceof Error ? err.message : "عملیات مرخصی ناموفق بود"); }
  }

  async function updateUserRole(userId: string, role: string) {
    try {
      await managerApi.updateUserRole(token, userId, role);
      setMessage("نقش کاربر بروزرسانی شد.");
      await loadData();
    } catch (err) { setError(err instanceof Error ? err.message : "تغییر نقش ناموفق بود"); }
  }

  async function saveProfile(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    if (!myId) return;
    try {
      const updated = await userApi.update(token, myId, {
        firstName: settingsFirstName,
        lastName: settingsLastName,
        email: settingsEmail,
      });
      setCurrentUser(updated);
      localStorage.setItem("taskino-user", JSON.stringify(updated));
      setMessage("پروفایل ذخیره شد.");
    } catch (err) { setError(err instanceof Error ? err.message : "ذخیره پروفایل ناموفق بود"); }
  }

  useEffect(() => {
    if (currentUser) {
      queueMicrotask(() => {
        setSettingsFirstName(currentUser.firstName ?? "");
        setSettingsLastName(currentUser.lastName ?? "");
        setSettingsEmail(currentUser.email ?? "");
      });
    }
  }, [currentUser]);

  async function loadExcelData(authToken = token) {
    if (!authToken || !myId) return;
    try {
      const [files, stats] = await Promise.all([
        excelApi.list(authToken, { limit: 50, createdBy: myId }),
        excelApi.statistics(authToken, myId),
      ]);
      setExcelFiles(normalizeList(files));
      setExcelStats(stats);
    } catch (err) {
      setError(err instanceof Error ? err.message : "دریافت فایل‌های اکسل ناموفق بود");
    }
  }

  useEffect(() => {
    if (token && activeView === "excel") queueMicrotask(() => void loadExcelData());
  }, [token, activeView]);

  async function loadProjectMembers(project: Project) {
    const people = [
      project.owner ? { id: "owner", user: project.owner, role: "manager" } : null,
      project.supervisorId ? { id: "supervisor", user: project.supervisorId, role: "supervisor" } : null,
      project.assigneeId ? { id: "assignee", user: project.assigneeId, role: "specialist" } : null,
    ].filter(Boolean) as Array<{ user?: string | User; role?: string; id?: string }>;
    setMembersProject(project);
    setProjectMembers(people);
  }

  async function handleExcelUpload(e: ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file || !myId) return;
    setExcelUploading(true);
    try {
      await excelApi.upload(token, file, myId, "import");
      setMessage("فایل اکسل آپلود شد.");
      await loadExcelData();
    } catch (err) {
      setError(err instanceof Error ? err.message : "آپلود فایل ناموفق بود");
    } finally {
      setExcelUploading(false);
      e.target.value = "";
    }
  }

  async function downloadExcelFile(record: ExcelFile) {
    try {
      await excelApi.download(token, getId(record), record.originalName || record.fileName || "file.xlsx");
    } catch (err) {
      setError(err instanceof Error ? err.message : "دانلود فایل ناموفق بود");
    }
  }

  async function processExcelFile(id: string) {
    try {
      await excelApi.process(token, id);
      setMessage("فایل پردازش شد.");
      await loadExcelData();
    } catch (err) {
      setError(err instanceof Error ? err.message : "پردازش فایل ناموفق بود");
    }
  }

  async function deleteExcelFile(id: string) {
    try {
      await excelApi.delete(token, id);
      setMessage("فایل حذف شد.");
      await loadExcelData();
    } catch (err) {
      setError(err instanceof Error ? err.message : "حذف فایل ناموفق بود");
    }
  }

  function openFixedTaskForm(ft?: FixedTask) {
    if (ft) {
      setEditingFixedTask(ft);
      setFtTitle(ft.title);
      setFtAssignee(getId(ft.assignedTo));
      setFtRecurrence(ft.recurrence);
      setFtDescription(ft.description ?? "");
      setFtProjectId(getId(ft.projectId));
    } else {
      setEditingFixedTask(null);
      setFtTitle(""); setFtAssignee(""); setFtRecurrence("daily");
      setFtDescription(""); setFtProjectId("");
    }
    setShowFixedTaskForm(true);
  }

  function closeFixedTaskForm() {
    setShowFixedTaskForm(false);
    setEditingFixedTask(null);
  }

  async function saveFixedTask(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    if (!ftTitle.trim() || !ftAssignee) return;
    const body = {
      title: ftTitle.trim(),
      assignedTo: ftAssignee,
      recurrence: ftRecurrence,
      description: ftDescription.trim() || undefined,
      projectId: ftProjectId || undefined,
    };
    try {
      if (editingFixedTask) {
        const updated = await fixedTaskApi.update(token, getId(editingFixedTask), body);
        setFixedTasks((prev) => prev.map((ft) => getId(ft) === getId(editingFixedTask) ? updated : ft));
        setMessage("الگوی ثابت بروزرسانی شد.");
      } else {
        const created = await fixedTaskApi.create(token, body);
        setFixedTasks((prev) => [created, ...prev]);
        setMessage("الگوی ثابت ساخته شد.");
      }
      closeFixedTaskForm();
    } catch (err) { setError(err instanceof Error ? err.message : "ذخیره الگو ناموفق بود"); }
  }

  async function toggleFixedTaskActive(ft: FixedTask) {
    try {
      const updated = await fixedTaskApi.update(token, getId(ft), { isActive: !ft.isActive });
      setFixedTasks((prev) => prev.map((item) => getId(item) === getId(ft) ? updated : item));
      setMessage(updated.isActive ? "الگو فعال شد." : "الگو غیرفعال شد.");
    } catch (err) { setError(err instanceof Error ? err.message : "تغییر وضعیت الگو ناموفق بود"); }
  }

  async function deleteFixedTask(id: string) {
    try {
      await fixedTaskApi.delete(token, id);
      setFixedTasks((prev) => prev.filter((ft) => getId(ft) !== id));
      setMessage("الگوی ثابت حذف شد.");
    } catch (err) { setError(err instanceof Error ? err.message : "حذف الگو ناموفق بود"); }
  }

  async function exportTasksToExcel() {
    if (!tasks.length) return;
    try {
      const data = tasks.map((t) => ({
        title: t.title,
        status: statusLabel(t.status),
        assignees: (t.assignedTo ?? []).map(userName).join(", "),
        dueDate: t.dueDate ?? "",
      }));
      await excelApi.generateExport(token, {
        data,
        columns: ["title", "status", "assignees", "dueDate"],
        sheetName: "Reports",
      }, "reports-export.xlsx");
      setMessage("خروجی اکسل گزارش‌ها دانلود شد.");
      await loadExcelData();
    } catch (err) {
      setError(err instanceof Error ? err.message : "خروجی اکسل ناموفق بود");
    }
  }

  function onDragEnd(result: DropResult) {
    const { destination, source, draggableId } = result;
    if (!destination) return;
    if (destination.droppableId === source.droppableId && destination.index === source.index) return;
    if (destination.droppableId !== source.droppableId) {
      void moveTask(draggableId, destination.droppableId);
    }
  }

  function logout() {
    setToken(""); setCurrentUser(null); setUsers([]); setTasks([]); setProjects([]);
    setNotifications([]); setUnreadCount(0); setLeaveRequests([]); setManagerStats(null);
    setProjectProgress({}); setExcelFiles([]); setExcelStats(null); setMembersProject(null);
    setShowNotifications(false); setActiveView("dashboard");
    localStorage.removeItem("taskino-token"); localStorage.removeItem("taskino-user");
  }

  // ─── Login ─────────────────────────────────────────────────────────────────
  if (!token) {
    return (
      <main className="relative min-h-screen overflow-hidden bg-white dark:bg-slate-950">
        <Toast message={error || message} type={error ? "error" : "success"} onClose={() => error ? setError("") : setMessage("")} />
        <div className="pointer-events-none absolute inset-0 overflow-hidden">
          <div className="absolute -right-32 -top-32 h-96 w-96 rounded-full bg-[#1f7a8c]/8 blur-3xl dark:bg-[#1f7a8c]/15" />
          <div className="absolute -left-20 bottom-0 h-80 w-80 rounded-full bg-[#1f7a8c]/5 blur-3xl" />
        </div>

        <div className="relative flex min-h-screen">
          <div className="hidden flex-col justify-between bg-gradient-to-br from-[#165e6d] to-[#1f7a8c] p-12 lg:flex lg:w-[45%]">
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-white/15 backdrop-blur">
                <Sparkles size={20} className="text-white" />
              </div>
              <span className="text-xl font-bold text-white tracking-tight">مدیریت واحد بهبود عملیات و برنامه ریزی</span>
            </div>
            <div>
              <h1 className="text-5xl font-extrabold leading-tight text-white">تیمت رو<br /><span className="opacity-80">منظم نگه دار</span></h1>
              <p className="mt-6 text-lg leading-8 text-[#a8dde5]">پروژه‌ها، گزارش‌ها و تیم رو از یک داشبورد پیگیری کن.</p>
              <div className="mt-10 space-y-3">
                {["📋  برد کانبان با drag & drop", "🎯  اولویت‌بندی گزارش‌ها", "👥  مدیریت تیم و تخصیص وظایف"].map((t) => (
                  <div key={t} className="flex items-center gap-3 rounded-xl bg-white/10 px-4 py-3 text-sm font-medium text-white/90 backdrop-blur-sm">{t}</div>
                ))}
              </div>
            </div>
            <p className="text-xs text-[#7ec5cf]">مدیریت واحد بهبود عملیات و برنامه ریزی</p>
          </div>

          <div className="flex flex-1 items-center justify-center px-6 py-12">
            <div className="w-full max-w-md">
              <div className="mb-8 flex items-center gap-2 lg:hidden">
                <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-[#1f7a8c] text-white"><Sparkles size={18} /></div>
                <span className="text-lg font-bold">مدیریت واحد بهبود عملیات و برنامه ریزی</span>
              </div>
              <h2 className="text-3xl font-bold">{authMode === "login" ? "خوش برگشتی" : "بریم شروع کنیم"}</h2>
              <p className="mt-2 text-sm text-slate-500 dark:text-slate-400">{authMode === "login" ? "با حساب‌ات وارد شو" : "یه حساب رایگان بساز"}</p>
              <div className="mt-8 flex rounded-xl bg-slate-100 dark:bg-slate-800 p-1">
                <button className={`flex-1 rounded-lg py-2.5 text-sm font-semibold transition-all ${authMode === "login" ? "bg-white dark:bg-slate-700 text-[#1f7a8c] shadow-sm" : "text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-300"}`} onClick={() => setAuthMode("login")} type="button">ورود</button>
                <button className={`flex-1 rounded-lg py-2.5 text-sm font-semibold transition-all ${authMode === "register" ? "bg-white dark:bg-slate-700 text-[#1f7a8c] shadow-sm" : "text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-300"}`} onClick={() => setAuthMode("register")} type="button">ثبت‌نام</button>
              </div>
              <form className="mt-6 space-y-4" onSubmit={handleAuth}>
                {authMode === "register" && (
                  <>
                    <div className="grid grid-cols-2 gap-3">
                      <Field label="نام" name="firstName" required />
                      <Field label="نام خانوادگی" name="lastName" required />
                    </div>
                    <Field label="ایمیل" name="email" required type="email" placeholder="you@example.com" />
                    <label className="block">
                      <span className="mb-1.5 block text-xs font-semibold text-[--text-2]">حوزه کاری</span>
                      <select className="h-10 w-full rounded-lg border border-[--border] bg-[--surface] px-3 text-sm text-[--text] outline-none transition focus:border-[#1f7a8c] focus:ring-2 focus:ring-[#1f7a8c]/15" name="workField" defaultValue="operations">
                        {WORK_FIELDS.map(([value, label]) => <option key={value} value={value}>{label}</option>)}
                      </select>
                    </label>
                  </>
                )}
                <Field label="موبایل" name="mobile" required placeholder="09xxxxxxxxx" />
                <Field label="رمز عبور" name="password" required type="password" placeholder="حداقل ۶ کاراکتر" />
                <button className="flex h-12 w-full items-center justify-center gap-2 rounded-xl bg-[#1f7a8c] text-sm font-semibold text-white shadow-lg shadow-[#1f7a8c]/25 transition-all hover:bg-[#196b7b] active:scale-[0.98] disabled:opacity-60" disabled={authLoading} type="submit">
                  {authLoading ? <Loader2 className="animate-spin" size={18} /> : <LogIn size={18} />}
                  {authMode === "login" ? "ورود به سامانه" : "ساخت حساب رایگان"}
                </button>
              </form>
              <p className="mt-6 text-center text-xs text-slate-400">
                {authMode === "login" ? "حساب نداری؟ " : "قبلاً ثبت‌نام کردی؟ "}
                <button className="font-semibold text-[#1f7a8c] hover:underline" onClick={() => setAuthMode(authMode === "login" ? "register" : "login")} type="button">
                  {authMode === "login" ? "ثبت‌نام کن" : "وارد شو"}
                </button>
              </p>
            </div>
          </div>
        </div>
      </main>
    );
  }

  // ─── Dashboard ─────────────────────────────────────────────────────────────
  const sideW = sidebarCollapsed ? 64 : 248;

  return (
    <div className="min-h-screen bg-[--bg] text-[--text]">
      <Toast message={error || message} type={error ? "error" : "success"} onClose={() => error ? setError("") : setMessage("")} />

      {/* Topbar */}
      <header className="sticky top-0 z-40 border-b border-[--border] bg-[--surface]/95 backdrop-blur-md">
        <div className="flex items-center gap-3 px-3 py-2.5">
          <div className="flex items-center gap-2">
            <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-[#1f7a8c] text-white shadow-sm">
              <Sparkles size={15} />
            </div>
            {!sidebarCollapsed && <span className="font-bold tracking-tight">مدیریت واحد بهبود عملیات و برنامه ریزی</span>}
          </div>

          <div className="flex flex-1 items-center justify-end gap-2">
            <div className="relative hidden sm:block">
              <Search className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-[--text-3]" size={15} />
              <input
                className="h-9 w-52 rounded-lg border border-[--border] bg-[--surface-2] pr-9 pl-3 text-sm outline-none transition-all placeholder:text-[--text-3] focus:w-72 focus:border-[#1f7a8c] focus:bg-[--surface] focus:ring-2 focus:ring-[#1f7a8c]/15"
                onChange={(e) => setTaskQuery(e.target.value)} placeholder="جستجوی گزارش…" value={taskQuery}
              />
              {taskQuery && <button className="absolute left-2.5 top-1/2 -translate-y-1/2 text-[--text-3] hover:text-[--text-2]" onClick={() => setTaskQuery("")} type="button"><X size={13} /></button>}
            </div>

            <div className="relative">
              <button
                className="relative flex h-9 w-9 items-center justify-center rounded-lg border border-[--border] bg-[--surface] text-[--text-2] transition hover:bg-[--surface-2]"
                onClick={() => setShowNotifications(!showNotifications)} type="button"
              >
                <Bell size={16} />
                {unreadCount > 0 && (
                  <span className="absolute -left-1 -top-1 flex h-4 min-w-4 items-center justify-center rounded-full bg-red-500 px-1 text-[10px] font-bold text-white">
                    {unreadCount > 9 ? "9+" : unreadCount}
                  </span>
                )}
              </button>
              {showNotifications && (
                <div className="absolute left-0 top-11 z-50 w-80 overflow-hidden rounded-xl border border-[--border] bg-[--surface] shadow-xl">
                  <div className="flex items-center justify-between border-b border-[--border] px-4 py-3">
                    <span className="text-sm font-bold">اعلان‌ها</span>
                    {unreadCount > 0 && (
                      <button className="text-xs font-semibold text-[#1f7a8c] hover:underline" onClick={() => void markAllNotificationsRead()} type="button">
                        همه خوانده شد
                      </button>
                    )}
                  </div>
                  <div className="max-h-72 overflow-y-auto">
                    {notifications.length === 0 ? (
                      <p className="p-4 text-center text-xs text-[--text-3]">اعلان جدیدی نیست</p>
                    ) : (
                      notifications.map((n) => (
                        <button
                          key={getId(n)}
                          className="flex w-full flex-col gap-0.5 border-b border-[--border] px-4 py-3 text-right transition hover:bg-[--surface-2]"
                          onClick={() => void markNotificationRead(getId(n))} type="button"
                        >
                          <span className="text-sm font-semibold">{n.title}</span>
                          <span className="text-xs text-[--text-3] line-clamp-2">{n.message}</span>
                        </button>
                      ))
                    )}
                  </div>
                </div>
              )}
            </div>

            {/* Dark mode toggle */}
            <button
              className="flex h-9 w-9 items-center justify-center rounded-lg border border-[--border] bg-[--surface] text-[--text-2] transition hover:bg-[--surface-2]"
              onClick={() => setDarkMode(!darkMode)} type="button"
            >
              {darkMode ? <Sun size={16} /> : <Moon size={16} />}
            </button>

            <button
              className="flex h-9 items-center gap-1.5 rounded-lg border border-[--border] bg-[--surface] px-3 text-sm font-medium text-[--text-2] transition hover:bg-[--surface-2] disabled:opacity-50"
              onClick={() => loadData()} disabled={loadingData} type="button"
            >
              {loadingData ? <Loader2 className="animate-spin" size={15} /> : <RefreshCw size={15} />}
              <span className="hidden sm:inline">بروزرسانی</span>
            </button>

            <button
              className="flex h-9 items-center gap-2 rounded-lg border border-[--border] bg-[--surface] px-2.5 text-sm font-medium text-[--text] transition hover:bg-[--surface-2]"
              onClick={logout} type="button"
            >
              <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-[#1f7a8c] text-[10px] font-bold text-white">
                {initials(currentUser ?? undefined)}
              </span>
              <span className="hidden max-w-[90px] truncate sm:block text-xs">{userName(currentUser ?? undefined)}</span>
              <LogOut size={14} className="text-[--text-3]" />
            </button>
          </div>
        </div>
      </header>

      <div className="flex">
        {/* ─ Sidebar ──────────────────────────────────────────────────────── */}
        <aside
          className="sticky top-[53px] h-[calc(100vh-53px)] shrink-0 overflow-y-auto overflow-x-hidden border-l border-[--border] bg-[--surface] transition-all duration-200"
          style={{ width: sideW }}
        >
          {/* Collapse toggle */}
          <div className="flex items-center justify-between border-b border-[--border] px-3 py-2">
            {!sidebarCollapsed && <span className="text-xs font-semibold text-[--text-3]">منو</span>}
            <button
              className="mr-auto flex h-7 w-7 items-center justify-center rounded-lg text-[--text-3] transition hover:bg-[--surface-2] hover:text-[--text]"
              onClick={() => setSidebarCollapsed(!sidebarCollapsed)} type="button"
            >
              {sidebarCollapsed ? <ChevronLeft size={15} /> : <ChevronRight size={15} />}
            </button>
          </div>

          <nav className="space-y-0.5 p-2">
            {isSupervisor ? (
              <>
                <SideItem active={activeView === "dashboard"} icon={LayoutDashboard} label="داشبورد" collapsed={sidebarCollapsed} onClick={() => setActiveView("dashboard")} />
                <SideItem active={activeView === "supervisor-projects"} icon={FolderKanban} label="پروژه‌های من" meta={supervisorProjects.length} collapsed={sidebarCollapsed} onClick={() => setActiveView("supervisor-projects")} />
                <SideItem active={activeView === "supervisor-team"} icon={UsersRound} label="عملکرد تیم" collapsed={sidebarCollapsed} onClick={() => setActiveView("supervisor-team")} />
                <SideItem active={activeView === "tasks"} icon={ClipboardList} label="گزارش‌ها" meta={tasks.length || overdueTasks.length || undefined} collapsed={sidebarCollapsed} onClick={() => setActiveView("tasks")} />
                <div className="my-1.5 border-t border-[--border]" />
                <SideItem active={activeView === "settings"} icon={Settings} label="تنظیمات" collapsed={sidebarCollapsed} onClick={() => setActiveView("settings")} />
              </>
            ) : isManager ? (
              <>
                <SideItem active={activeView === "dashboard"} icon={LayoutDashboard} label="داشبورد" collapsed={sidebarCollapsed} onClick={() => setActiveView("dashboard")} />
                <SideItem active={activeView === "projects"} icon={FolderKanban} label="پروژه‌ها" meta={statsProjects} collapsed={sidebarCollapsed} onClick={() => setActiveView("projects")} />
                <SideItem active={activeView === "tasks"} icon={ClipboardList} label="گزارش‌ها" meta={tasks.length} collapsed={sidebarCollapsed} onClick={() => setActiveView("tasks")} />
                <SideItem active={activeView === "analytics"} icon={BarChart2} label="آنالیتیکس" collapsed={sidebarCollapsed} onClick={() => setActiveView("analytics")} />
                <SideItem active={activeView === "fixed-reports"} icon={ClipboardList} label="گزارش‌های ثابت" meta={incompleteFixedTasks.length || fixedTasks.length || undefined} collapsed={sidebarCollapsed} onClick={() => setActiveView("fixed-reports")} />
                <SideItem active={activeView === "team"} icon={UsersRound} label="تیم" meta={statsUsers} collapsed={sidebarCollapsed} onClick={() => setActiveView("team")} />
                <SideItem active={activeView === "excel"} icon={FileSpreadsheet} label="اکسل" collapsed={sidebarCollapsed} onClick={() => setActiveView("excel")} />
                <div className="my-1.5 border-t border-[--border]" />
                <SideItem active={activeView === "settings"} icon={Settings} label="تنظیمات" collapsed={sidebarCollapsed} onClick={() => setActiveView("settings")} />
              </>
            ) : (
              <>
                <SideItem active={activeView === "dashboard"} icon={LayoutDashboard} label="داشبورد" collapsed={sidebarCollapsed} onClick={() => setActiveView("dashboard")} />
                <SideItem active={activeView === "projects"} icon={FolderKanban} label="پروژه‌ها" meta={statsProjects} collapsed={sidebarCollapsed} onClick={() => setActiveView("projects")} />
                <SideItem active={activeView === "tasks"} icon={ClipboardList} label="گزارش‌ها" meta={tasks.length} collapsed={sidebarCollapsed} onClick={() => setActiveView("tasks")} />
                <SideItem active={activeView === "team"} icon={UsersRound} label="تیم" meta={statsUsers} collapsed={sidebarCollapsed} onClick={() => setActiveView("team")} />
                <SideItem active={activeView === "excel"} icon={FileSpreadsheet} label="اکسل" meta={excelFiles.length} collapsed={sidebarCollapsed} onClick={() => setActiveView("excel")} />
                <div className="my-1.5 border-t border-[--border]" />
                <SideItem active={activeView === "settings"} icon={Settings} label="تنظیمات" collapsed={sidebarCollapsed} onClick={() => setActiveView("settings")} />
              </>
            )}
          </nav>

          {!sidebarCollapsed && (
            <>
              {/* User card */}
              <div className="mx-2 mt-2 rounded-xl border border-[--border] bg-[--surface-2] p-3">
                <div className="flex items-center gap-2.5">
                  <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-[#1f7a8c] to-[#165e6d] text-xs font-bold text-white">
                    {initials(currentUser ?? undefined)}
                  </div>
                  <div className="min-w-0">
                    <p className="truncate text-sm font-semibold">{userName(currentUser ?? undefined)}</p>
                    <p className="truncate text-xs text-[--text-3]">{currentUser?.mobile ?? currentUser?.email}</p>
                  </div>
                </div>
                <div className="mt-2 flex items-center gap-1.5 rounded-lg bg-[--surface] px-2.5 py-1.5 text-xs font-medium text-[--text-2]">
                  <ShieldCheck size={12} />
                  {currentUser?.roles || "specialist"}
                </div>
              </div>

              {/* Quick task — only for manager */}
              {isManager && (
                <div className="mx-2 mt-2 rounded-xl border border-[--border] bg-[--surface-2] p-3">
                  <div className="mb-2.5 flex items-center gap-1.5">
                    <Zap size={13} className="text-[#1f7a8c]" />
                    <span className="text-xs font-semibold text-[--text]">گزارش سریع</span>
                  </div>
                  <form className="space-y-2" onSubmit={createTask}>
                    <Field label="" name="taskTitle" id="quick-task-title" value={taskTitle} onChange={setTaskTitle} required placeholder="عنوان گزارش…" />
                    <Select label="" value={taskProjectId} onChange={setTaskProjectId} options={projects.map((p) => [getId(p), p.title])} placeholder="پروژه (اختیاری)" />
                    <Select label="" value={taskAssignee} onChange={setTaskAssignee} options={users.map((u) => [getId(u), userName(u)])} placeholder="بدون مسئول" />
                    <button className="flex h-9 w-full items-center justify-center gap-1.5 rounded-lg bg-[#1f7a8c] text-xs font-semibold text-white transition hover:bg-[#196b7b] active:scale-[0.98] disabled:opacity-60" disabled={!taskTitle.trim()} type="submit">
                      <Plus size={14} /> افزودن گزارش
                    </button>
                  </form>
                </div>
              )}

              {/* Progress mini */}
              <div className="mx-2 my-2 rounded-xl border border-[--border] bg-[--surface-2] p-3">
                <div className="mb-2 flex items-center justify-between">
                  <span className="text-xs font-semibold text-[--text-2]">پیشرفت</span>
                  <span className="text-sm font-bold text-[--text]">{progress}%</span>
                </div>
                <div className="h-1.5 overflow-hidden rounded-full bg-[--border]">
                  <div className="h-full rounded-full bg-gradient-to-l from-[#1f7a8c] to-[#2a9db2] transition-all duration-700" style={{ width: `${progress}%` }} />
                </div>
                <div className="mt-2.5 grid grid-cols-3 gap-1 text-center">
                  {[{ l: "باز", v: todoCount, c: "text-[--text]" }, { l: "جاری", v: inProgressTasks, c: "text-[#1f7a8c]" }, { l: "تمام", v: doneTasks, c: "text-emerald-500" }].map((s) => (
                    <div key={s.l} className="rounded-lg bg-[--surface] py-1.5">
                      <p className={`text-base font-bold ${s.c}`}>{s.v}</p>
                      <p className="text-[10px] text-[--text-3]">{s.l}</p>
                    </div>
                  ))}
                </div>
              </div>
            </>
          )}
        </aside>

        {/* ─ Main ─────────────────────────────────────────────────────────── */}
        <main className="min-w-0 flex-1 space-y-4 overflow-auto p-4" onClick={() => showNotifications && setShowNotifications(false)}>

          {/* ─── Supervisor Dashboard ─────────────────────────────────────────── */}
          {isSupervisor && activeView === "dashboard" && (
            <section className="space-y-4">
              {/* Welcome banner */}
              <div className="relative overflow-hidden rounded-2xl bg-gradient-to-l from-violet-700 via-violet-600 to-violet-500 px-6 py-5 text-white shadow-lg shadow-violet-500/15">
                <div className="pointer-events-none absolute -left-6 -top-6 h-36 w-36 rounded-full bg-white/5" />
                <div className="relative flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <p className="text-sm font-medium text-violet-200">سلام، {userName(currentUser ?? undefined).split(" ")[0]}</p>
                    <h1 className="mt-0.5 text-xl font-bold">داشبورد سرپرست</h1>
                    <p className="mt-1 text-sm text-violet-200">{overdueTasks.length > 0 ? `${overdueTasks.length} گزارش معوق نیاز به بررسی دارد` : "همه گزارش‌ها در موعد هستند"}</p>
                  </div>
                  <div className="flex shrink-0 items-center gap-5">
                    {[{ n: supervisorStats?.supervisedProjects ?? 0, l: "پروژه" }, { n: supervisorInProgressReports, l: "در حال انجام" }, { n: supervisorOwnDoneReports, l: "تکمیل شده" }].map((s, i) => (
                      <div key={i} className="text-center">
                        <p className="text-2xl font-extrabold">{s.n}</p>
                        <p className="text-[11px] text-violet-200">{s.l}</p>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Stats cards */}
              <div className="grid grid-cols-2 gap-3 xl:grid-cols-4">
                {[
                  { label: "پروژه‌های تحت نظر", value: supervisorStats?.supervisedProjects ?? 0, sub: "پروژه", icon: FolderKanban, a: "bg-violet-50 text-violet-600 ring-violet-100 dark:bg-violet-950/40 dark:text-violet-400 dark:ring-violet-900" },
                  { label: "گزارش‌های جاری", value: supervisorInProgressReports, sub: "در حال انجام", icon: Activity, a: "bg-[#e8f4f7] text-[#1f7a8c] ring-[#1f7a8c]/10 dark:bg-[#0f3040] dark:text-[#4fc3d5] dark:ring-[#1f7a8c]/20" },
                  { label: "گزارش‌های موفق پروژه", value: supervisorProjectDoneReports, sub: "تکمیل شده", icon: CheckCircle2, a: "bg-emerald-50 text-emerald-600 ring-emerald-100 dark:bg-emerald-950/40 dark:text-emerald-400 dark:ring-emerald-900" },
                  { label: "گزارش‌های معوق", value: overdueTasks.length, sub: "نیاز به بررسی", icon: AlertTriangle, a: "bg-amber-50 text-amber-600 ring-amber-100 dark:bg-amber-950/40 dark:text-amber-400 dark:ring-amber-900" },
                ].map((s) => (
                  <div key={s.label} className="group rounded-xl border border-[--border] bg-[--surface] p-4 transition-all hover:-translate-y-0.5 hover:shadow-md hover:shadow-black/5">
                    <div className="flex items-start justify-between gap-2">
                      <p className="text-xs font-medium text-[--text-2]">{s.label}</p>
                      <span className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg ring-4 ${s.a}`}><s.icon size={15} /></span>
                    </div>
                    <p className="mt-2.5 text-3xl font-bold">{s.value}</p>
                    <p className="mt-0.5 text-xs text-[--text-3]">{s.sub}</p>
                  </div>
                ))}
              </div>

              {/* Overdue tasks */}
              {overdueTasks.length > 0 && (
                <div className="overflow-hidden rounded-2xl border border-amber-200 dark:border-amber-900 bg-[--surface] shadow-sm">
                  <div className="flex items-center gap-3 border-b border-amber-100 dark:border-amber-900/50 bg-gradient-to-l from-amber-50 to-white dark:from-amber-950/30 dark:to-transparent px-5 py-4">
                    <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-amber-100 text-amber-600 dark:bg-amber-950/50 dark:text-amber-400">
                      <AlertTriangle size={17} />
                    </div>
                    <div>
                      <h2 className="font-bold text-[--text]">گزارش‌های معوق</h2>
                      <p className="text-[11px] text-[--text-3]">{overdueTasks.length} گزارش از موعد گذشته</p>
                    </div>
                  </div>
                  <div className="divide-y divide-[--border]">
                    {overdueTasks.slice(0, 5).map((t) => (
                      <div key={getId(t)} className="flex items-center justify-between px-5 py-3">
                        <div>
                          <p className="text-sm font-semibold">{t.title}</p>
                          <p className="text-xs text-[--text-3]">مهلت: {formatDate(t.dueDate)}</p>
                        </div>
                        <span className="rounded-full bg-amber-100 dark:bg-amber-950/50 px-2.5 py-0.5 text-xs font-semibold text-amber-700 dark:text-amber-400">{statusLabel(t.status)}</span>
                      </div>
                    ))}
                  </div>
                  {overdueTasks.length > 5 && (
                    <button className="flex w-full items-center justify-center py-3 text-xs font-semibold text-[#1f7a8c] hover:bg-[--surface-2]" onClick={() => setActiveView("tasks")} type="button">
                      مشاهده همه ({overdueTasks.length})
                    </button>
                  )}
                </div>
              )}

              {/* Team performance summary */}
              {teamAssignees.length > 0 && (
                <div className="overflow-hidden rounded-2xl border border-[--border] bg-[--surface]">
                  <div className="flex items-center justify-between border-b border-[--border] px-5 py-4">
                    <div className="flex items-center gap-3">
                      <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-violet-500 to-violet-600 text-white">
                        <UsersRound size={17} />
                      </div>
                      <div>
                        <h2 className="font-bold">عملکرد تیم</h2>
                        <p className="text-[11px] text-[--text-3]">{teamAssigneeCount} عضو · {teamPerformance.completionRate ?? 0}% نرخ تکمیل</p>
                      </div>
                    </div>
                    <button className="text-xs font-semibold text-[#1f7a8c] hover:underline" onClick={() => setActiveView("supervisor-team")} type="button">مشاهده کامل</button>
                  </div>
                  <div className="divide-y divide-[--border]">
                    {teamAssignees.slice(0, 5).map((m: any) => {
                      const total = m.totalTasks ?? 0;
                      const done = m.doneTasks ?? m.completedTasks ?? 0;
                      const rate = total ? Math.round((done / total) * 100) : 0;
                      return (
                        <div key={m.userId ?? m._id} className="flex items-center gap-4 px-5 py-3">
                          <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-violet-400 to-violet-600 text-xs font-bold text-white">
                            {(m.firstName?.[0] ?? m.fullName?.[0] ?? "؟").toUpperCase()}
                          </div>
                          <div className="min-w-0 flex-1">
                            <p className="text-sm font-semibold truncate">{(m.fullName ?? `${m.firstName ?? ""} ${m.lastName ?? ""}`.trim()) || "نامشخص"}</p>
                            <div className="mt-1 flex items-center gap-2">
                              <div className="h-1.5 flex-1 overflow-hidden rounded-full bg-[--border]">
                                <div className="h-full rounded-full bg-violet-500" style={{ width: `${rate}%` }} />
                              </div>
                              <span className="shrink-0 text-[11px] text-[--text-3]">{rate}%</span>
                            </div>
                          </div>
                          <div className="text-right shrink-0">
                            <p className="text-sm font-bold">{done}/{total}</p>
                            <p className="text-[10px] text-[--text-3]">گزارش</p>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </section>
          )}

          {/* ─── Supervisor Projects ─────────────────────────────────────────────── */}
          {isSupervisor && activeView === "supervisor-projects" && (
            <section className="space-y-4">
              <div className="overflow-hidden rounded-2xl border border-violet-200 dark:border-violet-900 bg-[--surface]">
                <div className="flex items-center gap-3 border-b border-violet-100 dark:border-violet-900/50 bg-gradient-to-l from-violet-50 to-white dark:from-violet-950/30 dark:to-transparent px-5 py-4">
                  <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-violet-500 to-violet-600 text-white">
                    <FolderKanban size={17} />
                  </div>
                  <div>
                    <h2 className="font-bold">پروژه‌های تحت نظر</h2>
                    <p className="text-[11px] text-[--text-3]">{supervisorProjects.length} پروژه</p>
                  </div>
                </div>

                {supervisorProjects.length === 0 ? (
                  <div className="flex flex-col items-center py-12 text-center">
                    <FolderKanban size={32} className="text-[--text-3]" />
                    <p className="mt-3 font-semibold text-[--text]">پروژه‌ای یافت نشد</p>
                  </div>
                ) : (
                  <div className="grid gap-4 p-4 md:grid-cols-2">
                    {supervisorProjects.map((proj: any, i) => {
                      const pid = proj.projectId ?? getId(proj);
                      const total = proj.totalTasks ?? 0;
                      const done = proj.doneTasks ?? 0;
                      const rate = total ? Math.round((done / total) * 100) : 0;
                      return (
                        <article key={pid} className="overflow-hidden rounded-xl border border-[--border] bg-[--surface] transition-all hover:-translate-y-0.5 hover:shadow-lg hover:shadow-black/5">
                          <div className={`relative h-16 bg-gradient-to-l ${PROJECT_COVERS[i % PROJECT_COVERS.length]} px-4 pt-3`}>
                            <div className="absolute inset-0 opacity-10" style={{ backgroundImage: "radial-gradient(circle at 20% 50%, white 1px, transparent 1px)", backgroundSize: "20px 20px" }} />
                            <div className="relative flex items-start justify-between">
                              <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-white/20 text-white border border-white/30 backdrop-blur-sm"><FolderKanban size={17} /></div>
                              <div className="flex gap-1.5">
                                <span className="rounded-full bg-white/20 px-2.5 py-0.5 text-xs font-semibold text-white border border-white/20 backdrop-blur-sm">{rate}%</span>
                                <span className="rounded-full bg-white/20 px-2.5 py-0.5 text-xs font-semibold text-white border border-white/20 backdrop-blur-sm">{statusLabel(proj.status)}</span>
                              </div>
                            </div>
                          </div>
                          <div className="p-4">
                            <h3 className="font-bold">{proj.title}</h3>
                            <div className="mt-2 h-1.5 overflow-hidden rounded-full bg-[--border]">
                              <div className="h-full rounded-full bg-violet-500" style={{ width: `${rate}%` }} />
                            </div>
                            <div className="mt-3 grid grid-cols-3 gap-2 text-center">
                              {[{ l: "کل", v: total, c: "text-[--text]" }, { l: "جاری", v: proj.inProgressTasks ?? 0, c: "text-[#1f7a8c]" }, { l: "تمام", v: done, c: "text-emerald-500" }].map((s) => (
                                <div key={s.l} className="rounded-lg bg-[--surface-2] py-2">
                                  <p className={`text-base font-bold ${s.c}`}>{s.v}</p>
                                  <p className="text-[10px] text-[--text-3]">{s.l}</p>
                                </div>
                              ))}
                            </div>
                            <div className="mt-3 flex items-center justify-between border-t border-[--border] pt-3">
                              <span className="text-xs text-[--text-3]">👥 {proj.membersCount ?? 0} عضو</span>
                              <button
                                className="flex items-center gap-1.5 rounded-lg bg-violet-50 dark:bg-violet-950/40 px-3 py-1.5 text-xs font-semibold text-violet-600 dark:text-violet-400 transition hover:bg-violet-100"
                                onClick={async () => {
                                  setSelectedSupervisorProjectId(pid);
                                  await loadSupervisorProject(pid);
                                }}
                                type="button"
                              >
                                <UsersRound size={12} />عملکرد اعضا
                              </button>
                            </div>
                          </div>
                        </article>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Selected project detail */}
              {selectedSupervisorProjectId && (
                <div className="overflow-hidden rounded-2xl border border-[--border] bg-[--surface]">
                  <div className="flex items-center justify-between border-b border-[--border] px-5 py-4">
                    <h3 className="font-bold">عملکرد اعضای پروژه</h3>
                    {supervisorProjectReport && (
                      <div className="flex gap-3 text-xs text-[--text-2]">
                        <span>کل: <strong>{supervisorProjectReport.totalTasks ?? 0}</strong></span>
                        <span>تکمیل: <strong className="text-emerald-600">{supervisorProjectReport.doneTasks ?? 0}</strong></span>
                        <span>معوق: <strong className="text-amber-600">{(supervisorProjectReport as any).overdueTasks ?? (supervisorProjectReport as any).overdueCount ?? 0}</strong></span>
                        <span>نرخ: <strong>{supervisorProjectReport.completionRate ?? 0}%</strong></span>
                      </div>
                    )}
                  </div>
                  {loadingSupervisorProject ? (
                    <div className="flex items-center justify-center py-10"><Loader2 className="animate-spin text-violet-500" size={24} /></div>
                  ) : supervisorProjectMembersPerf.length === 0 ? (
                    <p className="py-8 text-center text-sm text-[--text-3]">داده‌ای یافت نشد</p>
                  ) : (
                    <div className="divide-y divide-[--border]">
                      {supervisorProjectMembersPerf.map((m) => {
                        const total = m.totalTasks ?? 0;
                        const done = m.completedTasks ?? m.completedCount ?? 0;
                        const rate = m.completionRate ?? (total ? Math.round((done / total) * 100) : 0);
                        return (
                          <div key={m.userId ?? getId(m)} className="flex items-center gap-4 px-5 py-4">
                            <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-violet-400 to-violet-600 text-sm font-bold text-white">
                              {(m.firstName?.[0] ?? "؟").toUpperCase()}
                            </div>
                            <div className="min-w-0 flex-1">
                              <p className="font-semibold">{(m.fullName ?? `${m.firstName ?? ""} ${m.lastName ?? ""}`.trim()) || "نامشخص"}</p>
                              <div className="mt-1 flex items-center gap-2">
                                <div className="h-1.5 flex-1 overflow-hidden rounded-full bg-[--border]">
                                  <div className="h-full rounded-full bg-violet-500 transition-all" style={{ width: `${rate}%` }} />
                                </div>
                                <span className="shrink-0 text-xs text-[--text-3]">{rate}%</span>
                              </div>
                            </div>
                            <div className="flex gap-3 shrink-0 text-xs text-center">
                              <div><p className="font-bold">{total}</p><p className="text-[--text-3]">کل</p></div>
                              <div><p className="font-bold text-emerald-600">{done}</p><p className="text-[--text-3]">تمام</p></div>
                              <div><p className="font-bold text-amber-600">{m.score ?? 0}</p><p className="text-[--text-3]">امتیاز</p></div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              )}
            </section>
          )}

          {/* ─── Supervisor Team Performance ────────────────────────────────────── */}
          {isSupervisor && activeView === "supervisor-team" && (
            <section className="space-y-4">
              <div className="overflow-hidden rounded-2xl border border-violet-200 dark:border-violet-900 bg-[--surface]">
                <div className="flex items-center gap-3 border-b border-violet-100 dark:border-violet-900/50 bg-gradient-to-l from-violet-50 to-white dark:from-violet-950/30 dark:to-transparent px-5 py-4">
                  <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-violet-500 to-violet-600 text-white">
                    <Award size={17} />
                  </div>
                  <div>
                    <h2 className="font-bold">عملکرد کل تیم</h2>
                    <p className="text-[11px] text-[--text-3]">
                      {teamAssigneeCount} عضو · {teamPerformance?.completionRate ?? 0}% نرخ تکمیل · {teamPerformance?.doneTasks ?? 0}/{teamPerformance?.totalTasks ?? 0} گزارش
                    </p>
                  </div>
                </div>

                {!teamAssignees.length ? (
                  <div className="flex flex-col items-center py-12 text-center">
                    <UsersRound size={32} className="text-[--text-3]" />
                    <p className="mt-3 font-semibold">داده‌ای یافت نشد</p>
                  </div>
                ) : (
                  <div className="divide-y divide-[--border]">
                    {teamAssignees
                      .sort((a: any, b: any) => (b.completionRate ?? 0) - (a.completionRate ?? 0))
                      .map((m: any, i: number) => {
                        const total = m.totalTasks ?? 0;
                        const done = m.doneTasks ?? m.completedTasks ?? 0;
                        const rate = m.completionRate ?? (total ? Math.round((done / total) * 100) : 0);
                        return (
                          <div key={m.userId ?? m._id ?? i} className="flex items-center gap-4 px-5 py-4">
                            <div className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-xs font-bold ${i === 0 ? "bg-amber-100 text-amber-600" : i === 1 ? "bg-slate-100 text-slate-600" : i === 2 ? "bg-orange-100 text-orange-600" : "bg-[--surface-2] text-[--text-3]"}`}>
                              {i + 1}
                            </div>
                            <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-violet-400 to-violet-600 text-xs font-bold text-white">
                              {(m.firstName?.[0] ?? m.fullName?.[0] ?? "؟").toUpperCase()}
                            </div>
                            <div className="min-w-0 flex-1">
                              <p className="font-semibold">{(m.fullName ?? `${m.firstName ?? ""} ${m.lastName ?? ""}`.trim()) || "نامشخص"}</p>
                              <div className="mt-1 flex items-center gap-2">
                                <div className="h-1.5 flex-1 overflow-hidden rounded-full bg-[--border]">
                                  <div className="h-full rounded-full bg-violet-500 transition-all duration-700" style={{ width: `${rate}%` }} />
                                </div>
                                <span className="shrink-0 text-xs font-semibold text-[--text-2]">{rate}%</span>
                              </div>
                            </div>
                            <div className="flex gap-4 shrink-0 text-center text-xs">
                              <div><p className="font-bold">{total}</p><p className="text-[--text-3]">کل</p></div>
                              <div><p className="font-bold text-emerald-600">{done}</p><p className="text-[--text-3]">تمام</p></div>
                              <div><p className="font-bold text-[#1f7a8c]">{m.inProgressTasks ?? 0}</p><p className="text-[--text-3]">جاری</p></div>
                              <div><p className="font-bold text-amber-600">{m.overdueTasks ?? 0}</p><p className="text-[--text-3]">معوق</p></div>
                              <div><p className="font-bold">{m.score ?? 0}</p><p className="text-[--text-3]">امتیاز</p></div>
                            </div>
                          </div>
                        );
                      })}
                  </div>
                )}
              </div>
            </section>
          )}

          {/* ─── Manager Analytics ──────────────────────────────────────────────── */}
          {isManager && activeView === "analytics" && (
            <section className="space-y-4">
              {/* Task status overview */}
              <div className="grid grid-cols-3 gap-3">
                {[
                  { label: "در انتظار", value: managerTaskStatus?.todoTasks ?? managerTaskStatus?.todo ?? 0, color: "text-slate-600", bg: "bg-slate-50 dark:bg-slate-900/40", ring: "ring-slate-200 dark:ring-slate-700", dot: "bg-slate-400" },
                  { label: "در حال انجام", value: managerTaskStatus?.inProgressTasks ?? managerTaskStatus?.inProgress ?? managerTaskStatus?.in_progress ?? 0, color: "text-[#1f7a8c]", bg: "bg-[#e8f4f7] dark:bg-[#0f3040]", ring: "ring-[#1f7a8c]/20", dot: "bg-[#1f7a8c]" },
                  { label: "تکمیل شده", value: managerTaskStatus?.doneTasks ?? managerTaskStatus?.done ?? 0, color: "text-emerald-600", bg: "bg-emerald-50 dark:bg-emerald-950/40", ring: "ring-emerald-200 dark:ring-emerald-900", dot: "bg-emerald-500" },
                ].map((s) => (
                  <div key={s.label} className={`rounded-xl border border-[--border] ${s.bg} p-5 text-center ring-2 ${s.ring}`}>
                    <div className={`mx-auto mb-2 h-3 w-3 rounded-full ${s.dot}`} />
                    <p className={`text-3xl font-extrabold ${s.color}`}>{s.value}</p>
                    <p className="mt-1 text-sm font-medium text-[--text-2]">{s.label}</p>
                  </div>
                ))}
              </div>

              {/* Projects progress */}
              {managerProjectsProgress.length > 0 && (
                <div className="overflow-hidden rounded-2xl border border-[--border] bg-[--surface]">
                  <div className="flex items-center gap-3 border-b border-[--border] px-5 py-4">
                    <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-[#1f7a8c] to-[#165e6d] text-white">
                      <Target size={17} />
                    </div>
                    <h2 className="font-bold">پیشرفت پروژه‌ها</h2>
                  </div>
                  <div className="divide-y divide-[--border]">
                    {managerProjectsProgress.map((proj) => {
                      const rate = proj.progressPercentage ?? 0;
                      return (
                        <div key={proj.projectId ?? proj.title} className="flex items-center gap-4 px-5 py-3">
                          <div className="min-w-0 flex-1">
                            <div className="flex items-center justify-between">
                              <p className="font-semibold">{proj.projectName ?? proj.title}</p>
                              <span className="shrink-0 text-sm font-bold text-[#1f7a8c]">{rate}%</span>
                            </div>
                            <div className="mt-1.5 h-2 overflow-hidden rounded-full bg-[--border]">
                              <div className="h-full rounded-full bg-gradient-to-l from-[#1f7a8c] to-[#2a9db2] transition-all duration-700" style={{ width: `${rate}%` }} />
                            </div>
                            <div className="mt-1 flex gap-3 text-[11px] text-[--text-3]">
                              <span>کل: {proj.totalTasks ?? 0}</span>
                              <span className="text-emerald-600">تمام: {proj.completedTasks ?? 0}</span>
                              <span className="text-[#1f7a8c]">جاری: {proj.inProgressTasks ?? 0}</span>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* User task counts */}
              {managerUserCounts.length > 0 && (
                <div className="overflow-hidden rounded-2xl border border-[--border] bg-[--surface]">
                  <div className="flex items-center gap-3 border-b border-[--border] px-5 py-4">
                    <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-amber-400 to-orange-500 text-white">
                      <UsersRound size={17} />
                    </div>
                    <h2 className="font-bold">گزارش‌ها به تفکیک کاربر</h2>
                  </div>
                  <div className="divide-y divide-[--border]">
                    {managerUserCounts.map((u) => {
                      const total = u.totalTasks ?? u.total ?? 0;
                      const done = u.doneTasks ?? u.done ?? 0;
                      const name = (u.fullName ?? u.userName ?? `${u.firstName ?? ""} ${u.lastName ?? ""}`.trim()) || "نامشخص";
                      return (
                        <div key={u.userId ?? getId(u)} className="flex items-center gap-4 px-5 py-3">
                          <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-amber-400 to-orange-500 text-xs font-bold text-white">
                            {name[0]?.toUpperCase() ?? "؟"}
                          </div>
                          <div className="min-w-0 flex-1">
                            <p className="font-semibold">{name}</p>
                            <div className="mt-1 flex gap-2 text-[11px]">
                              <span className="text-slate-500">باز: {u.todoTasks ?? u.todo ?? 0}</span>
                              <span className="text-[#1f7a8c]">جاری: {u.inProgressTasks ?? u.inProgress ?? u.in_progress ?? 0}</span>
                              <span className="text-emerald-600">تمام: {done}</span>
                            </div>
                          </div>
                          <div className="shrink-0 text-right">
                            <p className="text-lg font-bold">{total}</p>
                            <p className="text-[11px] text-[--text-3]">گزارش</p>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* Monthly performance */}
              {managerMonthlyPerf.length > 0 && (
                <div className="overflow-hidden rounded-2xl border border-[--border] bg-[--surface]">
                  <div className="flex items-center gap-3 border-b border-[--border] px-5 py-4">
                    <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-emerald-500 to-teal-600 text-white">
                      <TrendingUp size={17} />
                    </div>
                    <h2 className="font-bold">عملکرد ماهانه</h2>
                  </div>
                  <div className="divide-y divide-[--border]">
                    {managerMonthlyPerf
                      .sort((a, b) => (b.completedTasks ?? 0) - (a.completedTasks ?? 0))
                      .map((u, i) => {
                        const name = (u.fullName ?? u.userName ?? `${u.firstName ?? ""} ${u.lastName ?? ""}`.trim()) || "نامشخص";
                        return (
                          <div key={u.userId ?? getId(u) ?? i} className="flex items-center gap-4 px-5 py-3">
                            <div className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-xs font-bold ${i === 0 ? "bg-amber-100 text-amber-600" : i === 1 ? "bg-slate-100 text-slate-600" : i === 2 ? "bg-orange-100 text-orange-600" : "bg-[--surface-2] text-[--text-3]"}`}>
                              {i + 1}
                            </div>
                            <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-emerald-400 to-teal-600 text-xs font-bold text-white">
                              {name[0]?.toUpperCase() ?? "؟"}
                            </div>
                            <div className="min-w-0 flex-1">
                              <p className="font-semibold">{name}</p>
                              <p className="text-xs text-[--text-3]">امتیاز: {u.score ?? 0}</p>
                            </div>
                            <div className="shrink-0 text-right">
                              <p className="text-xl font-bold text-emerald-600">{u.completedTasks ?? 0}</p>
                              <p className="text-[11px] text-[--text-3]">گزارش تکمیل</p>
                            </div>
                          </div>
                        );
                      })}
                  </div>
                </div>
              )}
            </section>
          )}

          {activeView === "team" && (
            <section className="space-y-4">
              <div className="rounded-2xl border border-[--border] bg-[--surface] p-5">
                <h2 className="font-bold">اعضای تیم</h2>
                <div className="mt-4 space-y-2">
                  {users.map((u) => (
                    <div key={getId(u)} className="flex items-center justify-between rounded-xl border border-[--border] bg-[--surface-2] px-4 py-3">
                      <div className="flex items-center gap-3">
                        <div className="flex h-9 w-9 items-center justify-center rounded-full bg-[#1f7a8c] text-xs font-bold text-white">{initials(u)}</div>
                        <div>
                          <p className="text-sm font-semibold">{userName(u)}</p>
                          <p className="text-xs text-[--text-3]">{u.mobile ?? u.email}</p>
                        </div>
                      </div>
                      {isManager ? (
                        <select
                          className="h-8 rounded-lg border border-[--border] bg-[--surface] px-2 text-xs"
                          value={u.roles ?? "specialist"}
                          onChange={(e) => void updateUserRole(getId(u), e.target.value)}
                        >
                          <option value="specialist">specialist</option>
                          <option value="supervisor">supervisor</option>
                          <option value="manager">manager</option>
                        </select>
                      ) : (
                        <span className="rounded-lg bg-[--surface] px-2.5 py-1 text-xs font-medium">{u.roles ?? "specialist"}</span>
                      )}
                    </div>
                  ))}
                </div>
              </div>
              <div className="rounded-2xl border border-[--border] bg-[--surface] p-5">
                <h2 className="font-bold">{isManager ? "درخواست‌های مرخصی" : "مرخصی من"}</h2>
                <form className="mt-4 grid gap-3 sm:grid-cols-2" onSubmit={createLeaveRequest}>
                  <Field label="از تاریخ" name="leaveStart" type="date" value={leaveStart} onChange={setLeaveStart} required />
                  <Field label="تا تاریخ" name="leaveEnd" type="date" value={leaveEnd} onChange={setLeaveEnd} required />
                  <Field label="دلیل" name="leaveReason" value={leaveReason} onChange={setLeaveReason} placeholder="اختیاری" />
                  <div className="flex items-end">
                    <button className="h-10 rounded-lg bg-[#1f7a8c] px-4 text-sm font-semibold text-white" type="submit">ثبت درخواست</button>
                  </div>
                </form>
                <div className="mt-4 space-y-2">
                  {leaveRequests.map((lr) => (
                    <div key={getId(lr)} className="flex items-center justify-between rounded-xl border border-[--border] px-4 py-3 text-sm">
                      <div>
                        <p className="font-semibold">{userName(lr.user)} · {statusLabel(lr.status)}</p>
                        <p className="text-xs text-[--text-3]">{formatDate(lr.startDate)} تا {formatDate(lr.endDate)}</p>
                      </div>
                      {isManager && lr.status === "pending" && (
                        <div className="flex gap-2">
                          <button className="rounded-lg bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-600" onClick={() => void handleLeaveAction(getId(lr), "approve")} type="button">تأیید</button>
                          <button className="rounded-lg bg-red-50 px-3 py-1 text-xs font-semibold text-red-600" onClick={() => void handleLeaveAction(getId(lr), "reject")} type="button">رد</button>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            </section>
          )}

          {isManager && activeView === "fixed-reports" && (
            <section className="space-y-4">
              {/* Stats */}
              <div className="grid grid-cols-2 gap-3 xl:grid-cols-4">
                {[
                  { label: "الگوهای ثابت", value: fixedTasks.length },
                  { label: "انجام‌نشده", value: incompleteFixedTasks.length },
                  { label: "فعال", value: fixedTasks.filter((item) => item.isActive !== false).length },
                  { label: "مهلت‌گذشته", value: incompleteFixedTasks.filter((item) => item.deadlineStatus === "overdue").length },
                ].map((s) => (
                  <div key={s.label} className="rounded-xl border border-[--border] bg-[--surface] p-4">
                    <p className="text-xs text-[--text-3]">{s.label}</p>
                    <p className="mt-1 text-2xl font-bold">{s.value}</p>
                  </div>
                ))}
              </div>

              {/* Tabs */}
              <div className="flex rounded-xl border border-[--border] bg-[--surface-2] p-1 w-fit">
                <button
                  className={`rounded-lg px-5 py-2 text-sm font-semibold transition-all ${fixedReportsTab === "templates" ? "bg-[--surface] text-[#1f7a8c] shadow-sm" : "text-[--text-2] hover:text-[--text]"}`}
                  onClick={() => setFixedReportsTab("templates")} type="button"
                >الگوهای ثابت ({fixedTasks.length})</button>
                <button
                  className={`rounded-lg px-5 py-2 text-sm font-semibold transition-all ${fixedReportsTab === "incomplete" ? "bg-[--surface] text-[#1f7a8c] shadow-sm" : "text-[--text-2] hover:text-[--text]"}`}
                  onClick={() => setFixedReportsTab("incomplete")} type="button"
                >انجام‌نشده ({incompleteFixedTasks.length})</button>
              </div>

              {/* ─ Templates tab ─ */}
              {fixedReportsTab === "templates" && (
                <div className="overflow-hidden rounded-2xl border border-[--border] bg-[--surface]">
                  <div className="flex items-center justify-between border-b border-[--border] px-5 py-4">
                    <div className="flex items-center gap-3">
                      <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-[#1f7a8c] to-[#165e6d] text-white">
                        <ClipboardList size={17} />
                      </div>
                      <div>
                        <h2 className="font-bold">مدیریت الگوهای ثابت</h2>
                        <p className="text-[11px] text-[--text-3]">{fixedTasks.length} الگو تعریف شده</p>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <button className="flex h-9 items-center gap-2 rounded-lg border border-[--border] bg-[--surface-2] px-4 text-sm font-semibold transition hover:bg-[--surface]" onClick={() => void loadManagerAnalytics()} type="button">
                        <RefreshCw size={15} />
                      </button>
                      <button
                        className="flex h-9 items-center gap-1.5 rounded-xl bg-gradient-to-l from-[#1f7a8c] to-[#2491a5] px-4 text-sm font-semibold text-white shadow-sm transition hover:shadow-md active:scale-[0.98]"
                        onClick={() => openFixedTaskForm()} type="button"
                      >
                        <Plus size={15} />الگوی جدید
                      </button>
                    </div>
                  </div>

                  {/* Create / Edit form */}
                  {showFixedTaskForm && (
                    <div className="border-b border-[--border] bg-[--surface-2]/70 p-4">
                      <p className="mb-3 text-sm font-bold">{editingFixedTask ? "ویرایش الگو" : "الگوی ثابت جدید"}</p>
                      <form className="grid gap-3 sm:grid-cols-2 xl:grid-cols-[1fr_180px_180px_1fr_auto]" onSubmit={saveFixedTask}>
                        <Field label="عنوان *" name="ftTitle" value={ftTitle} onChange={setFtTitle} required placeholder="مثلاً: گزارش روزانه" />
                        <label className="block">
                          <span className="mb-1.5 block text-xs font-semibold text-[--text-2]">توالی *</span>
                          <select
                            className="h-10 w-full rounded-lg border border-[--border] bg-[--surface] px-3 text-sm text-[--text] outline-none transition focus:border-[#1f7a8c] focus:ring-2 focus:ring-[#1f7a8c]/15"
                            value={ftRecurrence}
                            onChange={(e) => setFtRecurrence(e.target.value as "daily" | "weekly" | "monthly")}
                          >
                            <option value="daily">روزانه</option>
                            <option value="weekly">هفتگی</option>
                            <option value="monthly">ماهانه</option>
                          </select>
                        </label>
                        <Select label="مسئول *" value={ftAssignee} onChange={setFtAssignee}
                          options={users.filter((u) => u.roles === "specialist" || u.roles === "supervisor").map((u) => [getId(u), userName(u)])}
                          placeholder="انتخاب مسئول"
                        />
                        <Select label="پروژه" value={ftProjectId} onChange={setFtProjectId}
                          options={projects.map((p) => [getId(p), p.title])}
                          placeholder="مستقل (بدون پروژه)"
                        />
                        <div className="flex items-end gap-2">
                          <button className="flex h-10 flex-1 items-center justify-center gap-1.5 rounded-lg bg-[#1f7a8c] px-4 text-sm font-semibold text-white disabled:opacity-50 transition hover:bg-[#196b7b]" disabled={!ftTitle.trim() || !ftAssignee} type="submit">
                            {editingFixedTask ? "ذخیره" : <><Plus size={14} />ایجاد</>}
                          </button>
                          <button className="flex h-10 w-10 items-center justify-center rounded-lg border border-[--border] bg-[--surface] text-[--text-2] transition hover:bg-[--surface-2]" onClick={closeFixedTaskForm} type="button"><X size={15} /></button>
                        </div>
                      </form>
                      <Field label="توضیحات" name="ftDescription" value={ftDescription} onChange={setFtDescription} placeholder="اختیاری" />
                    </div>
                  )}

                  {/* Templates list */}
                  <div className="divide-y divide-[--border]">
                    {fixedTasks.length === 0 ? (
                      <div className="flex flex-col items-center py-12 text-center">
                        <ClipboardList size={32} className="text-[--text-3]" />
                        <p className="mt-3 font-semibold text-[--text]">الگویی تعریف نشده</p>
                        <button className="mt-4 flex h-9 items-center gap-2 rounded-lg bg-[#1f7a8c] px-4 text-sm font-semibold text-white" onClick={() => openFixedTaskForm()} type="button">
                          <Plus size={15} />اولین الگو را بساز
                        </button>
                      </div>
                    ) : (
                      fixedTasks.map((ft) => (
                        <div key={getId(ft)} className="flex flex-wrap items-center justify-between gap-3 px-5 py-4">
                          <div className="min-w-0 flex-1">
                            <div className="flex items-center gap-2">
                              <p className="font-semibold">{ft.title}</p>
                              <span className={`rounded-full px-2 py-0.5 text-[10px] font-bold ${ft.isActive !== false ? "bg-emerald-100 text-emerald-700 dark:bg-emerald-950/50 dark:text-emerald-400" : "bg-slate-100 text-slate-500 dark:bg-slate-800 dark:text-slate-400"}`}>
                                {ft.isActive !== false ? "فعال" : "غیرفعال"}
                              </span>
                              <span className="rounded-full bg-[--surface-2] px-2 py-0.5 text-[10px] font-semibold text-[--text-2]">
                                {ft.recurrence === "daily" ? "روزانه" : ft.recurrence === "weekly" ? "هفتگی" : "ماهانه"}
                              </span>
                            </div>
                            <p className="mt-0.5 text-xs text-[--text-3]">
                              {ft.assignedTo ? `مسئول: ${userName(ft.assignedTo)}` : ""}
                              {ft.projectId ? ` · پروژه: ${typeof ft.projectId === "string" ? ft.projectId : ft.projectId.title}` : ""}
                              {ft.description ? ` · ${ft.description}` : ""}
                            </p>
                          </div>
                          <div className="flex items-center gap-2 shrink-0">
                            <button
                              className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition ${ft.isActive !== false ? "bg-slate-100 text-slate-600 hover:bg-slate-200 dark:bg-slate-800 dark:text-slate-300" : "bg-emerald-50 text-emerald-600 hover:bg-emerald-100 dark:bg-emerald-950/40 dark:text-emerald-400"}`}
                              onClick={() => void toggleFixedTaskActive(ft)} type="button"
                            >
                              {ft.isActive !== false ? "غیرفعال کن" : "فعال کن"}
                            </button>
                            <button
                              className="rounded-lg border border-[--border] bg-[--surface] px-3 py-1.5 text-xs font-semibold text-[--text-2] transition hover:bg-[--surface-2]"
                              onClick={() => openFixedTaskForm(ft)} type="button"
                            >ویرایش</button>
                            <button
                              className="flex h-8 w-8 items-center justify-center rounded-lg text-red-500 transition hover:bg-red-50 dark:hover:bg-red-950/40"
                              onClick={() => void deleteFixedTask(getId(ft))} type="button"
                            ><Trash2 size={14} /></button>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              )}

              {/* ─ Incomplete tab ─ */}
              {fixedReportsTab === "incomplete" && (
                <div className="overflow-hidden rounded-2xl border border-[--border] bg-[--surface]">
                  <div className="flex items-center justify-between border-b border-[--border] px-5 py-4">
                    <div className="flex items-center gap-3">
                      <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-amber-100 text-amber-600 dark:bg-amber-950/50 dark:text-amber-400">
                        <AlertTriangle size={17} />
                      </div>
                      <div>
                        <h2 className="font-bold">گزارش‌های ثابت انجام‌نشده</h2>
                        <p className="text-[11px] text-[--text-3]">{incompleteFixedTasks.length} مورد نیاز به بررسی</p>
                      </div>
                    </div>
                    <button className="flex h-9 items-center gap-2 rounded-lg border border-[--border] bg-[--surface-2] px-4 text-sm font-semibold transition hover:bg-[--surface]" onClick={() => void loadManagerAnalytics()} type="button">
                      <RefreshCw size={15} />بروزرسانی
                    </button>
                  </div>
                  <div className="divide-y divide-[--border]">
                    {incompleteFixedTasks.length === 0 ? (
                      <p className="py-10 text-center text-sm text-[--text-3]">گزارش ثابت انجام‌نشده‌ای یافت نشد</p>
                    ) : (
                      incompleteFixedTasks.map((item) => (
                        <div key={getId(item)} className="flex flex-wrap items-center justify-between gap-3 px-5 py-4">
                          <div>
                            <p className="font-semibold">{item.title}</p>
                            <p className="mt-1 text-xs text-[--text-3]">
                              {item.recurrence === "daily" ? "روزانه" : item.recurrence === "weekly" ? "هفتگی" : "ماهانه"}
                              {item.assignedTo ? ` · مسئول: ${userName(item.assignedTo)}` : ""}
                              {item.projectId ? ` · پروژه: ${typeof item.projectId === "string" ? item.projectId : item.projectId.title}` : ""}
                              {item.deadline ? ` · مهلت: ${formatDate(item.deadline)}` : ""}
                            </p>
                          </div>
                          <span className={`rounded-lg px-3 py-1 text-xs font-semibold ${item.deadlineStatus === "overdue" ? "bg-amber-50 text-amber-600 dark:bg-amber-950/40 dark:text-amber-400" : "bg-[#e8f4f7] text-[#1f7a8c] dark:bg-[#0f3040] dark:text-[#4fc3d5]"}`}>
                            {item.deadlineStatus === "overdue" ? "مهلت گذشته" : "در مهلت"}
                          </span>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              )}
            </section>
          )}

          {activeView === "excel" && (
            <section className="space-y-4">
              <div className="grid grid-cols-2 gap-3 xl:grid-cols-5">
                {[
                  { label: "کل فایل‌ها", value: excelStats?.totalFiles ?? 0 },
                  { label: "ایمپورت", value: excelStats?.totalImports ?? 0 },
                  { label: "اکسپورت", value: excelStats?.totalExports ?? 0 },
                  { label: "موفق", value: excelStats?.completedImports ?? 0 },
                  { label: "ناموفق", value: excelStats?.failedImports ?? 0 },
                ].map((s) => (
                  <div key={s.label} className="rounded-xl border border-[--border] bg-[--surface] p-4">
                    <p className="text-xs text-[--text-3]">{s.label}</p>
                    <p className="mt-1 text-2xl font-bold">{s.value}</p>
                  </div>
                ))}
              </div>

              <div className="rounded-2xl border border-[--border] bg-[--surface] p-5">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <h2 className="font-bold">مدیریت فایل‌های اکسل</h2>
                  <div className="flex flex-wrap gap-2">
                    <label className="flex h-9 cursor-pointer items-center gap-2 rounded-lg bg-[#1f7a8c] px-4 text-sm font-semibold text-white transition hover:bg-[#196b7b]">
                      {excelUploading ? <Loader2 className="animate-spin" size={15} /> : <Upload size={15} />}
                      آپلود فایل
                      <input accept=".xlsx,.xls,.csv" className="hidden" disabled={excelUploading} onChange={handleExcelUpload} type="file" />
                    </label>
                    <button
                      className="flex h-9 items-center gap-2 rounded-lg border border-[--border] bg-[--surface-2] px-4 text-sm font-semibold transition hover:bg-[--surface]"
                      disabled={!tasks.length}
                      onClick={() => void exportTasksToExcel()} type="button"
                    >
                      <Download size={15} />خروجی گزارش‌ها
                    </button>
                  </div>
                </div>

                <div className="mt-4 space-y-2">
                  {excelFiles.length === 0 ? (
                    <p className="rounded-xl border border-dashed border-[--border] p-8 text-center text-sm text-[--text-3]">هنوز فایلی آپلود نشده</p>
                  ) : (
                    excelFiles.map((f) => (
                      <div key={getId(f)} className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-[--border] bg-[--surface-2] px-4 py-3">
                        <div>
                          <p className="text-sm font-semibold">{f.originalName || f.fileName}</p>
                          <p className="text-xs text-[--text-3]">
                            {f.type === "import" ? "ایمپورت" : "اکسپورت"} · {statusLabel(f.status)} · {f.totalRows ?? 0} ردیف
                          </p>
                        </div>
                        <div className="flex gap-2">
                          <button className="flex h-8 items-center gap-1 rounded-lg border border-[--border] bg-[--surface] px-3 text-xs font-semibold" onClick={() => void downloadExcelFile(f)} type="button">
                            <Download size={13} />دانلود
                          </button>
                          {f.type === "import" && f.status === "pending" && (
                            <button className="flex h-8 items-center rounded-lg bg-emerald-50 px-3 text-xs font-semibold text-emerald-600" onClick={() => void processExcelFile(getId(f))} type="button">
                              پردازش
                            </button>
                          )}
                          <button className="flex h-8 w-8 items-center justify-center rounded-lg text-red-500 hover:bg-red-50" onClick={() => void deleteExcelFile(getId(f))} type="button">
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </section>
          )}

          {activeView === "settings" && (
            <section className="rounded-2xl border border-[--border] bg-[--surface] p-5">
              <h2 className="font-bold">تنظیمات پروفایل</h2>
              <form className="mt-4 grid max-w-lg gap-3" onSubmit={saveProfile}>
                <Field label="نام" name="settingsFirstName" value={settingsFirstName} onChange={setSettingsFirstName} />
                <Field label="نام خانوادگی" name="settingsLastName" value={settingsLastName} onChange={setSettingsLastName} />
                <Field label="ایمیل" name="settingsEmail" type="email" value={settingsEmail} onChange={setSettingsEmail} />
                <button className="h-10 w-fit rounded-lg bg-[#1f7a8c] px-5 text-sm font-semibold text-white" type="submit">ذخیره تغییرات</button>
              </form>
            </section>
          )}

          {((!isSupervisor && (activeView === "dashboard" || activeView === "tasks" || activeView === "projects")) || (isSupervisor && activeView === "tasks")) && (<>

          {/* Welcome banner */}
          <div className={`relative overflow-hidden rounded-2xl px-6 py-5 text-white shadow-lg ${isManager ? "bg-gradient-to-l from-indigo-700 via-indigo-600 to-indigo-500 shadow-indigo-500/15" : "bg-gradient-to-l from-[#1a6b7c] via-[#1f7a8c] to-[#2491a5] shadow-[#1f7a8c]/15"}`}>
            <div className="pointer-events-none absolute -left-6 -top-6 h-36 w-36 rounded-full bg-white/5" />
            <div className="pointer-events-none absolute -bottom-8 left-24 h-28 w-28 rounded-full bg-white/5" />
            <div className="relative flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <p className="text-sm font-medium opacity-75">سلام، {userName(currentUser ?? undefined).split(" ")[0]}</p>
                <h1 className="mt-0.5 text-xl font-bold">{isManager ? "داشبورد مدیر" : "داشبورد گزارش‌ها"}</h1>
                <p className="mt-1 text-sm opacity-75">{isManager ? `${managerStats?.activeProjects ?? projects.length} پروژه فعال · ${managerStats?.activeUsers ?? users.length} کاربر` : activeTasks === 0 ? "همه گزارش‌ها تکمیل شده‌اند" : `${activeTasks} گزارش باز داری`}</p>
              </div>
              <div className="flex shrink-0 items-center gap-5">
                {isManager
                  ? [{ n: managerStats?.activeProjects ?? projects.length, l: "پروژه فعال" }, { n: managerStats?.openTasks ?? tasks.length, l: "گزارش باز" }, { n: managerStats?.activeUsers ?? users.length, l: "کاربر فعال" }].map((s, i) => (
                      <div key={i} className="text-center"><p className="text-2xl font-extrabold">{s.n}</p><p className="text-[11px] opacity-75">{s.l}</p></div>
                    ))
                  : [{ n: tasks.length, l: "کل گزارش" }, { n: projects.length, l: "پروژه" }, { n: `${progress}%`, l: "پیشرفت" }].map((s, i) => (
                      <div key={i} className="text-center"><p className="text-2xl font-extrabold">{s.n}</p><p className="text-[11px] opacity-75">{s.l}</p></div>
                    ))
                }
              </div>
            </div>
          </div>

          {/* Stats */}
          <div className="grid grid-cols-2 gap-3 xl:grid-cols-4">
            {(isManager
              ? [
                  { label: "پروژه‌ها", value: statsProjects, sub: "فعال", icon: FolderKanban, a: "bg-indigo-50 text-indigo-600 ring-indigo-100 dark:bg-indigo-950/40 dark:text-indigo-400 dark:ring-indigo-900", onClick: () => setActiveView("projects") },
                  { label: "گزارش‌های باز", value: managerStats?.openTasks ?? activeTasks, sub: `${inProgressTasks} جاری`, icon: ClipboardList, a: "bg-[#e8f4f7] text-[#1f7a8c] ring-[#1f7a8c]/10 dark:bg-[#0f3040] dark:text-[#4fc3d5] dark:ring-[#1f7a8c]/20", onClick: () => setActiveView("tasks") },
                  { label: "کاربران فعال", value: managerStats?.activeUsers ?? statsUsers, sub: "کاربر", icon: UsersRound, a: "bg-amber-50 text-amber-600 ring-amber-100 dark:bg-amber-950/40 dark:text-amber-400 dark:ring-amber-900", onClick: () => setActiveView("team") },
                  { label: "آنالیتیکس", value: managerTaskStatus?.doneTasks ?? managerTaskStatus?.done ?? doneTasks, sub: "گزارش تکمیل", icon: BarChart2, a: "bg-emerald-50 text-emerald-600 ring-emerald-100 dark:bg-emerald-950/40 dark:text-emerald-400 dark:ring-emerald-900", onClick: () => setActiveView("analytics") },
                ]
              : [
                  { label: "پروژه‌ها", value: statsProjects, sub: "فعال", icon: FolderKanban, a: "bg-violet-50 text-violet-600 ring-violet-100 dark:bg-violet-950/40 dark:text-violet-400 dark:ring-violet-900", onClick: undefined },
                  { label: "گزارش‌های باز", value: activeTasks, sub: `${inProgressTasks} جاری`, icon: ClipboardList, a: "bg-[#e8f4f7] text-[#1f7a8c] ring-[#1f7a8c]/10 dark:bg-[#0f3040] dark:text-[#4fc3d5] dark:ring-[#1f7a8c]/20", onClick: undefined },
                  { label: "اعضای تیم", value: statsUsers, sub: "کاربر", icon: UsersRound, a: "bg-amber-50 text-amber-600 ring-amber-100 dark:bg-amber-950/40 dark:text-amber-400 dark:ring-amber-900", onClick: undefined },
                  { label: "تکمیل شده", value: doneTasks, sub: `${progress}% پیشرفت`, icon: TrendingUp, a: "bg-emerald-50 text-emerald-600 ring-emerald-100 dark:bg-emerald-950/40 dark:text-emerald-400 dark:ring-emerald-900", onClick: undefined },
                ]
            ).map((s) => (
              <div key={s.label} className={`group rounded-xl border border-[--border] bg-[--surface] p-4 transition-all hover:-translate-y-0.5 hover:shadow-md hover:shadow-black/5 ${s.onClick ? "cursor-pointer" : ""}`} onClick={s.onClick}>
                <div className="flex items-start justify-between gap-2">
                  <p className="text-xs font-medium text-[--text-2]">{s.label}</p>
                  <span className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg ring-4 ${s.a}`}><s.icon size={15} /></span>
                </div>
                <p className="mt-2.5 text-3xl font-bold">{s.value}</p>
                <p className="mt-0.5 text-xs text-[--text-3]">{s.sub}</p>
              </div>
            ))}
          </div>

          {/* Manager: Pending leave requests on dashboard */}
          {isManager && activeView === "dashboard" && leaveRequests.filter((lr) => lr.status === "pending").length > 0 && (
            <div className="overflow-hidden rounded-2xl border border-amber-200 dark:border-amber-900 bg-[--surface]">
              <div className="flex items-center justify-between border-b border-amber-100 dark:border-amber-900/50 bg-gradient-to-l from-amber-50 to-white dark:from-amber-950/30 dark:to-transparent px-5 py-4">
                <div className="flex items-center gap-3">
                  <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-amber-100 text-amber-600 dark:bg-amber-950/50 dark:text-amber-400">
                    <CalendarDays size={17} />
                  </div>
                  <div>
                    <h2 className="font-bold">درخواست‌های مرخصی در انتظار</h2>
                    <p className="text-[11px] text-[--text-3]">{leaveRequests.filter((lr) => lr.status === "pending").length} درخواست نیاز به بررسی دارد</p>
                  </div>
                </div>
                <button className="text-xs font-semibold text-[#1f7a8c] hover:underline" onClick={() => setActiveView("team")} type="button">مشاهده همه</button>
              </div>
              <div className="divide-y divide-[--border]">
                {leaveRequests.filter((lr) => lr.status === "pending").slice(0, 5).map((lr) => (
                  <div key={getId(lr)} className="flex items-center justify-between px-5 py-3">
                    <div>
                      <p className="text-sm font-semibold">{userName(lr.user)}</p>
                      <p className="text-xs text-[--text-3]">{formatDate(lr.startDate)} تا {formatDate(lr.endDate)}{lr.reason ? ` · ${lr.reason}` : ""}</p>
                    </div>
                    <div className="flex gap-2">
                      <button className="rounded-lg bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-600 hover:bg-emerald-100 transition" onClick={() => void handleLeaveAction(getId(lr), "approve")} type="button">تأیید</button>
                      <button className="rounded-lg bg-red-50 px-3 py-1 text-xs font-semibold text-red-600 hover:bg-red-100 transition" onClick={() => void handleLeaveAction(getId(lr), "reject")} type="button">رد</button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Kanban board */}
          {(activeView === "dashboard" || activeView === "tasks") && <div className="overflow-hidden rounded-2xl border border-[#b8dfe8] dark:border-[#1f5060] bg-[--surface] shadow-md shadow-[#1f7a8c]/8">
            <div className="flex flex-col gap-3 border-b border-[#cce8ef] dark:border-[#1f5060] bg-gradient-to-l from-[#e0f4f8] to-[#f0fafb] dark:from-[#0f2535] dark:to-[#0f172a] px-5 py-4 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex items-center gap-3">
                <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-[#1f7a8c] to-[#165e6d] text-white shadow-sm shadow-[#1f7a8c]/20">
                  <ClipboardList size={17} />
                </div>
                <div>
                  <h2 className="font-bold text-[--text]">برد گزارش ها</h2>
                  <p className="text-[11px] text-[--text-3]">کانبان با drag & drop · {tasks.length} گزارش</p>
                </div>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <span className="flex h-8 items-center gap-1.5 rounded-lg bg-[--surface-2] px-3 text-xs font-medium text-[--text-2]">
                  <ListFilter size={12} />{filteredTasks.length}/{tasks.length}
                </span>
                {hasTaskFilters && (
                  <button
                    className="flex h-8 items-center gap-1 rounded-lg border border-red-200 bg-red-50 px-2.5 text-xs font-medium text-red-600 transition hover:bg-red-100 dark:border-red-900 dark:bg-red-950/40 dark:text-red-400"
                    onClick={() => {
                      setTaskQuery("");
                      setSelectedProjectFilter("");
                      setSelectedStatusFilter("");
                      setSelectedAssigneeFilter("");
                      setSelectedPriorityFilter("");
                      setSelectedPeriodFilter("");
                    }}
                    type="button"
                  >
                    <X size={11} /> پاک فیلترها
                  </button>
                )}
              </div>
            </div>

            <div className="flex gap-2 overflow-x-auto border-b border-[--border] bg-[--surface-2]/50 px-4 py-2.5 scrollbar-hide">
              <FilterChip active={!selectedProjectFilter} label="همه" count={tasks.length} onClick={() => setSelectedProjectFilter("")} />
              {projects.slice(0, 6).map((p) => {
                const pid = getId(p);
                const count = tasks.filter((t) => getId(t.projectId) === pid).length;
                return (
                  <FilterChip
                    key={pid}
                    active={selectedProjectFilter === pid}
                    label={p.title}
                    count={count || projectProgress[pid]?.totalTasks || 0}
                    onClick={() => setSelectedProjectFilter(selectedProjectFilter === pid ? "" : pid)}
                  />
                );
              })}
            </div>

            <div className="grid gap-3 border-b border-[--border] bg-[--surface] px-4 py-3 sm:grid-cols-2 xl:grid-cols-4">
              <Select
                label="وضعیت"
                value={selectedStatusFilter}
                onChange={setSelectedStatusFilter}
                options={COLUMNS.map((col) => [col.status, col.title])}
                placeholder="همه وضعیت‌ها"
              />
              <Select
                label="مسئول"
                value={selectedAssigneeFilter}
                onChange={setSelectedAssigneeFilter}
                options={users.map((u) => [getId(u), userName(u)])}
                placeholder="همه مسئول‌ها"
              />
              <Select
                label="اولویت"
                value={selectedPriorityFilter}
                onChange={(value) => setSelectedPriorityFilter(value as Priority | "")}
                options={(Object.keys(PRIORITY) as Priority[]).map((key) => [key, PRIORITY[key].label])}
                placeholder="همه اولویت‌ها"
              />
              <Select
                label="دوره گزارش"
                value={selectedPeriodFilter}
                onChange={(value) => setSelectedPeriodFilter(value as TaskPeriod | "")}
                options={TASK_PERIODS}
                placeholder="همه دوره‌ها"
              />
            </div>

            <DragDropContext onDragEnd={onDragEnd}>
              <div className="grid gap-4 bg-[--surface-2]/40 p-4 lg:grid-cols-3">
                {COLUMNS.map((col) => {
                  const colTasks = filteredTasks.filter((t) => (t.status ?? "todo") === col.status);
                  return (
                    <div key={col.status} className={`flex flex-col rounded-2xl border ${col.border} ${col.colBg}`}>
                      <div className={`flex items-center justify-between rounded-t-2xl bg-gradient-to-l ${col.headerGrad} px-4 py-3`}>
                        <div className="flex items-center gap-2">
                          <span className={`h-2.5 w-2.5 rounded-full ${col.dot}`} />
                          <h3 className={`text-sm font-bold ${col.headerText}`}>{col.title}</h3>
                        </div>
                        <span className={`rounded-full px-2.5 py-0.5 text-xs font-bold ${col.badge}`}>{colTasks.length}</span>
                      </div>

                      <Droppable droppableId={col.status}>
                        {(provided, snapshot) => (
                          <div
                            ref={provided.innerRef}
                            {...provided.droppableProps}
                            className={`flex flex-col gap-2.5 p-2.5 min-h-[120px] transition-colors ${snapshot.isDraggingOver ? "bg-white/30 dark:bg-white/5" : ""}`}
                          >
                            {colTasks.length === 0 && !snapshot.isDraggingOver ? (
                              <div className={`flex flex-col items-center justify-center rounded-xl border-2 border-dashed ${col.emptyBorder} bg-white/50 dark:bg-black/10 py-10`}>
                                <CircleDashed size={28} className={`${col.headerText} opacity-30`} />
                                <p className="mt-2 text-xs text-[--text-3]">گزارشی اینجا نیست</p>
                              </div>
                            ) : (
                              colTasks.map((task, index) => {
                                const tid = getId(task);
                                const priority = taskPriorities[tid];
                                const canClaim = isSpecialist && isUnassignedTask(task);
                                return (
                                  <Draggable key={tid} draggableId={tid} index={index}>
                                    {(prov, snap) => (
                                      <article
                                        ref={prov.innerRef}
                                        {...prov.draggableProps}
                                        className={`rounded-xl border border-[--border] border-t-[3px] ${col.cardBorder} bg-[--surface] p-3.5 shadow-sm transition-all cursor-pointer ${snap.isDragging ? "shadow-xl shadow-black/15 rotate-1 scale-[1.02]" : "hover:-translate-y-0.5 hover:shadow-md hover:shadow-black/5"}`}
                                        onClick={() => setSelectedTask(task)}
                                      >
                                        <div className="flex items-center justify-between">
                                          <div className="flex items-center gap-1.5">
                                            <div {...prov.dragHandleProps} className="text-[--text-3] hover:text-[--text-2] cursor-grab active:cursor-grabbing" onClick={(e) => e.stopPropagation()}>
                                              <GripVertical size={14} />
                                            </div>
                                            {priority && (
                                              <span className={`rounded-md px-1.5 py-0.5 text-[10px] font-bold ${PRIORITY[priority].color}`}>
                                                {PRIORITY[priority].label}
                                              </span>
                                            )}
                                          </div>
                                          <button className="flex h-6 w-6 items-center justify-center rounded-lg text-[--text-3] transition hover:bg-[--surface-2] hover:text-[--text-2]" type="button" onClick={(e) => { e.stopPropagation(); setSelectedTask(task); }}>
                                            <MoreHorizontal size={14} />
                                          </button>
                                        </div>

                                        <div className="mt-2.5 flex items-start gap-2">
                                          {task.status === "done"
                                            ? <CheckCircle2 size={15} className="mt-0.5 shrink-0 text-emerald-500" />
                                            : task.status === "in_progress"
                                            ? <div className="mt-1 h-3 w-3 shrink-0 rounded-full border-2 border-[#1f7a8c] border-t-transparent animate-spin" />
                                            : <CircleDashed size={15} className="mt-0.5 shrink-0 text-[--text-3]" />
                                          }
                                          <h4 className="text-sm font-semibold leading-snug">{task.title}</h4>
                                        </div>

                                        <div className="mt-3 flex items-center justify-between">
                                          <AssigneeStack users={task.assignedTo} />
                                          <div className="flex items-center gap-1 rounded-md bg-[--surface-2] px-2 py-1 text-[10px] text-[--text-3]">
                                            <CalendarDays size={10} />{formatDate(task.dueDate ?? task.createdAt)}
                                          </div>
                                        </div>

                                        {canClaim && (
                                          <button
                                            className="mt-3 flex h-8 w-full items-center justify-center gap-1.5 rounded-lg border border-[#b8dfe8] bg-[#e0f4f8] text-xs font-semibold text-[#1f7a8c] transition-all hover:bg-[#d0eef5] active:scale-[0.98] dark:border-[#1f5060] dark:bg-[#0f3040] dark:text-[#4fc3d5]"
                                            onClick={(e) => { e.stopPropagation(); void claimTask(tid); }}
                                            type="button"
                                          >
                                            <UserPlus size={12} /> برداشتن برای من
                                          </button>
                                        )}
                                        {!canClaim && task.status !== "done" && (
                                          <button
                                            className={`mt-3 flex h-8 w-full items-center justify-center gap-1.5 rounded-lg border border-[--border] bg-[--surface-2] text-xs font-semibold text-[--text-2] transition-all ${col.btnHover} active:scale-[0.98]`}
                                            onClick={(e) => { e.stopPropagation(); void moveTask(tid, nextStatus(task.status)); }} type="button"
                                          >
                                            {task.status === "todo" ? "شروع کار" : "تکمیل کردن"}
                                            <ChevronLeft size={12} />
                                          </button>
                                        )}
                                        {task.status === "done" && (
                                          <div className="mt-3 flex items-center justify-center gap-1.5 rounded-lg bg-emerald-50 dark:bg-emerald-950/40 py-1.5 text-[11px] font-semibold text-emerald-600 dark:text-emerald-400">
                                            <CheckCircle2 size={12} />تکمیل شد
                                          </div>
                                        )}
                                      </article>
                                    )}
                                  </Draggable>
                                );
                              })
                            )}
                            {provided.placeholder}
                          </div>
                        )}
                      </Droppable>
                    </div>
                  );
                })}
              </div>
            </DragDropContext>
          </div>}

          {/* Projects */}
          {(activeView === "dashboard" || activeView === "projects") && <div className="overflow-hidden rounded-2xl border border-violet-200 dark:border-violet-900 bg-[--surface] shadow-md shadow-violet-500/8">
            <div className="flex items-center justify-between border-b border-violet-100 dark:border-violet-900/50 bg-gradient-to-l from-violet-50 to-white dark:from-violet-950/30 dark:to-transparent px-5 py-4">
              <div className="flex items-center gap-3">
                <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-violet-500 to-violet-600 text-white shadow-sm shadow-violet-200 dark:shadow-violet-900">
                  <FolderKanban size={17} />
                </div>
                <div>
                  <h2 className="font-bold text-[--text]">پروژه‌ها</h2>
                  <p className="text-[11px] text-[--text-3]">{projects.length} پروژه فعال</p>
                </div>
              </div>
              {isManager && (
                <button className="flex h-9 items-center gap-1.5 rounded-xl bg-gradient-to-l from-[#1f7a8c] to-[#2491a5] px-4 text-sm font-semibold text-white shadow-sm shadow-[#1f7a8c]/20 transition-all active:scale-[0.98] hover:shadow-md" onClick={() => setShowProjectForm(!showProjectForm)} type="button">
                  <Plus size={15} />پروژه جدید
                </button>
              )}
            </div>

            {isManager && showProjectForm && (
              <div className="border-b border-[--border] bg-[--surface-2]/70 p-4">
                <form className="grid gap-3 sm:grid-cols-2 xl:grid-cols-[1fr_1fr_180px_200px_200px_auto]" onSubmit={createProject}>
                  <Field label="عنوان پروژه *" name="projectTitle" value={projectTitle} onChange={setProjectTitle} required placeholder="مثلاً: ری‌دیزاین اپ" />
                  <Field label="توضیحات" name="projectDescription" value={projectDescription} onChange={setProjectDescription} placeholder="اختیاری" />
                  <Select label="حوزه کاری" value={projectWorkField} onChange={(value) => setProjectWorkField(value as WorkField)} options={WORK_FIELDS} placeholder="انتخاب حوزه" />
                  <Select label="سرپرست *" value={projectSupervisorId} onChange={setProjectSupervisorId} options={users.filter((u) => u.roles === "supervisor").map((u) => [getId(u), userName(u)])} placeholder="انتخاب سرپرست" />
                  <Select label="متخصص مسئول" value={projectAssigneeId} onChange={setProjectAssigneeId} options={users.filter((u) => u.roles === "specialist").map((u) => [getId(u), userName(u)])} placeholder="پروژه عمومی" />
                  <div className="flex items-end gap-2">
                    <button className="flex h-10 flex-1 items-center justify-center gap-2 rounded-lg bg-[--text] px-4 text-sm font-semibold text-[--surface] transition hover:opacity-90 disabled:opacity-50 lg:flex-none" disabled={!projectTitle.trim() || !projectSupervisorId} type="submit"><Plus size={15} />ایجاد</button>
                    <button className="flex h-10 w-10 items-center justify-center rounded-lg border border-[--border] bg-[--surface] text-[--text-2] transition hover:bg-[--surface-2]" onClick={() => setShowProjectForm(false)} type="button"><X size={15} /></button>
                  </div>
                </form>
              </div>
            )}

            <div className="grid gap-4 p-4 md:grid-cols-2">
              {loadingData && projects.length === 0 ? (
                Array.from({ length: 4 }).map((_, i) => <SkeletonCard key={i} />)
              ) : projects.length === 0 ? (
                <EmptyState title="هنوز پروژه‌ای نداری" text={isManager ? "روی «پروژه جدید» بزن." : "هنوز پروژه‌ای تعریف نشده."} action={isManager ? { label: "ساخت اولین پروژه", onClick: () => setShowProjectForm(true) } : undefined} />
              ) : (
                projects.map((project, i) => {
                  const pid = getId(project);
                  const prog = projectProgress[pid];
                  return (
                  <article
                    key={pid}
                    className="group overflow-hidden rounded-xl border border-[--border] bg-[--surface] transition-all hover:-translate-y-0.5 hover:shadow-lg hover:shadow-black/5 cursor-pointer"
                    onClick={() => setSelectedProject(project)}
                  >
                    <div className={`relative h-16 bg-gradient-to-l ${PROJECT_COVERS[i % PROJECT_COVERS.length]} px-4 pt-3`}>
                      <div className="absolute inset-0 opacity-10" style={{ backgroundImage: "radial-gradient(circle at 20% 50%, white 1px, transparent 1px)", backgroundSize: "20px 20px" }} />
                      <div className="relative flex items-start justify-between">
                        <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-white/20 text-white border border-white/30 backdrop-blur-sm"><FolderKanban size={17} /></div>
                        <div className="flex items-center gap-2">
                          {prog && <span className="rounded-full bg-white/20 px-2.5 py-0.5 text-xs font-semibold text-white border border-white/20 backdrop-blur-sm">{prog.progressPercentage}%</span>}
                          <span className="rounded-full bg-white/20 px-2.5 py-0.5 text-xs font-semibold text-white border border-white/20 backdrop-blur-sm">{statusLabel(project.status)}</span>
                        </div>
                      </div>
                    </div>
                    <div className="p-4">
                      <h3 className="font-bold">{project.title}</h3>
                      <p className="mt-0.5 line-clamp-1 text-xs text-[--text-3]">{project.description || "بدون توضیحات"}</p>
                      {prog && (
                        <div className="mt-3 h-1.5 overflow-hidden rounded-full bg-[--border]">
                          <div className="h-full rounded-full bg-[#1f7a8c]" style={{ width: `${prog.progressPercentage}%` }} />
                        </div>
                      )}
                      <div className="mt-3 flex items-center justify-between border-t border-[--border] pt-3">
                        <AssigneeStack users={[project.assigneeId, project.supervisorId].filter(Boolean) as Array<string | User>} fallback={project.owner} />
                        <div className="flex flex-wrap gap-1.5" onClick={(e) => e.stopPropagation()}>
                          {isManager && (
                            <button className="flex items-center gap-1 rounded-lg bg-[#e8f4f7] px-2.5 py-1 text-[11px] font-semibold text-[#1f7a8c]" onClick={() => void loadProjectMembers(project)} type="button">
                              <UserPlus size={11} />افراد
                            </button>
                          )}
                          <span className="rounded-lg bg-[--surface-2] px-2.5 py-1 text-[11px] font-semibold text-[--text-2]">👥 {[project.owner, project.supervisorId, project.assigneeId].filter(Boolean).length}</span>
                          <span className="rounded-lg bg-[--surface-2] px-2.5 py-1 text-[11px] font-semibold text-[--text-2]">✅ {prog?.totalTasks ?? project.tasks?.length ?? 0}</span>
                          {isManager && (
                            <>
                              <button
                                className={`rounded-lg px-2 py-1 text-[11px] font-semibold transition ${project.isArchived ? "bg-emerald-50 text-emerald-600" : "bg-slate-100 text-slate-600"}`}
                                onClick={() => void managerApi.setProjectActivation(token, pid, !!project.isArchived).then(() => { setMessage("وضعیت پروژه تغییر کرد."); void loadData(); }).catch(() => setError("تغییر وضعیت ناموفق بود"))}
                                type="button"
                              >
                                {project.isArchived ? "فعال‌سازی" : "آرشیو"}
                              </button>
                              <button className="rounded-lg bg-red-50 px-2 py-1 text-[11px] font-semibold text-red-600" onClick={() => void deleteProject(pid)} type="button">حذف</button>
                            </>
                          )}
                        </div>
                      </div>
                    </div>
                  </article>
                  );
                })
              )}
            </div>
          </div>}

          </>)}

        </main>
      </div>

      {/* ─ Project Members Panel ──────────────────────────────────────────── */}
      {membersProject && (
        <>
          <div className="fixed inset-0 z-50 bg-black/30 backdrop-blur-sm" onClick={() => setMembersProject(null)} />
          <div className="fixed inset-y-0 left-0 z-50 flex w-full max-w-md flex-col bg-[--surface] shadow-2xl">
            <div className="flex items-center justify-between border-b border-[--border] px-5 py-4">
              <div>
                <h3 className="font-bold">افراد پروژه</h3>
                <p className="text-xs text-[--text-3]">{membersProject.title}</p>
              </div>
              <button className="flex h-8 w-8 items-center justify-center rounded-lg text-[--text-3] hover:bg-[--surface-2]" onClick={() => setMembersProject(null)} type="button">
                <X size={18} />
              </button>
            </div>
            <div className="flex-1 overflow-y-auto p-5 space-y-4">
              <div className="space-y-2">
                {projectMembers.length === 0 ? (
                  <p className="text-center text-xs text-[--text-3] py-4">فردی ثبت نشده</p>
                ) : (
                  projectMembers.map((m) => {
                    const uid = getId(m.user);
                    return (
                      <div key={m.id ?? uid} className="flex items-center justify-between rounded-xl border border-[--border] px-4 py-3">
                        <div className="flex items-center gap-3">
                          <div className="flex h-8 w-8 items-center justify-center rounded-full bg-[#1f7a8c] text-xs font-bold text-white">{initials(m.user)}</div>
                          <div>
                            <p className="text-sm font-semibold">{userName(m.user)}</p>
                            <p className="text-xs text-[--text-3]">{m.role ?? "member"}</p>
                          </div>
                        </div>
                        <span className="rounded-lg bg-[--surface-2] px-2.5 py-1 text-xs font-medium">{workFieldLabel(typeof m.user === "string" ? undefined : m.user?.workField)}</span>
                      </div>
                    );
                  })
                )}
              </div>
            </div>
          </div>
        </>
      )}

      {/* ─ Task Detail Panel ──────────────────────────────────────────────── */}
      {selectedProject && (
        <ProjectPanel
          project={selectedProject}
          progress={projectProgress[getId(selectedProject)]}
          tasks={tasks.filter((t) => getId(t.projectId) === getId(selectedProject))}
          users={users}
          coverClass={PROJECT_COVERS[projects.findIndex((p) => getId(p) === getId(selectedProject)) % PROJECT_COVERS.length]}
          isManager={isManager}
          onArchive={() => {
            const pid = getId(selectedProject);
            void managerApi.setProjectActivation(token, pid, !!selectedProject.isArchived)
              .then(() => { setMessage("وضعیت پروژه تغییر کرد."); void loadData(); setSelectedProject(null); })
              .catch(() => setError("تغییر وضعیت ناموفق بود"));
          }}
          onDelete={() => { void deleteProject(getId(selectedProject)); setSelectedProject(null); }}
          onManageMembers={() => { void loadProjectMembers(selectedProject); setSelectedProject(null); }}
          onClose={() => setSelectedProject(null)}
        />
      )}

      {selectedTask && (
        <TaskPanel
          task={selectedTask}
          users={users}
          priority={taskPriorities[getId(selectedTask)]}
          canEditAssignments={isManager}
          canClaim={isSpecialist && isUnassignedTask(selectedTask)}
          onPriorityChange={(p) => {
            setTaskPriorities((prev) => ({ ...prev, [getId(selectedTask)]: p }));
            void updateTask(getId(selectedTask), { taskComment: `priority:${p}` });
          }}
          onClaim={() => void claimTask(getId(selectedTask))}
          onStatusChange={(s) => void moveTask(getId(selectedTask), s)}
          onDescriptionChange={(d) => void updateTask(getId(selectedTask), { description: d })}
          onAssign={(userId) => {
            const current = (selectedTask.assignedTo ?? []).map((u) => typeof u === "string" ? u : getId(u));
            if (!current.includes(userId)) void updateTask(getId(selectedTask), { assignedTo: [...current, userId] });
          }}
          onUnassign={(userId) => {
            const current = (selectedTask.assignedTo ?? []).map((u) => typeof u === "string" ? u : getId(u));
            void updateTask(getId(selectedTask), { assignedTo: current.filter((id) => id !== userId) });
          }}
          onDelete={() => void deleteTask(getId(selectedTask))}
          onClose={() => setSelectedTask(null)}
        />
      )}
    </div>
  );
}

// ─── Project Detail Panel ─────────────────────────────────────────────────────
function ProjectPanel({
  project, progress, tasks, users, coverClass, isManager,
  onArchive, onDelete, onManageMembers, onClose,
}: {
  project: Project;
  progress?: ProjectProgress;
  tasks: Task[];
  users: User[];
  coverClass: string;
  isManager: boolean;
  onArchive: () => void;
  onDelete: () => void;
  onManageMembers: () => void;
  onClose: () => void;
}) {
  const doneTasks = tasks.filter((t) => t.status === "done").length;
  const inProgressTasks = tasks.filter((t) => t.status === "in_progress").length;
  const todoTasks = tasks.filter((t) => (t.status ?? "todo") === "todo").length;
  const progressPct = progress?.progressPercentage ?? (tasks.length ? Math.round((doneTasks / tasks.length) * 100) : 0);

  return (
    <>
      <div className="fixed inset-0 z-50 bg-black/30 backdrop-blur-sm dark:bg-black/50" onClick={onClose} />
      <div className="fixed inset-y-0 left-0 z-50 flex w-full max-w-sm flex-col bg-[--surface] shadow-2xl">

        {/* Cover header */}
        <div className={`relative bg-gradient-to-l ${coverClass} px-5 pt-5 pb-10`}>
          <div className="absolute inset-0 opacity-10" style={{ backgroundImage: "radial-gradient(circle at 20% 50%, white 1px, transparent 1px)", backgroundSize: "20px 20px" }} />
          <div className="relative flex items-start justify-between">
            <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-white/20 text-white border border-white/30 backdrop-blur-sm">
              <FolderKanban size={20} />
            </div>
            <button className="flex h-8 w-8 items-center justify-center rounded-lg bg-white/20 text-white hover:bg-white/30 transition" onClick={onClose} type="button">
              <X size={16} />
            </button>
          </div>
          <div className="relative mt-3">
            <h2 className="text-lg font-bold text-white leading-snug">{project.title}</h2>
            <div className="mt-1 flex items-center gap-2">
              <span className="rounded-full bg-white/20 border border-white/20 px-2.5 py-0.5 text-xs font-semibold text-white">{statusLabel(project.status)}</span>
              {project.isArchived && <span className="rounded-full bg-black/20 border border-white/10 px-2.5 py-0.5 text-xs font-semibold text-white/80">آرشیو شده</span>}
            </div>
          </div>
        </div>

        {/* Progress bar floating over cover */}
        <div className="mx-5 -mt-5 relative z-10 rounded-xl bg-[--surface] border border-[--border] shadow-lg p-4">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-semibold text-[--text-2]">پیشرفت کلی</span>
            <span className="text-sm font-bold text-[#1f7a8c]">{progressPct}%</span>
          </div>
          <div className="h-2 overflow-hidden rounded-full bg-[--border]">
            <div className="h-full rounded-full bg-gradient-to-l from-[#1f7a8c] to-[#2a9db2] transition-all duration-700" style={{ width: `${progressPct}%` }} />
          </div>
          <div className="mt-3 grid grid-cols-3 gap-2 text-center">
            {[
              { l: "در انتظار", v: todoTasks, c: "text-slate-500" },
              { l: "جاری", v: inProgressTasks, c: "text-[#1f7a8c]" },
              { l: "تمام", v: doneTasks, c: "text-emerald-500" },
            ].map((s) => (
              <div key={s.l} className="rounded-lg bg-[--surface-2] py-2">
                <p className={`text-base font-bold ${s.c}`}>{s.v}</p>
                <p className="text-[10px] text-[--text-3]">{s.l}</p>
              </div>
            ))}
          </div>
        </div>

        <div className="flex-1 overflow-y-auto p-5 space-y-5">

          {/* Description */}
          {project.description && (
            <div>
              <p className="mb-1.5 text-xs font-semibold text-[--text-3]">توضیحات</p>
              <p className="rounded-xl bg-[--surface-2] px-3 py-2.5 text-sm text-[--text] leading-relaxed">{project.description}</p>
            </div>
          )}

          {/* Members */}
          <div>
            <div className="mb-2 flex items-center justify-between">
              <p className="text-xs font-semibold text-[--text-3]">افراد ({[project.owner, project.supervisorId, project.assigneeId].filter(Boolean).length})</p>
              {isManager && (
                <button className="text-xs font-semibold text-[#1f7a8c] hover:underline" onClick={onManageMembers} type="button">
                  مشاهده افراد
                </button>
              )}
            </div>
            {[project.owner, project.supervisorId, project.assigneeId].filter(Boolean).length > 0 ? (
              <div className="space-y-2">
                {([project.owner, project.supervisorId, project.assigneeId].filter(Boolean) as Array<string | User>).slice(0, 5).map((m, i) => {
                  const id = typeof m === "string" ? m : getId(m);
                  const fullUser = users.find((u) => getId(u) === id) ?? (typeof m === "object" ? m : undefined);
                  return (
                    <div key={i} className="flex items-center gap-3 rounded-xl bg-[--surface-2] px-3 py-2.5">
                      <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-[#1f7a8c] to-[#165e6d] text-xs font-bold text-white">
                        {initials(fullUser)}
                      </div>
                      <div>
                        <p className="text-sm font-semibold">{userName(fullUser)}</p>
                        <p className="text-xs text-[--text-3]">{fullUser?.roles || "specialist"}</p>
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="rounded-xl border border-dashed border-[--border] p-3 text-center text-xs text-[--text-3]">فردی ندارد</div>
            )}
          </div>

          {/* Tasks list */}
          <div>
            <p className="mb-2 text-xs font-semibold text-[--text-3]">گزارش‌ها ({tasks.length})</p>
            {tasks.length === 0 ? (
              <div className="rounded-xl border border-dashed border-[--border] p-3 text-center text-xs text-[--text-3]">گزارشی ندارد</div>
            ) : (
              <div className="space-y-1.5">
                {tasks.map((t) => (
                  <div key={getId(t)} className="flex items-center gap-3 rounded-xl border border-[--border] bg-[--surface-2] px-3 py-2.5">
                    {t.status === "done"
                      ? <CheckCircle2 size={14} className="shrink-0 text-emerald-500" />
                      : t.status === "in_progress"
                      ? <div className="h-3 w-3 shrink-0 rounded-full border-2 border-[#1f7a8c] border-t-transparent animate-spin" />
                      : <CircleDashed size={14} className="shrink-0 text-[--text-3]" />
                    }
                    <p className="flex-1 truncate text-sm font-medium">{t.title}</p>
                    <span className={`shrink-0 rounded-md px-1.5 py-0.5 text-[10px] font-bold ${COLUMNS.find((c) => c.status === t.status)?.badge ?? "bg-slate-100 text-slate-600"}`}>
                      {statusLabel(t.status)}
                    </span>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Meta */}
          <div className="rounded-xl bg-[--surface-2] p-4 space-y-3 text-xs">
            <div className="flex items-center justify-between">
              <span className="text-[--text-3]">مالک پروژه</span>
              <span className="font-semibold">{userName(project.owner)}</span>
            </div>
            <div className="flex items-center justify-between">
              <span className="text-[--text-3]">وضعیت</span>
              <span className="font-semibold">{statusLabel(project.status)}</span>
            </div>
            <div className="flex items-center justify-between">
              <span className="text-[--text-3]">کل گزارش‌ها</span>
              <span className="font-semibold">{progress?.totalTasks ?? tasks.length}</span>
            </div>
          </div>
        </div>

        {/* Footer — manager only */}
        {isManager && (
          <div className="border-t border-[--border] p-4 space-y-2">
            <button
              className={`flex h-10 w-full items-center justify-center gap-2 rounded-xl border text-sm font-semibold transition hover:opacity-80 ${project.isArchived ? "border-emerald-200 text-emerald-600 bg-emerald-50 dark:bg-emerald-950/30" : "border-slate-200 text-slate-600 bg-slate-50 dark:bg-slate-800 dark:border-slate-700"}`}
              onClick={onArchive} type="button"
            >
              {project.isArchived ? "فعال‌سازی پروژه" : "آرشیو پروژه"}
            </button>
            <button
              className="flex h-9 w-full items-center justify-center rounded-xl border border-red-200 text-sm font-semibold text-red-600 transition hover:bg-red-50 dark:border-red-900 dark:hover:bg-red-950/40"
              onClick={onDelete} type="button"
            >
              حذف پروژه
            </button>
          </div>
        )}
      </div>
    </>
  );
}

// ─── Task Detail Panel ────────────────────────────────────────────────────────
function TaskPanel({
  task, users, priority, canEditAssignments, canClaim, onPriorityChange, onStatusChange, onDescriptionChange, onAssign, onUnassign, onClaim, onDelete, onClose,
}: {
  task: Task; users: User[]; priority?: Priority;
  canEditAssignments: boolean;
  canClaim: boolean;
  onPriorityChange: (p: Priority) => void;
  onStatusChange: (s: string) => void;
  onDescriptionChange: (d: string) => void;
  onAssign: (userId: string) => void;
  onUnassign: (userId: string) => void;
  onClaim: () => void;
  onDelete: () => void;
  onClose: () => void;
}) {
  const overlayRef = useRef<HTMLDivElement>(null);
  const [desc, setDesc] = useState(task.description ?? "");
  const [descEditing, setDescEditing] = useState(false);
  const [assignSelect, setAssignSelect] = useState("");

  const assignedIds = (task.assignedTo ?? []).map((u) => typeof u === "string" ? u : getId(u));
  const unassignedUsers = users.filter((u) => !assignedIds.includes(getId(u)));

  return (
    <>
      {/* Backdrop */}
      <div
        ref={overlayRef}
        className="fixed inset-0 z-50 bg-black/30 backdrop-blur-sm dark:bg-black/50"
        onClick={onClose}
      />
      {/* Panel */}
      <div className="fixed inset-y-0 left-0 z-50 flex w-full max-w-sm flex-col bg-[--surface] shadow-2xl">
        {/* Panel header */}
        <div className="flex items-center justify-between border-b border-[--border] px-5 py-4">
          <h3 className="font-bold text-[--text]">جزئیات گزارش</h3>
          <button className="flex h-8 w-8 items-center justify-center rounded-lg text-[--text-3] transition hover:bg-[--surface-2] hover:text-[--text]" onClick={onClose} type="button">
            <X size={18} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-5 space-y-5">
          {/* Title */}
          <div>
            <div className="flex items-start gap-2">
              {task.status === "done"
                ? <CheckCircle2 size={18} className="mt-0.5 shrink-0 text-emerald-500" />
                : task.status === "in_progress"
                ? <div className="mt-1 h-4 w-4 shrink-0 rounded-full border-2 border-[#1f7a8c] border-t-transparent animate-spin" />
                : <CircleDashed size={18} className="mt-0.5 shrink-0 text-[--text-3]" />
              }
              <h2 className="text-lg font-bold leading-snug text-[--text]">{task.title}</h2>
            </div>
          </div>

          {/* Description */}
          <div>
            <div className="mb-2 flex items-center justify-between">
              <p className="text-xs font-semibold text-[--text-3]">توضیحات</p>
              {!descEditing && (
                <button className="text-xs font-semibold text-[#1f7a8c] hover:underline" onClick={() => setDescEditing(true)} type="button">
                  {desc ? "ویرایش" : "افزودن"}
                </button>
              )}
            </div>
            {descEditing ? (
              <div className="space-y-2">
                <textarea
                  autoFocus
                  className="h-24 w-full resize-none rounded-lg border border-[--border] bg-[--surface] px-3 py-2 text-sm text-[--text] outline-none transition placeholder:text-[--text-3] focus:border-[#1f7a8c] focus:ring-2 focus:ring-[#1f7a8c]/15"
                  value={desc}
                  onChange={(e) => setDesc(e.target.value)}
                  placeholder="توضیحات گزارش را وارد کنید…"
                />
                <div className="flex gap-2">
                  <button
                    className="flex h-8 flex-1 items-center justify-center rounded-lg bg-[#1f7a8c] text-xs font-semibold text-white transition hover:bg-[#196b7b]"
                    onClick={() => { onDescriptionChange(desc); setDescEditing(false); }}
                    type="button"
                  >ذخیره</button>
                  <button
                    className="flex h-8 w-8 items-center justify-center rounded-lg border border-[--border] text-[--text-2] transition hover:bg-[--surface-2]"
                    onClick={() => { setDesc(task.description ?? ""); setDescEditing(false); }}
                    type="button"
                  ><X size={13} /></button>
                </div>
              </div>
            ) : desc ? (
              <p className="rounded-xl bg-[--surface-2] px-3 py-2.5 text-sm text-[--text] leading-relaxed">{desc}</p>
            ) : (
              <button
                className="w-full rounded-xl border border-dashed border-[--border] p-3 text-center text-xs text-[--text-3] transition hover:border-[#1f7a8c]/40 hover:text-[#1f7a8c]"
                onClick={() => setDescEditing(true)}
                type="button"
              >+ افزودن توضیحات</button>
            )}
          </div>

          {/* Status */}
          <div>
            <p className="mb-2 text-xs font-semibold text-[--text-3]">وضعیت</p>
            <div className="flex gap-2 flex-wrap">
              {COLUMNS.map((col) => (
                <button
                  key={col.status}
                  className={`flex items-center gap-1.5 rounded-lg border px-3 py-1.5 text-xs font-semibold transition-all ${task.status === col.status ? `${col.badge} border-transparent` : "border-[--border] text-[--text-2] hover:bg-[--surface-2]"}`}
                  onClick={() => onStatusChange(col.status)}
                  type="button"
                >
                  <span className={`h-2 w-2 rounded-full ${col.dot}`} />
                  {col.title}
                </button>
              ))}
            </div>
          </div>

          {/* Priority */}
          <div>
            <p className="mb-2 text-xs font-semibold text-[--text-3]">اولویت</p>
            <div className="flex flex-wrap gap-2">
              {(Object.keys(PRIORITY) as Priority[]).map((p) => (
                <button
                  key={p}
                  className={`flex items-center gap-1.5 rounded-lg border px-3 py-1.5 text-xs font-semibold transition-all ${priority === p ? PRIORITY[p].color : "border-[--border] text-[--text-2] hover:bg-[--surface-2]"}`}
                  onClick={() => onPriorityChange(p)}
                  type="button"
                >
                  <Flag size={11} />
                  {PRIORITY[p].label}
                </button>
              ))}
            </div>
          </div>

          {/* Assignees */}
          <div>
            <p className="mb-2 text-xs font-semibold text-[--text-3]">مسئولان</p>
            {canClaim && (
              <button
                className="mb-2 flex h-9 w-full items-center justify-center gap-1.5 rounded-xl bg-[#1f7a8c] text-xs font-semibold text-white transition hover:bg-[#196b7b]"
                onClick={onClaim}
                type="button"
              >
                <UserPlus size={13} /> برداشتن این گزارش برای من
              </button>
            )}
            {task.assignedTo && task.assignedTo.length > 0 ? (
              <div className="space-y-2">
                {task.assignedTo.map((u, i) => {
                  const id = typeof u === "string" ? u : getId(u);
                  const fullUser = users.find((usr) => getId(usr) === id) ?? (typeof u === "object" ? u : undefined);
                  return (
                    <div key={i} className="flex items-center justify-between gap-3 rounded-xl bg-[--surface-2] px-3 py-2.5">
                      <div className="flex items-center gap-3">
                        <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-[#1f7a8c] to-[#165e6d] text-xs font-bold text-white">
                          {initials(fullUser)}
                        </div>
                        <div>
                          <p className="text-sm font-semibold">{userName(fullUser)}</p>
                          <p className="text-xs text-[--text-3]">{fullUser?.roles || "specialist"}</p>
                        </div>
                      </div>
                      {canEditAssignments && (
                        <button
                          className="flex h-6 w-6 items-center justify-center rounded-md text-[--text-3] transition hover:bg-red-50 hover:text-red-500"
                          onClick={() => onUnassign(id)}
                          type="button"
                          title="حذف مسئول"
                        ><X size={12} /></button>
                      )}
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="rounded-xl border border-dashed border-[--border] p-3 text-center text-xs text-[--text-3]">بدون مسئول</div>
            )}
            {canEditAssignments && unassignedUsers.length > 0 && (
              <div className="mt-2 flex gap-2">
                <select
                  className="h-9 flex-1 rounded-lg border border-[--border] bg-[--surface] px-2 text-sm text-[--text] outline-none transition focus:border-[#1f7a8c] focus:ring-2 focus:ring-[#1f7a8c]/15"
                  value={assignSelect}
                  onChange={(e) => setAssignSelect(e.target.value)}
                >
                  <option value="">اساین به…</option>
                  {unassignedUsers.map((u) => <option key={getId(u)} value={getId(u)}>{userName(u)}</option>)}
                </select>
                <button
                  className="flex h-9 items-center gap-1 rounded-lg bg-[#1f7a8c] px-3 text-xs font-semibold text-white transition hover:bg-[#196b7b] disabled:opacity-40"
                  disabled={!assignSelect}
                  onClick={() => { if (assignSelect) { onAssign(assignSelect); setAssignSelect(""); } }}
                  type="button"
                ><UserPlus size={13} />اساین</button>
              </div>
            )}
          </div>

          {/* Meta */}
          <div className="rounded-xl bg-[--surface-2] p-4 space-y-3">
            <div className="flex items-center justify-between text-xs">
              <span className="text-[--text-3]">تاریخ ایجاد</span>
              <span className="font-medium text-[--text]">{formatDate(task.createdAt)}</span>
            </div>
            {task.dueDate && (
              <div className="flex items-center justify-between text-xs">
                <span className="text-[--text-3]">مهلت</span>
                <span className="font-medium text-[--text]">{formatDate(task.dueDate)}</span>
              </div>
            )}
            <div className="flex items-center justify-between text-xs">
              <span className="text-[--text-3]">وضعیت</span>
              <span className={`rounded-full px-2.5 py-0.5 text-[11px] font-bold ${COLUMNS.find((c) => c.status === task.status)?.badge ?? ""}`}>
                {statusLabel(task.status)}
              </span>
            </div>
            {priority && (
              <div className="flex items-center justify-between text-xs">
                <span className="text-[--text-3]">اولویت</span>
                <span className={`rounded-md px-2 py-0.5 text-[11px] font-bold ${PRIORITY[priority].color}`}>{PRIORITY[priority].label}</span>
              </div>
            )}
          </div>
        </div>

        {/* Panel footer */}
        <div className="border-t border-[--border] p-4 space-y-2">
          {task.status !== "done" ? (
            <button
              className="flex h-10 w-full items-center justify-center gap-2 rounded-xl bg-[#1f7a8c] text-sm font-semibold text-white transition hover:bg-[#196b7b] active:scale-[0.98]"
              onClick={() => { onStatusChange(nextStatus(task.status)); onClose(); }}
              type="button"
            >
              {task.status === "todo" ? "شروع کار" : "تکمیل کردن"}
              <ChevronLeft size={15} />
            </button>
          ) : (
            <div className="flex items-center justify-center gap-2 rounded-xl bg-emerald-50 dark:bg-emerald-950/40 py-2.5 text-sm font-semibold text-emerald-600 dark:text-emerald-400">
              <CheckCircle2 size={16} />تکمیل شده
            </div>
          )}
          <button
            className="flex h-9 w-full items-center justify-center rounded-xl border border-red-200 text-sm font-semibold text-red-600 transition hover:bg-red-50 dark:border-red-900 dark:hover:bg-red-950/40"
            onClick={onDelete} type="button"
          >
            حذف گزارش
          </button>
        </div>
      </div>
    </>
  );
}

// ─── Shared components ────────────────────────────────────────────────────────
function Field({ label, name, id, required, type = "text", value, onChange, placeholder }: {
  label: string; name: string; id?: string; required?: boolean; type?: string;
  value?: string; onChange?: (v: string) => void; placeholder?: string;
}) {
  return (
    <label className="block">
      {label && <span className="mb-1.5 block text-xs font-semibold text-[--text-2]">{label}</span>}
      <input
        id={id}
        className="h-10 w-full rounded-lg border border-[--border] bg-[--surface] px-3 text-sm text-[--text] outline-none transition placeholder:text-[--text-3] focus:border-[#1f7a8c] focus:ring-2 focus:ring-[#1f7a8c]/15"
        name={name} onChange={(e) => onChange?.(e.target.value)} required={required} type={type} value={value} placeholder={placeholder}
      />
    </label>
  );
}

function Select({ label, value, onChange, options, placeholder }: {
  label: string; value: string; onChange: (v: string) => void; options: Array<[string, string]>; placeholder?: string;
}) {
  return (
    <label className="block">
      {label && <span className="mb-1.5 block text-xs font-semibold text-[--text-2]">{label}</span>}
      <select className="h-10 w-full rounded-lg border border-[--border] bg-[--surface] px-3 text-sm text-[--text] outline-none transition focus:border-[#1f7a8c] focus:ring-2 focus:ring-[#1f7a8c]/15" onChange={(e) => onChange(e.target.value)} value={value}>
        <option value="">{placeholder ?? "انتخاب نشده"}</option>
        {options.map(([id, l]) => <option key={id} value={id}>{l}</option>)}
      </select>
    </label>
  );
}

function SideItem({ active, icon: Icon, label, meta, collapsed, onClick }: {
  active?: boolean; icon: typeof LayoutDashboard; label: string; meta?: number; collapsed?: boolean; onClick?: () => void;
}) {
  return (
    <button
      className={`flex w-full items-center rounded-lg px-2.5 py-2 text-sm font-medium transition-all ${active ? "bg-[#e8f4f7] text-[#1f7a8c] dark:bg-[#0f3040] dark:text-[#4fc3d5]" : "text-[--text-2] hover:bg-[--surface-2] hover:text-[--text]"} ${collapsed ? "justify-center" : "justify-between"}`}
      onClick={onClick} type="button"
      title={collapsed ? label : undefined}
    >
      <span className="flex items-center gap-2.5"><Icon size={16} />{!collapsed && label}</span>
      {!collapsed && meta !== undefined && (
        <span className={`rounded-md px-2 py-0.5 text-[11px] font-semibold ${active ? "bg-[#1f7a8c]/12 text-[#1f7a8c] dark:text-[#4fc3d5]" : "bg-[--surface-2] text-[--text-3]"}`}>{meta}</span>
      )}
    </button>
  );
}

function FilterChip({ active, label, count, onClick }: { active?: boolean; label: string; count: number; onClick?: () => void }) {
  return (
    <button
      className={`flex h-7 shrink-0 items-center gap-1.5 rounded-full border px-3 text-xs font-semibold transition-all ${active ? "border-[#1f7a8c]/30 bg-[#e8f4f7] text-[#1f7a8c] dark:bg-[#0f3040] dark:text-[#4fc3d5] dark:border-[#1f5060]" : "border-[--border] bg-[--surface] text-[--text-2] hover:bg-[--surface-2]"}`}
      onClick={onClick} type="button"
    >
      <span className="max-w-24 truncate">{label}</span>
      <span className={`rounded-full px-1.5 text-[10px] ${active ? "bg-[#1f7a8c]/10" : "bg-[--surface-2]"}`}>{count}</span>
    </button>
  );
}

function AssigneeStack({ users, fallback }: { users?: Array<string | User>; fallback?: string | User }) {
  const visible = users?.length ? users.slice(0, 3) : fallback ? [fallback] : [];
  if (!visible.length) return <span className="text-xs text-[--text-3]">بدون مسئول</span>;
  return (
    <div className="flex items-center gap-1.5">
      <div className="flex -space-x-1.5 space-x-reverse">
        {visible.map((u, i) => (
          <span key={`${userName(u)}-${i}`} className="flex h-6 w-6 items-center justify-center rounded-full border-2 border-[--surface] bg-gradient-to-br from-[#e8f4f7] to-[#d0edf3] dark:from-[#0f3040] dark:to-[#1f5060] text-[9px] font-bold text-[#1f7a8c] dark:text-[#4fc3d5] shadow-sm" title={userName(u)}>
            {initials(u)}
          </span>
        ))}
      </div>
      {visible.length === 1 && <span className="text-xs font-medium text-[--text-2]">{userName(visible[0])}</span>}
    </div>
  );
}

function Toast({ message, type, onClose }: { message: string; type: "success" | "error"; onClose: () => void }) {
  if (!message) return null;
  const ok = type === "success";
  return (
    <div className="fixed left-4 top-4 z-[60] w-[calc(100%-2rem)] max-w-sm" role="status" aria-live="polite">
      <div className={`flex items-center gap-3 rounded-xl border bg-[--surface] px-4 py-3 shadow-xl ${ok ? "border-emerald-200 dark:border-emerald-900" : "border-red-200 dark:border-red-900"}`}>
        <div className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-lg ${ok ? "bg-emerald-50 text-emerald-600 dark:bg-emerald-950/50 dark:text-emerald-400" : "bg-red-50 text-red-600 dark:bg-red-950/50 dark:text-red-400"}`}>
          {ok ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />}
        </div>
        <p className="flex-1 text-sm font-medium text-[--text]">{message}</p>
        <button className="flex h-7 w-7 items-center justify-center rounded-lg text-[--text-3] transition hover:bg-[--surface-2]" onClick={onClose} type="button"><X size={13} /></button>
      </div>
    </div>
  );
}

function SkeletonCard() {
  return (
    <div className="overflow-hidden rounded-xl border border-[--border] bg-[--surface]">
      <div className="skeleton h-16 w-full" />
      <div className="p-4 space-y-3">
        <div className="skeleton h-4 w-2/3 rounded-lg" />
        <div className="skeleton h-3 w-1/2 rounded-lg" />
        <div className="mt-3 flex items-center justify-between border-t border-[--border] pt-3">
          <div className="skeleton h-6 w-16 rounded-full" />
          <div className="skeleton h-5 w-12 rounded-lg" />
        </div>
      </div>
    </div>
  );
}

function EmptyState({ title, text, action }: { title: string; text: string; action?: { label: string; onClick: () => void } }) {
  return (
    <div className="col-span-full flex flex-col items-center rounded-xl border-2 border-dashed border-[--border] p-10 text-center">
      <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-[--surface-2] text-[--text-3]"><FolderKanban size={26} /></div>
      <p className="mt-3 font-semibold text-[--text]">{title}</p>
      <p className="mt-1 text-sm text-[--text-3]">{text}</p>
      {action && (
        <button className="mt-4 flex h-9 items-center gap-2 rounded-lg bg-[#1f7a8c] px-4 text-sm font-semibold text-white shadow-sm transition hover:bg-[#196b7b]" onClick={action.onClick} type="button">
          <Plus size={15} />{action.label}
        </button>
      )}
    </div>
  );
}
