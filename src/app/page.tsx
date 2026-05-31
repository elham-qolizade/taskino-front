"use client";

import {
  CheckCircle2,
  CircleDashed,
  ClipboardList,
  FolderKanban,
  Loader2,
  LogIn,
  LogOut,
  Plus,
  RefreshCw,
  ShieldCheck,
  Sparkles,
  UserRound,
  UsersRound,
  X,
} from "lucide-react";
import { FormEvent, useEffect, useMemo, useState } from "react";

type User = {
  _id?: string;
  id?: string;
  firstName?: string;
  lastName?: string;
  email?: string;
  mobile?: string;
  roles?: string;
  isActive?: boolean;
};

type Task = {
  _id?: string;
  id?: string;
  title: string;
  createdBy?: string | User;
  assignedTo?: Array<string | User>;
  status?: string;
};

type Project = {
  _id?: string;
  id?: string;
  title: string;
  description?: string;
  owner?: string | User;
  members?: Array<string | User>;
  tasks?: Array<string | Task>;
  status?: string;
  isArchived?: boolean;
};

type AuthResponse = {
  user: User;
  accessToken: string;
};

const apiUrl = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:3000/api";

const taskStatuses = ["todo", "in_progress", "done"];
function getId(item?: { _id?: string; id?: string } | string) {
  if (!item) return "";
  return typeof item === "string" ? item : item._id ?? item.id ?? "";
}

function userName(user?: User | string) {
  if (!user) return "نامشخص";
  if (typeof user === "string") return user;
  return [user.firstName, user.lastName].filter(Boolean).join(" ") || user.mobile || user.email || getId(user);
}

function statusLabel(status?: string) {
  const labels: Record<string, string> = {
    todo: "در انتظار",
    in_progress: "در حال انجام",
    done: "انجام شده",
    pending: "در انتظار",
    completed: "تکمیل شده",
  };

  return labels[status ?? ""] ?? status ?? "نامشخص";
}

async function request<T>(path: string, options: RequestInit = {}, token?: string): Promise<T> {
  let res: Response;

  try {
    res = await fetch(`${apiUrl}${path}`, {
      ...options,
      headers: {
        "Content-Type": "application/json",
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        ...options.headers,
      },
    });
  } catch {
    throw new Error(`اتصال به بک‌اند برقرار نشد. مطمئن شو API روی ${apiUrl} اجرا شده است.`);
  }

  const body = await res.text();
  const data = body ? JSON.parse(body) : null;

  if (!res.ok) {
    const message = data?.message;
    throw new Error(Array.isArray(message) ? message.join("، ") : message ?? "درخواست ناموفق بود");
  }

  return data as T;
}

function normalizeList<T>(value: T[] | { data?: T[]; items?: T[]; docs?: T[] }) {
  if (Array.isArray(value)) return value;
  return value.data ?? value.items ?? value.docs ?? [];
}

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
  const [projectTitle, setProjectTitle] = useState("");
  const [projectDescription, setProjectDescription] = useState("");
  const [projectMember, setProjectMember] = useState("");

  useEffect(() => {
    queueMicrotask(() => {
      const savedToken = localStorage.getItem("taskino-token") ?? "";
      const savedUser = localStorage.getItem("taskino-user");
      if (savedToken) setToken(savedToken);
      if (savedUser) setCurrentUser(JSON.parse(savedUser) as User);
    });
  }, []);

  const myId = getId(currentUser ?? undefined);

  const stats = useMemo(
    () => [
      { label: "پروژه‌ها", value: projects.length, icon: FolderKanban },
      { label: "تسک‌ها", value: tasks.length, icon: ClipboardList },
      { label: "کاربرها", value: users.length, icon: UsersRound },
      { label: "تکمیل شده", value: tasks.filter((task) => task.status === "done").length, icon: CheckCircle2 },
    ],
    [projects.length, tasks, users.length],
  );

  useEffect(() => {
    if (!error && !message) return;

    const timeout = window.setTimeout(() => {
      setError("");
      setMessage("");
    }, 3500);

    return () => window.clearTimeout(timeout);
  }, [error, message]);

  async function loadData(authToken = token) {
    if (!authToken) return;
    setLoadingData(true);
    setError("");

    try {
      const [usersRes, tasksRes, projectsRes] = await Promise.all([
        request<User[] | { data?: User[]; items?: User[]; docs?: User[] }>("/users?page=1&limit=50", {}, authToken),
        request<Task[] | { data?: Task[]; items?: Task[]; docs?: Task[] }>("/tasks?page=1&limit=50", {}, authToken),
        request<Project[] | { data?: Project[]; items?: Project[]; docs?: Project[] }>("/projects", {}, authToken),
      ]);

      setUsers(normalizeList(usersRes));
      setTasks(normalizeList(tasksRes));
      setProjects(normalizeList(projectsRes));
    } catch (err) {
      setError(err instanceof Error ? err.message : "دریافت اطلاعات ناموفق بود");
    } finally {
      setLoadingData(false);
    }
  }

  useEffect(() => {
    queueMicrotask(() => {
      void loadData();
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token]);

  async function handleAuth(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setAuthLoading(true);
    setError("");
    setMessage("");

    const form = new FormData(event.currentTarget);
    const payload =
      authMode === "login"
        ? {
            mobile: String(form.get("mobile") ?? ""),
            password: String(form.get("password") ?? ""),
          }
        : {
            firstName: String(form.get("firstName") ?? ""),
            lastName: String(form.get("lastName") ?? ""),
            email: String(form.get("email") ?? ""),
            mobile: String(form.get("mobile") ?? ""),
            password: String(form.get("password") ?? ""),
            roles: "specialist",
          };

    try {
      const data = await request<AuthResponse>(`/auth/${authMode}`, {
        method: "POST",
        body: JSON.stringify(payload),
      });

      setToken(data.accessToken);
      setCurrentUser(data.user);
      localStorage.setItem("taskino-token", data.accessToken);
      localStorage.setItem("taskino-user", JSON.stringify(data.user));
      setMessage(authMode === "login" ? "با موفقیت وارد شدی." : "ثبت‌نام انجام شد.");
      await loadData(data.accessToken);
    } catch (err) {
      setError(err instanceof Error ? err.message : "ورود ناموفق بود");
    } finally {
      setAuthLoading(false);
    }
  }

  async function createTask(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!myId || !taskTitle.trim()) return;
    setError("");

    try {
      await request<Task>(
        "/tasks",
        {
          method: "POST",
          body: JSON.stringify({
            title: taskTitle.trim(),
            createdBy: myId,
            assignedTo: taskAssignee ? [taskAssignee] : [],
            status: "todo",
          }),
        },
        token,
      );
      setTaskTitle("");
      setTaskAssignee("");
      setMessage("تسک جدید ساخته شد.");
      await loadData();
    } catch (err) {
      setError(err instanceof Error ? err.message : "ساخت تسک ناموفق بود");
    }
  }

  async function createProject(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!myId || !projectTitle.trim()) return;
    setError("");

    try {
      await request<Project>(
        "/projects",
        {
          method: "POST",
          body: JSON.stringify({
            title: projectTitle.trim(),
            description: projectDescription.trim() || undefined,
            owner: myId,
            members: projectMember ? [projectMember] : [],
            status: "pending",
          }),
        },
        token,
      );
      setProjectTitle("");
      setProjectDescription("");
      setProjectMember("");
      setMessage("پروژه جدید ساخته شد.");
      await loadData();
    } catch (err) {
      setError(err instanceof Error ? err.message : "ساخت پروژه ناموفق بود");
    }
  }

  async function updateTaskStatus(task: Task) {
    const currentIndex = taskStatuses.indexOf(task.status ?? "todo");
    const nextStatus = taskStatuses[(currentIndex + 1) % taskStatuses.length];
    setError("");

    try {
      await request(`/tasks/${getId(task)}/status`, {
        method: "PATCH",
        body: JSON.stringify({ status: nextStatus }),
      }, token);
      await loadData();
    } catch (err) {
      setError(err instanceof Error ? err.message : "تغییر وضعیت ناموفق بود");
    }
  }

  function logout() {
    setToken("");
    setCurrentUser(null);
    setUsers([]);
    setTasks([]);
    setProjects([]);
    localStorage.removeItem("taskino-token");
    localStorage.removeItem("taskino-user");
  }

  if (!token) {
    return (
      <main className="min-h-screen px-4 py-8 text-[#1d2733] sm:px-6 lg:px-10">
        <Toast message={error || message} type={error ? "error" : "success"} onClose={() => (error ? setError("") : setMessage(""))} />
        <section className="mx-auto grid min-h-[calc(100vh-4rem)] w-full max-w-6xl items-center gap-8 lg:grid-cols-[1.1fr_0.9fr]">
          <div className="space-y-8">
            <div className="flex items-center gap-4">
              <div className="flex h-12 w-12 items-center justify-center rounded-lg bg-[#1f7a8c] text-white shadow-soft">
                <Sparkles size={22} />
              </div>
              <div>
                <p className="text-sm font-semibold text-[#5f6f7d]">Taskino</p>
                <h1 className="mt-2 text-3xl font-bold leading-tight text-[#15202b] sm:text-5xl">مدیریت کارهای تیمی</h1>
              </div>
            </div>
          </div>

          <div className="rounded-lg border border-[#d8e0e7] bg-white/95 p-5 shadow-soft ring-1 ring-white/70 backdrop-blur">
            <div className="mb-5 grid grid-cols-2 rounded-lg bg-[#edf2f6] p-1">
              <button
                className={`rounded-md px-4 py-2 text-sm font-semibold transition ${authMode === "login" ? "bg-white text-[#1f7a8c] shadow-sm" : "text-[#5d6b78]"}`}
                onClick={() => setAuthMode("login")}
                type="button"
              >
                ورود
              </button>
              <button
                className={`rounded-md px-4 py-2 text-sm font-semibold transition ${authMode === "register" ? "bg-white text-[#1f7a8c] shadow-sm" : "text-[#5d6b78]"}`}
                onClick={() => setAuthMode("register")}
                type="button"
              >
                ثبت‌نام
              </button>
            </div>

            <form className="space-y-4" onSubmit={handleAuth}>
              {authMode === "register" && (
                <div className="grid gap-3 sm:grid-cols-2">
                  <Field label="نام" name="firstName" required />
                  <Field label="نام خانوادگی" name="lastName" required />
                  <Field label="ایمیل" name="email" required type="email" />
                </div>
              )}
              <Field label="موبایل" name="mobile" required />
              <Field label="رمز عبور" name="password" required type="password" />
              <button
                className="flex h-11 w-full items-center justify-center gap-2 rounded-lg bg-[#1f7a8c] px-4 font-semibold text-white transition hover:bg-[#196b7b] disabled:cursor-not-allowed disabled:opacity-60"
                disabled={authLoading}
                type="submit"
              >
                {authLoading ? <Loader2 className="animate-spin" size={18} /> : <LogIn size={18} />}
                {authMode === "login" ? "ورود به تسکینو" : "ساخت حساب"}
              </button>
            </form>
          </div>
        </section>
      </main>
    );
  }

  return (
    <main className="min-h-screen text-[#1d2733]">
      <Toast message={error || message} type={error ? "error" : "success"} onClose={() => (error ? setError("") : setMessage(""))} />
      <header className="sticky top-0 z-40 border-b border-[#dce3e9] bg-white/90 shadow-[0_1px_0_rgba(255,255,255,0.7)] backdrop-blur">
        <div className="mx-auto flex max-w-7xl flex-col gap-4 px-4 py-3 sm:px-6 lg:flex-row lg:items-center lg:justify-between lg:px-8">
          <div className="flex items-center gap-3">
            <div className="flex h-11 w-11 items-center justify-center rounded-lg bg-[#1f7a8c] text-white shadow-soft">
              <FolderKanban size={20} />
            </div>
            <div>
              <h1 className="text-2xl font-bold tracking-normal">Taskino</h1>
              <p className="text-sm text-[#687887]">سلام {userName(currentUser ?? undefined)}</p>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button
              className="flex h-10 items-center gap-2 rounded-lg border border-[#cfd9e2] bg-white px-3 text-sm font-semibold text-[#314150] transition hover:bg-[#f3f6f8]"
              onClick={() => loadData()}
              type="button"
            >
              {loadingData ? <Loader2 className="animate-spin" size={17} /> : <RefreshCw size={17} />}
              به‌روزرسانی
            </button>
            <button
              className="flex h-10 items-center gap-2 rounded-lg bg-[#233142] px-3 text-sm font-semibold text-white transition hover:bg-[#1c2836]"
              onClick={logout}
              type="button"
            >
              <LogOut size={17} />
              خروج
            </button>
          </div>
        </div>
      </header>

      <div className="mx-auto grid max-w-7xl gap-5 px-4 py-6 sm:px-6 lg:grid-cols-[300px_1fr] lg:px-8">
        <aside className="space-y-4 lg:sticky lg:top-24 lg:self-start">
          <section className="overflow-hidden rounded-lg border border-[#d8e0e7] bg-white/95 shadow-soft ring-1 ring-white/70">
            <div className="h-1 bg-[#1f7a8c]" />
            <div className="p-4">
            <div className="flex items-center gap-3">
              <div className="flex h-12 w-12 items-center justify-center rounded-lg bg-[#e7f4f2] text-[#1f7a8c]">
                <UserRound size={20} />
              </div>
              <div className="min-w-0">
                <p className="truncate font-semibold">{userName(currentUser ?? undefined)}</p>
                <p className="truncate text-sm text-[#687887]">{currentUser?.mobile ?? currentUser?.email}</p>
              </div>
            </div>
            <div className="mt-4 flex items-center gap-2 rounded-md bg-[#f3f6f8] px-3 py-2 text-sm text-[#4f6172]">
              <ShieldCheck size={16} />
              {currentUser?.roles || "specialist"}
            </div>
            </div>
          </section>

          <section className="rounded-lg border border-[#d8e0e7] bg-white/95 p-4 shadow-soft ring-1 ring-white/70">
            <h2 className="border-b border-[#edf1f5] pb-3 text-lg font-bold">ساخت تسک</h2>
            <form className="mt-4 space-y-3" onSubmit={createTask}>
              <Field label="عنوان تسک" name="taskTitle" value={taskTitle} onChange={setTaskTitle} required />
              <Select label="مسئول" value={taskAssignee} onChange={setTaskAssignee} options={users.map((user) => [getId(user), userName(user)])} />
              <button className="flex h-11 w-full items-center justify-center gap-2 rounded-lg bg-[#1f7a8c] px-3 text-sm font-semibold text-white shadow-sm transition hover:bg-[#196b7b]" type="submit">
                <Plus size={17} />
                افزودن تسک
              </button>
            </form>
          </section>
        </aside>

        <div className="space-y-5">
          <section className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
            {stats.map((stat) => (
              <div key={stat.label} className="group overflow-hidden rounded-lg border border-[#d8e0e7] bg-white/95 shadow-soft ring-1 ring-white/70 transition hover:-translate-y-0.5 hover:border-[#bdd1db]">
                <div className="h-1 bg-[#1f7a8c]" />
                <div className="p-4">
                <div className="flex items-center justify-between">
                  <p className="text-sm text-[#657483]">{stat.label}</p>
                  <span className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#e7f4f2] text-[#1f7a8c]">
                    <stat.icon size={19} />
                  </span>
                </div>
                <p className="mt-3 text-3xl font-bold">{stat.value}</p>
                </div>
              </div>
            ))}
          </section>

          <section className="grid gap-5 xl:grid-cols-[1fr_360px]">
            <div className="overflow-hidden rounded-lg border border-[#d8e0e7] bg-white/95 shadow-soft ring-1 ring-white/70">
              <div className="flex items-center justify-between border-b border-[#e5ebf0] bg-[#fbfcfd] p-4">
                <h2 className="text-lg font-bold">تسک‌ها</h2>
                <span className="text-sm text-[#687887]">{tasks.length} مورد</span>
              </div>
              <div className="divide-y divide-[#edf1f5]">
                {tasks.length === 0 ? (
                  <EmptyState title="هنوز تسکی نیست" text="اولین تسک را از فرم سمت راست بساز." />
                ) : (
                  tasks.map((task) => (
                    <div key={getId(task)} className="flex flex-col gap-3 p-4 transition hover:bg-[#f8fbfc] md:flex-row md:items-center md:justify-between">
                      <div className="min-w-0">
                        <p className="truncate font-semibold">{task.title}</p>
                        <p className="mt-1 text-sm text-[#687887]">
                          مسئول: {task.assignedTo?.length ? task.assignedTo.map((user) => userName(user)).join("، ") : "نامشخص"}
                        </p>
                      </div>
                      <button
                        className="flex h-9 shrink-0 items-center justify-center gap-2 rounded-lg border border-[#cfd9e2] bg-white px-3 text-sm font-semibold text-[#314150] transition hover:border-[#1f7a8c]/40 hover:bg-[#f3f8fa]"
                        onClick={() => updateTaskStatus(task)}
                        type="button"
                      >
                        {task.status === "done" ? <CheckCircle2 size={16} /> : <CircleDashed size={16} />}
                        {statusLabel(task.status)}
                      </button>
                    </div>
                  ))
                )}
              </div>
            </div>

            <section className="rounded-lg border border-[#d8e0e7] bg-white/95 p-4 shadow-soft ring-1 ring-white/70">
              <h2 className="border-b border-[#edf1f5] pb-3 text-lg font-bold">ساخت پروژه</h2>
              <form className="mt-4 space-y-3" onSubmit={createProject}>
                <Field label="عنوان پروژه" name="projectTitle" value={projectTitle} onChange={setProjectTitle} required />
                <Field label="توضیحات" name="projectDescription" value={projectDescription} onChange={setProjectDescription} />
                <Select label="عضو اولیه" value={projectMember} onChange={setProjectMember} options={users.map((user) => [getId(user), userName(user)])} />
                <button className="flex h-11 w-full items-center justify-center gap-2 rounded-lg bg-[#233142] px-3 text-sm font-semibold text-white shadow-sm transition hover:bg-[#1c2836]" type="submit">
                  <Plus size={17} />
                  افزودن پروژه
                </button>
              </form>
            </section>
          </section>

          <section className="overflow-hidden rounded-lg border border-[#d8e0e7] bg-white/95 shadow-soft ring-1 ring-white/70">
            <div className="flex items-center justify-between border-b border-[#e5ebf0] bg-[#fbfcfd] p-4">
              <h2 className="text-lg font-bold">پروژه‌ها</h2>
              <span className="text-sm text-[#687887]">{projects.length} مورد</span>
            </div>
            <div className="grid gap-3 p-4 md:grid-cols-2">
              {projects.length === 0 ? (
                <EmptyState title="پروژه‌ای ثبت نشده" text="یک پروژه بساز تا لیست اینجا پر شود." />
              ) : (
                projects.map((project) => (
                  <article key={getId(project)} className="rounded-lg border border-[#dce3e9] border-r-4 border-r-[#1f7a8c] bg-white p-4 transition hover:border-[#bdd1db] hover:shadow-soft">
                    <div className="flex items-start justify-between gap-3">
                      <div className="min-w-0">
                        <h3 className="truncate font-bold">{project.title}</h3>
                        <p className="mt-2 line-clamp-2 text-sm text-[#687887]">{project.description || "بدون توضیحات"}</p>
                      </div>
                      <span className="rounded-md bg-[#e7f4f2] px-2 py-1 text-xs font-semibold text-[#1f7a8c]">{statusLabel(project.status)}</span>
                    </div>
                    <div className="mt-4 flex flex-wrap gap-2 text-xs text-[#526575]">
                      <span className="rounded-md bg-[#f3f6f8] px-2 py-1">مالک: {userName(project.owner)}</span>
                      <span className="rounded-md bg-[#f3f6f8] px-2 py-1">اعضا: {project.members?.length ?? 0}</span>
                      <span className="rounded-md bg-[#f3f6f8] px-2 py-1">تسک‌ها: {project.tasks?.length ?? 0}</span>
                    </div>
                  </article>
                ))
              )}
            </div>
          </section>
        </div>
      </div>
    </main>
  );
}

function Field({
  label,
  name,
  required,
  type = "text",
  value,
  onChange,
}: {
  label: string;
  name: string;
  required?: boolean;
  type?: string;
  value?: string;
  onChange?: (value: string) => void;
}) {
  return (
    <label className="block">
      <span className="mb-1 block text-sm font-medium text-[#4f6172]">{label}</span>
      <input
        className="h-11 w-full rounded-lg border border-[#cfd9e2] bg-white px-3 text-[#1d2733] outline-none transition placeholder:text-[#96a3af] focus:border-[#1f7a8c] focus:ring-4 focus:ring-[#1f7a8c]/10"
        name={name}
        onChange={(event) => onChange?.(event.target.value)}
        required={required}
        type={type}
        value={value}
      />
    </label>
  );
}

function Select({
  label,
  value,
  onChange,
  options,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: Array<[string, string]>;
}) {
  return (
    <label className="block">
      <span className="mb-1 block text-sm font-medium text-[#4f6172]">{label}</span>
      <select
        className="h-11 w-full rounded-lg border border-[#cfd9e2] bg-white px-3 text-[#1d2733] outline-none transition focus:border-[#1f7a8c] focus:ring-4 focus:ring-[#1f7a8c]/10"
        onChange={(event) => onChange(event.target.value)}
        value={value}
      >
        <option value="">انتخاب نشده</option>
        {options.map(([id, labelText]) => (
          <option key={id} value={id}>
            {labelText}
          </option>
        ))}
      </select>
    </label>
  );
}

function Toast({
  message,
  type,
  onClose,
}: {
  message: string;
  type: "success" | "error";
  onClose: () => void;
}) {
  if (!message) return null;

  const isSuccess = type === "success";

  return (
    <div className="fixed left-4 top-5 z-50 w-[calc(100%-2rem)] max-w-sm sm:left-6 sm:top-6" role="status" aria-live="polite">
      <div
        className={`flex items-start gap-3 rounded-lg border bg-white px-4 py-3 shadow-soft ${
          isSuccess ? "border-[#cfe9dd] text-[#247a56]" : "border-[#f3c9c1] text-[#b23b2e]"
        }`}
      >
        <div className={`mt-0.5 flex h-8 w-8 shrink-0 items-center justify-center rounded-lg ${isSuccess ? "bg-[#edf8f3]" : "bg-[#fff0ed]"}`}>
          {isSuccess ? <CheckCircle2 size={18} /> : <X size={18} />}
        </div>
        <p className="min-w-0 flex-1 pt-1 text-sm font-medium leading-6">{message}</p>
        <button
          aria-label="بستن اعلان"
          className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-[#657483] transition hover:bg-[#f3f6f8] hover:text-[#1d2733]"
          onClick={onClose}
          type="button"
        >
          <X size={16} />
        </button>
      </div>
    </div>
  );
}

function EmptyState({ title, text }: { title: string; text: string }) {
  return (
    <div className="col-span-full p-6 text-center">
      <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-lg bg-[#edf2f6] text-[#657483]">
        <ClipboardList size={22} />
      </div>
      <p className="mt-3 font-semibold">{title}</p>
      <p className="mt-1 text-sm text-[#687887]">{text}</p>
    </div>
  );
}
