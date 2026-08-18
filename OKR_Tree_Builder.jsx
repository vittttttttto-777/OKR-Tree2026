import React, { useState, useEffect, useCallback, useRef, useContext, createContext } from "react";
import * as XLSX from "xlsx";
import { createClient as createSupabaseClient } from "@supabase/supabase-js";
import {
  ChevronDown, ChevronRight, Plus, Trash2, Target, Flag,
  Layers, CalendarRange, ListChecks, CheckCircle2, Circle, Sparkles, Users, User, Network, Briefcase, Search,
  Download, Upload, Activity, BarChart3, PieChart, ShieldCheck, FileUp, Eye
} from "lucide-react";

// ---- Language toggle (RU / EN) — translates app UI chrome; user-entered content stays as typed ----
const LanguageCtx = createContext("ru");
function useLang() { return useContext(LanguageCtx); }
function useT() {
  const lang = useLang();
  return (s) => (lang === "en" && EN_DICT[s]) ? EN_DICT[s] : s;
}
const EN_DICT = {
  "OKR-конструктор трека": "OKR Tree Builder",
  "Волна 1–6 мес. декомпозируется сразу. Волны 7–12 и 13–18 стартуют как цель-ориентир — раскладывайте их на контрольной точке, ближе к делу. У каждого Objective — владелец, у каждой задачи — ответственный из справочника.":
    "Wave 1–6 mo. is broken down right away. Waves 7–12 and 13–18 start as a target direction — break them down at the checkpoint, closer to the time. Every Objective has an owner; every task has someone responsible from the directory.",
  "Общее дерево OKR": "Combined OKR Tree",
  "Проекты": "Projects",
  "Дашборд трекинга": "Tracking Dashboard",
  "Дашборд проектов": "Projects Dashboard",
  "Пока нет проектов": "No projects yet",
  "Всего проектов": "Total projects",
  "Активны в этом месяце": "Active this month",
  "Загрузка по дивизионам": "Load by division",
  "Распределение статусов": "Status distribution",
  "Активных проектов / месяц": "Active projects / month",
  "Топ функций по кол-ву проектов": "Top functions by project count",
  "Гант": "Gantt",
  "Главные цели": "Main goals",
  "Волны": "Waves",
  "Кварталы": "Quarters",
  "Месяцы": "Months",
  "Задачи": "Tasks",
  "план (без трекинга)": "planned (not tracked)",
  "свернуть": "less",
  "ещё": "more",
  "Outcome — итоговый результат": "Outcome — the end result",
  "Инициатива — программа работ": "Initiative — the work program",
  "Output — результат программы": "Output — the program's result",
  "Волна — полугодовой этап": "Wave — half-year stage",
  "Квартал — операционный план": "Quarter — the operating plan",
  "Пока нет заполненных целей ни в одном треке.": "No goals filled in yet in any track.",
  "Дерево ещё не открыто": "The tree hasn't been opened yet",
  "Войдите под своей ролью — и увидите свою ветку: от цели на 18 месяцев до задач этой недели.":
    "Sign in with your role to see your branch — from the 18-month goal down to this week's tasks.",
  "Этот раздел не по вашей роли": "This section isn't part of your role",
  "У вашей роли нет доступа сюда — переключитесь на раздел, который вам доступен, или попросите администратора расширить права.":
    "Your role doesn't have access here — switch to a section you can view, or ask an administrator to extend your permissions.",

  "Google Таблица": "Google Sheet",
  "не подключено": "not connected",
  "Управление привилегиями": "Manage privileges",
  "Клик по ячейке переключает право по кругу: — → Видит → Меняет → —. Изменения сохраняются сразу.": "Click a cell to cycle the right: — → View → Edit → —. Changes save immediately.",
  "Модуль": "Module",
  "Видит": "View",
  "Меняет": "Edit",
  "Пароли ролей": "Role passwords",
  "новый пароль": "new password",
  "Сохранить": "Save",
  "Новая роль": "New role",
  "Название роли": "Role name",
  "Треки (все)": "All tracks",
  "Трек Health OS": "Health OS track",
  "Трек GrowthModel": "GrowthModel track",
  "Трек PrimeGrowth": "PrimeGrowth track",
  "Трек CoralEVO": "CoralEVO track",
  "Трек ITModel": "ITModel track",
  "Общее дерево OKR (связи)": "Combined OKR tree (links)",
  "Вход": "Sign in",
  "Выберите роль": "Choose a role",
  "Пароль": "Password",
  "Войти": "Sign in",
  "Входим…": "Signing in…",
  "Неверная роль или пароль.": "Wrong role or password.",
  "Без входа приложение доступно только на просмотр. Роль и пароль вам выдаёт администратор.": "Without signing in, the app is view-only. Your admin gives you a role and password.",
  "Войдите, чтобы редактировать": "Sign in to edit",
  "Доступен только просмотр": "View only",
  "Этот раздел вам недоступен — войдите под ролью с нужными правами.": "This section isn't available to you — sign in with a role that has access.",
  "Войти через Google": "Sign in with Google",
  "выйти": "sign out",
  "Администратор": "Admin",
  "Редактор": "Editor",
  "Только просмотр": "View only",
  "ID таблицы Google Sheets (из ссылки на таблицу)": "Google Sheets ID (from the sheet's link)",
  "Подключить": "Connect",
  "Подключение к таблице…": "Connecting to the sheet…",
  "Без входа приложение работает локально в этом браузере (демо-режим). Со входом — данные общие для всей команды, через вашу Google Таблицу.":
    "Without signing in, the app works locally in this browser (demo mode). Once signed in, data is shared with the whole team via your Google Sheet.",

  "Скачать Excel": "Download Excel",
  "Загрузить Excel": "Upload Excel",
  "Готово — файл скачан.": "Done — file downloaded.",
  "Ошибка экспорта: ": "Export error: ",
  "Не удалось прочитать файл: ": "Failed to read the file: ",
  "Ошибка импорта: ": "Import error: ",
  "Заменить": "Replace",
  "Отмена": "Cancel",
  "Импортировано — данные обновлены.": "Imported — data updated.",
  "Файл": "File",
  "прочитан. Импорт заменит ВСЕ текущие данные (все треки, проекты, справочник, связи). Продолжить?":
    "has been read. Importing will replace ALL current data (every track, project, directory entry and link). Continue?",

  "Справочник: компания / отделы / функции": "Directory: company / departments / functions",
  "Компания": "Company",
  "Отдел": "Department",
  "Функция": "Function",
  "Название (например: Отдел маркетинга)": "Name (e.g. Marketing department)",
  "Справочник пуст — добавьте функции, отделы или компанию": "The directory is empty — add functions, departments, or the company",
  "Справочник редактирует администратор": "Only an admin can edit the directory",

  "Загрузка…": "Loading…",
  "Заполнено полей": "Fields filled",
  "Миссия трека": "Track mission",
  "Зачем трек существует": "Why the track exists",
  "Главная цель": "Main goal",
  "из": "of",
  "Ultimate Objective (18 мес.)": "Ultimate Objective (18 mo.)",
  "Одна из главных стратегических целей трека на 18 месяцев": "One of the track's main strategic goals over 18 months",
  "Владелец Objective:": "Objective owner:",
  "Без ответственного": "No owner",
  "Импорт задач из Bitrix24": "Import tasks from Bitrix24",
  "Файл экспорта задач из Bitrix24 (.xls) — колонки «Название», «Крайний срок», «Теги».":
    "Bitrix24 task export file (.xls) — columns \u201cName\u201d, \u201cDeadline\u201d, \u201cTags\u201d.",
  "Загрузить файл Bitrix": "Upload Bitrix file",
  "Загружено строк: ": "Rows loaded: ",
  "не нашёл таблицу с колонкой «Название» в файле": "couldn't find a table with a \u201cName\u201d column in the file",
  "Загрузите файл выгрузки задач из Bitrix24, чтобы начать проверку.": "Upload a Bitrix24 task export file to start reviewing.",
  "Отмечено к применению:": "Marked for applying:",
  "Принять все": "Accept all",
  "Применить отмеченное": "Apply marked",
  "Трек…": "Track…",
  "срок": "due",
  "не размещено": "not placed yet",
  "Текст": "Text",
  "Главная цель…": "Ultimate objective…",
  "новая Главная цель": "new ultimate objective",
  "KR Outcome…": "KR Outcome…",
  "новый KR Outcome": "new KR Outcome",
  "KR Outcome появится после создания цели": "KR Outcome becomes available once the goal is created",
  "KR Output…": "KR Output…",
  "Волна…": "Wave…",
  "Квартал…": "Quarter…",
  "Месяц…": "Month…",
  "Загрузка трека…": "Loading track…",
  "Ничего не применено — проверьте лимиты (например, максимум Главных целей/KR Outcome/задач в месяце).":
    "Nothing was applied — check the limits (e.g. max ultimate objectives / KR Outcomes / tasks per month).",
  "Применено: ": "Applied: ",
  "Ошибка применения: ": "Apply error: ",
  "Импорт из Bitrix24": "Import from Bitrix24",
  "Импорт доступен администратору": "Import is available to the administrator",
  "черновик": "draft",
  "Objective волны (6 мес.)": "Wave objective (6 mo.)",
  "Инициатива квартала (3 мес.)": "Quarter initiative (3 mo.)",
  "Готовы к применению:": "Ready to apply:",
  "на конфликте (остаются черновиком в Ганте):": "in conflict (stay as a Gantt draft):",
  "конфликтных строк оставлено на доработку: ": "conflicting rows left for cleanup: ",
  "конфликт — сначала виден только в Ганте": "conflict — visible only in the Gantt for now",
  "будет объединено рядом с существующим при применении": "will be merged alongside the existing text when applied",
  "применить всё равно (дописать рядом)": "apply anyway (append alongside)",
  "применить всё равно (добавить рядом)": "apply anyway (add alongside)",
  "Похожая задача уже есть в этом месяце": "A similar task already exists in this month",
  "Сейчас там:": "Currently there:",
  "Выберите место в дереве трека": "Choose a spot in the track's tree",
  "Поиск по дереву…": "Search the tree…",
  "Обзор дерева": "Browse tree",
  "место подставлено по коду": "placement filled in from the code",
  "Файл экспорта задач из Bitrix24 (.xls) — колонки «Название», «Крайний срок», «Теги», «№ заявки».":
    "Bitrix24 task export file (.xls) — columns \u201cName\u201d, \u201cDeadline\u201d, \u201cTags\u201d, \u201cTask #\u201d.",
  "Совет: добавьте в выгрузку колонку с номером задачи Bitrix (ID / № заявки) — после того как вы один раз разместите задачу вручную, её номер сохранится, и при повторном импорте того же файла приложение само узнает эту задачу и не предложит разместить её ещё раз.":
    "Tip: include a column with the Bitrix task number (ID / task #) in the export — once you place a task by hand, its number is saved, and re-importing the same file later will recognize that task automatically instead of asking you to place it again.",
  "уже импортирована по ID": "already imported (matched by ID)",
  "Уже в дереве без изменений — можно пропустить.": "Already in the tree, unchanged — can be skipped.",
  "В дереве сейчас другой текст:": "The tree currently has different text:",
  "если в Bitrix задачу переименовали, отметьте чекбокс, чтобы дописать новый текст рядом.":
    "if the task was renamed in Bitrix, check the box to append the new text alongside.",
  "Ответственный": "Owner",
  "Назначить ответственного": "Assign owner",
  "Ответственный за задачу": "Task owner",
  "Добавить KR Outcome": "Add KR Outcome",
  "Добавить главную цель": "Add main goal",
  "Ключевой результат стратегической цели (18 мес.)": "Key result of the strategic goal (18 mo.)",
  "Метрика": "Metric",
  "Как измеряется (число, %, статус)": "How it's measured (number, %, status)",
  "Стратегическая инициатива на 18 мес.": "Strategic initiative (18 mo.)",
  "Программа работ, обеспечивающая KR Outcome": "Work program that delivers the KR Outcome",
  "Objective": "Objective",
  "Цель инициативы": "Goal of the initiative",
  "Формулировка ключевого результата (output) на 18 мес.": "Wording of the key result (output) over 18 months",

  "Трекинг KR (по неделям)": "KR tracking (weekly)",
  "добавить трекинг прогресса": "add progress tracking",
  "убрать трекинг": "remove tracking",
  "Тип метрики": "Metric type",
  "Накопительный (прирост)": "Cumulative (increment)",
  "Уровень (срез)": "Level (snapshot)",
  "Исходное": "Baseline",
  "Целевое": "Target",
  "добавить неделю": "add week",
  "прирост за неделю": "increment this week",
  "значение на конец недели": "value at week end",
  "Факт": "Actual",
  "Поставлено": "Set on",
  "Дата начала трека": "Track start date",
  "Срок": "Due",
  "осталось": "left",
  "дн.": "d.",
  "просрочено на": "overdue by",
  "Просрочено": "Overdue",
  "Критично": "Critical",
  "Скоро": "Soon",
  "В графике": "On schedule",
  "изменить срок": "change deadline",
  "сброс": "reset",
  "готово": "done",
  "сегодня": "today",
  "Уверенность 1-10": "Confidence 1-10",
  "На треке": "On track",
  "Есть риск": "At risk",
  "Отстаёт": "Behind",

  "Волна": "Wave",
  "Objective на 6 мес.": "Objective (6 mo.)",
  "Цель полугодия": "Half-year goal",
  "Цель-ориентир (без декомпозиции)": "Target direction (not broken down)",
  "Чего нужно достичь в этом полугодии — детали позже, на контрольной точке": "What needs to be achieved this half-year — details later, at the checkpoint",
  "разложить волну": "break down wave",
  "свернуть в цель": "collapse to target",

  "Квартал": "Quarter",
  "веха + инициатива 3 мес": "milestone + 3-mo initiative",
  "KR-веха квартала": "Quarter KR milestone",
  "KR-веха": "KR milestone",
  "KR месяца": "Month KR",
  "мес.": "mo.",
  "Что считается достигнутым к концу квартала": "What counts as achieved by quarter end",
  "Инициатива на 3 мес.": "3-month initiative",
  "Операционный проект квартала": "The quarter's operating project",

  "KR месяца — что должно быть достигнуто": "Month KR — what should be achieved",
  "задача": "task",
  "Задача": "Task",

  "Единая цель Coral Club из": "Coral Club's unified goal across",
  "треков": "tracks",
  "уместить все треки": "fit all tracks",
  "Конфликт": "Conflict",
  "Дополнение": "Complement",
  "Ожидание": "Waiting",
  "Синхронизация": "Sync",
  "Связи между треками": "Cross-track links",
  "Пока ни одна волна ни в одном треке не заполнена": "No wave in any track has been filled in yet",
  "Выбрано:": "Selected:",
  "Кликните 🔗 у другого узла (цели, KR, задачи или проекта), чтобы связать.": "Click 🔗 on another node (goal, KR, task, or project) to link them.",
  "отмена": "cancel",
  "Тип связи между целями": "Link type between goals",
  "пока нет заполненных целей": "no goals filled in yet",
  "волны ещё не разложены": "waves not broken down yet",

  "Проект/Инициатива": "Project/Initiative",
  "Проектов": "Projects",
  "трека": "tracks",
  "Проекты / инициативы трека": "Track projects / initiatives",
  "без названия": "untitled",
  "старт": "starts",
  "добавить проект": "add project",
  "проектов": "projects",
  "связаны с треками": "linked to tracks",
  "Все дивизионы": "All divisions",
  "Поиск по названию, функции, дивизиону…": "Search by name, function, division…",
  "Ничего не найдено": "Nothing found",
  "Описание проекта": "Project description",
  "Тип проекта": "Project type",
  "Статусы проектов": "Project statuses",
  "Новый статус (например: На паузе)": "New status (e.g. On hold)",
  "Статус": "Status",
  "Дата начала": "Start date",
  "Ожидаемая дата конца": "Expected end date",
  "RUN — текущая операционная деятельность": "RUN — ongoing operational activity",
  "Change — разовая инициатива изменений": "Change — one-off change initiative",
  "Кратко опишите проект…": "Briefly describe the project…",
  "Каким трекам помогает": "Which tracks it supports",
  "KR проекта": "Project KRs",
  "добавить KR": "add KR",
  "Удалить проект": "Delete project",
  "Проекты редактирует администратор": "Only an admin can edit projects",

  "KR под трекингом": "KRs tracked",
  "средний прогресс": "average progress",
  "на треке": "on track",
  "отстают": "behind",
  "Пока нет KR с включённым трекингом. Кнопка «добавить трекинг прогресса» доступна у KR Outcome, KR Output, KR-вехи квартала и KR месяца в любом треке.":
    "No KRs are being tracked yet. The \u201cadd progress tracking\u201d button is available on any KR Outcome, KR Output, quarter KR milestone, or month KR in any track.",

  "Войдите через Google, чтобы редактировать": "Sign in with Google to edit",
  "Доступен только просмотр — это не ваш трек": "View only — this isn't your track",
  "Ваш email не назначен ни на один трек — обратитесь к администратору": "Your email isn't assigned to any track — contact an admin",
  "Не удалось войти: ": "Failed to sign in: ",
  "Вошли, но не удалось получить email.": "Signed in, but couldn't retrieve the email.",
  "Не удалось загрузить Google Sign-In (проверьте интернет-соединение).": "Failed to load Google Sign-In (check your internet connection).",
  "Нет прав на изменение этих данных — обратитесь к администратору.": "You don't have permission to change this data — contact an admin.",
  "Не удалось сохранить данные в Google Таблице.": "Failed to save data to the Google Sheet.",
  "Не удалось подключиться к Google Таблице.": "Failed to connect to the Google Sheet.",
  "Email": "Email",
  "не найден на листе": "wasn't found on the sheet",
  "обратитесь к администратору.": "contact an admin.",
  "Связать с другой целью, KR, задачей или проектом": "Link to another goal, KR, task, or project",
  "Активность по месяцам 2026": "2026 monthly activity",
};

// ---- Data translation (translates the user's own OKR text via a free public API) ----
// Unlike EN_DICT above (static UI copy), this handles arbitrary text the user typed —
// goal names, descriptions, tasks, etc. Only used in read-only views (Combined tree,
// Dashboard, project cards); edit forms always show the real saved text, untranslated,
// so nothing gets corrupted on save.
const TRANSLATE_CACHE_KEY = "okr-translate-cache";
function loadTranslateCache() {
  try { return JSON.parse(localStorage.getItem(TRANSLATE_CACHE_KEY) || "{}"); } catch { return {}; }
}
const translateCache = loadTranslateCache();
function saveTranslateCache() {
  try { localStorage.setItem(TRANSLATE_CACHE_KEY, JSON.stringify(translateCache)); } catch {}
}
const pendingTranslations = new Map();

async function translateViaApi(text) {
  try {
    const res = await fetch(`https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=ru|en`);
    const data = await res.json();
    const out = data && data.responseData && data.responseData.translatedText;
    return out && !/^PLEASE SELECT/i.test(out) ? out : null;
  } catch {
    return null;
  }
}

function useDataT(text) {
  const lang = useLang();
  const [, bump] = useState(0);

  useEffect(() => {
    if (lang !== "en" || !text || !text.trim()) return;
    if (translateCache[text] !== undefined) return;
    if (pendingTranslations.has(text)) return;
    const p = translateViaApi(text).then((result) => {
      translateCache[text] = result || text;
      saveTranslateCache();
      pendingTranslations.delete(text);
      bump((n) => n + 1);
    });
    pendingTranslations.set(text, p);
  }, [text, lang]);

  if (lang !== "en" || !text) return text;
  return translateCache[text] !== undefined ? translateCache[text] : text;
}

function FieldPreview({ value }) {
  const lang = useLang();
  const preview = useDataT(value);
  if (lang !== "en" || !value || !value.trim() || !preview || preview === value) return null;
  return <div className="text-xs text-sky-600 italic mt-0.5">🌐 {preview}</div>;
}

function TranslatedText({ text }) {
  return useDataT(text);
}

const TRACKS = [
  { id: "health-os", name: "Health OS" },
  { id: "growth-model", name: "GrowthModel" },
  { id: "prime-growth", name: "PrimeGrowth" },
  { id: "coral-evo", name: "CoralEVO" },
  { id: "it-model", name: "ITModel" },
];

const TRACK_COLORS = {
  "health-os": "#16a34a",
  "growth-model": "#ca8a04",
  "prime-growth": "#b45309",
  "coral-evo": "#8b5cf6",
  "it-model": "#f87171",
};

const MAX_PROJECT_TRACKS = 3;

const STATUS_COLORS = { G: "#16a34a", Y: "#ca8a04", R: "#dc2626", N: "#94a3b8" };

const MAX_TASKS = 25;

const PROJECTS_SEED = [
  {
    "id": "seed-0",
    "division": "GLOBAL",
    "function": "ДП",
    "name": "Описание процессов",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-1",
    "division": "GLOBAL",
    "function": "ДП",
    "name": "Стартовая программа для новичков",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": "03.26",
    "activity": [
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-2",
    "division": "GLOBAL",
    "function": "ДП",
    "name": "Друзья Клуба",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": "02.26",
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-3",
    "division": "GLOBAL",
    "function": "ДП",
    "name": "Быстрый старт",
    "status": "—",
    "statusCode": "N",
    "launchMonth": "06.26",
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-4",
    "division": "GLOBAL",
    "function": "ДП",
    "name": "Программы лояльности",
    "status": "—",
    "statusCode": "N",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-5",
    "division": "GLOBAL",
    "function": "ДРД",
    "name": "Digital kit",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-6",
    "division": "GLOBAL",
    "function": "CX",
    "name": "CRM Bitrix (IP телефония для КЦ)",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-7",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "Refiner",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": "07.26",
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-8",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "Промокоды",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-9",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "Адреса на сайте",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-10",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "Курьерские службы",
    "status": "Откл. >10%",
    "statusCode": "R",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-11",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "Верификация",
    "status": "Откл. >10%",
    "statusCode": "R",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-12",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "База знаний",
    "status": "Откл. >10%",
    "statusCode": "R",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-13",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "Упрощенная регистрация",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-14",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "Статусная модель заказов",
    "status": "—",
    "statusCode": "N",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-15",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "ЛК Дистрибьютора — Back office",
    "status": "—",
    "statusCode": "N",
    "launchMonth": "06.26",
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-16",
    "division": "GLOBAL",
    "function": "Ecom",
    "name": "Навигация по продуктам",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-17",
    "division": "GLOBAL",
    "function": "GM",
    "name": "Пульс — Apache Superset (BI)",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": "03.26",
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-18",
    "division": "GLOBAL",
    "function": "Маркетинг",
    "name": "AI ассистент для ЧК и Дистрибьюторов",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-19",
    "division": "GLOBAL",
    "function": "HR",
    "name": "Бизнес процессы и заявки",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-20",
    "division": "GLOBAL",
    "function": "HR",
    "name": "Инструменты для менеджеров (дашборды)",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-21",
    "division": "GLOBAL",
    "function": "HR",
    "name": "Обучение сотрудников — новая платформа",
    "status": "—",
    "statusCode": "N",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-22",
    "division": "GLOBAL",
    "function": "HR",
    "name": "ИИ-ассистент Коралина",
    "status": "—",
    "statusCode": "N",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-23",
    "division": "GLOBAL",
    "function": "Training",
    "name": "Бизнес обучение + CBA (LMS)",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": "04.26",
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-24",
    "division": "GLOBAL",
    "function": "Events",
    "name": "Event — замена платформ",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-25",
    "division": "GLOBAL",
    "function": "Finance",
    "name": "Внедрение 1С ERP",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": "08.26",
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-26",
    "division": "GLOBAL",
    "function": "Finance",
    "name": "Интеграция с локальными бух. системами",
    "status": "—",
    "statusCode": "N",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-27",
    "division": "GLOBAL",
    "function": "Finance",
    "name": "Фискализация по всем рынкам",
    "status": "—",
    "statusCode": "N",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-28",
    "division": "GLOBAL",
    "function": "Logistics",
    "name": "Штрихкодирование",
    "status": "Откл. >10%",
    "statusCode": "R",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-29",
    "division": "GLOBAL",
    "function": "Marketing",
    "name": "Эффективная коммуникация (смс, имейл)",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-30",
    "division": "GLOBAL",
    "function": "Marketing",
    "name": "Контент платформа для трафика",
    "status": "—",
    "statusCode": "N",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-31",
    "division": "GLOBAL",
    "function": "Marketing",
    "name": "Подписка",
    "status": "—",
    "statusCode": "N",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-32",
    "division": "GLOBAL",
    "function": "Operations",
    "name": "РЦ — внедрение новых систем (1С)",
    "status": "—",
    "statusCode": "N",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-33",
    "division": "GLOBAL",
    "function": "Planning",
    "name": "Региональные скидки",
    "status": "Откл. >10%",
    "statusCode": "R",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-34",
    "division": "GLOBAL",
    "function": "Logistics",
    "name": "Сервис логистики",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": "04.26",
    "activity": [
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-35",
    "division": "РФ+РБ",
    "function": "HR",
    "name": "КЭДО",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-36",
    "division": "РФ+РБ",
    "function": "Operations",
    "name": "Бизнес модель РФ (франчайзи, портал)",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-37",
    "division": "РФ+РБ",
    "function": "Operations",
    "name": "Поставки в РФ, включая маркировку",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-38",
    "division": "РФ+РБ",
    "function": "Operations",
    "name": "Честный знак",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-39",
    "division": "РФ+РБ",
    "function": "Operations",
    "name": "Производство паучей",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-40",
    "division": "РФ+РБ",
    "function": "Finance",
    "name": "ЭДО",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-41",
    "division": "РФ+РБ",
    "function": "Operations",
    "name": "Транспортные документы ЭДО",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-42",
    "division": "Азия",
    "function": "Ecom",
    "name": "Верификация Дистрибьюторов",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-43",
    "division": "Азия",
    "function": "Operations",
    "name": "Маркировка Казахстан",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-44",
    "division": "Азия",
    "function": "Operations",
    "name": "Маркировка Таджикистан / Армения",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-45",
    "division": "Азия",
    "function": "ДПД",
    "name": "Азербайджан",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": "07.26",
    "activity": [
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-46",
    "division": "Америка+Канада",
    "function": "ДПД",
    "name": "Изменение маркетинг плана",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-47",
    "division": "Америка+Канада",
    "function": "ДПД",
    "name": "Расширение географии — Пуэрто-Рико",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-48",
    "division": "Америка+Канада",
    "function": "ДПД",
    "name": "Расширение географии — Мексика",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-49",
    "division": "Европа-Север",
    "function": "Operations",
    "name": "3PL в Ирландии",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-50",
    "division": "Европа-Север",
    "function": "Operations",
    "name": "DDP для Великобритании",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-51",
    "division": "Европа-Север",
    "function": "Finance",
    "name": "Норвегия — переход на Евро",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-52",
    "division": "Европа-Центр",
    "function": "ДПД",
    "name": "Румыния",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-53",
    "division": "Европа-Центр",
    "function": "ДПД",
    "name": "Подписка",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": "07.26",
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-54",
    "division": "Европа-Запад",
    "function": "Operations",
    "name": "Склад в Германии",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": "11.26",
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-55",
    "division": "Европа-Запад",
    "function": "Operations",
    "name": "DDP для доставки из Германии",
    "status": "Откл. до 10%",
    "statusCode": "Y",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-56",
    "division": "Европа-Юг",
    "function": "Finance",
    "name": "Испания — документооборот с налоговой",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": "10.26",
    "activity": [
      1,
      1,
      1,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-57",
    "division": "Европа-Юг",
    "function": "Finance",
    "name": "Болгария — Евро",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": "08.26",
    "activity": [
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      1,
      0,
      0,
      0,
      0
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  },
  {
    "id": "seed-58",
    "division": "Back office",
    "function": "Finance",
    "name": "Автоматизация учета рабочего времени",
    "status": "В плане",
    "statusCode": "G",
    "launchMonth": null,
    "activity": [
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      1,
      1,
      1,
      1,
      1
    ],
    "description": "",
    "krs": [],
    "trackIds": []
  }
];


const newId = () => `id_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;

function createTask() {
  return { id: newId(), text: "", ownerId: "" };
}
function createMonthlyKR(label) {
  return { id: newId(), label, text: "", tracking: null, tasks: [createTask(), createTask()] };
}
function createQuarter(label, monthLabels) {
  return {
    id: newId(),
    label,
    milestone: "",
    initiative3mo: "",
    objective: "",
    objectiveOwner: "",
    tracking: null,
    monthlyKRs: monthLabels.map(createMonthlyKR),
  };
}
function createWave(periodLabel, decomposed) {
  return {
    id: newId(),
    periodLabel,
    status: decomposed ? "decomposed" : "target",
    targetText: "",
    objective6mo: "",
    objective6moOwner: "",
    quarters: [
      createQuarter("1–3 мес", ["Мес. 1", "Мес. 2", "Мес. 3"]),
      createQuarter("4–6 мес", ["Мес. 4", "Мес. 5", "Мес. 6"]),
    ],
  };
}
function createKROutput(n) {
  return {
    id: newId(),
    label: `KR_${n} OUTPUT`,
    text: "",
    tracking: null,
    waves: [
      createWave("1–6 мес из 18", true),
      createWave("7–12 мес из 18", false),
      createWave("13–18 мес из 18", false),
    ],
  };
}
function createKROutcome(n) {
  return {
    id: newId(),
    label: `KR Outcome ${n}`,
    text: "",
    metric: "",
    initiativeText: "",
    objectiveText: "",
    objectiveOwner: "",
    tracking: null,
    krOutputs: [createKROutput(1), createKROutput(2), createKROutput(3)],
  };
}
const MAX_ULTIMATE_OBJECTIVES = 2;
const MAX_KR_OUTCOMES = 4;
function createUltimateObjective(n) {
  return { id: newId(), label: `Главная цель ${n}`, text: "", owner: "", krOutcomes: [] };
}
function migrateTrackData(data) {
  if (!data) return createTrackData();
  if (Array.isArray(data.ultimateObjectives)) {
    if (!data.startDate) data.startDate = todayISO();
    return data;
  }
  // legacy shape: single ultimateObjective/Owner + top-level krOutcomes
  return {
    mission: data.mission || "",
    startDate: todayISO(),
    ultimateObjectives: [{
      id: newId(),
      label: "Главная цель 1",
      text: data.ultimateObjective || "",
      owner: data.ultimateObjectiveOwner || "",
      krOutcomes: data.krOutcomes || [],
    }],
  };
}

function createTrackData() {
  return { mission: "", startDate: todayISO(), ultimateObjectives: [createUltimateObjective(1)] };
}

function createTrackerWeek(n) {
  return { id: newId(), label: `Неделя ${n}`, value: 0, confidence: 5 };
}
function createTracker() {
  return { measureType: "cumulative", target: 0, baseline: 0, weeks: [], startDate: todayISO(), deadlineOverride: null };
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function addMonthsISO(dateStr, months) {
  const d = new Date((dateStr || todayISO()) + "T00:00:00");
  d.setMonth(d.getMonth() + months);
  return d.toISOString().slice(0, 10);
}
function daysBetweenISO(fromISO, toISO) {
  const a = new Date(fromISO + "T00:00:00");
  const b = new Date(toISO + "T00:00:00");
  return Math.round((b - a) / 86400000);
}
function formatDateRu(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y}`;
}
function deadlineUrgency(daysLeft) {
  if (daysLeft < 0) return { label: "Просрочено", color: "#dc2626" };
  if (daysLeft <= 14) return { label: "Критично", color: "#dc2626" };
  if (daysLeft <= 45) return { label: "Скоро", color: "#ca8a04" };
  return { label: "В графике", color: "#16a34a" };
}
function clamp01to100(n) {
  return Math.min(100, Math.max(0, n));
}
function buildMonthTicks(rangeStartISO, rangeEndISO) {
  const start = new Date(rangeStartISO + "T00:00:00");
  const end = new Date(rangeEndISO + "T00:00:00");
  const cur = new Date(start.getFullYear(), start.getMonth(), 1);
  const ticks = [];
  while (cur <= end) {
    if (cur >= start) {
      const iso = cur.toISOString().slice(0, 10);
      const label = cur.toLocaleDateString("ru-RU", { month: "short", year: "2-digit" });
      ticks.push({ iso, label });
    }
    cur.setMonth(cur.getMonth() + 1);
  }
  return ticks;
}
function computeTracking(t) {
  if (!t) return { fact: 0, pct: 0 };
  const weeks = t.weeks || [];
  const base = Number(t.baseline) || 0;
  const target = Number(t.target) || 0;
  const denom = target - base;
  const sumDelta = weeks.reduce((s, w) => s + (Number(w.value) || 0), 0);
  let fact, pct;
  if (t.measureType === "level") {
    fact = weeks.length ? Number(weeks[weeks.length - 1].value) || 0 : base;
    pct = denom === 0 ? (fact >= target ? 1 : 0) : (fact - base) / denom;
  } else {
    fact = sumDelta + base;
    pct = denom === 0 ? (sumDelta >= denom ? 1 : 0) : sumDelta / denom;
  }
  return { fact, pct };
}
function trackingStatus(pct) {
  if (pct >= 0.9) return { label: "На треке", color: "#16a34a" };
  if (pct >= 0.5) return { label: "Есть риск", color: "#ca8a04" };
  return { label: "Отстаёт", color: "#dc2626" };
}

function normKey(s) {
  return String(s || "").trim().toLowerCase();
}
function ownerNameById(directory, id) {
  if (!id) return "";
  const d = (directory || []).find((d) => d.id === id);
  return d ? d.name : "";
}
function resolveOwnerId(name, directory, extra) {
  const trimmed = String(name || "").trim();
  if (!trimmed) return "";
  const found =
    (directory || []).find((d) => normKey(d.name) === normKey(trimmed)) ||
    extra.find((d) => normKey(d.name) === normKey(trimmed));
  if (found) return found.id;
  const created = { id: newId(), name: trimmed, type: "function" };
  extra.push(created);
  return created.id;
}
function firstNonEmpty(rows, field) {
  for (const r of rows) {
    const v = r[field];
    if (v !== undefined && v !== null && String(v).trim() !== "") return String(v);
  }
  return "";
}
function groupByNum(rows, field) {
  const map = new Map();
  rows.forEach((r) => {
    const n = Number(r[field]);
    if (!n) return;
    if (!map.has(n)) map.set(n, []);
    map.get(n).push(r);
  });
  return map;
}

// ---- Export: build human-fillable rows (names/numbers, no opaque IDs) ----

function buildTrackRow(trackName, data) {
  return { "Трек": trackName, "Миссия": data.mission };
}

function buildStructureRows(trackName, data, directory) {
  const rows = [];
  (data.ultimateObjectives || []).forEach((uo, ui) => {
    (uo.krOutcomes || []).forEach((o, oi) => {
      (o.krOutputs || []).forEach((ko, ki) => {
        (ko.waves || []).forEach((w, wi) => {
          const base = {
            "Трек": trackName,
            "Главная цель №": ui + 1,
            "Ultimate Objective": uo.text,
            "Владелец Ultimate Objective": ownerNameById(directory, uo.owner),
            "KR Outcome №": oi + 1,
            "KR Outcome": o.text,
            "Метрика": o.metric,
            "Стратегическая инициатива": o.initiativeText,
            "Objective инициативы": o.objectiveText,
            "Владелец Objective инициативы": ownerNameById(directory, o.objectiveOwner),
            "KR Output №": ki + 1,
            "KR Output": ko.text,
            "Волна №": wi + 1,
          };
          if (w.status === "decomposed") {
            w.quarters.forEach((q, qi) => {
              rows.push({
                ...base,
                "Цель-ориентир волны": "",
                "Objective на 6 мес.": w.objective6mo,
                "Владелец Objective 6 мес.": ownerNameById(directory, w.objective6moOwner),
                "Квартал (1-3 / 4-6)": qi === 0 ? "1-3" : "4-6",
                "KR-веха квартала": q.milestone,
                "Инициатива на 3 мес.": q.initiative3mo,
                "Objective квартала": q.objective,
                "Владелец Objective квартала": ownerNameById(directory, q.objectiveOwner),
              });
            });
          } else {
            rows.push({
              ...base,
              "Цель-ориентир волны": w.targetText,
              "Objective на 6 мес.": "", "Владелец Objective 6 мес.": "",
              "Квартал (1-3 / 4-6)": "", "KR-веха квартала": "", "Инициатива на 3 мес.": "",
              "Objective квартала": "", "Владелец Objective квартала": "",
            });
          }
        });
      });
    });
  });
  return rows;
}

function buildTaskRows(trackName, data, directory) {
  const rows = [];
  (data.ultimateObjectives || []).forEach((uo, ui) => {
    (uo.krOutcomes || []).forEach((o, oi) => {
      (o.krOutputs || []).forEach((ko, ki) => {
        (ko.waves || []).forEach((w, wi) => {
          if (w.status !== "decomposed") return;
          w.quarters.forEach((q, qi) => {
            q.monthlyKRs.forEach((m, mi) => {
              const monthNo = qi * 3 + mi + 1;
              const base = {
                "Трек": trackName, "Главная цель №": ui + 1, "KR Outcome №": oi + 1,
                "KR Output №": ki + 1, "Волна №": wi + 1, "Месяц №": monthNo,
              };
              const filledTasks = m.tasks.filter((t) => t.text);
              if (!m.text && filledTasks.length === 0) return;
              if (filledTasks.length === 0) {
                rows.push({ ...base, "KR месяца": m.text, "Задача": "", "Ответственный": "" });
              } else {
                filledTasks.forEach((t) => {
                  rows.push({ ...base, "KR месяца": m.text, "Задача": t.text, "Ответственный": ownerNameById(directory, t.ownerId) });
                });
              }
            });
          });
        });
      });
    });
  });
  return rows;
}

const STATUS_TEXT_TO_CODE = { "В плане": "G", "Откл. до 10%": "Y", "Откл. >10%": "R", "—": "N" };

function buildProjectRows(projects) {
  return projects.map((p) => ({
    "Проект": p.name, "Тип (RUN/Change)": (p.projectType || "run") === "change" ? "Change" : "RUN",
    "Дивизион": p.division, "Функция": p.function,
    "Статус": (DEFAULT_PROJECT_STATUSES.find((s) => s.id === p.statusId) || {}).name || p.status || "",
    "Дата начала": p.startDate || "", "Ожидаемая дата конца": p.endDate || "",
    "Месяц запуска": p.launchMonth || "", "Описание": p.description,
    "Треки (через ;)": p.trackIds.map((id) => (TRACKS.find((t) => t.id === id) || {}).name || id).join("; "),
    "Активность мес.1-12 (через запятую)": p.activity.join(","),
  }));
}
function buildProjectKrRows(projects) {
  const rows = [];
  projects.forEach((p) => p.krs.forEach((k) => { if (k.text) rows.push({ "Проект": p.name, "KR": k.text }); }));
  return rows;
}
function buildDirectoryRows(directory) {
  return directory.map((d) => ({ "Название": d.name, "Тип": TYPE_LABEL[d.type] || d.type }));
}
function buildLinkRows(links) {
  return links.map((l) => {
    const cfg = LINK_TYPES.find((t) => t.key === l.type);
    return {
      "ID": l.id, "Тип": cfg ? cfg.label : l.type,
      "Трек А": l.aTrackName, "Узел А": l.aLabel, "ID узла А": l.aItemId,
      "Трек Б": l.bTrackName, "Узел Б": l.bLabel, "ID узла Б": l.bItemId,
    };
  });
}

function buildInstructionsRows() {
  return [
    { "Раздел": "Общее", "Пояснение": "Не переименовывайте листы и заголовки колонок — импорт ищет их по названию." },
    { "Раздел": "Общее", "Пояснение": "Пустая колонка = пусто в приложении. Не нужно ничего вписывать, если поле не заполнено." },
    { "Раздел": "1. Треки", "Пояснение": "Ровно 5 строк — по одной на трек. Названия треков менять нельзя, они должны совпадать с приложением. Здесь — только миссия." },
    { "Раздел": "2. OKR - Структура", "Пояснение": "У трека может быть до 2 главных целей (Ultimate Objective). «Главная цель №» — 1 или 2, у каждой — свой текст, владелец и своя нумерация KR Outcome (начинается заново с 1 для каждой цели, максимум 4 KR Outcome на одну главную цель)." },
    { "Раздел": "2. OKR - Структура", "Пояснение": "Одна строка = один квартал (или одна строка на волну, если волна ещё НЕ разложена)." },
    { "Раздел": "2. OKR - Структура", "Пояснение": "«KR Outcome №» и «KR Output №» — нумеруйте с 1 подряд внутри своей главной цели, без пропусков. Новые добавляйте В КОНЕЦ списка — так сохранятся уже созданные связи 🔗." },
    { "Раздел": "2. OKR - Структура", "Пояснение": "«Волна №» всегда 1, 2 или 3 (1 = 1–6 мес, 2 = 7–12 мес, 3 = 13–18 мес из 18)." },
    { "Раздел": "2. OKR - Структура", "Пояснение": "Если волна ещё НЕ разложена — заполните только «Цель-ориентир волны» и оставьте «Квартал» пустым (1 строка на волну)." },
    { "Раздел": "2. OKR - Структура", "Пояснение": "Если волна разложена — сделайте 2 строки (Квартал = 1-3 и Квартал = 4-6), «Цель-ориентир волны» не нужен." },
    { "Раздел": "3. OKR - Задачи", "Пояснение": "Одна строка = одна задача. «Месяц №» — абсолютный номер месяца волны, от 1 до 6 (не 1–3 внутри квартала)." },
    { "Раздел": "3. OKR - Задачи", "Пояснение": "«KR месяца» можно повторять на каждой строке задачи этого месяца — берётся первое непустое значение." },
    { "Раздел": "3. OKR - Задачи", "Пояснение": "Если хотите указать только KR месяца без конкретных задач — сделайте одну строку с пустой «Задача»." },
    { "Раздел": "Владельцы", "Пояснение": "Впишите имя/название прямо текстом (например «Отдел маркетинга»). Если такого нет в справочнике — он будет создан автоматически при импорте." },
    { "Раздел": "4-5. Проекты", "Пояснение": "«Треки (через ;)» — названия треков через точку с запятой, не больше 3. KR проекта — на отдельном листе «5. KR проектов», по одной строке на KR, со ссылкой на «Проект» по названию." },
    { "Раздел": "6. Справочник", "Пояснение": "Компания / Отдел / Функция — список организаций, которые можно назначать владельцами и ответственными." },
    { "Раздел": "7. Связи", "Пояснение": "Служебный лист (связи между целями/проектами с типом конфликт/дополнение/ожидание/синхронизация). Создавайте и редактируйте связи в самом приложении — этот лист лучше не заполнять вручную." },
  ];
}

// ---- Import: rebuild nested data from the sheets, preserving IDs by position so existing 🔗-links survive edits ----

function rebuildTrackFromSheets(trackName, trackRow, structureRows, taskRows, existing, directory, newDirEntries) {
  const data = createTrackData();
  if (trackRow) {
    data.mission = trackRow["Миссия"] || "";
  }

  const myStruct = structureRows.filter((r) => normKey(r["Трек"]) === normKey(trackName));
  const myTasks = taskRows.filter((r) => normKey(r["Трек"]) === normKey(trackName));

  const uoGroups = groupByNum(myStruct, "Главная цель №");
  const uoNos = Array.from(uoGroups.keys()).sort((a, b) => a - b).slice(0, MAX_ULTIMATE_OBJECTIVES);
  const finalUoNos = uoNos.length ? uoNos : [1];

  data.ultimateObjectives = finalUoNos.map((uoNo, ui) => {
    const uoRows = uoGroups.get(uoNo) || [];
    const existingUo = (existing.ultimateObjectives || [])[ui];
    const uoTasks = myTasks.filter((r) => (Number(r["Главная цель №"]) || 1) === uoNo);

    const outcomeGroups = groupByNum(uoRows, "KR Outcome №");
    const outcomeNos = Array.from(outcomeGroups.keys()).sort((a, b) => a - b).slice(0, MAX_KR_OUTCOMES);

    return {
      id: existingUo ? existingUo.id : newId(),
      label: existingUo ? existingUo.label : `Главная цель ${ui + 1}`,
      text: firstNonEmpty(uoRows, "Ultimate Objective"),
      owner: resolveOwnerId(firstNonEmpty(uoRows, "Владелец Ultimate Objective"), directory, newDirEntries),
      krOutcomes: outcomeNos.map((oNo, oi) => {
        const oRows = outcomeGroups.get(oNo);
        const existingOutcome = existingUo && existingUo.krOutcomes[oi];
        const outputGroups = groupByNum(oRows, "KR Output №");
        const outputNos = Array.from(outputGroups.keys()).sort((a, b) => a - b);

        return {
          id: existingOutcome ? existingOutcome.id : newId(),
          label: existingOutcome ? existingOutcome.label : `KR Outcome ${oi + 1}`,
          text: firstNonEmpty(oRows, "KR Outcome"),
          metric: firstNonEmpty(oRows, "Метрика"),
          initiativeText: firstNonEmpty(oRows, "Стратегическая инициатива"),
          objectiveText: firstNonEmpty(oRows, "Objective инициативы"),
          objectiveOwner: resolveOwnerId(firstNonEmpty(oRows, "Владелец Objective инициативы"), directory, newDirEntries),
          krOutputs: outputNos.map((koNo, ki) => {
            const koRows = outputGroups.get(koNo);
            const existingOutput = existingOutcome && existingOutcome.krOutputs[ki];
            const waveGroups = groupByNum(koRows, "Волна №");

            const waves = [0, 1, 2].map((wi) => {
              const wRows = waveGroups.get(wi + 1) || [];
              const existingWave = existingOutput && existingOutput.waves[wi];
              const periodLabel = wi === 0 ? "1–6 мес из 18" : wi === 1 ? "7–12 мес из 18" : "13–18 мес из 18";
              const hasQuarterData = wRows.some((r) => String(r["Квартал (1-3 / 4-6)"] || "").trim() !== "");
              const qLabels = ["1–3 мес", "4–6 мес"];
              const monthLabelsFor = [["Мес. 1", "Мес. 2", "Мес. 3"], ["Мес. 4", "Мес. 5", "Мес. 6"]];

              const quarters = [0, 1].map((qi) => {
                const qKey = qi === 0 ? "1-3" : "4-6";
                const qRow = wRows.find((r) => normKey(r["Квартал (1-3 / 4-6)"]) === qKey);
                const existingQuarter = existingWave && existingWave.quarters[qi];

                const monthlyKRs = [0, 1, 2].map((mi) => {
                  const absMonth = qi * 3 + mi + 1;
                  const monthTaskRows = uoTasks.filter(
                    (r) => Number(r["KR Outcome №"]) === oNo && Number(r["KR Output №"]) === koNo &&
                      Number(r["Волна №"]) === wi + 1 && Number(r["Месяц №"]) === absMonth
                  );
                  const existingMonth = existingQuarter && existingQuarter.monthlyKRs[mi];
                  const taskEntries = monthTaskRows.filter((r) => String(r["Задача"] || "").trim() !== "");
                  return {
                    id: existingMonth ? existingMonth.id : newId(),
                    label: monthLabelsFor[qi][mi],
                    text: firstNonEmpty(monthTaskRows, "KR месяца"),
                    tasks: taskEntries.map((r, ti) => {
                      const existingTask = existingMonth && existingMonth.tasks[ti];
                      return {
                        id: existingTask ? existingTask.id : newId(),
                        text: r["Задача"] || "",
                        ownerId: resolveOwnerId(r["Ответственный"], directory, newDirEntries),
                      };
                    }),
                  };
                });

                return {
                  id: existingQuarter ? existingQuarter.id : newId(),
                  label: qLabels[qi],
                  milestone: qRow ? (qRow["KR-веха квартала"] || "") : "",
                  initiative3mo: qRow ? (qRow["Инициатива на 3 мес."] || "") : "",
                  objective: qRow ? (qRow["Objective квартала"] || "") : "",
                  objectiveOwner: resolveOwnerId(qRow ? qRow["Владелец Objective квартала"] : "", directory, newDirEntries),
                  monthlyKRs,
                };
              });

              return {
                id: existingWave ? existingWave.id : newId(),
                periodLabel,
                status: hasQuarterData ? "decomposed" : "target",
                targetText: firstNonEmpty(wRows, "Цель-ориентир волны"),
                objective6mo: firstNonEmpty(wRows, "Objective на 6 мес."),
                objective6moOwner: resolveOwnerId(firstNonEmpty(wRows, "Владелец Objective 6 мес."), directory, newDirEntries),
                quarters,
              };
            });

            return {
              id: existingOutput ? existingOutput.id : newId(),
              label: existingOutput ? existingOutput.label : `KR_${ki + 1} OUTPUT`,
              text: firstNonEmpty(koRows, "KR Output"),
              waves,
            };
          }),
        };
      }),
    };
  });

  return data;
}

function rebuildProjectsFromSheets(projectRows, krRows, existingProjects) {
  const existingByName = new Map((existingProjects || []).map((p) => [normKey(p.name), p]));
  return projectRows
    .filter((r) => String(r["Проект"] || "").trim())
    .map((r) => {
      const name = String(r["Проект"]).trim();
      const existing = existingByName.get(normKey(name));
      const trackNames = String(r["Треки (через ;)"] || "").split(/[;,]/).map((s) => s.trim()).filter(Boolean);
      const trackIds = trackNames
        .map((n) => { const t = TRACKS.find((t) => normKey(t.name) === normKey(n) || t.id === n); return t ? t.id : null; })
        .filter(Boolean)
        .slice(0, MAX_PROJECT_TRACKS);
      const activity = String(r["Активность мес.1-12 (через запятую)"] || "").split(",").map((x) => parseInt(x, 10) || 0);
      while (activity.length < 12) activity.push(0);
      return {
        id: existing ? existing.id : newId(),
        name, division: r["Дивизион"] || "", function: r["Функция"] || "",
        statusId: (DEFAULT_PROJECT_STATUSES.find((s) => normKey(s.name) === normKey(r["Статус"])) || {}).id
          || (existing ? existing.statusId : "not-started"),
        startDate: r["Дата начала"] || (existing ? existing.startDate : null),
        endDate: r["Ожидаемая дата конца"] || (existing ? existing.endDate : null),
        projectType: normKey(r["Тип (RUN/Change)"]) === "change" ? "change" : "run",
        launchMonth: r["Месяц запуска"] ? String(r["Месяц запуска"]) : null,
        description: r["Описание"] || "", trackIds, activity: activity.slice(0, 12),
        krs: krRows.filter((k) => normKey(k["Проект"]) === normKey(name)).map((k, ki) => {
          const existingKr = existing && existing.krs[ki];
          return { id: existingKr ? existingKr.id : newId(), text: k["KR"] || "" };
        }),
      };
    });
}

function rebuildDirectoryFromSheet(dirRows, existing) {
  const TYPE_FROM_LABEL = { "Компания": "company", "Отдел": "department", "Функция": "function" };
  const existingByName = new Map((existing || []).map((d) => [normKey(d.name), d]));
  return dirRows
    .filter((r) => String(r["Название"] || "").trim())
    .map((r) => {
      const name = String(r["Название"]).trim();
      const existingEntry = existingByName.get(normKey(name));
      return { id: existingEntry ? existingEntry.id : newId(), name, type: TYPE_FROM_LABEL[String(r["Тип"] || "").trim()] || "function" };
    });
}

function rebuildLinksFromSheet(rows) {
  const LABEL_TO_KEY = Object.fromEntries(LINK_TYPES.map((t) => [t.label, t.key]));
  return rows
    .filter((r) => r["ID узла А"] && r["ID узла Б"])
    .map((r) => ({
      id: String(r["ID"] || newId()),
      type: LABEL_TO_KEY[r["Тип"]] || "sync",
      aTrackId: (TRACKS.find((t) => normKey(t.name) === normKey(r["Трек А"])) || {}).id || "",
      aTrackName: r["Трек А"] || "", aItemId: String(r["ID узла А"]), aLabel: r["Узел А"] || "",
      bTrackId: (TRACKS.find((t) => normKey(t.name) === normKey(r["Трек Б"])) || {}).id || "",
      bTrackName: r["Трек Б"] || "", bItemId: String(r["ID узла Б"]), bLabel: r["Узел Б"] || "",
    }));
}


const UpdateCtx = createContext(() => {});
const DirectoryCtx = createContext([]);
function useUpdate() { return useContext(UpdateCtx); }
function useDirectoryList() { return useContext(DirectoryCtx); }

function countFilled(track) {
  let total = 1, filled = 0;
  if (track.mission) filled++;
  (track.ultimateObjectives || []).forEach((uo) => {
    total += 2;
    if (uo.text) filled++;
    if (uo.owner) filled++;
    uo.krOutcomes.forEach((o) => {
      total += 5;
      if (o.text) filled++;
      if (o.metric) filled++;
      if (o.initiativeText) filled++;
      if (o.objectiveText) filled++;
      if (o.objectiveOwner) filled++;
      o.krOutputs.forEach((ko) => {
        total += 1;
        if (ko.text) filled++;
        ko.waves.forEach((w) => {
          if (w.status === "target") {
            total += 1;
            if (w.targetText) filled++;
          } else {
            total += 2;
            if (w.objective6mo) filled++;
            if (w.objective6moOwner) filled++;
            w.quarters.forEach((q) => {
              total += 4;
              if (q.milestone) filled++;
              if (q.initiative3mo) filled++;
              if (q.objective) filled++;
              if (q.objectiveOwner) filled++;
              q.monthlyKRs.forEach((m) => {
                total += 1;
                if (m.text) filled++;
                m.tasks.forEach((t) => {
                  total += 2;
                  if (t.text) filled++;
                  if (t.ownerId) filled++;
                });
              });
            });
          }
        });
      });
    });
  });
  return { filled, total };
}

function monthHasContent(m) {
  return !!m.text || m.tasks.some((t) => !!t.text);
}
function quarterHasContent(q) {
  return !!q.milestone || !!q.initiative3mo || !!q.objective || q.monthlyKRs.some(monthHasContent);
}
function waveHasContent(w) {
  return w.status === "decomposed" && (!!w.objective6mo || w.quarters.some(quarterHasContent));
}
function outputHasContent(ko) {
  return !!ko.text || ko.waves.some(waveHasContent);
}
function outcomeHasContent(o) {
  return !!o.text || !!o.metric || !!o.initiativeText || !!o.objectiveText || o.krOutputs.some(outputHasContent);
}
function ultimateObjectiveHasContent(uo) {
  return !!uo.text || uo.krOutcomes.some(outcomeHasContent);
}
function trackHasContent(t) {
  return !!t.mission || (t.ultimateObjectives || []).some(ultimateObjectiveHasContent);
}
function ownerName(directory, id) {
  if (!id) return null;
  const found = directory.find((d) => d.id === id);
  return found ? found.name : null;
}

const LINK_TYPES = [
  { key: "conflict", label: "Конфликт", color: "#dc2626", icon: "⚔" },
  { key: "complement", label: "Дополнение", color: "#15803d", icon: "➕" },
  { key: "waiting", label: "Ожидание", color: "#a16207", icon: "⏳" },
  { key: "sync", label: "Синхронизация", color: "#0369a1", icon: "⇄" },
];

function Truncated({ text, limit = 80 }) {
  const t = useT();
  const [open, setOpen] = useState(false);
  if (!text) return null;
  const isLong = text.length > limit;
  if (!isLong) return <span className="whitespace-pre-wrap">{text}</span>;
  return (
    <span className="whitespace-pre-wrap">
      {open ? text : truncate(text, limit)}{" "}
      <button
        onClick={(e) => { e.stopPropagation(); setOpen((o) => !o); }}
        className="text-sky-600 hover:text-sky-800 t11 font-medium shrink-0"
      >
        {open ? `· ${t("свернуть")}` : `· ${t("ещё")}`}
      </button>
    </span>
  );
}

function truncate(s, n = 64) {
  if (!s) return s;
  return s.length > n ? s.slice(0, n) + "…" : s;
}

const LinkCtx = createContext(null);
function useLinkCtx() { return useContext(LinkCtx); }

function Linkable({ link, children }) {
  const t = useT();
  const ctx = useLinkCtx();
  const authCtx = useAuthCtx();
  const canEditLinks = canEditModule(authCtx.permissions, "combined_tree_links");
  if (!ctx || !link) return children;
  const { pending, pick, links, removeLink, registerRef } = ctx;
  const isPending = pending && pending.itemId === link.itemId;
  const myLinks = links.filter((l) => l.aItemId === link.itemId || l.bItemId === link.itemId);
  return (
    <div
      ref={(el) => registerRef && registerRef(link.itemId, el)}
      className={`relative min-w-0 ${isPending ? "ring-2 ring-amber-400 rounded-xl" : ""}`}
    >
      {children}
      {canEditLinks && (
        <button
          onClick={(e) => { e.stopPropagation(); pick(link); }}
          title={t("Связать с другой целью, KR, задачей или проектом")}
          className={`absolute -top-2 -right-2 w-5 h-5 rounded-full bg-white border t10 flex items-center justify-center shadow-sm z-20 hover:bg-amber-100 hover:text-amber-700 hover:border-amber-300 ${
            isPending ? "border-amber-400 text-amber-700 opacity-100" : "border-neutral-300 text-neutral-400 opacity-60"
          }`}
        >
          🔗
        </button>
      )}
      {myLinks.length > 0 && (
        <div className="flex flex-wrap gap-1 mt-1 mb-1">
          {myLinks.map((l) => {
            const mine = l.aItemId === link.itemId;
            const otherTrack = mine ? l.bTrackName : l.aTrackName;
            const otherLabel = mine ? l.bLabel : l.aLabel;
            const cfg = LINK_TYPES.find((t) => t.key === l.type);
            return (
              <span
                key={l.id}
                title={otherLabel}
                className="inline-flex items-center gap-1 t10 px-1.5 py-0.5 rounded-full border max-w-full"
                style={{ background: cfg.color + "14", color: cfg.color, borderColor: cfg.color + "55" }}
              >
                <span className="shrink-0">{cfg.icon} {t(cfg.label)} ·</span>
                <span className="truncate">{otherTrack}</span>
                {canEditLinks && (
                  <button onClick={() => removeLink(l.id)} className="ml-0.5 opacity-60 hover:opacity-100 shrink-0">×</button>
                )}
              </span>
            );
          })}
        </div>
      )}
    </div>
  );
}


function Field({ value, onChange, placeholder, multiline, small, frame }) {
  const lang = useLang();
  const preview = useDataT(value);
  const cls =
    "w-full bg-transparent outline-none placeholder-neutral-400 focus:placeholder-neutral-300 text-sm";
  const showPreview = lang === "en" && value && value.trim() && preview && preview !== value;
  const input = multiline ? (
    <textarea
      className={cls + " resize-none"}
      rows={2}
      value={value}
      placeholder={placeholder}
      title={value || undefined}
      onChange={(e) => onChange(e.target.value)}
    />
  ) : (
    <input
      className={cls}
      value={value}
      placeholder={placeholder}
      title={value || undefined}
      onChange={(e) => onChange(e.target.value)}
    />
  );
  return (
    <div className={frame ? `rounded-md border ${frame} px-1.5 py-0.5` : undefined}>
      {input}
      {showPreview && <div className="text-xs text-sky-600 italic mt-0.5">🌐 {preview}</div>}
    </div>
  );
}

function OwnerSelect({ value, onChange, compact }) {
  const t = useT();
  const directory = useDirectoryList();
  const groups = { company: [], department: [], function: [] };
  directory.forEach((d) => { if (groups[d.type]) groups[d.type].push(d); });
  const typeLabel = { company: t("Компания"), department: t("Отделы"), function: t("Функции") };
  return (
    <select
      value={value || ""}
      onChange={(e) => onChange(e.target.value)}
      className={
        "bg-neutral-50 border border-neutral-200 rounded-md outline-none text-neutral-600 " +
        (compact ? "t11 py-0.5 px-1 w-32 shrink-0" : "text-xs py-1 px-1.5")
      }
    >
      <option value="">{t("Без ответственного")}</option>
      {Object.entries(groups).map(([type, items]) =>
        items.length ? (
          <optgroup key={type} label={typeLabel[type]}>
            {items.map((it) => (
              <option key={it.id} value={it.id}>{it.name}</option>
            ))}
          </optgroup>
        ) : null
      )}
    </select>
  );
}

function TrackerEditor({ tracker, onChange, horizonMonths }) {
  const tt = useT();
  const authCtx = useAuthCtx();
  const t = tracker;
  const { fact, pct } = computeTracking(t);
  const update = (field, value) => onChange({ ...t, [field]: value });
  const addWeek = () => onChange({ ...t, weeks: [...t.weeks, createTrackerWeek(t.weeks.length + 1)] });
  const updateWeek = (id, field, value) => onChange({ ...t, weeks: t.weeks.map((w) => (w.id === id ? { ...w, [field]: value } : w)) });
  const removeWeek = (id) => onChange({ ...t, weeks: t.weeks.filter((w) => w.id !== id) });
  const status = trackingStatus(pct);

  const startDate = t.startDate || todayISO();
  const autoDeadline = addMonthsISO(startDate, horizonMonths || 1);
  const deadline = t.deadlineOverride || autoDeadline;
  const daysLeft = daysBetweenISO(todayISO(), deadline);
  const urgency = deadlineUrgency(daysLeft);
  const [editingDeadline, setEditingDeadline] = useState(false);

  return (
    <div className="rounded-lg border border-sky-200 bg-sky-50 p-2.5 space-y-2">
      <div className="flex items-center gap-1.5 text-xs text-sky-700 font-medium">
        <Activity size={12} /> {tt("Трекинг KR (по неделям)")}
      </div>

      <div className="rounded-md bg-white border border-sky-200 px-2 py-1.5 flex items-center gap-1.5 flex-wrap t11">
        <CalendarRange size={12} className="text-sky-400 shrink-0" />
        <span className="text-neutral-500">{tt("Поставлено")} {formatDateRu(startDate)}</span>
        <span className="text-neutral-300">·</span>
        <span className="text-neutral-500">{tt("Срок")} {formatDateRu(deadline)}</span>
        <span className="text-neutral-300">·</span>
        <span style={{ color: urgency.color }} className="font-medium">
          {daysLeft < 0 ? `${tt("просрочено на")} ${Math.abs(daysLeft)} ${tt("дн.")}` : `${tt("осталось")} ${daysLeft} ${tt("дн.")}`} · {tt(urgency.label)}
        </span>
        {authCtx.isAdmin && !editingDeadline && (
          <button onClick={() => setEditingDeadline(true)} className="ml-auto text-sky-500 hover:text-sky-700 shrink-0">{tt("изменить срок")}</button>
        )}
        {authCtx.isAdmin && editingDeadline && (
          <span className="flex items-center gap-1 ml-auto shrink-0">
            <input
              type="date" value={deadline}
              onChange={(e) => update("deadlineOverride", e.target.value)}
              className="text-xs border border-sky-200 rounded px-1 py-0.5"
            />
            {t.deadlineOverride && (
              <button onClick={() => update("deadlineOverride", null)} className="text-neutral-400 hover:text-red-500">{tt("сброс")}</button>
            )}
            <button onClick={() => setEditingDeadline(false)} className="text-sky-600 hover:text-sky-800">{tt("готово")}</button>
          </span>
        )}
      </div>

      <div className="grid grid-cols-3 gap-1.5">
        <label className="t11 text-sky-600 space-y-0.5">
          <span>{tt("Тип метрики")}</span>
          <select
            value={t.measureType}
            onChange={(e) => update("measureType", e.target.value)}
            className="w-full text-xs border border-sky-200 rounded px-1 py-0.5 bg-white"
          >
            <option value="cumulative">{tt("Накопительный (прирост)")}</option>
            <option value="level">{tt("Уровень (срез)")}</option>
          </select>
        </label>
        <label className="t11 text-sky-600 space-y-0.5">
          <span>{tt("Исходное")}</span>
          <input
            type="number" value={t.baseline}
            onChange={(e) => update("baseline", parseFloat(e.target.value) || 0)}
            className="w-full text-xs border border-sky-200 rounded px-1 py-0.5 bg-white"
          />
        </label>
        <label className="t11 text-sky-600 space-y-0.5">
          <span>{tt("Целевое")}</span>
          <input
            type="number" value={t.target}
            onChange={(e) => update("target", parseFloat(e.target.value) || 0)}
            className="w-full text-xs border border-sky-200 rounded px-1 py-0.5 bg-white"
          />
        </label>
      </div>

      <div className="space-y-1">
        {t.weeks.map((w) => (
          <div key={w.id} className="flex items-center gap-1.5">
            <input
              value={w.label} onChange={(e) => updateWeek(w.id, "label", e.target.value)}
              className="w-16 shrink-0 t11 text-sky-600 bg-transparent border-b border-dashed border-sky-200 outline-none"
            />
            <input
              type="number" value={w.value}
              onChange={(e) => updateWeek(w.id, "value", parseFloat(e.target.value) || 0)}
              placeholder={t.measureType === "cumulative" ? tt("прирост за неделю") : tt("значение на конец недели")}
              className="flex-1 text-xs border border-sky-200 rounded px-1 py-0.5 bg-white"
            />
            <select
              value={w.confidence} onChange={(e) => updateWeek(w.id, "confidence", parseInt(e.target.value, 10))}
              title={tt("Уверенность 1-10")} className="text-xs border border-sky-200 rounded px-1 py-0.5 bg-white w-12"
            >
              {Array.from({ length: 10 }, (_, k) => k + 1).map((n) => (
                <option key={n} value={n}>{n}</option>
              ))}
            </select>
            <button onClick={() => removeWeek(w.id)} className="text-sky-300 hover:text-red-500 shrink-0"><Trash2 size={11} /></button>
          </div>
        ))}
        <button onClick={addWeek} className="flex items-center gap-1 t11 text-sky-600 hover:text-sky-800">
          <Plus size={11} /> {tt("добавить неделю")}
        </button>
      </div>

      <div className="flex items-center gap-2 pt-1.5 border-t border-sky-200">
        <div className="flex-1 h-1.5 bg-sky-100 rounded-full overflow-hidden">
          <div className="h-full rounded-full" style={{ width: `${Math.min(100, Math.max(0, pct * 100))}%`, background: status.color }} />
        </div>
        <span className="text-xs font-medium w-36 text-right" style={{ color: status.color }}>
          {tt("Факт")} {fact.toFixed(1)} · {(pct * 100).toFixed(0)}% · {tt(status.label)}
        </span>
      </div>
    </div>
  );
}

function TrackingSection({ tracking, onChange, horizonMonths }) {
  const t = useT();
  if (!tracking) {
    return (
      <button onClick={() => onChange(createTracker())} className="flex items-center gap-1 t11 text-sky-600 hover:text-sky-800">
        <Activity size={11} /> {t("добавить трекинг прогресса")}
      </button>
    );
  }
  return (
    <div className="space-y-1">
      <TrackerEditor tracker={tracking} onChange={onChange} horizonMonths={horizonMonths} />
      <button onClick={() => onChange(null)} className="t11 text-neutral-400 hover:text-red-500">{t("убрать трекинг")}</button>
    </div>
  );
}

function ObjectiveField({ label, labelColor, textPath, ownerPath, text, ownerId, placeholder }) {
  const t = useT();
  const update = useUpdate();
  return (
    <div>
      <div className={`text-xs mb-0.5 ${labelColor}`}>{label}</div>
      <Field value={text} placeholder={placeholder} onChange={(v) => update(textPath, v)} />
      <div className="flex items-center gap-1.5 mt-1">
        <span className="t11 text-neutral-400">{t("Владелец Objective:")}</span>
        <OwnerSelect value={ownerId} onChange={(v) => update(ownerPath, v)} compact />
      </div>
    </div>
  );
}

function Collapsible({ title, subtitle, icon, tone, accentColor, defaultOpen = true, right, children }) {
  const [open, setOpen] = useState(defaultOpen);
  const outerStyle = accentColor ? { borderColor: accentColor } : { borderColor: tone.border, borderStyle: tone.dashed ? "dashed" : "solid" };
  const headerStyle = accentColor ? { background: accentColor + "10", color: accentColor } : { background: tone.bg, color: tone.text };
  return (
    <div className="rounded-xl border overflow-hidden bg-white" style={outerStyle}>
      <div className="flex items-start gap-2 px-3 py-2 min-w-0" style={headerStyle}>
        <button onClick={() => setOpen((o) => !o)} className="flex items-start gap-2 flex-1 min-w-0 text-left">
          {open ? <ChevronDown size={15} className="shrink-0 mt-0.5" /> : <ChevronRight size={15} className="shrink-0 mt-0.5" />}
          <span className="shrink-0 mt-0.5">{icon}</span>
          <span className="min-w-0 flex-1 flex flex-wrap items-baseline gap-x-1">
            <span className="text-xs font-medium tracking-wide uppercase break-words">{title}</span>
            {subtitle && <span className="text-xs opacity-70 break-words">· {subtitle}</span>}
          </span>
        </button>
        {right}
      </div>
      {open && <div className="px-3 pb-3 bg-white">{children}</div>}
    </div>
  );
}

const TRACK_TONE_TEXT = {
  "health-os": "text-emerald-700", "growth-model": "text-amber-700", "prime-growth": "text-orange-700",
  "coral-evo": "text-violet-700", "it-model": "text-red-500",
};

// Palette per the reference: bold solid header for the "big" strategic levels (18/6-month
// horizons), quiet neutral for the "detail" levels nested one horizon down (Output/Milestone),
// and each level keeps ONE color everywhere it appears — track editor cards, combined tree,
// and the Gantt. Collapsible only paints the header strip with this — card body stays plain
// white so nested fields underneath stay fully readable regardless of how bold the tone is.
// Exact hex values (not Tailwind steps) so this matches the reference the team already knows.
const TONES = {
  mission: { bg: "#375623", border: "#375623", text: "#FFFFFF" },
  outcome: { bg: "#F4B6B6", border: "#F4B6B6", text: "#262626" },
  initiative: { bg: "#375623", border: "#375623", text: "#FFFFFF" },
  output: { bg: "#E7E6E6", border: "#E7E6E6", text: "#262626" },
  wave: { bg: "#1F3864", border: "#1F3864", text: "#FFFFFF" },
  waveTarget: { bg: "#FAFAFA", border: "#E5E5E5", text: "#737373", dashed: true },
  quarter: { bg: "#E7E6E6", border: "#E7E6E6", text: "#262626" },
  month: { bg: "#70AD47", border: "#70AD47", text: "#FFFFFF" },
  task: { bg: "#FFF2CC", border: "#FDE68A", text: "#262626" },
  project: { bg: "#EEF2FF", border: "#C7D2FE", text: "#3730A3" },
};

// Same tones as the track editor, indexed by Gantt row depth (0=Главная цель … 6=Задача) —
// so the hierarchy reads the same way in the Gantt as it does in the editor.
const GANTT_DEPTH_TONE = [TONES.mission, TONES.outcome, TONES.output, TONES.wave, TONES.quarter, TONES.month, TONES.task];

function OwnerPickerModal({ value, onChange, onClose, title }) {
  const t = useT();
  const directory = useDirectoryList();
  const groups = { company: [], department: [], function: [] };
  directory.forEach((d) => { if (groups[d.type]) groups[d.type].push(d); });
  const typeLabel = { company: t("Компания"), department: t("Отделы"), function: t("Функции") };
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white rounded-2xl p-4 w-full max-w-xs space-y-1.5 max-h-[80vh] overflow-auto" onClick={(e) => e.stopPropagation()}>
        <div className="text-sm font-semibold text-neutral-800 mb-1.5">{title || t("Ответственный")}</div>
        <button
          onClick={() => onChange("")}
          className={
            "w-full text-left text-xs rounded-lg px-2.5 py-1.5 border " +
            (!value ? "border-neutral-800 bg-neutral-50 font-medium" : "border-neutral-200 text-neutral-500 hover:bg-neutral-50")
          }
        >
          {t("Без ответственного")}
        </button>
        {Object.entries(groups).map(([type, items]) =>
          items.length ? (
            <div key={type} className="pt-1.5 space-y-1">
              <div className="t11 text-neutral-400 uppercase tracking-wide px-0.5">{typeLabel[type]}</div>
              {items.map((it) => (
                <button
                  key={it.id}
                  onClick={() => onChange(it.id)}
                  className={
                    "w-full text-left text-xs rounded-lg px-2.5 py-1.5 border " +
                    (value === it.id ? "border-neutral-800 bg-neutral-50 font-medium" : "border-neutral-200 hover:bg-neutral-50")
                  }
                >
                  {it.name}
                </button>
              ))}
            </div>
          ) : null
        )}
        <button onClick={onClose} className="w-full text-xs text-neutral-400 hover:text-neutral-600 pt-2">
          {t("Отмена")}
        </button>
      </div>
    </div>
  );
}

function OwnerIconButton({ value, onChange, title }) {
  const t = useT();
  const directory = useDirectoryList();
  const [open, setOpen] = useState(false);
  const owner = directory.find((d) => d.id === value);
  return (
    <>
      <button
        type="button"
        onClick={() => setOpen(true)}
        title={owner ? owner.name : t("Без ответственного")}
        className={
          "shrink-0 flex items-center justify-center w-6 h-6 rounded-md border " +
          (owner
            ? "border-neutral-300 text-neutral-500 hover:bg-neutral-50"
            : "border-red-300 text-red-400 hover:bg-red-50")
        }
      >
        <User size={12} />
      </button>
      {open && (
        <OwnerPickerModal
          value={value}
          title={title}
          onChange={(v) => { onChange(v); setOpen(false); }}
          onClose={() => setOpen(false)}
        />
      )}
    </>
  );
}

function TaskRow({ path, task, idx, onRemove }) {
  const t = useT();
  const update = useUpdate();
  const hasOwner = !!task.ownerId;
  const isDup = task.text.startsWith(IMPORT_TASK_DUPLICATE_PREFIX);
  return (
    <div className="flex items-center gap-1.5 pl-6 py-0.5">
      <Circle size={6} className="text-neutral-300 shrink-0" />
      {isDup && (
        <span className="t10 shrink-0 px-1 rounded font-medium" style={{ background: "#FED7AA", color: "#9A3412" }} title={t("Возможный дубль — сравните с похожей задачей рядом и удалите лишнее")}>
          ⚠
        </span>
      )}
      <div className="flex-1 min-w-0">
        <Field
          small
          value={task.text}
          placeholder={`${t("Задача")} ${idx + 1}`}
          onChange={(v) => update([...path, "text"], v)}
          frame={isDup ? "border-orange-300" : hasOwner ? "border-neutral-300" : "border-red-300"}
        />
      </div>
      <OwnerIconButton value={task.ownerId} onChange={(v) => update([...path, "ownerId"], v)} title={t("Ответственный за задачу")} />
      <button onClick={onRemove} className="text-neutral-300 hover:text-red-500 p-0.5 shrink-0">
        <Trash2 size={12} />
      </button>
    </div>
  );
}

function MonthlyKRCard({ path, month }) {
  const t = useT();
  const update = useUpdate();
  const addTask = () => {
    if (month.tasks.length >= MAX_TASKS) return;
    update([...path, "tasks"], [...month.tasks, createTask()]);
  };
  const removeTask = (idx) => {
    update([...path, "tasks"], month.tasks.filter((_, i) => i !== idx));
  };
  return (
    <Collapsible title={month.label} icon={<ListChecks size={13} />} tone={TONES.month} defaultOpen={false}>
      <div className="pl-1 pt-1 space-y-1">
        <Field value={month.text} placeholder={t("KR месяца — что должно быть достигнуто")}
          onChange={(v) => update([...path, "text"], v)} />
        <TrackingSection tracking={month.tracking} onChange={(v) => update([...path, "tracking"], v)} horizonMonths={1} />
        <div className="pt-1 space-y-1">
          {month.tasks.map((tk, i) => (
            <TaskRow key={tk.id} path={[...path, "tasks", i]} task={tk} idx={i} onRemove={() => removeTask(i)} />
          ))}
        </div>
        <div className="pl-6 flex items-center gap-2 pt-0.5">
          <button
            onClick={addTask}
            disabled={month.tasks.length >= MAX_TASKS}
            className="flex items-center gap-1 t11 text-neutral-500 hover:text-neutral-800 disabled:opacity-40 disabled:cursor-not-allowed"
          >
            <Plus size={11} /> {t("задача")}
          </button>
          <span className="t10 text-neutral-300">{month.tasks.length}/{MAX_TASKS}</span>
        </div>
      </div>
    </Collapsible>
  );
}

function QuarterCard({ path, quarter }) {
  const t = useT();
  const update = useUpdate();
  return (
    <Collapsible title={`${t("Квартал")} ${quarter.label}`} subtitle={t("веха + инициатива 3 мес")} icon={<CalendarRange size={13} />} tone={TONES.quarter}>
      <div className="space-y-2">
        <div>
          <div className="text-xs text-rose-600 mb-0.5">{t("KR-веха квартала")}</div>
          <Field value={quarter.milestone} placeholder={t("Что считается достигнутым к концу квартала")}
            onChange={(v) => update([...path, "milestone"], v)} />
        </div>
        <TrackingSection tracking={quarter.tracking} onChange={(v) => update([...path, "tracking"], v)} horizonMonths={3} />
        <div>
          <div className="text-xs text-rose-600 mb-0.5">{t("Инициатива на 3 мес.")}</div>
          <Field value={quarter.initiative3mo} placeholder={t("Операционный проект квартала")}
            onChange={(v) => update([...path, "initiative3mo"], v)} />
        </div>
        <ObjectiveField
          label={t("Objective")} labelColor="text-rose-600"
          textPath={[...path, "objective"]} ownerPath={[...path, "objectiveOwner"]}
          text={quarter.objective} ownerId={quarter.objectiveOwner}
          placeholder={t("Цель инициативы")}
        />
        <div className="pt-1 space-y-1.5">
          {quarter.monthlyKRs.map((m, i) => (
            <MonthlyKRCard key={m.id} path={[...path, "monthlyKRs", i]} month={m} />
          ))}
        </div>
      </div>
    </Collapsible>
  );
}

function WaveCard({ path, wave, waveIndex }) {
  const t = useT();
  const update = useUpdate();
  const isFirst = waveIndex === 0;
  const decomposed = wave.status === "decomposed";
  const tone = decomposed ? TONES.wave : TONES.waveTarget;
  const expand = () => update([...path, "status"], "decomposed");
  const collapse = () => update([...path, "status"], "target");

  return (
    <Collapsible
      title={`${t("Волна")} ${waveIndex + 1}`} subtitle={wave.periodLabel} icon={<Layers size={13} />} tone={tone} defaultOpen={isFirst}
      right={
        decomposed ? (
          !isFirst && <button onClick={collapse} className="t11 text-neutral-400 hover:text-neutral-600 px-2 py-0.5">{t("свернуть в цель")}</button>
        ) : (
          <button onClick={expand} className="flex items-center gap-1 t11 text-emerald-700 bg-emerald-100 hover:bg-emerald-200 px-2 py-1 rounded-md">
            <Sparkles size={11} /> {t("разложить волну")}
          </button>
        )
      }
    >
      {decomposed ? (
        <div className="space-y-2">
          <ObjectiveField
            label={t("Objective на 6 мес.")} labelColor="text-emerald-700"
            textPath={[...path, "objective6mo"]} ownerPath={[...path, "objective6moOwner"]}
            text={wave.objective6mo} ownerId={wave.objective6moOwner}
            placeholder={t("Цель полугодия")}
          />
          <div className="grid gap-1.5 sm:grid-cols-2">
            {wave.quarters.map((q, i) => (
              <QuarterCard key={q.id} path={[...path, "quarters", i]} quarter={q} />
            ))}
          </div>
        </div>
      ) : (
        <div>
          <div className="text-xs text-neutral-400 mb-0.5">{t("Цель-ориентир (без декомпозиции)")}</div>
          <Field value={wave.targetText} placeholder={t("Чего нужно достичь в этом полугодии — детали позже, на контрольной точке")}
            onChange={(v) => update([...path, "targetText"], v)} />
        </div>
      )}
    </Collapsible>
  );
}

function KROutputCard({ path, output }) {
  const t = useT();
  const update = useUpdate();
  return (
    <Collapsible title={output.label} icon={<Target size={13} />} tone={TONES.output}>
      <div className="space-y-2">
        <Field value={output.text} placeholder={t("Формулировка ключевого результата (output) на 18 мес.")}
          onChange={(v) => update([...path, "text"], v)} />
        <TrackingSection tracking={output.tracking} onChange={(v) => update([...path, "tracking"], v)} horizonMonths={18} />
        <div className="space-y-1.5 pt-1">
          {output.waves.map((w, i) => (
            <WaveCard key={w.id} path={[...path, "waves", i]} wave={w} waveIndex={i} />
          ))}
        </div>
      </div>
    </Collapsible>
  );
}

function KROutcomeCard({ path, outcome, onRemove }) {
  const t = useT();
  const update = useUpdate();
  return (
    <Collapsible
      title={outcome.label} icon={<Flag size={14} />} tone={TONES.outcome}
      right={<button onClick={onRemove} className="text-neutral-300 hover:text-red-500 p-1"><Trash2 size={13} /></button>}
    >
      <div className="space-y-2">
        <div>
          <div className="text-xs text-yellow-700 mb-0.5">KR Outcome</div>
          <Field value={outcome.text} placeholder={t("Ключевой результат стратегической цели (18 мес.)")}
            onChange={(v) => update([...path, "text"], v)} />
        </div>
        <div>
          <div className="text-xs text-yellow-700 mb-0.5">{t("Метрика")}</div>
          <Field small value={outcome.metric} placeholder={t("Как измеряется (число, %, статус)")}
            onChange={(v) => update([...path, "metric"], v)} />
        </div>
        <TrackingSection tracking={outcome.tracking} onChange={(v) => update([...path, "tracking"], v)} horizonMonths={18} />

        <Collapsible title={t("Стратегическая инициатива на 18 мес.")} icon={<Layers size={13} />} tone={TONES.initiative}>
          <div className="space-y-2">
            <Field value={outcome.initiativeText} placeholder={t("Программа работ, обеспечивающая KR Outcome")}
              onChange={(v) => update([...path, "initiativeText"], v)} />
            <ObjectiveField
              label={t("Objective")} labelColor="text-blue-700"
              textPath={[...path, "objectiveText"]} ownerPath={[...path, "objectiveOwner"]}
              text={outcome.objectiveText} ownerId={outcome.objectiveOwner}
              placeholder={t("Цель инициативы")}
            />
            <div className="space-y-1.5 pt-1">
              {outcome.krOutputs.map((ko, i) => (
                <KROutputCard key={ko.id} path={[...path, "krOutputs", i]} output={ko} />
              ))}
            </div>
          </div>
        </Collapsible>
      </div>
    </Collapsible>
  );
}

function TrackEditor({ trackId }) {
  const t = useT();
  const authCtx = useAuthCtx();
  const [data, setData] = useState(createTrackData());
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    let cancelled = false;
    setLoaded(false);
    (async () => {
      try {
        const res = await window.storage.get(`okr-track:${trackId}`, false);
        if (!cancelled) setData(res ? migrateTrackData(JSON.parse(res.value)) : createTrackData());
      } catch {
        if (!cancelled) setData(createTrackData());
      }
      if (!cancelled) setLoaded(true);
    })();
    return () => { cancelled = true; };
  }, [trackId]);

  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => {
      window.storage.set(`okr-track:${trackId}`, JSON.stringify(data), false).catch(() => {});
    }, 400);
    return () => clearTimeout(t);
  }, [data, loaded, trackId]);

  const update = useCallback((path, value) => {
    setData((prev) => {
      const draft = JSON.parse(JSON.stringify(prev));
      let node = draft;
      for (let i = 0; i < path.length - 1; i++) node = node[path[i]];
      node[path[path.length - 1]] = value;
      return draft;
    });
  }, []);

  const addOutcome = (uoIdx) => {
    setData((prev) => {
      const draft = JSON.parse(JSON.stringify(prev));
      const uo = draft.ultimateObjectives[uoIdx];
      if (uo.krOutcomes.length >= MAX_KR_OUTCOMES) return prev;
      uo.krOutcomes.push(createKROutcome(uo.krOutcomes.length + 1));
      return draft;
    });
  };
  const removeOutcome = (uoIdx, id) => {
    setData((prev) => {
      const draft = JSON.parse(JSON.stringify(prev));
      const uo = draft.ultimateObjectives[uoIdx];
      uo.krOutcomes = uo.krOutcomes.filter((o) => o.id !== id);
      return draft;
    });
  };
  const addUltimateObjective = () => {
    setData((prev) => {
      if (prev.ultimateObjectives.length >= MAX_ULTIMATE_OBJECTIVES) return prev;
      return { ...prev, ultimateObjectives: [...prev.ultimateObjectives, createUltimateObjective(prev.ultimateObjectives.length + 1)] };
    });
  };
  const removeUltimateObjective = (id) => {
    setData((prev) => {
      if (prev.ultimateObjectives.length <= 1) return prev;
      return { ...prev, ultimateObjectives: prev.ultimateObjectives.filter((u) => u.id !== id) };
    });
  };

  if (!loaded) {
    return <div className="text-sm text-neutral-400 py-12 text-center">{t("Загрузка…")}</div>;
  }

  const { filled, total } = countFilled(data);
  const pct = total ? Math.round((filled / total) * 100) : 0;
  const trackName = (TRACKS.find((t) => t.id === trackId) || {}).name || trackId;
  const trackColor = TRACK_COLORS[trackId] || "#525252";

  return (
    <UpdateCtx.Provider value={update}>
      <div className="space-y-3">
        <div className="flex items-center gap-2 px-1">
          <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: trackColor }} />
          <span className="text-sm font-semibold" style={{ color: trackColor }}>{trackName}</span>
        </div>
        <div className="flex items-center justify-between px-1">
          <div className="text-xs text-neutral-400">{t("Заполнено полей")}: {filled} / {total}</div>
          <div className="flex items-center gap-2 w-40">
            <div className="h-1.5 flex-1 bg-neutral-100 rounded-full overflow-hidden">
              <div className="h-full rounded-full transition-all" style={{ width: `${pct}%`, background: trackColor }} />
            </div>
            <span className="text-xs text-neutral-400 w-8 text-right">{pct}%</span>
          </div>
        </div>

        <div className="grid grid-cols-1 xl:grid-cols-2 gap-3 items-start">
          <Collapsible title={t("Миссия трека")} icon={<CheckCircle2 size={14} />} tone={TONES.mission} accentColor={trackColor}>
            <div className="space-y-2">
              <div>
                <div className={`text-xs mb-0.5 ${TRACK_TONE_TEXT[trackId] || "text-amber-700"}`}>{t("Миссия трека")}</div>
                <Field value={data.mission} placeholder={t("Зачем трек существует")} onChange={(v) => update(["mission"], v)} />
              </div>
              <div className="flex items-center gap-1.5 t11 text-neutral-400">
                <CalendarRange size={11} className="shrink-0" />
                <span>{t("Дата начала трека")}:</span>
                {authCtx.isAdmin ? (
                  <input
                    type="date" value={data.startDate || todayISO()}
                    onChange={(e) => update(["startDate"], e.target.value)}
                    className="text-xs border border-neutral-200 rounded px-1 py-0.5"
                  />
                ) : (
                  <span className="text-neutral-600">{formatDateRu(data.startDate || todayISO())}</span>
                )}
              </div>
            </div>
          </Collapsible>

          <TrackProjects trackId={trackId} trackName={TRACKS.find((t) => t.id === trackId)?.name} />
        </div>

        {data.ultimateObjectives.map((uo, ui) => (
          <Collapsible
            key={uo.id}
            title={`${t("Главная цель")} ${ui + 1} ${t("из")} ${data.ultimateObjectives.length}`}
            icon={<Flag size={14} />} tone={TONES.mission}
            right={
              data.ultimateObjectives.length > 1 && (
                <button onClick={() => removeUltimateObjective(uo.id)} className="text-neutral-300 hover:text-red-500 p-1">
                  <Trash2 size={13} />
                </button>
              )
            }
          >
            <div className="space-y-3">
              <ObjectiveField
                label={t("Ultimate Objective (18 мес.)")} labelColor={TRACK_TONE_TEXT[trackId] || "text-amber-700"}
                textPath={["ultimateObjectives", ui, "text"]} ownerPath={["ultimateObjectives", ui, "owner"]}
                text={uo.text} ownerId={uo.owner}
                placeholder={t("Одна из главных стратегических целей трека на 18 месяцев")}
              />

              <div className="grid grid-cols-1 xl:grid-cols-2 gap-3">
                {uo.krOutcomes.map((o, i) => (
                  <KROutcomeCard
                    key={o.id} path={["ultimateObjectives", ui, "krOutcomes", i]} outcome={o}
                    onRemove={() => removeOutcome(ui, o.id)}
                  />
                ))}
              </div>

              {uo.krOutcomes.length < MAX_KR_OUTCOMES && (
                <button
                  onClick={() => addOutcome(ui)}
                  className="w-full flex items-center justify-center gap-1.5 text-sm text-yellow-700 border border-dashed border-yellow-300 rounded-xl py-2.5 hover:bg-yellow-50"
                >
                  <Plus size={15} /> {t("Добавить KR Outcome")} ({uo.krOutcomes.length}/{MAX_KR_OUTCOMES})
                </button>
              )}
            </div>
          </Collapsible>
        ))}

        {data.ultimateObjectives.length < MAX_ULTIMATE_OBJECTIVES && (
          <button
            onClick={addUltimateObjective}
            className="w-full flex items-center justify-center gap-1.5 text-sm border border-dashed rounded-xl py-2.5"
            style={{ color: trackColor, borderColor: trackColor + "80" }}
          >
            <Plus size={15} /> {t("Добавить главную цель")} ({data.ultimateObjectives.length}/{MAX_ULTIMATE_OBJECTIVES})
          </button>
        )}
      </div>
    </UpdateCtx.Provider>
  );
}

function Owner({ id }) {
  const directory = useDirectoryList();
  const name = ownerName(directory, id);
  if (!name) return null;
  return <span className="t11 text-neutral-400 font-normal">· {name}</span>;
}

function StaticRow({ label, labelColor, text, ownerId, link }) {
  const displayText = useDataT(text);
  if (!text) return null;
  const body = (
    <div>
      <div className={`text-xs mb-0.5 ${labelColor}`}>
        {label} <Owner id={ownerId} />
      </div>
      <div className="text-sm text-neutral-800"><Truncated text={displayText} /></div>
    </div>
  );
  return link ? <Linkable link={link}>{body}</Linkable> : body;
}

function CombinedMonthCompact({ month, trackId, trackName }) {
  const monthText = useDataT(month.text);
  if (!monthHasContent(month)) return null;
  const tasks = month.tasks.filter((t) => !!t.text);
  return (
    <div>
      {month.text && (
        <Linkable link={{ trackId, trackName, itemId: month.id, itemLabel: `KR месяца (${month.label}) — ${truncate(month.text)}` }}>
          <div className="t11 font-medium text-neutral-700 flex items-center gap-1">
            <ListChecks size={10} className="text-neutral-400 shrink-0" />
            <span className="truncate">{month.label} · <Truncated text={monthText} limit={40} /></span>
          </div>
        </Linkable>
      )}
      {tasks.length > 0 && (
        <div className="pl-3.5 space-y-0.5 mt-0.5">
          {tasks.map((t) => {
            const isDup = t.text.startsWith(IMPORT_TASK_DUPLICATE_PREFIX);
            return (
              <Linkable key={t.id} link={{ trackId, trackName, itemId: t.id, itemLabel: `Задача — ${truncate(t.text)}` }}>
                <div className="flex items-center gap-1.5">
                  <Circle size={5} className={isDup ? "text-orange-400 shrink-0" : "text-neutral-300 shrink-0"} />
                  <span className={isDup ? "t11 text-orange-700 truncate flex-1" : "t11 text-neutral-600 truncate flex-1"}>
                    <TranslatedText text={t.text} />
                  </span>
                  <Owner id={t.ownerId} />
                </div>
              </Linkable>
            );
          })}
        </div>
      )}
    </div>
  );
}

function CombinedQuarter({ quarter, trackId, trackName }) {
  const t = useT();
  if (!quarterHasContent(quarter)) return null;
  const months = quarter.monthlyKRs.filter(monthHasContent);
  return (
    <div className="rounded-lg border p-2 space-y-1.5 min-w-0" style={{ borderColor: TONES.quarter.border, background: TONES.quarter.bg }}>
      <div className="t11 font-semibold uppercase tracking-wide flex items-center gap-1" style={{ color: TONES.quarter.text }}>
        <CalendarRange size={11} className="shrink-0" /> {t("Квартал")} {quarter.label}
      </div>
      {quarter.milestone && (
        <Linkable link={{ trackId, trackName, itemId: `${quarter.id}-milestone`, itemLabel: `KR-веха (${quarter.label}) — ${truncate(quarter.milestone)}` }}>
          <div className="text-xs text-neutral-800"><Truncated text={quarter.milestone} limit={60} /></div>
        </Linkable>
      )}
      {months.length > 0 && (
        <div className="space-y-1 pt-1 border-t border-neutral-200">
          {months.map((m) => <CombinedMonthCompact key={m.id} month={m} trackId={trackId} trackName={trackName} />)}
        </div>
      )}
    </div>
  );
}

function CombinedWave({ wave, waveIndex, trackId, trackName }) {
  const t = useT();
  if (!waveHasContent(wave)) return null;
  const quarters = wave.quarters.filter(quarterHasContent);
  return (
    <Collapsible title={`${t("Волна")} ${waveIndex + 1}`} subtitle={wave.periodLabel} icon={<Layers size={13} />} tone={TONES.wave} defaultOpen={waveIndex === 0}>
      <div className="space-y-2">
        <StaticRow
          label={t("Objective на 6 мес.")} labelColor="text-emerald-700" text={wave.objective6mo} ownerId={wave.objective6moOwner}
          link={{ trackId, trackName, itemId: wave.id, itemLabel: `Objective 6 мес. (Волна ${waveIndex + 1}) — ${truncate(wave.objective6mo)}` }}
        />
        {quarters.length > 0 && (
          <div className="grid grid-cols-2 gap-1.5">
            {quarters.map((q) => <CombinedQuarter key={q.id} quarter={q} trackId={trackId} trackName={trackName} />)}
          </div>
        )}
      </div>
    </Collapsible>
  );
}

function CombinedOutput({ output, trackId, trackName }) {
  const t = useT();
  const outputText = useDataT(output.text);
  const waves = output.waves.filter(waveHasContent);
  return (
    <Linkable link={{ trackId, trackName, itemId: output.id, itemLabel: `${output.label} — ${truncate(output.text)}` }}>
      <div className="rounded-lg border p-2.5 space-y-2" style={{ borderColor: TONES.output.border, background: TONES.output.bg }}>
        {output.text && <div className="text-sm text-neutral-800"><Truncated text={outputText} /></div>}
        {waves.length > 0 ? (
          <div className="space-y-1.5">
            {output.waves.map((w, i) => <CombinedWave key={w.id} wave={w} waveIndex={i} trackId={trackId} trackName={trackName} />)}
          </div>
        ) : (
          <div className="text-xs text-neutral-400 italic">{t("волны ещё не разложены")}</div>
        )}
      </div>
    </Linkable>
  );
}

function CombinedOutputBranch({ outputs, trackId, trackName }) {
  const [selected, setSelected] = useState(outputs[0]?.id);
  const active = outputs.find((o) => o.id === selected) || outputs[0];
  if (outputs.length === 1) return <CombinedOutput output={outputs[0]} trackId={trackId} trackName={trackName} />;
  return (
    <div className="space-y-2">
      <div className="flex flex-wrap gap-1.5">
        {outputs.map((o) => (
          <button
            key={o.id}
            onClick={() => setSelected(o.id)}
            className="t11 font-medium px-2 py-1 rounded-full border flex items-center gap-1"
            style={
              o.id === active.id
                ? { borderColor: TONES.output.border, background: TONES.output.bg, color: TONES.output.text }
                : { borderColor: "#e5e5e5", color: "#a3a3a3" }
            }
          >
            <Target size={10} className="shrink-0" /> {o.label}
          </button>
        ))}
      </div>
      {active && <CombinedOutput output={active} trackId={trackId} trackName={trackName} />}
    </div>
  );
}

function CombinedOutcome({ outcome, trackId, trackName }) {
  const t = useT();
  const initiativeText = useDataT(outcome.initiativeText);
  if (!outcomeHasContent(outcome)) return null;
  const outputs = outcome.krOutputs.filter(outputHasContent);
  return (
    <Linkable link={{ trackId, trackName, itemId: outcome.id, itemLabel: `KR Outcome — ${truncate(outcome.text)}` }}>
      <Collapsible title={outcome.label} subtitle={truncate(outcome.text, 44)} icon={<Flag size={14} />} tone={TONES.outcome} defaultOpen={false}>
        <div className="space-y-2">
          <StaticRow label="KR Outcome" labelColor="text-yellow-700" text={outcome.text} />
          {outcome.metric && (
            <span className="inline-flex items-center gap-1 t11 font-medium text-yellow-800 bg-yellow-100 px-2 py-0.5 rounded-full">
              <BarChart3 size={10} /> {outcome.metric}
            </span>
          )}
          {(outcome.initiativeText || outcome.objectiveText || outputs.length > 0) && (
            <Collapsible
              title={t("Стратегическая инициатива на 18 мес.")} subtitle={truncate(outcome.initiativeText, 36)}
              icon={<Layers size={13} />} tone={TONES.initiative} defaultOpen={false}
            >
              <div className="space-y-2">
                {outcome.initiativeText && <div className="text-sm text-neutral-800"><Truncated text={initiativeText} /></div>}
                <StaticRow
                  label={t("Objective")} labelColor="text-blue-700" text={outcome.objectiveText} ownerId={outcome.objectiveOwner}
                  link={{ trackId, trackName, itemId: `${outcome.id}-init-objective`, itemLabel: `Objective инициативы — ${truncate(outcome.objectiveText)}` }}
                />
                {outputs.length > 0 && (
                  <div className="pt-1">
                    <CombinedOutputBranch outputs={outputs} trackId={trackId} trackName={trackName} />
                  </div>
                )}
              </div>
            </Collapsible>
          )}
        </div>
      </Collapsible>
    </Linkable>
  );
}

function MainGoalBlock({ uo, ui, trackId, trackName, color }) {
  const t = useT();
  const [open, setOpen] = useState(false);
  const outcomes = uo.krOutcomes.filter(outcomeHasContent);
  const uoText = useDataT(uo.text);
  const label = `${t("Главная цель")} ${ui + 1}`;
  const body = (
    <div>
      <span
        className="inline-block text-xs mb-0.5 px-1.5 py-0.5 rounded"
        style={{ background: TONES.mission.bg, color: TONES.mission.text }}
      >
        {label} <Owner id={uo.owner} />
      </span>
      {uo.text && <div className="text-sm text-neutral-800 mt-0.5"><Truncated text={uoText} /></div>}
    </div>
  );
  return (
    <div className="rounded-lg bg-white border border-neutral-100 px-3 py-2 space-y-1.5" style={{ borderLeftWidth: 3, borderLeftColor: color }}>
      {uo.text ? (
        <Linkable link={{ trackId, trackName, itemId: `ultimate-${uo.id}`, itemLabel: `Главная цель ${ui + 1} — ${truncate(uo.text)}` }}>
          {body}
        </Linkable>
      ) : body}
      {outcomes.length > 0 && (
        <>
          <button onClick={() => setOpen((o) => !o)} className="t11 text-amber-700 hover:text-amber-900 flex items-center gap-1 font-medium">
            {open ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
            {outcomes.length} KR Outcome
          </button>
          {open && (
            <div className="space-y-1.5 pt-1">
              {outcomes.map((o) => <CombinedOutcome key={o.id} outcome={o} trackId={trackId} trackName={trackName} />)}
            </div>
          )}
        </>
      )}
    </div>
  );
}

function TrackProjectsCollapse({ projects, trackId, trackName }) {
  const t = useT();
  const [open, setOpen] = useState(false);
  if (projects.length === 0) return null;
  return (
    <div className="space-y-1.5">
      <button
        onClick={() => setOpen((o) => !o)}
        className="w-full flex items-center gap-1.5 text-xs font-medium text-indigo-700 bg-indigo-50 border border-indigo-200 rounded-lg px-3 py-2 hover:bg-indigo-100"
      >
        {open ? <ChevronDown size={13} className="shrink-0" /> : <ChevronRight size={13} className="shrink-0" />}
        <Briefcase size={12} className="shrink-0" />
        <span>{t("Проектов")}: {projects.length}</span>
      </button>
      {open && (
        <div className="space-y-1.5">
          {projects.map((p) => (
            <ProjectMiniCard
              key={p.id}
              project={p}
              link={{ trackId, trackName, itemId: `project-${p.id}`, itemLabel: `Проект/Инициатива — ${truncate(p.name)}` }}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function CombinedTrackSection({ trackId, trackName, data, projects }) {
  const t = useT();
  const missionText = useDataT(data.mission);
  const color = TRACK_COLORS[trackId] || "#525252";
  const { filled, total } = countFilled(data);
  const pct = total ? Math.round((filled / total) * 100) : 0;
  const ultimateObjectives = (data.ultimateObjectives || []).filter(ultimateObjectiveHasContent);
  const trackProjects = (projects || []).filter((p) => p.trackIds.includes(trackId));

  return (
    <div className="rounded-xl border border-neutral-200 overflow-hidden" style={{ borderLeftWidth: 4, borderLeftColor: color }}>
      <div className="px-3 py-2.5 bg-neutral-50 flex items-center gap-3">
        <span className="text-sm font-semibold" style={{ color }}>{trackName}</span>
        <div className="flex items-center gap-2 ml-auto w-32">
          <div className="h-1.5 flex-1 bg-neutral-200 rounded-full overflow-hidden">
            <div className="h-full rounded-full" style={{ width: `${pct}%`, background: color }} />
          </div>
          <span className="text-xs text-neutral-400 w-8 text-right">{pct}%</span>
        </div>
      </div>
      <div className="p-3 space-y-2">
        {data.mission && (
          <div className="rounded-lg bg-white border border-neutral-100 px-3 py-2" style={{ borderLeftWidth: 3, borderLeftColor: color }}>
            <div className="text-xs text-neutral-500"><Truncated text={missionText} limit={100} /></div>
          </div>
        )}
        {trackProjects.length > 0 && (
          <TrackProjectsCollapse projects={trackProjects} trackId={trackId} trackName={trackName} />
        )}
        {ultimateObjectives.length > 0 ? (
          <div className="space-y-2">
            {ultimateObjectives.map((uo, ui) => (
              <MainGoalBlock key={uo.id} uo={uo} ui={ui} trackId={trackId} trackName={trackName} color={color} />
            ))}
          </div>
        ) : (
          <div className="text-xs text-neutral-400 italic px-1">{t("пока нет заполненных целей")}</div>
        )}
      </div>
    </div>
  );
}

function LinkBanner({ pending, onCancel }) {
  const t = useT();
  if (!pending) return null;
  return (
    <div className="flex items-center justify-between gap-3 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-800">
      <span className="min-w-0 break-words">
        {t("Выбрано:")} <b>{pending.trackName}</b> · {truncate(pending.itemLabel, 70)}. {t("Кликните 🔗 у другого узла (цели, KR, задачи или проекта), чтобы связать.")}
      </span>
      <button onClick={onCancel} className="text-amber-700 hover:text-amber-900 underline shrink-0">{t("отмена")}</button>
    </div>
  );
}

function LinkTypeModal({ pair, onChoose, onCancel }) {
  const t = useT();
  if (!pair) return null;
  const { a, b } = pair;
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4" onClick={onCancel}>
      <div className="bg-white rounded-2xl p-4 w-full max-w-sm space-y-3" onClick={(e) => e.stopPropagation()}>
        <div className="text-sm font-semibold text-neutral-800">{t("Тип связи между целями")}</div>
        <div className="text-xs bg-neutral-50 border border-neutral-200 rounded-lg p-2.5 space-y-1">
          <div><b style={{ color: TRACK_COLORS[a.trackId] }}>{a.trackName}</b> · {truncate(a.itemLabel, 80)}</div>
          <div className="text-neutral-400">↕</div>
          <div><b style={{ color: TRACK_COLORS[b.trackId] }}>{b.trackName}</b> · {truncate(b.itemLabel, 80)}</div>
        </div>
        <div className="grid grid-cols-2 gap-2">
          {LINK_TYPES.map((lt) => (
            <button
              key={lt.key}
              onClick={() => onChoose(lt.key)}
              className="flex items-center gap-1.5 text-xs font-medium rounded-lg px-2.5 py-2 border hover:brightness-95"
              style={{ borderColor: lt.color, color: lt.color, background: lt.color + "10" }}
            >
              <span>{lt.icon}</span> {t(lt.label)}
            </button>
          ))}
        </div>
        <button onClick={onCancel} className="w-full text-xs text-neutral-400 hover:text-neutral-600 py-1">
          {t("Отмена")}
        </button>
      </div>
    </div>
  );
}

function LinksPanel({ links, removeLink }) {
  const t = useT();
  if (links.length === 0) return null;
  return (
    <div className="border border-neutral-200 rounded-xl p-3 space-y-1.5">
      <div className="text-xs font-medium uppercase tracking-wide text-neutral-500 mb-1">
        {t("Связи между треками")} ({links.length})
      </div>
      {links.map((l) => {
        const cfg = LINK_TYPES.find((lt) => lt.key === l.type);
        return (
          <div key={l.id} className="flex items-center gap-2 text-xs bg-neutral-50 rounded-lg px-2.5 py-1.5">
            <span
              className="px-1.5 py-0.5 rounded-full border shrink-0"
              style={{ background: cfg.color + "14", color: cfg.color, borderColor: cfg.color + "55" }}
            >
              {cfg.icon} {t(cfg.label)}
            </span>
            <span className="flex-1 truncate">
              <b style={{ color: TRACK_COLORS[l.aTrackId] }}>{l.aTrackName}</b> · {truncate(l.aLabel, 40)}
              <span className="text-neutral-400"> ↔ </span>
              <b style={{ color: TRACK_COLORS[l.bTrackId] }}>{l.bTrackName}</b> · {truncate(l.bLabel, 40)}
            </span>
            <button onClick={() => removeLink(l.id)} className="text-neutral-300 hover:text-red-500 shrink-0">
              <Trash2 size={12} />
            </button>
          </div>
        );
      })}
    </div>
  );
}

function useLinks() {
  const [links, setLinks] = useState([]);
  const [loaded, setLoaded] = useState(false);
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("okr-links", false);
        setLinks(res ? JSON.parse(res.value) : []);
      } catch {
        setLinks([]);
      }
      setLoaded(true);
    })();
  }, []);
  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => {
      window.storage.set("okr-links", JSON.stringify(links), false).catch(() => {});
    }, 400);
    return () => clearTimeout(t);
  }, [links, loaded]);
  return [links, setLinks];
}

function useAllTracksData() {
  const [tracksData, setTracksData] = useState(null);
  const [loaded, setLoaded] = useState(false);
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const result = {};
      for (const t of TRACKS) {
        try {
          const res = await window.storage.get(`okr-track:${t.id}`, false);
          result[t.id] = res ? migrateTrackData(JSON.parse(res.value)) : createTrackData();
        } catch {
          result[t.id] = createTrackData();
        }
      }
      if (!cancelled) { setTracksData(result); setLoaded(true); }
    })();
    return () => { cancelled = true; };
  }, []);
  return [tracksData, loaded];
}

function ConnectorLayer({ paths }) {
  return (
    <svg className="absolute inset-0 w-full h-full" style={{ pointerEvents: "none", zIndex: 50, overflow: "visible" }}>
      {paths.map((p) => {
        const cfg = LINK_TYPES.find((t) => t.key === p.type);
        const dx = Math.max(50, Math.abs(p.bx - p.ax) / 2);
        const d = `M ${p.ax},${p.ay} C ${p.ax + dx},${p.ay} ${p.bx - dx},${p.by} ${p.bx},${p.by}`;
        return (
          <g key={p.id}>
            <path
              d={d} fill="none" stroke="#ffffff" strokeWidth="5.5"
              strokeLinecap="round" opacity="0.9"
            />
            <path
              d={d} fill="none" stroke={cfg.color} strokeWidth="2.5"
              strokeLinecap="round"
              strokeDasharray={p.type === "waiting" ? "7 5" : "none"}
            />
            <circle cx={p.ax} cy={p.ay} r="4.5" fill="#ffffff" />
            <circle cx={p.ax} cy={p.ay} r="3.5" fill={cfg.color} />
            <circle cx={p.bx} cy={p.by} r="4.5" fill="#ffffff" />
            <circle cx={p.bx} cy={p.by} r="3.5" fill={cfg.color} />
          </g>
        );
      })}
    </svg>
  );
}

function useZoom() {
  const [zoom, setZoom] = useState(1);
  const outerRef = useRef(null);
  const contentRef = useRef(null);

  const zoomIn = () => setZoom((z) => Math.min(1.2, +(z + 0.1).toFixed(2)));
  const zoomOut = () => setZoom((z) => Math.max(0.3, +(z - 0.1).toFixed(2)));
  const zoomReset = () => setZoom(1);
  const zoomFit = () => {
    if (!outerRef.current || !contentRef.current) return;
    const avail = outerRef.current.clientWidth - 8;
    const naturalW = contentRef.current.scrollWidth / zoom;
    if (avail > 0 && naturalW > 0) {
      setZoom(Math.max(0.3, Math.min(1, avail / naturalW)));
    }
  };

  return { zoom, zoomIn, zoomOut, zoomReset, zoomFit, outerRef, contentRef };
}

function useConnectorLines(links) {
  const stageRef = useRef(null);
  const refsMap = useRef(new Map());
  const [paths, setPaths] = useState([]);

  const registerRef = useCallback((itemId, el) => {
    if (el) refsMap.current.set(itemId, el);
    else refsMap.current.delete(itemId);
  }, []);

  const recompute = useCallback(() => {
    const stage = stageRef.current;
    if (!stage) return;
    const stageRect = stage.getBoundingClientRect();
    const next = [];
    links.forEach((l) => {
      const aEl = refsMap.current.get(l.aItemId);
      const bEl = refsMap.current.get(l.bItemId);
      if (!aEl || !bEl) return;
      const ar = aEl.getBoundingClientRect();
      const br = bEl.getBoundingClientRect();
      next.push({
        id: l.id,
        type: l.type,
        ax: ar.right - stageRect.left,
        ay: ar.top - stageRect.top + 6,
        bx: br.right - stageRect.left,
        by: br.top - stageRect.top + 6,
      });
    });
    setPaths(next);
  }, [links]);

  useEffect(() => {
    recompute();
    const stage = stageRef.current;
    const ro = new ResizeObserver(() => recompute());
    if (stage) ro.observe(stage);
    window.addEventListener("resize", recompute);
    const t1 = setTimeout(recompute, 60);
    const t2 = setTimeout(recompute, 300);
    return () => {
      ro.disconnect();
      window.removeEventListener("resize", recompute);
      clearTimeout(t1);
      clearTimeout(t2);
    };
  }, [recompute]);

  return { stageRef, registerRef, paths, recompute };
}

function CombinedTree() {
  const t = useT();
  const [tracksData, loaded] = useAllTracksData();
  const [projects, , projectsLoaded] = useProjects();
  const [links, setLinks] = useLinks();
  const [pending, setPending] = useState(null);
  const [modalPair, setModalPair] = useState(null);
  const { stageRef, registerRef, paths, recompute } = useConnectorLines(links);
  const { zoom, zoomIn, zoomOut, zoomReset, zoomFit, outerRef, contentRef } = useZoom();

  const pick = useCallback((item) => {
    setPending((prev) => {
      if (!prev) return item;
      if (prev.itemId === item.itemId) return null;
      setModalPair({ a: prev, b: item });
      return null;
    });
  }, []);

  const removeLink = useCallback((id) => {
    setLinks((prev) => prev.filter((l) => l.id !== id));
  }, [setLinks]);

  const confirmLink = (type) => {
    if (!modalPair) return;
    const { a, b } = modalPair;
    setLinks((prev) => [
      ...prev,
      {
        id: newId(), type,
        aTrackId: a.trackId, aTrackName: a.trackName, aItemId: a.itemId, aLabel: a.itemLabel,
        bTrackId: b.trackId, bTrackName: b.trackName, bItemId: b.itemId, bLabel: b.itemLabel,
      },
    ]);
    setModalPair(null);
    setTimeout(recompute, 60);
  };

  if (!loaded || !projectsLoaded) {
    return <div className="text-sm text-neutral-400 py-12 text-center">{t("Загрузка…")}</div>;
  }

  const withContent = TRACKS.filter(
    (tr) => trackHasContent(tracksData[tr.id]) || (projects || []).some((p) => p.trackIds.includes(tr.id))
  );

  return (
    <LinkCtx.Provider value={{ pending, pick, links, removeLink, registerRef }}>
      <div className="space-y-3">
        <div className="text-center">
          <div className="inline-block bg-neutral-100 rounded-xl px-4 py-2 text-sm font-semibold text-neutral-700">
            {t("Единая цель Coral Club из")} {TRACKS.length} {t("треков")}
          </div>
          <div className="text-neutral-300 text-lg leading-none mt-1">↓</div>
        </div>

        <div className="flex gap-2 flex-wrap justify-center pb-1">
          {TRACKS.map((tr) => {
            const { filled, total } = countFilled(tracksData[tr.id]);
            const pct = total ? Math.round((filled / total) * 100) : 0;
            return (
              <div key={tr.id} className="flex items-center gap-1.5 text-xs bg-white border border-neutral-200 rounded-full px-2.5 py-1">
                <span className="w-2 h-2 rounded-full shrink-0" style={{ background: TRACK_COLORS[tr.id] }} />
                <span className="text-neutral-600">{tr.name}</span>
                <span className="text-neutral-400">{pct}%</span>
              </div>
            );
          })}
        </div>

        <div className="flex items-center gap-3 flex-wrap justify-center t10 text-neutral-500 border-t border-neutral-100 pt-2">
          <span className="inline-flex items-center gap-1"><Flag size={11} className="text-yellow-600" /> {t("Outcome — итоговый результат")}</span>
          <span className="inline-flex items-center gap-1"><Layers size={11} className="text-blue-600" /> {t("Инициатива — программа работ")}</span>
          <span className="inline-flex items-center gap-1"><Target size={11} className="text-teal-600" /> {t("Output — результат программы")}</span>
          <span className="inline-flex items-center gap-1"><Layers size={11} className="text-emerald-600" /> {t("Волна — полугодовой этап")}</span>
          <span className="inline-flex items-center gap-1"><CalendarRange size={11} className="text-rose-600" /> {t("Квартал — операционный план")}</span>
        </div>

        <div className="flex items-center gap-3 flex-wrap justify-center t10 text-neutral-400">
          {LINK_TYPES.map((lt) => (
            <span key={lt.key} className="inline-flex items-center gap-1">
              <span className="inline-block w-3 h-0.5 rounded" style={{ background: lt.color }} />
              {lt.icon} {t(lt.label)}
            </span>
          ))}
        </div>

        <LinkBanner pending={pending} onCancel={() => setPending(null)} />

        {withContent.length === 0 ? (
          <div className="text-sm text-neutral-400 text-center py-10 border border-dashed border-neutral-200 rounded-xl">
            {t("Пока ни одна волна ни в одном треке не заполнена")}
          </div>
        ) : (
          <>
            <div className="flex items-center justify-center gap-1 pb-1">
              <button onClick={zoomOut} className="w-7 h-7 rounded-md border border-neutral-200 text-neutral-500 hover:bg-neutral-100 text-sm">−</button>
              <button onClick={zoomReset} className="px-2 h-7 rounded-md border border-neutral-200 text-neutral-500 hover:bg-neutral-100 text-xs tabular-nums w-14">
                {Math.round(zoom * 100)}%
              </button>
              <button onClick={zoomIn} className="w-7 h-7 rounded-md border border-neutral-200 text-neutral-500 hover:bg-neutral-100 text-sm">+</button>
              <button onClick={zoomFit} className="ml-1 px-2.5 h-7 rounded-md border border-neutral-200 text-neutral-500 hover:bg-neutral-100 text-xs">
                {t("уместить все треки")}
              </button>
            </div>

            <div ref={outerRef} className="overflow-auto border border-neutral-100 rounded-xl" style={{ maxHeight: "70vh" }}>
              <div
                ref={(el) => { contentRef.current = el; stageRef.current = el; }}
                className="relative flex gap-3 w-max"
                style={{ zoom }}
              >
                {withContent.map((tr) => (
                  <div key={tr.id} className="shrink-0 overflow-hidden" style={{ width: 340, maxWidth: 340 }}>
                    <CombinedTrackSection trackId={tr.id} trackName={tr.name} data={tracksData[tr.id]} projects={projects} />
                  </div>
                ))}
                <div className="absolute inset-0" style={{ pointerEvents: "none" }}>
                  <ConnectorLayer paths={paths} />
                </div>
              </div>
            </div>
          </>
        )}

        <LinksPanel links={links} removeLink={removeLink} />
      </div>

      <LinkTypeModal pair={modalPair} onChoose={confirmLink} onCancel={() => setModalPair(null)} />
    </LinkCtx.Provider>
  );
}

const DEFAULT_PROJECT_STATUSES = [
  { id: "not-started", name: "Не начат", color: "#94a3b8" },
  { id: "in-progress", name: "В работе", color: "#16a34a" },
  { id: "off-track", name: "Отклонение от срока", color: "#ca8a04" },
  { id: "frozen", name: "Заморожен", color: "#3b82f6" },
];
const LEGACY_STATUS_CODE_MAP = { G: "in-progress", Y: "off-track", R: "frozen", N: "not-started" };

function migrateProject(p) {
  if (p.statusId) return p;
  return {
    ...p,
    statusId: LEGACY_STATUS_CODE_MAP[p.statusCode] || "not-started",
    startDate: p.startDate || null,
    endDate: p.endDate || null,
  };
}

function useProjectStatuses() {
  const [statuses, setStatuses] = useState(DEFAULT_PROJECT_STATUSES);
  const [loaded, setLoaded] = useState(false);
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("okr-project-statuses", false);
        setStatuses(res ? JSON.parse(res.value) : DEFAULT_PROJECT_STATUSES);
      } catch {
        setStatuses(DEFAULT_PROJECT_STATUSES);
      }
      setLoaded(true);
    })();
  }, []);
  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => {
      window.storage.set("okr-project-statuses", JSON.stringify(statuses), false).catch(() => {});
    }, 400);
    return () => clearTimeout(t);
  }, [statuses, loaded]);
  return [statuses, setStatuses, loaded];
}

function useProjects() {
  const [projects, setProjects] = useState(null);
  const [loaded, setLoaded] = useState(false);
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("okr-projects", false);
        const raw = res ? JSON.parse(res.value) : PROJECTS_SEED;
        setProjects(raw.map(migrateProject));
      } catch {
        setProjects(PROJECTS_SEED.map(migrateProject));
      }
      setLoaded(true);
    })();
  }, []);
  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => {
      window.storage.set("okr-projects", JSON.stringify(projects), false).catch(() => {});
    }, 400);
    return () => clearTimeout(t);
  }, [projects, loaded]);
  return [projects, setProjects, loaded];
}

function projectShareTone(count) {
  if (count >= 4) return { border: "border-red-300", bg: "bg-red-50", label: "text-red-700", dot: "bg-red-400" };
  if (count === 3) return { border: "border-amber-300", bg: "bg-amber-50", label: "text-amber-700", dot: "bg-amber-400" };
  if (count === 2) return { border: "border-emerald-300", bg: "bg-emerald-50", label: "text-emerald-700", dot: "bg-emerald-400" };
  return { border: "border-indigo-200", bg: "bg-indigo-50", label: "text-indigo-700", dot: "bg-indigo-300" };
}

function ProjectMiniCard({ project, link }) {
  const t = useT();
  const description = useDataT(project.description);
  const filledKrs = project.krs.filter((k) => !!k.text);
  const shareCount = (project.trackIds || []).length;
  const tone = projectShareTone(shareCount);
  const body = (
    <div className={`rounded-lg border ${tone.border} ${tone.bg} px-3 py-2 space-y-1 min-w-0`}>
      <div className={`flex items-center gap-1.5 text-xs font-medium ${tone.label} min-w-0`}>
        <Briefcase size={12} className="shrink-0" />
        <span className="break-words min-w-0 flex-1">{t("Проект/Инициатива")} — {project.name || t("без названия")}</span>
        {shareCount > 1 && (
          <span className={`t10 px-1.5 py-0.5 rounded-full shrink-0 font-semibold text-white ${tone.dot}`}>
            {shareCount} {t("трека")}
          </span>
        )}
      </div>
      {(project.division || project.function) && (
        <div className="t11 text-indigo-400 break-words">
          {[project.division, project.function].filter(Boolean).join(" · ")}
        </div>
      )}
      {project.description && <div className="text-sm text-neutral-700 whitespace-pre-wrap break-words">{description}</div>}
      {filledKrs.length > 0 && (
        <div className="space-y-0.5 pt-0.5">
          {filledKrs.map((k) => (
            <div key={k.id} className="flex items-start gap-1.5 text-xs text-neutral-600 min-w-0">
              <Circle size={5} className="text-indigo-300 shrink-0 mt-1" />
              <span className="break-words min-w-0"><TranslatedText text={k.text} /></span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
  return link ? <Linkable link={link}>{body}</Linkable> : body;
}

function TrackProjects({ trackId, trackName }) {
  const t = useT();
  const [projects, , loaded] = useProjects();
  if (!loaded || !projects) return null;
  const list = projects.filter((p) => p.trackIds.includes(trackId));
  if (list.length === 0) return null;
  return (
    <Collapsible title={t("Проекты / инициативы трека")} icon={<Briefcase size={13} />} tone={TONES.project}>
      <div className="space-y-1.5">
        {list.map((p) => (
          <ProjectMiniCard
            key={p.id}
            project={p}
            link={{ trackId, trackName, itemId: `project-${p.id}`, itemLabel: `Проект/Инициатива — ${truncate(p.name)}` }}
          />
        ))}
      </div>
    </Collapsible>
  );
}


function MonthSpark({ activity }) {
  const t = useT();
  const MLABELS = ["Я", "Ф", "М", "А", "М", "И", "И", "А", "С", "О", "Н", "Д"];
  return (
    <div className="flex items-end gap-0.5 h-4" title={t("Активность по месяцам 2026")}>
      {activity.map((v, i) => (
        <div
          key={i}
          className={`rounded-sm ${v ? "bg-neutral-400" : "bg-neutral-150"}`}
          style={{ width: 7, height: v ? "100%" : "25%", background: v ? undefined : "#e5e5e5" }}
          title={MLABELS[i]}
        />
      ))}
    </div>
  );
}

function ProjectCard({ project, onChange, onDelete, statuses }) {
  const t = useT();
  const [open, setOpen] = useState(false);
  const statusObj = statuses.find((s) => s.id === project.statusId) || statuses[0] || { name: "", color: "#94a3b8" };
  const color = statusObj.color;

  const update = (field, value) => onChange({ ...project, [field]: value });

  const addKr = () => update("krs", [...project.krs, { id: newId(), text: "" }]);
  const updateKr = (id, text) => update("krs", project.krs.map((k) => (k.id === id ? { ...k, text } : k)));
  const removeKr = (id) => update("krs", project.krs.filter((k) => k.id !== id));

  const toggleTrack = (trackId) => {
    const has = project.trackIds.includes(trackId);
    if (has) {
      update("trackIds", project.trackIds.filter((t) => t !== trackId));
    } else if (project.trackIds.length < MAX_PROJECT_TRACKS) {
      update("trackIds", [...project.trackIds, trackId]);
    }
  };

  return (
    <div className="rounded-xl border border-neutral-200 overflow-hidden bg-white">
      <button onClick={() => setOpen((o) => !o)} className="w-full flex items-start gap-2.5 px-3 py-2.5 text-left hover:bg-neutral-50">
        {open ? <ChevronDown size={15} className="text-neutral-400 shrink-0 mt-0.5" /> : <ChevronRight size={15} className="text-neutral-400 shrink-0 mt-0.5" />}
        <span className="w-2 h-2 rounded-full shrink-0 mt-1.5" style={{ background: color }} title={statusObj.name} />
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-1.5 flex-wrap">
            <span className="text-sm font-medium text-neutral-800 break-words">{project.name}</span>
            <span className="t10 px-1.5 py-0.5 rounded-full shrink-0 font-medium" style={{ background: color + "1A", color }}>
              {statusObj.name}
            </span>
            <span
              className="t10 px-1.5 py-0.5 rounded-full shrink-0 font-medium"
              style={
                (project.projectType || "run") === "change"
                  ? { background: "#ede9fe", color: "#6d28d9" }
                  : { background: "#e0f2fe", color: "#0369a1" }
              }
            >
              {(project.projectType || "run") === "change" ? "Change" : "RUN"}
            </span>
          </div>
          <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 mt-0.5 t11 text-neutral-400">
            <span>{project.division}</span>
            <span>·</span>
            <span>{project.function}</span>
            {(project.startDate || project.endDate) && (
              <>
                <span>·</span>
                <span>{formatDateRu(project.startDate)}{project.endDate ? ` — ${formatDateRu(project.endDate)}` : ""}</span>
              </>
            )}
          </div>
        </div>
        <div className="flex items-center gap-2 shrink-0 mt-1">
          {project.trackIds.map((tid) => (
            <span key={tid} className="w-2 h-2 rounded-full" style={{ background: TRACK_COLORS[tid] }} title={TRACKS.find((t) => t.id === tid)?.name} />
          ))}
          <MonthSpark activity={project.activity} />
        </div>
      </button>

      {open && (
        <div className="px-3 pb-3 pt-1 space-y-3 border-t border-neutral-100">
          <div>
            <div className="text-xs text-neutral-500 mb-1">{t("Тип проекта")}</div>
            <div className="flex gap-1.5">
              <button
                onClick={() => update("projectType", "run")}
                className="flex items-center gap-1.5 text-xs px-2.5 py-1 rounded-full border"
                style={
                  (project.projectType || "run") === "run"
                    ? { background: "#e0f2fe", borderColor: "#0369a1", color: "#0369a1" }
                    : { borderColor: "#e5e5e5", color: "#737373" }
                }
              >
                RUN
              </button>
              <button
                onClick={() => update("projectType", "change")}
                className="flex items-center gap-1.5 text-xs px-2.5 py-1 rounded-full border"
                style={
                  (project.projectType || "run") === "change"
                    ? { background: "#ede9fe", borderColor: "#6d28d9", color: "#6d28d9" }
                    : { borderColor: "#e5e5e5", color: "#737373" }
                }
              >
                Change
              </button>
            </div>
            <div className="t11 text-neutral-400 mt-1">
              {(project.projectType || "run") === "change" ? t("Change — разовая инициатива изменений") : t("RUN — текущая операционная деятельность")}
            </div>
          </div>

          <div className="grid grid-cols-3 gap-2">
            <div>
              <div className="text-xs text-neutral-500 mb-1">{t("Статус")}</div>
              <select
                value={project.statusId || statuses[0]?.id}
                onChange={(e) => update("statusId", e.target.value)}
                className="w-full text-sm bg-neutral-50 border border-neutral-200 rounded-md px-2 py-1.5 outline-none"
              >
                {statuses.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </div>
            <div>
              <div className="text-xs text-neutral-500 mb-1">{t("Дата начала")}</div>
              <input
                type="date" value={project.startDate || ""}
                onChange={(e) => update("startDate", e.target.value || null)}
                className="w-full text-sm bg-neutral-50 border border-neutral-200 rounded-md px-2 py-1.5 outline-none"
              />
            </div>
            <div>
              <div className="text-xs text-neutral-500 mb-1">{t("Ожидаемая дата конца")}</div>
              <input
                type="date" value={project.endDate || ""}
                onChange={(e) => update("endDate", e.target.value || null)}
                className="w-full text-sm bg-neutral-50 border border-neutral-200 rounded-md px-2 py-1.5 outline-none"
              />
            </div>
          </div>

          <div>
            <div className="text-xs text-neutral-500 mb-1">{t("Описание проекта")}</div>
            <textarea
              className="w-full text-sm bg-neutral-50 border border-neutral-200 rounded-lg p-2 outline-none resize-none"
              rows={2}
              placeholder={t("Кратко опишите проект…")}
              value={project.description}
              onChange={(e) => update("description", e.target.value)}
            />
            <FieldPreview value={project.description} />
          </div>

          <div>
            <div className="text-xs text-neutral-500 mb-1">
              {t("Каким трекам помогает")} ({project.trackIds.length}/{MAX_PROJECT_TRACKS})
            </div>
            <div className="flex flex-wrap gap-1.5">
              {TRACKS.map((tr) => {
                const active = project.trackIds.includes(tr.id);
                const disabled = !active && project.trackIds.length >= MAX_PROJECT_TRACKS;
                return (
                  <button
                    key={tr.id}
                    disabled={disabled}
                    onClick={() => toggleTrack(tr.id)}
                    className="flex items-center gap-1.5 text-xs px-2 py-1 rounded-full border disabled:opacity-35 disabled:cursor-not-allowed"
                    style={
                      active
                        ? { background: TRACK_COLORS[tr.id] + "18", borderColor: TRACK_COLORS[tr.id], color: TRACK_COLORS[tr.id] }
                        : { borderColor: "#e5e5e5", color: "#737373" }
                    }
                  >
                    <span className="w-1.5 h-1.5 rounded-full" style={{ background: TRACK_COLORS[tr.id] }} />
                    {tr.name}
                    {active && <span className="ml-0.5">✓</span>}
                  </button>
                );
              })}
            </div>
          </div>

          <div>
            <div className="text-xs text-neutral-500 mb-1">{t("KR проекта")}</div>
            <div className="space-y-1">
              {project.krs.map((k, i) => (
                <div key={k.id}>
                  <div className="flex items-center gap-1.5">
                    <Circle size={6} className="text-neutral-300 shrink-0" />
                    <input
                      className="flex-1 text-sm bg-neutral-50 border border-neutral-200 rounded-md px-2 py-1 outline-none"
                      placeholder={`KR ${i + 1}`}
                      value={k.text}
                      onChange={(e) => updateKr(k.id, e.target.value)}
                    />
                    <button onClick={() => removeKr(k.id)} className="text-neutral-300 hover:text-red-500 p-0.5 shrink-0">
                      <Trash2 size={12} />
                    </button>
                  </div>
                  <div className="pl-4"><FieldPreview value={k.text} /></div>
                </div>
              ))}
            </div>
            <button onClick={addKr} className="mt-1.5 flex items-center gap-1 text-xs text-neutral-500 hover:text-neutral-800">
              <Plus size={12} /> {t("добавить KR")}
            </button>
          </div>

          <button onClick={onDelete} className="text-xs text-neutral-300 hover:text-red-500">
            {t("Удалить проект")}
          </button>
        </div>
      )}
    </div>
  );
}

const STATUS_COLOR_SWATCHES = ["#94a3b8", "#16a34a", "#ca8a04", "#dc2626", "#3b82f6", "#8b5cf6", "#ec4899", "#0d9488"];

function StatusManager({ statuses, setStatuses }) {
  const t = useT();
  const [name, setName] = useState("");
  const [open, setOpen] = useState(false);

  const add = () => {
    if (!name.trim()) return;
    const usedColors = statuses.map((s) => s.color);
    const color = STATUS_COLOR_SWATCHES.find((c) => !usedColors.includes(c)) || STATUS_COLOR_SWATCHES[0];
    setStatuses((prev) => [...prev, { id: newId(), name: name.trim(), color }]);
    setName("");
  };
  const remove = (id) => setStatuses((prev) => prev.filter((s) => s.id !== id));
  const recolor = (id, color) => setStatuses((prev) => prev.map((s) => (s.id === id ? { ...s, color } : s)));

  return (
    <div className="mb-4 border border-neutral-200 rounded-xl overflow-hidden">
      <button onClick={() => setOpen((o) => !o)} className="w-full flex items-center gap-2 px-3 py-2 bg-neutral-50 hover:bg-neutral-100">
        {open ? <ChevronDown size={15} className="text-neutral-500" /> : <ChevronRight size={15} className="text-neutral-500" />}
        <Flag size={14} className="text-neutral-500" />
        <span className="text-xs font-medium uppercase tracking-wide text-neutral-600">{t("Статусы проектов")}</span>
        <span className="text-xs text-neutral-400 ml-auto">{statuses.length}</span>
      </button>
      {open && (
        <div className="p-3 space-y-2">
          <div className="flex gap-1.5">
            <input
              className="flex-1 text-sm border border-neutral-200 rounded-md px-2 py-1 outline-none"
              placeholder={t("Новый статус (например: На паузе)")}
              value={name}
              onChange={(e) => setName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && add()}
            />
            <button onClick={add} className="px-2.5 rounded-md bg-neutral-900 text-white text-sm hover:bg-neutral-800">
              <Plus size={14} />
            </button>
          </div>
          <div className="space-y-1.5">
            {statuses.map((s) => (
              <div key={s.id} className="flex items-center gap-2 px-2 py-1.5 rounded-md hover:bg-neutral-50">
                <div className="flex items-center gap-1">
                  {STATUS_COLOR_SWATCHES.map((c) => (
                    <button
                      key={c} onClick={() => recolor(s.id, c)}
                      className="w-4 h-4 rounded-full shrink-0"
                      style={{ background: c, outline: s.color === c ? "2px solid #171717" : "none", outlineOffset: 1 }}
                    />
                  ))}
                </div>
                <span className="text-sm flex-1 min-w-0 break-words" style={{ color: s.color }}>{s.name}</span>
                <button onClick={() => remove(s.id)} className="text-neutral-300 hover:text-red-500 p-0.5 shrink-0">
                  <Trash2 size={12} />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function ProjectsModule() {
  const t = useT();
  const [projects, setProjects, loaded] = useProjects();
  const [statuses, setStatuses, statusesLoaded] = useProjectStatuses();
  const [query, setQuery] = useState("");
  const [divisionFilter, setDivisionFilter] = useState("");
  const [trackFilter, setTrackFilter] = useState("");
  const [typeFilter, setTypeFilter] = useState("");

  if (!loaded || !statusesLoaded) {
    return <div className="text-sm text-neutral-400 py-12 text-center">{t("Загрузка…")}</div>;
  }

  const updateProject = (updated) => {
    setProjects((prev) => prev.map((p) => (p.id === updated.id ? updated : p)));
  };
  const deleteProject = (id) => {
    setProjects((prev) => prev.filter((p) => p.id !== id));
  };
  const addProject = () => {
    setProjects((prev) => [
      {
        id: newId(), division: "", function: "", name: "", statusId: statuses[0]?.id || "not-started",
        startDate: null, endDate: null, launchMonth: null, activity: Array(12).fill(0),
        description: "", krs: [], trackIds: [], projectType: "run",
      },
      ...prev,
    ]);
  };

  const divisions = Array.from(new Set(projects.map((p) => p.division).filter(Boolean))).sort();

  const filtered = projects.filter((p) => {
    if (divisionFilter && p.division !== divisionFilter) return false;
    if (trackFilter && !p.trackIds.includes(trackFilter)) return false;
    if (typeFilter && (p.projectType || "run") !== typeFilter) return false;
    if (query && !(`${p.name} ${p.function} ${p.division}`.toLowerCase().includes(query.toLowerCase()))) return false;
    return true;
  });

  const linkedCount = projects.filter((p) => p.trackIds.length > 0).length;

  return (
    <div className="space-y-3">
      <StatusManager statuses={statuses} setStatuses={setStatuses} />

      <div className="flex items-center justify-between px-1">
        <div className="text-xs text-neutral-400">
          {projects.length} {t("проектов")} · {t("связаны с треками")} — {linkedCount}
        </div>
        <button onClick={addProject} className="flex items-center gap-1 text-xs text-yellow-700 hover:text-yellow-900">
          <Plus size={13} /> {t("добавить проект")}
        </button>
      </div>

      <div className="flex flex-wrap gap-1.5 justify-center">
        {TRACKS.map((tr) => {
          const count = projects.filter((p) => p.trackIds.includes(tr.id)).length;
          const active = trackFilter === tr.id;
          return (
            <button
              key={tr.id}
              onClick={() => setTrackFilter(active ? "" : tr.id)}
              className="flex items-center gap-1.5 text-xs rounded-full px-2.5 py-1 border"
              style={
                active
                  ? { background: TRACK_COLORS[tr.id] + "18", borderColor: TRACK_COLORS[tr.id], color: TRACK_COLORS[tr.id] }
                  : { borderColor: "#e5e5e5", color: "#737373" }
              }
            >
              <span className="w-2 h-2 rounded-full shrink-0" style={{ background: TRACK_COLORS[tr.id] }} />
              {tr.name} <span className="text-neutral-400">{count}</span>
            </button>
          );
        })}
      </div>

      <div className="flex gap-1.5 justify-center">
        <button
          onClick={() => setTypeFilter(typeFilter === "run" ? "" : "run")}
          className="text-xs px-2.5 py-1 rounded-full border"
          style={typeFilter === "run" ? { background: "#e0f2fe", borderColor: "#0369a1", color: "#0369a1" } : { borderColor: "#e5e5e5", color: "#737373" }}
        >
          RUN
        </button>
        <button
          onClick={() => setTypeFilter(typeFilter === "change" ? "" : "change")}
          className="text-xs px-2.5 py-1 rounded-full border"
          style={typeFilter === "change" ? { background: "#ede9fe", borderColor: "#6d28d9", color: "#6d28d9" } : { borderColor: "#e5e5e5", color: "#737373" }}
        >
          Change
        </button>
      </div>

      <div className="flex gap-2">
        <div className="flex-1 flex items-center gap-1.5 bg-neutral-50 border border-neutral-200 rounded-lg px-2.5 py-1.5">
          <Search size={13} className="text-neutral-400 shrink-0" />
          <input
            className="flex-1 bg-transparent outline-none text-sm placeholder-neutral-400"
            placeholder={t("Поиск по названию, функции, дивизиону…")}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <select
          value={divisionFilter}
          onChange={(e) => setDivisionFilter(e.target.value)}
          className="text-sm border border-neutral-200 rounded-lg px-2 bg-neutral-50"
        >
          <option value="">{t("Все дивизионы")}</option>
          {divisions.map((d) => (
            <option key={d} value={d}>{d}</option>
          ))}
        </select>
      </div>

      <div className="space-y-1.5">
        {filtered.length === 0 ? (
          <div className="text-sm text-neutral-400 text-center py-8 border border-dashed border-neutral-200 rounded-xl">
            {t("Ничего не найдено")}
          </div>
        ) : (
          filtered.map((p) => (
            <ProjectCard key={p.id} project={p} onChange={updateProject} onDelete={() => deleteProject(p.id)} statuses={statuses} />
          ))
        )}
      </div>
    </div>
  );
}

function rangeFromOffset(trackStart, offsetMonths, spanMonths) {
  return { start: addMonthsISO(trackStart, offsetMonths), end: addMonthsISO(trackStart, offsetMonths + spanMonths) };
}
function resolveRange(tracking, structuralRange, horizonMonths) {
  if (tracking) {
    const s = tracking.startDate || structuralRange.start;
    const e = tracking.deadlineOverride || addMonthsISO(s, horizonMonths || 1);
    return { start: s, end: e, tracked: true };
  }
  return { start: structuralRange.start, end: structuralRange.end, tracked: false };
}
function taskPointDate(monthStart, monthEnd, idx, total) {
  const totalDays = Math.max(1, daysBetweenISO(monthStart, monthEnd));
  const offsetDays = Math.round(((idx + 0.5) / Math.max(1, total)) * totalDays);
  const d = new Date(monthStart + "T00:00:00");
  d.setDate(d.getDate() + offsetDays);
  return d.toISOString().slice(0, 10);
}
function waveHasAnyContent(w) {
  return !!w.targetText || waveHasContent(w);
}

// Walks the whole tree for one track — mission down to tasks — producing flat Gantt rows.
// maxDepth caps how deep to recurse: 0=главные цели … 6=задачи.
function lastVisibleIndex(arr, hasContentFn) {
  let last = -1;
  arr.forEach((item, i) => { if (hasContentFn(item)) last = i; });
  return last;
}

// ---- Bitrix import preview overlay for the Gantt ----
function useBitrixDraftRows() {
  const [drafts, setDrafts] = useState([]);
  const [loaded, setLoaded] = useState(false);
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("okr-bitrix-import:draft", false);
        const d = res ? JSON.parse(res.value) : [];
        setDrafts(Array.isArray(d) ? d.filter((r) => r.include && r.trackId && pathIsComplete(r.levelKey, r.path)) : []);
      } catch { setDrafts([]); }
      setLoaded(true);
    })();
  }, []);
  return [drafts, loaded];
}

const IMPORT_EDIT_ROW_ID = {
  mission: (p) => `uo-${p.uoId}`, initiative: (p) => `o-${p.outcomeId}`,
  wave: (p) => `w-${p.waveId}`, quarterInitiative: (p) => `q-${p.quarterId}`, month: (p) => `m-${p.monthId}`,
};

// Merges reviewed-but-not-yet-applied Bitrix rows into one track's already-built Gantt rows —
// new tasks (and brand-new Главная цель / KR Outcome) get inserted as dashed "draft" rows;
// everything else (filling text on a node that already exists) highlights that existing row.
function injectDraftRows(trackRows, drafts, trackId, trackName) {
  const relevant = drafts.filter((d) => d.trackId === trackId);
  if (relevant.length === 0) return trackRows;
  let out = trackRows.slice();
  relevant.forEach((d) => {
    const text = d.editedText || d.rawText;
    if (d.levelKey === "task") {
      const idx = out.findIndex((r) => r.id === `m-${d.path.monthId}`);
      if (idx === -1) return;
      const parent = out[idx];
      out.splice(idx + 1, 0, {
        id: `draft-${d.rowId}`, depth: parent.depth + 1, guides: [...parent.guides, false], isLast: true,
        trackId, trackName, levelKey: "Задача", levelExtra: "",
        title: text, start: todayISO(), end: todayISO(), tracked: false, isTask: true, draft: true,
      });
      return;
    }
    if (d.levelKey === "mission" && d.path.uoId === "__new__") {
      out.push({
        id: `draft-${d.rowId}`, depth: 0, guides: [], isLast: true, trackId, trackName,
        levelKey: "Главная цель", levelExtra: "", title: text, start: todayISO(), end: todayISO(), tracked: false, draft: true,
      });
      return;
    }
    if (d.levelKey === "initiative" && d.path.outcomeId === "__new__") {
      const idx = out.findIndex((r) => r.id === `uo-${d.path.uoId}`);
      const depth = idx >= 0 ? out[idx].depth + 1 : 1;
      out.splice(idx >= 0 ? idx + 1 : out.length, 0, {
        id: `draft-${d.rowId}`, depth, guides: idx >= 0 ? [...out[idx].guides, false] : [], isLast: true, trackId, trackName,
        levelKey: "KR Outcome", levelExtra: "", title: text, start: todayISO(), end: todayISO(), tracked: false, draft: true,
      });
      return;
    }
    const idFn = IMPORT_EDIT_ROW_ID[d.levelKey];
    const targetId = idFn && idFn(d.path);
    const idx = targetId ? out.findIndex((r) => r.id === targetId) : -1;
    if (idx === -1) return;
    out[idx] = { ...out[idx], draftEdit: true, draftText: text };
  });
  return out;
}

function buildFullGanttRows(trackId, trackName, data, maxDepth) {
  const rows = [];
  const trackStart = data.startDate || todayISO();
  const uos = data.ultimateObjectives || [];
  const uoLastIdx = lastVisibleIndex(uos, ultimateObjectiveHasContent);

  uos.forEach((uo, ui) => {
    if (!ultimateObjectiveHasContent(uo)) return;
    const uoIsLast = ui === uoLastIdx;
    const uoGuides = [];
    const uoRange = rangeFromOffset(trackStart, 0, 18);
    rows.push({ id: `uo-${uo.id}`, depth: 0, guides: uoGuides, isLast: uoIsLast, trackId, trackName, levelKey: "Главная цель", levelExtra: `${ui + 1}`, title: uo.text, ...uoRange, tracked: false });
    if (maxDepth < 1) return;

    const outcomes = uo.krOutcomes;
    const oLastIdx = lastVisibleIndex(outcomes, outcomeHasContent);
    outcomes.forEach((o, oi) => {
      if (!outcomeHasContent(o)) return;
      const oIsLast = oi === oLastIdx;
      const oGuides = [...uoGuides, !uoIsLast];
      const oRange = resolveRange(o.tracking, rangeFromOffset(trackStart, 0, 18), 18);
      rows.push({ id: `o-${o.id}`, depth: 1, guides: oGuides, isLast: oIsLast, trackId, trackName, levelKey: "KR Outcome", levelExtra: "", title: o.text || o.label, ...oRange });
      if (maxDepth < 2) return;

      const outputs = o.krOutputs;
      const koLastIdx = lastVisibleIndex(outputs, outputHasContent);
      outputs.forEach((ko, ki) => {
        if (!outputHasContent(ko)) return;
        const koIsLast = ki === koLastIdx;
        const koGuides = [...oGuides, !oIsLast];
        const koRange = resolveRange(ko.tracking, rangeFromOffset(trackStart, 0, 18), 18);
        rows.push({ id: `ko-${ko.id}`, depth: 2, guides: koGuides, isLast: koIsLast, trackId, trackName, levelKey: ko.label, levelExtra: "", title: ko.text, ...koRange });
        if (maxDepth < 3) return;

        const waves = ko.waves;
        const wLastIdx = lastVisibleIndex(waves, waveHasAnyContent);
        waves.forEach((w, wi) => {
          if (!waveHasAnyContent(w)) return;
          const wIsLast = wi === wLastIdx;
          const wGuides = [...koGuides, !koIsLast];
          const wRange = rangeFromOffset(trackStart, wi * 6, 6);
          rows.push({ id: `w-${w.id}`, depth: 3, guides: wGuides, isLast: wIsLast, trackId, trackName, levelKey: "Волна", levelExtra: `${wi + 1}`, title: w.status === "decomposed" ? (w.objective6mo || "") : w.targetText, start: wRange.start, end: wRange.end, tracked: false });
          if (maxDepth < 4 || w.status !== "decomposed") return;

          const quarters = w.quarters;
          const qLastIdx = lastVisibleIndex(quarters, quarterHasContent);
          quarters.forEach((q, qi) => {
            if (!quarterHasContent(q)) return;
            const qIsLast = qi === qLastIdx;
            const qGuides = [...wGuides, !wIsLast];
            const qStructural = rangeFromOffset(trackStart, wi * 6 + qi * 3, 3);
            const qRange = resolveRange(q.tracking, qStructural, 3);
            rows.push({ id: `q-${q.id}`, depth: 4, guides: qGuides, isLast: qIsLast, trackId, trackName, levelKey: "Квартал", levelExtra: q.label, title: q.milestone || "", ...qRange });
            if (maxDepth < 5) return;

            const months = q.monthlyKRs;
            const mLastIdx = lastVisibleIndex(months, monthHasContent);
            months.forEach((m, mi) => {
              if (!monthHasContent(m)) return;
              const mIsLast = mi === mLastIdx;
              const mGuides = [...qGuides, !qIsLast];
              const mStructural = rangeFromOffset(trackStart, wi * 6 + qi * 3 + mi, 1);
              const mRange = resolveRange(m.tracking, mStructural, 1);
              rows.push({ id: `m-${m.id}`, depth: 5, guides: mGuides, isLast: mIsLast, trackId, trackName, levelKey: "KR месяца", levelExtra: m.label, title: m.text || "", ...mRange });
              if (maxDepth < 6) return;

              const tasks = m.tasks.filter((tk) => !!tk.text);
              tasks.forEach((tk, ti) => {
                const tIsLast = ti === tasks.length - 1;
                const tGuides = [...mGuides, !mIsLast];
                const point = taskPointDate(mStructural.start, mStructural.end, ti, tasks.length);
                rows.push({ id: `tk-${tk.id}`, depth: 6, guides: tGuides, isLast: tIsLast, trackId, trackName, levelKey: "Задача", levelExtra: "", title: tk.text, start: point, end: point, tracked: false, isTask: true });
              });
            });
          });
        });
      });
    });
  });

  return rows;
}

function collectTrackers(trackId, trackName, data, directory) {
  const items = [];
  (data.ultimateObjectives || []).forEach((uo) => {
  (uo.krOutcomes || []).forEach((o) => {
    if (o.tracking) {
      items.push({
        key: `outcome-${o.id}`, trackId, trackName, level: "KR Outcome (18 мес.)",
        title: o.text || o.label, owner: ownerNameById(directory, o.objectiveOwner), tracking: o.tracking, horizonMonths: 18,
      });
    }
    (o.krOutputs || []).forEach((ko) => {
      if (ko.tracking) {
        items.push({
          key: `output-${ko.id}`, trackId, trackName, level: `${ko.label} (18 мес.)`,
          title: ko.text || ko.label, owner: "", tracking: ko.tracking, horizonMonths: 18,
        });
      }
      (ko.waves || []).forEach((w) => {
        (w.quarters || []).forEach((q) => {
          if (q.tracking) {
            items.push({
              key: `quarter-${q.id}`, trackId, trackName, level: `KR-веха (${q.label})`,
              title: q.milestone || tSync("без названия"), owner: ownerNameById(directory, q.objectiveOwner), tracking: q.tracking, horizonMonths: 3,
            });
          }
          (q.monthlyKRs || []).forEach((m) => {
            if (m.tracking) {
              items.push({
                key: `month-${m.id}`, trackId, trackName, level: `KR месяца (${m.label})`,
                title: m.text || m.label, owner: "", tracking: m.tracking, horizonMonths: 1,
              });
            }
          });
        });
      });
    });
  });
  });
  return items;
}

function quarterHasTracking(q) {
  return !!q.tracking || (q.monthlyKRs || []).some((m) => !!m.tracking);
}
function outputHasTracking(ko) {
  return !!ko.tracking || (ko.waves || []).some((w) => (w.quarters || []).some(quarterHasTracking));
}
function outcomeHasTracking(o) {
  return !!o.tracking || (o.krOutputs || []).some(outputHasTracking);
}
function ultimateObjectiveHasTracking(uo) {
  return (uo.krOutcomes || []).some(outcomeHasTracking);
}

function TrackerBar({ tracking, horizonMonths }) {
  const t = useT();
  const { fact, pct } = computeTracking(tracking);
  const status = trackingStatus(pct);
  const startDate = tracking.startDate || todayISO();
  const deadline = tracking.deadlineOverride || addMonthsISO(startDate, horizonMonths || 1);
  const daysLeft = daysBetweenISO(todayISO(), deadline);
  const urgency = deadlineUrgency(daysLeft);
  return (
    <div className="space-y-0.5 mt-1">
      <div className="flex items-center gap-2">
        <div className="flex-1 h-1.5 bg-neutral-100 rounded-full overflow-hidden">
          <div className="h-full rounded-full" style={{ width: `${Math.min(100, Math.max(0, pct * 100))}%`, background: status.color }} />
        </div>
        <span className="text-xs shrink-0" style={{ color: status.color }}>
          {fact.toFixed(1)}/{tracking.target} · {(pct * 100).toFixed(0)}% · {t(status.label)}
        </span>
      </div>
      <div className="t11" style={{ color: urgency.color }}>
        {t("Срок")} {formatDateRu(deadline)} ·{" "}
        {daysLeft < 0 ? `${t("просрочено на")} ${Math.abs(daysLeft)} ${t("дн.")}` : `${t("осталось")} ${daysLeft} ${t("дн.")}`}
      </div>
    </div>
  );
}

function DashboardNode({ level, title, owner, tracking, horizonMonths, borderColor, children }) {
  const displayTitle = useDataT(title);
  return (
    <div className="pl-3 space-y-1.5" style={{ borderLeftWidth: 2, borderLeftColor: borderColor }}>
      <div>
        <div className="t11 text-neutral-400">{level}{owner && ` · ${owner}`}</div>
        <div className="text-sm text-neutral-800 break-words">{displayTitle}</div>
        {tracking && <TrackerBar tracking={tracking} horizonMonths={horizonMonths} />}
      </div>
      {children}
    </div>
  );
}

function DashboardQuarter({ quarter, waveLabel }) {
  const t = useT();
  if (!quarterHasTracking(quarter)) return null;
  const months = (quarter.monthlyKRs || []).filter((m) => !!m.tracking);
  return (
    <DashboardNode level={`${t("KR-веха")} · ${quarter.label} (${waveLabel})`} title={quarter.milestone || t("без названия")} tracking={quarter.tracking} horizonMonths={3} borderColor="#fecdd3">
      {months.map((m) => (
        <DashboardNode key={m.id} level={`${t("KR месяца")} · ${m.label}`} title={m.text || m.label} tracking={m.tracking} horizonMonths={1} borderColor="#e5e5e5" />
      ))}
    </DashboardNode>
  );
}

function DashboardOutput({ output }) {
  const t = useT();
  if (!outputHasTracking(output)) return null;
  const quarterBlocks = [];
  (output.waves || []).forEach((w) => {
    (w.quarters || []).forEach((q) => {
      if (quarterHasTracking(q)) quarterBlocks.push({ q, waveLabel: w.periodLabel });
    });
  });
  return (
    <DashboardNode level={`${output.label} · 18 ${t("мес.")}`} title={output.text || output.label} tracking={output.tracking} horizonMonths={18} borderColor="#e5e5e5">
      {quarterBlocks.map(({ q, waveLabel }) => <DashboardQuarter key={q.id} quarter={q} waveLabel={waveLabel} />)}
    </DashboardNode>
  );
}

function DashboardOutcome({ outcome, directory }) {
  const t = useT();
  if (!outcomeHasTracking(outcome)) return null;
  const outputs = (outcome.krOutputs || []).filter(outputHasTracking);
  return (
    <DashboardNode
      level={`KR Outcome · 18 ${t("мес.")}`} title={outcome.text || outcome.label}
      owner={ownerNameById(directory, outcome.objectiveOwner)} tracking={outcome.tracking} horizonMonths={18} borderColor="#fde68a"
    >
      {outputs.map((ko) => <DashboardOutput key={ko.id} output={ko} />)}
    </DashboardNode>
  );
}

const GANTT_GUIDE_STEP = 14;
function TreeGutter({ depth, guides, isLast }) {
  if (depth === 0) return null;
  const width = depth * GANTT_GUIDE_STEP;
  return (
    <div className="relative shrink-0" style={{ width, height: 30 }}>
      {guides.slice(0, depth - 1).map((cont, i) => cont && (
        <div key={i} className="absolute top-0 bottom-0" style={{ left: i * GANTT_GUIDE_STEP + GANTT_GUIDE_STEP / 2, width: 1, background: "#d4d4d4" }} />
      ))}
      <div className="absolute top-0" style={{ left: (depth - 1) * GANTT_GUIDE_STEP + GANTT_GUIDE_STEP / 2, width: 1, height: "50%", background: "#d4d4d4" }} />
      {!isLast && (
        <div className="absolute bottom-0" style={{ left: (depth - 1) * GANTT_GUIDE_STEP + GANTT_GUIDE_STEP / 2, width: 1, height: "50%", background: "#d4d4d4" }} />
      )}
      <div className="absolute" style={{ left: (depth - 1) * GANTT_GUIDE_STEP + GANTT_GUIDE_STEP / 2, top: "50%", width: GANTT_GUIDE_STEP / 2 + 4, height: 1, background: "#d4d4d4" }} />
    </div>
  );
}

const GANTT_DEPTH_LEVELS = [
  { value: 0, key: "Главные цели" },
  { value: 1, key: "KR Outcome" },
  { value: 2, key: "KR Output" },
  { value: 3, key: "Волны" },
  { value: 4, key: "Кварталы" },
  { value: 5, key: "Месяцы" },
  { value: 6, key: "Задачи" },
];

function GanttChart() {
  const t = useT();
  const [tracksData, loaded] = useAllTracksData();
  const [drafts, draftsLoaded] = useBitrixDraftRows();
  const [trackFilter, setTrackFilter] = useState("");
  const [maxDepth, setMaxDepth] = useState(6);

  if (!loaded) {
    return <div className="text-sm text-neutral-400 py-12 text-center">{t("Загрузка…")}</div>;
  }

  let rows = [];
  TRACKS.forEach((tr) => {
    let trackRows = buildFullGanttRows(tr.id, tr.name, tracksData[tr.id], maxDepth);
    if (draftsLoaded) trackRows = injectDraftRows(trackRows, drafts, tr.id, tr.name);
    if (trackRows.length > 0 && !trackFilter) {
      rows.push({ id: `header-${tr.id}`, isTrackHeader: true, trackId: tr.id, trackName: tr.name });
    }
    rows.push(...trackRows);
  });
  const filtered = trackFilter ? rows.filter((r) => r.trackId === trackFilter) : rows;
  const dataRows = filtered.filter((r) => !r.isTrackHeader);
  // no re-sort: rows come out of buildFullGanttRows already in strict parent→children order per track,
  // which is required for the tree-guide lines below to line up correctly.

  const depthControl = (
    <div className="flex flex-wrap gap-1.5 justify-center">
      {GANTT_DEPTH_LEVELS.map((lvl) => (
        <button
          key={lvl.value} onClick={() => setMaxDepth(lvl.value)}
          className="text-xs px-2.5 py-1 rounded-full border"
          style={maxDepth === lvl.value ? { background: "#171717", borderColor: "#171717", color: "#fff" } : { borderColor: "#e5e5e5", color: "#737373" }}
        >
          {t(lvl.key)}
        </button>
      ))}
    </div>
  );
  const trackControl = (
    <div className="flex flex-wrap gap-1.5 justify-center">
      {TRACKS.map((tr) => {
        const active = trackFilter === tr.id;
        return (
          <button
            key={tr.id} onClick={() => setTrackFilter(active ? "" : tr.id)}
            className="flex items-center gap-1.5 text-xs rounded-full px-2.5 py-1 border"
            style={active ? { background: TRACK_COLORS[tr.id] + "18", borderColor: TRACK_COLORS[tr.id], color: TRACK_COLORS[tr.id] } : { borderColor: "#e5e5e5", color: "#737373" }}
          >
            <span className="w-2 h-2 rounded-full shrink-0" style={{ background: TRACK_COLORS[tr.id] }} />
            {tr.name}
          </button>
        );
      })}
    </div>
  );

  if (dataRows.length === 0) {
    return (
      <div className="space-y-3">
        {trackControl}
        {depthControl}
        <div className="text-sm text-neutral-400 text-center py-10 border border-dashed border-neutral-200 rounded-xl">
          {t("Пока нет заполненных целей ни в одном треке.")}
        </div>
      </div>
    );
  }

  const rangeStartRaw = dataRows.reduce((min, r) => (r.start < min ? r.start : min), dataRows[0].start);
  const rangeEndRaw = dataRows.reduce((max, r) => (r.end > max ? r.end : max), dataRows[0].end);
  const padStart = new Date(rangeStartRaw + "T00:00:00"); padStart.setDate(padStart.getDate() - 7);
  const padEnd = new Date(rangeEndRaw + "T00:00:00"); padEnd.setDate(padEnd.getDate() + 7);
  const rangeStart = padStart.toISOString().slice(0, 10);
  const rangeEnd = padEnd.toISOString().slice(0, 10);
  const totalDays = Math.max(1, daysBetweenISO(rangeStart, rangeEnd));
  const todayPct = clamp01to100((daysBetweenISO(rangeStart, todayISO()) / totalDays) * 100);
  const ticks = buildMonthTicks(rangeStart, rangeEnd).map((tk) => ({ ...tk, pct: clamp01to100((daysBetweenISO(rangeStart, tk.iso) / totalDays) * 100) }));

  const Timeline = ({ children }) => (
    <div className="relative flex-1" style={{ height: 24 }}>
      {ticks.map((tk) => (
        <div key={tk.iso} className="absolute top-0 bottom-0 border-l border-neutral-100" style={{ left: `${tk.pct}%` }} />
      ))}
      <div className="absolute top-0 bottom-0 border-l" style={{ left: `${todayPct}%`, borderColor: "#171717" }} />
      {children}
    </div>
  );

  return (
    <div className="space-y-3">
      {trackControl}
      {depthControl}

      <div className="flex items-center gap-4 t11 flex-wrap justify-center">
        {["Главная цель", "KR Outcome", "KR Output", "Волна", "Квартал", "KR месяца", "Задача"].map((key, i) => {
          const tone = GANTT_DEPTH_TONE[i];
          return (
            <span
              key={key}
              className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded border"
              style={{ background: tone.bg, borderColor: tone.border, color: tone.text }}
            >
              {t(key)}
            </span>
          );
        })}
      </div>

      <div className="border border-neutral-200 rounded-xl overflow-hidden">
        <div className="flex items-end border-b border-neutral-200 bg-neutral-50 px-2 pt-2 pb-1">
          <div style={{ width: 270 }} className="shrink-0" />
          <Timeline>
            {ticks.map((tk) => (
              <span key={tk.iso} className="absolute t10 text-neutral-400" style={{ left: `${tk.pct}%`, bottom: 2, transform: "translateX(2px)" }}>
                {tk.label}
              </span>
            ))}
          </Timeline>
        </div>

        <div className="divide-y divide-neutral-100 overflow-y-auto" style={{ maxHeight: "70vh" }}>
          {filtered.map((r) => {
            if (r.isTrackHeader) {
              return (
                <div
                  key={r.id} className="flex items-center px-2 sticky top-0 z-10"
                  style={{ background: TRACK_COLORS[r.trackId] + "14", paddingTop: 6, paddingBottom: 6 }}
                >
                  <div className="flex items-center gap-1.5 text-xs font-semibold" style={{ color: TRACK_COLORS[r.trackId] }}>
                    <span className="w-2 h-2 rounded-full shrink-0" style={{ background: TRACK_COLORS[r.trackId] }} />
                    {r.trackName}
                  </div>
                </div>
              );
            }
            const barLeft = clamp01to100((daysBetweenISO(rangeStart, r.start) / totalDays) * 100);
            const barRight = clamp01to100((daysBetweenISO(rangeStart, r.end) / totalDays) * 100);
            const daysLeft = daysBetweenISO(todayISO(), r.end);
            const barColor = r.tracked ? deadlineUrgency(daysLeft).color : "#c4c4c4";
            const tone = r.draft ? { bg: "#FFFBEB", border: "#D97706", text: "#9A3412" } : (GANTT_DEPTH_TONE[r.depth] || TONES.task);
            return (
              <div key={r.id} className="flex items-center px-2 hover:bg-neutral-50">
                <TreeGutter depth={r.depth} guides={r.guides} isLast={r.isLast} />
                <div
                  style={{
                    width: 270 - r.depth * GANTT_GUIDE_STEP,
                    borderLeftColor: tone.border,
                    borderLeftStyle: r.draft ? "dashed" : "solid",
                    background: tone.bg,
                    color: tone.text,
                  }}
                  className="shrink-0 pr-2 pl-1.5 py-[5px] min-w-0 border-l-4"
                >
                  <div className="t10 flex items-center gap-1 opacity-80">
                    {r.draft && (
                      <span className="text-[9px] font-medium px-1 rounded shrink-0" style={{ background: "#FED7AA", color: "#9A3412" }}>
                        {t("черновик")}
                      </span>
                    )}
                    {r.depth === 0 && (
                      <span
                        className="w-1.5 h-1.5 rounded-full shrink-0"
                        style={{ background: TRACK_COLORS[r.trackId], boxShadow: "0 0 0 1.5px rgba(255,255,255,0.8)" }}
                      />
                    )}
                    <span className="truncate">{r.depth === 0 ? r.trackName + " · " : ""}{t(r.levelKey)}{r.levelExtra ? ` ${r.levelExtra}` : ""}</span>
                  </div>
                  {r.title && (
                    <div className="text-xs truncate" title={r.title}>
                      <TranslatedText text={r.title} />
                    </div>
                  )}
                  {r.draftEdit && (
                    <div className="text-xs truncate mt-0.5 flex items-center gap-1" style={{ color: "#9A3412" }} title={r.draftText}>
                      <span className="text-[9px] font-medium px-1 rounded shrink-0" style={{ background: "#FED7AA" }}>{t("черновик")}</span>
                      <span className="truncate">{r.draftText}</span>
                    </div>
                  )}
                </div>
                <Timeline>
                  {r.isTask ? (
                    <div
                      className="absolute rounded-full"
                      style={{ left: `calc(${barLeft}% - 4px)`, top: 6, width: 8, height: 8, background: barColor, transform: "rotate(45deg)" }}
                      title={formatDateRu(r.start)}
                    />
                  ) : (
                    <div
                      className="absolute rounded-full"
                      style={{ left: `${barLeft}%`, width: `${Math.max(0.6, barRight - barLeft)}%`, top: 5, height: 10, background: barColor }}
                      title={`${formatDateRu(r.start)} — ${formatDateRu(r.end)}${r.tracked ? " · " + (daysLeft < 0 ? t("просрочено на") + " " + Math.abs(daysLeft) : t("осталось") + " " + daysLeft) + " " + t("дн.") : ""}`}
                    />
                  )}
                </Timeline>
              </div>
            );
          })}
        </div>
      </div>

      <div className="flex items-center gap-4 t11 text-neutral-400 justify-center flex-wrap">
        <span><span className="inline-block w-2.5 h-2.5 rounded-full mr-1" style={{ background: "#16a34a" }} />{t("В графике")}</span>
        <span><span className="inline-block w-2.5 h-2.5 rounded-full mr-1" style={{ background: "#ca8a04" }} />{t("Скоро")}</span>
        <span><span className="inline-block w-2.5 h-2.5 rounded-full mr-1" style={{ background: "#dc2626" }} />{t("Критично")} / {t("Просрочено")}</span>
        <span><span className="inline-block w-2.5 h-2.5 rounded-full mr-1" style={{ background: "#c4c4c4" }} />{t("план (без трекинга)")}</span>
        <span><span className="inline-block w-2 h-2.5 mr-1" style={{ borderLeft: "1px solid #171717" }} />{t("сегодня")}</span>
      </div>
    </div>
  );
}

function ProjectsDashboard() {
  const t = useT();
  const [projects, , loaded] = useProjects();
  const [statuses, , statusesLoaded] = useProjectStatuses();

  if (!loaded || !statusesLoaded) {
    return <div className="text-sm text-neutral-400 py-12 text-center">{t("Загрузка…")}</div>;
  }
  if (projects.length === 0) {
    return <div className="text-sm text-neutral-400 text-center py-10 border border-dashed border-neutral-200 rounded-xl">{t("Пока нет проектов")}</div>;
  }

  const statusOf = (p) => statuses.find((s) => s.id === p.statusId) || statuses[0] || { id: "?", name: "—", color: "#94a3b8" };
  const byStatus = statuses.map((s) => ({ ...s, count: projects.filter((p) => p.statusId === s.id).length }));
  const divisions = Array.from(new Set(projects.map((p) => p.division).filter(Boolean))).sort();
  const functions = Array.from(new Set(projects.map((p) => p.function).filter(Boolean)));
  const topFunctions = functions
    .map((f) => ({ name: f, count: projects.filter((p) => p.function === f).length }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 6);
  const monthLabels = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"];
  const monthCounts = monthLabels.map((_, i) => projects.filter((p) => (p.activity || [])[i]).length);
  const maxMonth = Math.max(1, ...monthCounts);
  const currentMonthIdx = new Date().getMonth();
  const activeNow = monthCounts[currentMonthIdx];
  const maxDivisionTotal = Math.max(1, ...divisions.map((d) => projects.filter((p) => p.division === d).length));

  let donutOffset = 0;
  const donutStops = byStatus.filter((s) => s.count > 0).map((s) => {
    const pct = (s.count / projects.length) * 100;
    const stop = `${s.color} ${donutOffset}% ${donutOffset + pct}%`;
    donutOffset += pct;
    return stop;
  });

  return (
    <div className="space-y-4">
      <div className="grid gap-2" style={{ gridTemplateColumns: `repeat(${Math.min(5, byStatus.length + 2)}, minmax(0,1fr))` }}>
        <div className="rounded-xl border border-neutral-200 p-3">
          <div className="text-2xl font-semibold text-neutral-800">{projects.length}</div>
          <div className="t11 text-neutral-400 uppercase tracking-wide">{t("Всего проектов")}</div>
        </div>
        {byStatus.map((s) => (
          <div key={s.id} className="rounded-xl border border-neutral-200 p-3" style={{ borderLeftWidth: 3, borderLeftColor: s.color }}>
            <div className="text-2xl font-semibold" style={{ color: s.color }}>{s.count}</div>
            <div className="t11 text-neutral-400 uppercase tracking-wide truncate">{s.name}</div>
            <div className="t10 text-neutral-400">{projects.length ? Math.round((s.count / projects.length) * 100) : 0}%</div>
          </div>
        ))}
        <div className="rounded-xl border border-neutral-200 p-3">
          <div className="text-2xl font-semibold text-neutral-800">{activeNow}</div>
          <div className="t11 text-neutral-400 uppercase tracking-wide">{t("Активны в этом месяце")}</div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
        <div className="rounded-xl border border-neutral-200 p-3">
          <div className="t11 font-semibold uppercase tracking-wide text-neutral-500 mb-2">{t("Загрузка по дивизионам")}</div>
          <div className="space-y-1.5">
            {divisions.map((d) => {
              const dProjects = projects.filter((p) => p.division === d);
              return (
                <div key={d} className="flex items-center gap-2">
                  <span className="text-xs text-neutral-600 w-28 shrink-0 truncate">{d}</span>
                  <div className="flex-1 h-3 rounded-full overflow-hidden bg-neutral-100 flex" style={{ maxWidth: `${(dProjects.length / maxDivisionTotal) * 100}%` }}>
                    {statuses.map((s) => {
                      const cnt = dProjects.filter((p) => p.statusId === s.id).length;
                      if (!cnt) return null;
                      return <div key={s.id} style={{ width: `${(cnt / dProjects.length) * 100}%`, background: s.color }} title={s.name} />;
                    })}
                  </div>
                  <span className="t11 text-neutral-400 w-6 text-right shrink-0">{dProjects.length}</span>
                </div>
              );
            })}
          </div>
          <div className="flex flex-wrap gap-2 mt-2 pt-2 border-t border-neutral-100">
            {statuses.map((s) => (
              <span key={s.id} className="t10 text-neutral-400 flex items-center gap-1">
                <span className="w-2 h-2 rounded-full" style={{ background: s.color }} /> {s.name}
              </span>
            ))}
          </div>
        </div>

        <div className="rounded-xl border border-neutral-200 p-3 flex items-center gap-4">
          <div>
            <div className="t11 font-semibold uppercase tracking-wide text-neutral-500 mb-2">{t("Распределение статусов")}</div>
            <div
              className="relative rounded-full shrink-0"
              style={{ width: 130, height: 130, background: donutStops.length ? `conic-gradient(${donutStops.join(",")})` : "#e5e5e5" }}
            >
              <div className="absolute rounded-full bg-white flex flex-col items-center justify-center" style={{ inset: 22 }}>
                <span className="text-lg font-semibold text-neutral-800">{projects.length}</span>
                <span className="t10 text-neutral-400 uppercase">{t("проектов")}</span>
              </div>
            </div>
          </div>
          <div className="space-y-1 flex-1 min-w-0">
            {byStatus.map((s) => (
              <div key={s.id} className="flex items-center gap-1.5 t11">
                <span className="w-2 h-2 rounded-full shrink-0" style={{ background: s.color }} />
                <span className="text-neutral-600 truncate flex-1">{s.name}</span>
                <span className="font-medium text-neutral-800 shrink-0">{s.count}</span>
                <span className="text-neutral-400 shrink-0 w-10 text-right">{projects.length ? Math.round((s.count / projects.length) * 100) : 0}%</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
        <div className="rounded-xl border border-neutral-200 p-3">
          <div className="t11 font-semibold uppercase tracking-wide text-neutral-500 mb-2">{t("Активных проектов / месяц")}</div>
          <div className="flex items-end gap-1.5" style={{ height: 90 }}>
            {monthCounts.map((c, i) => (
              <div key={i} className="flex-1 flex flex-col items-center justify-end gap-1">
                <span className="t10 text-neutral-400">{c || ""}</span>
                <div
                  className="w-full rounded-sm"
                  style={{ height: `${(c / maxMonth) * 60 + (c ? 4 : 2)}px`, background: i === currentMonthIdx ? "#4f46e5" : "#d4d4d4" }}
                />
                <span className="t10 text-neutral-400">{monthLabels[i]}</span>
              </div>
            ))}
          </div>
        </div>

        <div className="rounded-xl border border-neutral-200 p-3">
          <div className="t11 font-semibold uppercase tracking-wide text-neutral-500 mb-2">{t("Топ функций по кол-ву проектов")}</div>
          <div className="space-y-1.5">
            {topFunctions.map((f) => (
              <div key={f.name} className="flex items-center gap-2">
                <span className="text-xs text-neutral-600 w-24 shrink-0 truncate">{f.name}</span>
                <div className="flex-1 h-3 bg-neutral-100 rounded-full overflow-hidden">
                  <div className="h-full rounded-full bg-indigo-400" style={{ width: `${(f.count / (topFunctions[0]?.count || 1)) * 100}%` }} />
                </div>
                <span className="t11 text-neutral-400 w-5 text-right shrink-0">{f.count}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

const PERM_LEVELS = ["none", "view", "edit"];
const PERM_LEVEL_STYLE = {
  none: { bg: "#f5f5f5", color: "#a3a3a3", label: "—" },
  view: { bg: "#dbeafe", color: "#1d4ed8", label: "Видит" },
  edit: { bg: "#dcfce7", color: "#15803d", label: "Меняет" },
};

function PrivilegesManager({ auth }) {
  const t = useT();
  const [roles, setRoles] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState("");
  const [pwDrafts, setPwDrafts] = useState({});
  const [newRoleName, setNewRoleName] = useState("");

  const reload = useCallback(async () => {
    const { data, error } = await supabase.rpc("okr_list_roles");
    if (!error && data) setRoles(data);
    setLoaded(true);
  }, []);

  useEffect(() => { reload(); }, [reload]);

  const cycleLevel = async (role, moduleKey) => {
    const current = (role.permissions && role.permissions[moduleKey]) || "none";
    const next = PERM_LEVELS[(PERM_LEVELS.indexOf(current) + 1) % PERM_LEVELS.length];
    const newPermissions = { ...role.permissions, [moduleKey]: next };
    setRoles((prev) => prev.map((r) => (r.id === role.id ? { ...r, permissions: newPermissions } : r)));
    const { error } = await supabase.rpc("okr_upsert_role", {
      p_id: role.id, p_name: role.name, p_permissions: newPermissions, p_sort_order: role.sort_order || 0,
    });
    if (error) setError(error.message);
  };

  const savePassword = async (roleId) => {
    const pw = (pwDrafts[roleId] || "").trim();
    if (!pw) return;
    const { error } = await supabase.rpc("okr_set_role_password", { p_role_id: roleId, p_new_password: pw });
    if (error) setError(error.message);
    else setPwDrafts((prev) => ({ ...prev, [roleId]: "" }));
  };

  const addRole = async () => {
    if (!newRoleName.trim()) return;
    const id = "role-" + newId();
    const emptyPerms = Object.fromEntries(MODULES.map((m) => [m.key, "none"]));
    const { error } = await supabase.rpc("okr_upsert_role", {
      p_id: id, p_name: newRoleName.trim(), p_permissions: emptyPerms, p_sort_order: roles.length + 1,
    });
    if (error) setError(error.message);
    else { setNewRoleName(""); reload(); }
  };

  const removeRole = async (roleId) => {
    const { error } = await supabase.rpc("okr_delete_role", { p_id: roleId });
    if (error) setError(error.message);
    else reload();
  };

  if (!loaded) {
    return <div className="text-sm text-neutral-400 py-12 text-center">{t("Загрузка…")}</div>;
  }

  return (
    <div className="space-y-4">
      <div className="text-sm text-neutral-500">
        {t("Клик по ячейке переключает право по кругу: — → Видит → Меняет → —. Изменения сохраняются сразу.")}
      </div>
      {error && <div className="text-xs text-red-600">{error}</div>}

      <div className="overflow-x-auto border border-neutral-200 rounded-xl">
        <table className="text-xs" style={{ borderCollapse: "collapse", minWidth: "100%" }}>
          <thead>
            <tr>
              <th className="text-left px-2 py-2 bg-neutral-50 sticky left-0" style={{ minWidth: 200 }}>{t("Модуль")}</th>
              {roles.map((r) => (
                <th key={r.id} className="px-2 py-2 bg-neutral-50 font-medium text-neutral-700" style={{ minWidth: 110 }}>
                  <div className="flex items-center justify-center gap-1">
                    <span className="text-center">{r.name}</span>
                    <button onClick={() => removeRole(r.id)} className="text-neutral-300 hover:text-red-500 shrink-0">
                      <Trash2 size={11} />
                    </button>
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {MODULES.map((m, mi) => (
              <tr key={m.key} className={mi % 2 === 0 ? "bg-white" : "bg-neutral-50"}>
                <td className="px-2 py-1.5 text-neutral-700 sticky left-0" style={{ background: "inherit" }}>{t(m.label)}</td>
                {roles.map((r) => {
                  const level = (r.permissions && r.permissions[m.key]) || "none";
                  const style = PERM_LEVEL_STYLE[level];
                  return (
                    <td key={r.id} className="px-2 py-1.5 text-center">
                      <button
                        onClick={() => cycleLevel(r, m.key)}
                        className="text-xs px-2 py-1 rounded-full font-medium w-full"
                        style={{ background: style.bg, color: style.color }}
                      >
                        {t(style.label)}
                      </button>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="border border-neutral-200 rounded-xl p-3 space-y-2">
        <div className="text-xs font-medium uppercase tracking-wide text-neutral-600">{t("Пароли ролей")}</div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
          {roles.map((r) => (
            <div key={r.id} className="flex items-center gap-1.5">
              <span className="text-xs text-neutral-600 flex-1 min-w-0 truncate">{r.name}</span>
              <input
                type="text" value={pwDrafts[r.id] || ""}
                onChange={(e) => setPwDrafts((prev) => ({ ...prev, [r.id]: e.target.value }))}
                placeholder={t("новый пароль")}
                className="text-xs border border-neutral-200 rounded-md px-2 py-1 bg-neutral-50"
                style={{ width: 120 }}
              />
              <button onClick={() => savePassword(r.id)} className="text-xs border border-neutral-200 rounded-md px-2 py-1 hover:bg-neutral-50 shrink-0">
                {t("Сохранить")}
              </button>
            </div>
          ))}
        </div>
      </div>

      <div className="border border-neutral-200 rounded-xl p-3">
        <div className="text-xs font-medium uppercase tracking-wide text-neutral-600 mb-2">{t("Новая роль")}</div>
        <div className="flex gap-1.5">
          <input
            value={newRoleName} onChange={(e) => setNewRoleName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && addRole()}
            placeholder={t("Название роли")}
            className="flex-1 text-sm border border-neutral-200 rounded-md px-2 py-1.5 outline-none"
          />
          <button onClick={addRole} className="px-2.5 rounded-md bg-neutral-900 text-white text-sm hover:bg-neutral-800">
            <Plus size={14} />
          </button>
        </div>
      </div>
    </div>
  );
}

function TrackingDashboard() {
  const t = useT();
  const [tracksData, loaded] = useAllTracksData();
  const directory = useDirectoryList();
  const [trackFilter, setTrackFilter] = useState("");

  if (!loaded) {
    return <div className="text-sm text-neutral-400 py-12 text-center">{t("Загрузка…")}</div>;
  }

  let items = [];
  TRACKS.forEach((tr) => {
    items.push(...collectTrackers(tr.id, tr.name, tracksData[tr.id], directory));
  });
  items = items.map((it) => ({ ...it, ...computeTracking(it.tracking) }));
  const filteredItems = trackFilter ? items.filter((it) => it.trackId === trackFilter) : items;

  const avgPct = filteredItems.length ? filteredItems.reduce((s, it) => s + it.pct, 0) / filteredItems.length : 0;
  const counts = { onTrack: 0, atRisk: 0, behind: 0 };
  filteredItems.forEach((it) => {
    if (it.pct >= 0.9) counts.onTrack++;
    else if (it.pct >= 0.5) counts.atRisk++;
    else counts.behind++;
  });

  const tracksToShow = TRACKS.filter((tr) => !trackFilter || tr.id === trackFilter).filter((tr) =>
    (tracksData[tr.id].ultimateObjectives || []).some(ultimateObjectiveHasTracking)
  );

  return (
    <div className="space-y-3">
      <div className="grid grid-cols-4 gap-2">
        <div className="rounded-xl border border-neutral-200 p-2.5 text-center">
          <div className="text-lg font-semibold text-neutral-800">{filteredItems.length}</div>
          <div className="t11 text-neutral-400">{t("KR под трекингом")}</div>
        </div>
        <div className="rounded-xl border border-neutral-200 p-2.5 text-center">
          <div className="text-lg font-semibold text-neutral-800">{Math.round(avgPct * 100)}%</div>
          <div className="t11 text-neutral-400">{t("средний прогресс")}</div>
        </div>
        <div className="rounded-xl border border-neutral-200 p-2.5 text-center">
          <div className="text-lg font-semibold text-emerald-600">{counts.onTrack}</div>
          <div className="t11 text-neutral-400">{t("на треке")}</div>
        </div>
        <div className="rounded-xl border border-neutral-200 p-2.5 text-center">
          <div className="text-lg font-semibold text-red-500">{counts.behind}</div>
          <div className="t11 text-neutral-400">{t("отстают")}</div>
        </div>
      </div>

      <div className="flex flex-wrap gap-1.5 justify-center">
        {TRACKS.map((tr) => {
          const active = trackFilter === tr.id;
          return (
            <button
              key={tr.id}
              onClick={() => setTrackFilter(active ? "" : tr.id)}
              className="flex items-center gap-1.5 text-xs rounded-full px-2.5 py-1 border"
              style={
                active
                  ? { background: TRACK_COLORS[tr.id] + "18", borderColor: TRACK_COLORS[tr.id], color: TRACK_COLORS[tr.id] }
                  : { borderColor: "#e5e5e5", color: "#737373" }
              }
            >
              <span className="w-2 h-2 rounded-full shrink-0" style={{ background: TRACK_COLORS[tr.id] }} />
              {tr.name}
            </button>
          );
        })}
      </div>

      {tracksToShow.length === 0 ? (
        <div className="text-sm text-neutral-400 text-center py-10 border border-dashed border-neutral-200 rounded-xl">
          {t("Пока нет KR с включённым трекингом. Кнопка «добавить трекинг прогресса» доступна у KR Outcome, KR Output, KR-вехи квартала и KR месяца в любом треке.")}
        </div>
      ) : (
        <div className="space-y-2">
          {tracksToShow.map((tr) => (
            <div key={tr.id} className="rounded-xl border border-neutral-200 p-3" style={{ borderLeftWidth: 4, borderLeftColor: TRACK_COLORS[tr.id] }}>
              <div className="text-sm font-semibold mb-2" style={{ color: TRACK_COLORS[tr.id] }}>{tr.name}</div>
              <div className="space-y-2">
                {(tracksData[tr.id].ultimateObjectives || []).filter(ultimateObjectiveHasTracking).map((uo) => (
                  uo.krOutcomes.filter(outcomeHasTracking).map((o) => (
                    <DashboardOutcome key={o.id} outcome={o} directory={directory} />
                  ))
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ExportImportBar({ directory, setDirectory, onImported }) {
  const t = useT();
  const fileInputRef = useRef(null);
  const [pendingImport, setPendingImport] = useState(null);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");

  const handleExport = async () => {
    setBusy(true);
    setMessage("");
    try {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildInstructionsRows()), "0. Инструкция");

      const trackRows = [];
      const structureRows = [];
      const taskRows = [];
      for (const t of TRACKS) {
        let data;
        try {
          const res = await window.storage.get(`okr-track:${t.id}`, false);
          data = res ? migrateTrackData(JSON.parse(res.value)) : createTrackData();
        } catch {
          data = createTrackData();
        }
        trackRows.push(buildTrackRow(t.name, data));
        structureRows.push(...buildStructureRows(t.name, data, directory));
        taskRows.push(...buildTaskRows(t.name, data, directory));
      }
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(trackRows), "1. Треки");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(structureRows), "2. OKR - Структура");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(taskRows), "3. OKR - Задачи");

      let projects = [];
      try {
        const res = await window.storage.get("okr-projects", false);
        projects = res ? JSON.parse(res.value) : PROJECTS_SEED;
      } catch {
        projects = PROJECTS_SEED;
      }
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildProjectRows(projects)), "4. Проекты");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildProjectKrRows(projects)), "5. KR проектов");

      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildDirectoryRows(directory || [])), "6. Справочник");

      let links = [];
      try {
        const res = await window.storage.get("okr-links", false);
        links = res ? JSON.parse(res.value) : [];
      } catch {
        links = [];
      }
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildLinkRows(links)), "7. Связи");

      XLSX.writeFile(wb, `Coral_Club_OKR_${new Date().toISOString().slice(0, 10)}.xlsx`);
      setMessage(t("Готово — файл скачан."));
    } catch (e) {
      setMessage(t("Ошибка экспорта: ") + e.message);
    }
    setBusy(false);
  };

  const handleFileSelected = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setBusy(true);
    setMessage("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = (name) => (wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: "" }) : []);

      const trackRows = sheet("1. Треки");
      const structureRows = sheet("2. OKR - Структура");
      const taskRows = sheet("3. OKR - Задачи");
      const projectRows = sheet("4. Проекты");
      const krRows = sheet("5. KR проектов");
      const dirRows = sheet("6. Справочник");
      const linkRows = sheet("7. Связи");

      let currentDirectory = directory || [];
      const newDirEntries = [];
      const baseDirectory = rebuildDirectoryFromSheet(dirRows, currentDirectory);

      const existingTracks = {};
      for (const t of TRACKS) {
        try {
          const res = await window.storage.get(`okr-track:${t.id}`, false);
          existingTracks[t.id] = res ? migrateTrackData(JSON.parse(res.value)) : createTrackData();
        } catch {
          existingTracks[t.id] = createTrackData();
        }
      }

      const tracksById = {};
      TRACKS.forEach((t) => {
        const trow = trackRows.find((r) => normKey(r["Трек"]) === normKey(t.name));
        tracksById[t.id] = rebuildTrackFromSheets(t.name, trow, structureRows, taskRows, existingTracks[t.id], baseDirectory, newDirEntries);
      });

      let existingProjects = [];
      try {
        const res = await window.storage.get("okr-projects", false);
        existingProjects = res ? JSON.parse(res.value) : [];
      } catch {
        existingProjects = [];
      }
      const projects = rebuildProjectsFromSheets(projectRows, krRows, existingProjects);

      const finalDirectory = [...baseDirectory, ...newDirEntries];
      const newLinks = rebuildLinksFromSheet(linkRows);

      setPendingImport({ tracksById, projects, directory: finalDirectory, links: newLinks, fileName: file.name });
    } catch (e) {
      setMessage(t("Не удалось прочитать файл: ") + e.message);
    }
    setBusy(false);
  };


  const applyImport = async () => {
    if (!pendingImport) return;
    setBusy(true);
    try {
      await Promise.all([
        ...TRACKS.map((t) => window.storage.set(`okr-track:${t.id}`, JSON.stringify(pendingImport.tracksById[t.id]), false)),
        window.storage.set("okr-projects", JSON.stringify(pendingImport.projects), false),
        window.storage.set("okr-directory", JSON.stringify(pendingImport.directory), false),
        window.storage.set("okr-links", JSON.stringify(pendingImport.links), false),
      ]);
      setDirectory(pendingImport.directory);
      setMessage(t("Импортировано — данные обновлены."));
      setPendingImport(null);
      onImported();
    } catch (e) {
      setMessage(t("Ошибка импорта: ") + e.message);
    }
    setBusy(false);
  };

  return (
    <div className="mb-4 border border-neutral-200 rounded-xl p-3 space-y-2">
      <div className="flex items-center gap-2 flex-wrap">
        <span className="text-xs font-medium uppercase tracking-wide text-neutral-600 mr-1">Excel</span>
        <button
          onClick={handleExport}
          disabled={busy}
          className="flex items-center gap-1.5 text-xs bg-neutral-900 text-white rounded-md px-2.5 py-1.5 disabled:opacity-50"
        >
          <Download size={12} /> {t("Скачать Excel")}
        </button>
        <button
          onClick={() => fileInputRef.current && fileInputRef.current.click()}
          disabled={busy}
          className="flex items-center gap-1.5 text-xs border border-neutral-200 rounded-md px-2.5 py-1.5 hover:bg-neutral-50 disabled:opacity-50"
        >
          <Upload size={12} /> {t("Загрузить Excel")}
        </button>
        <input ref={fileInputRef} type="file" accept=".xlsx" className="hidden" onChange={handleFileSelected} />
        {message && <span className="text-xs text-neutral-400">{message}</span>}
      </div>

      {pendingImport && (
        <div className="bg-amber-50 border border-amber-200 rounded-lg p-2.5 text-xs text-amber-800 flex items-center justify-between gap-3 flex-wrap">
          <span>
            {t("Файл")} «{pendingImport.fileName}» {t("прочитан. Импорт заменит ВСЕ текущие данные (все треки, проекты, справочник, связи). Продолжить?")}
          </span>
          <div className="flex gap-1.5 shrink-0">
            <button onClick={applyImport} disabled={busy} className="bg-amber-600 text-white rounded-md px-2.5 py-1">{t("Заменить")}</button>
            <button onClick={() => setPendingImport(null)} disabled={busy} className="border border-amber-300 rounded-md px-2.5 py-1">{t("Отмена")}</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ---- Google Identity Services loader ----
// ---- Supabase client ----
const SUPABASE_URL = "https://awwhrlenhqoreagzfymn.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_t94mAXN3a3ql6fGuoIf8Lg_HMrcSGiy";
const supabase = createSupabaseClient(SUPABASE_URL, SUPABASE_ANON_KEY);
const ROLE_SESSION_KEY = "okr-role-session";

// ---- Bitrix24 import: parsing + hierarchy targeting ----
// Bitrix's "export to Excel" is actually an HTML table saved with an .xls extension —
// this reads it with the browser's own HTML parser rather than treating it as a real xlsx.
function parseBitrixExport(text) {
  const doc = new DOMParser().parseFromString(text, "text/html");
  const table = doc.querySelector("table");
  if (!table) return [];
  const rows = [...table.querySelectorAll("tr")];
  if (rows.length === 0) return [];
  const header = [...rows[0].querySelectorAll("td,th")].map((c) => c.textContent.trim().toLowerCase());
  const idx = {
    name: header.findIndex((h) => h.includes("назв")),
    deadline: header.findIndex((h) => h.includes("срок") || h.includes("дедлайн")),
    tags: header.findIndex((h) => h.includes("тег")),
    bitrixId: header.findIndex((h) => h.includes("id") || h.includes("заявк") || h.includes("№")),
  };
  const out = [];
  rows.slice(1).forEach((r, i) => {
    const cells = [...r.querySelectorAll("td")].map((c) => c.textContent.trim());
    if (cells.length === 0) return;
    const rawText = idx.name >= 0 ? cells[idx.name] : cells[0];
    if (!rawText) return;
    const deadline = idx.deadline >= 0 ? cells[idx.deadline] : "";
    const tagsCell = idx.tags >= 0 ? cells[idx.tags] : "";
    const tags = tagsCell.split(",").map((s) => s.replace(/^#/, "").trim()).filter(Boolean);
    const bitrixTaskId = idx.bitrixId >= 0 ? cells[idx.bitrixId] : "";
    out.push({ rowId: `bx_${i}_${newId()}`, rawText, deadline, tags, bitrixTaskId });
  });
  return out;
}

const TAG_TRACK_ALIASES = {
  "health-os": "health-os", "healthos": "health-os", "health_os": "health-os",
  "growth-model": "growth-model", "growthmodel": "growth-model", "growth_model": "growth-model",
  "prime-growth": "prime-growth", "primegrowth": "prime-growth", "prime_growth": "prime-growth",
  "coral-evo": "coral-evo", "coralevo": "coral-evo", "coral_evo": "coral-evo", "coralevolution": "coral-evo",
  "it-model": "it-model", "itmodel": "it-model", "it_model": "it-model",
};
function guessTrackFromTags(tags) {
  for (const raw of tags) {
    const norm = raw.toLowerCase().replace(/[\s_-]+/g, "-");
    if (TAG_TRACK_ALIASES[norm]) return TAG_TRACK_ALIASES[norm];
    const squashed = raw.toLowerCase().replace(/[^a-zа-я0-9]/g, "");
    if (TAG_TRACK_ALIASES[squashed]) return TAG_TRACK_ALIASES[squashed];
  }
  return "";
}

// Depth = how many cascading pickers (UO → KR Outcome → KR Output → Волна → Квартал → Месяц)
// are needed to pin down where this level lives in the tree.
const IMPORT_LEVELS = [
  { key: "mission", label: "Главная цель", depth: 1 },
  { key: "initiative", label: "Стратегическая инициатива", depth: 2 },
  { key: "wave", label: "Objective волны (6 мес.)", depth: 4 },
  { key: "quarterInitiative", label: "Инициатива квартала (3 мес.)", depth: 5 },
  { key: "month", label: "KR месяца", depth: 6 },
  { key: "task", label: "Задача", depth: 6 },
];
function guessImportLevel(rawText) {
  const s = rawText.toLowerCase();
  if (s.includes("ultimate objective") || s.includes("стратегическая цель трека")) return "mission";
  if (s.includes("ultimate initiative") || s.includes("стратегическая инициатива")) return "initiative";
  if (s.includes("project-initiative") || s.includes("project- initiative") || s.includes("6 мес")) return "wave";
  if (s.includes("objective на 3 мес") || s.includes("объектив квартала")) return "quarterInitiative";
  if (s.includes("стратегическая миссия трека")) return "skip";
  return "task";
}

// Finds an already-imported task by its Bitrix task ID — walks every month of every branch in
// the track. Used so re-importing the same (or an updated) Bitrix export recognizes tasks that
// were already brought in before, instead of relying on fuzzy text matching or manual re-placement.
function findTaskByBitrixId(trackData, bitrixTaskId) {
  if (!bitrixTaskId) return null;
  for (const uo of trackData.ultimateObjectives) {
    for (const o of uo.krOutcomes) {
      for (const ko of o.krOutputs) {
        for (const w of ko.waves) {
          for (const q of w.quarters) {
            for (const m of q.monthlyKRs) {
              const task = m.tasks.find((tk) => tk.bitrixId && tk.bitrixId === bitrixTaskId);
              if (task) {
                return {
                  task,
                  path: { uoId: uo.id, outcomeId: o.id, outputId: ko.id, waveId: w.id, quarterId: q.id, monthId: m.id },
                };
              }
            }
          }
        }
      }
    }
  }
  return null;
}

// ---- Conflict handling: never silently overwrite existing content. If the target already
// has something different, the import text is appended (not replacing it) with a clear
// marker, so both versions stay visible everywhere that field renders — track editor,
// combined tree, Gantt — until a human compares them and deletes the one that's redundant.
const IMPORT_FIELD_CONFLICT_MARKER = "\n\n⚠ ИЗ BITRIX — сравните и уберите лишнее:\n";
const IMPORT_TASK_DUPLICATE_PREFIX = "⚠ Возможный дубль — сравните: ";

function normText(s) {
  return (s || "").trim().toLowerCase().replace(/\s+/g, " ");
}
// For field-level levels (mission/initiative/wave/quarterInitiative/month): is there already
// different, non-empty content sitting at this exact target?
function detectFieldConflict(trackData, levelKey, path) {
  if (levelKey === "task") return null;
  const uo = trackData.ultimateObjectives.find((u) => u.id === path.uoId);
  if (levelKey === "mission") return uo && uo.text ? uo.text : null;
  if (path.uoId === "__new__" || !uo) return null;
  const outcome = uo.krOutcomes.find((o) => o.id === path.outcomeId);
  if (levelKey === "initiative") return outcome && outcome.initiativeText ? outcome.initiativeText : null;
  if (path.outcomeId === "__new__" || !outcome) return null;
  const output = outcome.krOutputs.find((k) => k.id === path.outputId);
  const wave = output && output.waves.find((w) => w.id === path.waveId);
  if (levelKey === "wave") return wave && wave.objective6mo ? wave.objective6mo : null;
  if (!wave) return null;
  const quarter = wave.quarters.find((q) => q.id === path.quarterId);
  if (levelKey === "quarterInitiative") return quarter && quarter.initiative3mo ? quarter.initiative3mo : null;
  if (!quarter) return null;
  const month = quarter.monthlyKRs.find((m) => m.id === path.monthId);
  if (levelKey === "month") return month && month.text ? month.text : null;
  return null;
}
// For "task": is there already a task with the same (normalized) text in the target month?
function detectTaskDuplicate(trackData, path, text) {
  const uo = trackData.ultimateObjectives.find((u) => u.id === path.uoId);
  const outcome = uo && uo.krOutcomes.find((o) => o.id === path.outcomeId);
  const output = outcome && outcome.krOutputs.find((k) => k.id === path.outputId);
  const wave = output && output.waves.find((w) => w.id === path.waveId);
  const quarter = wave && wave.quarters.find((q) => q.id === path.quarterId);
  const month = quarter && quarter.monthlyKRs.find((m) => m.id === path.monthId);
  if (!month) return false;
  const n = normText(text);
  return month.tasks.some((tk) => normText(tk.text) === n && n !== "");
}

function emptyTargetPath() {
  return { uoId: "", outcomeId: "", outputId: "", waveId: "", quarterId: "", monthId: "" };
}
function pathIsComplete(levelKey, path) {
  const level = IMPORT_LEVELS.find((l) => l.key === levelKey);
  const depth = level ? level.depth : 6;
  if (depth >= 1 && !path.uoId) return false;
  if (depth >= 2 && !path.outcomeId) return false;
  if (depth >= 4 && !path.outputId) return false;
  if (depth >= 4 && !path.waveId) return false;
  if (depth >= 5 && !path.quarterId) return false;
  if (depth >= 6 && !path.monthId) return false;
  return true;
}
// Writes one reviewed Bitrix row into a (cloned) track's data tree. Returns null if the
// chosen placement no longer resolves (e.g. hit a cap) so the caller can skip it safely.
function applyImportRow(trackData, levelKey, path, text, bitrixTaskId) {
  const data = JSON.parse(JSON.stringify(trackData));
  let uo;
  if (path.uoId === "__new__") {
    if (data.ultimateObjectives.length >= MAX_ULTIMATE_OBJECTIVES) return null;
    uo = createUltimateObjective(data.ultimateObjectives.length + 1);
    data.ultimateObjectives.push(uo);
  } else {
    uo = data.ultimateObjectives.find((u) => u.id === path.uoId);
  }
  if (!uo) return null;
  if (levelKey === "mission") {
    uo.text = uo.text && uo.text.trim() && uo.text.trim() !== text.trim() ? uo.text + IMPORT_FIELD_CONFLICT_MARKER + text : text;
    return data;
  }

  let outcome;
  if (path.outcomeId === "__new__") {
    if (uo.krOutcomes.length >= MAX_KR_OUTCOMES) return null;
    outcome = createKROutcome(uo.krOutcomes.length + 1);
    uo.krOutcomes.push(outcome);
  } else {
    outcome = uo.krOutcomes.find((o) => o.id === path.outcomeId);
  }
  if (!outcome) return null;
  if (levelKey === "initiative") {
    outcome.initiativeText =
      outcome.initiativeText && outcome.initiativeText.trim() && outcome.initiativeText.trim() !== text.trim()
        ? outcome.initiativeText + IMPORT_FIELD_CONFLICT_MARKER + text
        : text;
    return data;
  }

  const output = outcome.krOutputs.find((k) => k.id === path.outputId);
  if (!output) return null;
  const wave = output.waves.find((w) => w.id === path.waveId);
  if (!wave) return null;
  if (levelKey === "wave") {
    wave.objective6mo =
      wave.objective6mo && wave.objective6mo.trim() && wave.objective6mo.trim() !== text.trim()
        ? wave.objective6mo + IMPORT_FIELD_CONFLICT_MARKER + text
        : text;
    return data;
  }

  const quarter = wave.quarters.find((q) => q.id === path.quarterId);
  if (!quarter) return null;
  if (levelKey === "quarterInitiative") {
    quarter.initiative3mo =
      quarter.initiative3mo && quarter.initiative3mo.trim() && quarter.initiative3mo.trim() !== text.trim()
        ? quarter.initiative3mo + IMPORT_FIELD_CONFLICT_MARKER + text
        : text;
    return data;
  }

  const month = quarter.monthlyKRs.find((m) => m.id === path.monthId);
  if (!month) return null;
  if (levelKey === "month") {
    month.text =
      month.text && month.text.trim() && month.text.trim() !== text.trim()
        ? month.text + IMPORT_FIELD_CONFLICT_MARKER + text
        : text;
    return data;
  }
  if (levelKey === "task") {
    if (month.tasks.length >= MAX_TASKS) return null;
    const isDup = detectTaskDuplicate(data, path, text);
    month.tasks.push({ id: newId(), text: isDup ? IMPORT_TASK_DUPLICATE_PREFIX + text : text, ownerId: "", bitrixId: bitrixTaskId || "" });
    return data;
  }
  return null;
}

// Flat, searchable list of every valid placement for a given level — each entry shows the
// real breadcrumb (Главная цель › KR Outcome › … ) with the actual live text at each step,
// so browsing feels like browsing the track itself, not guessing through blind dropdowns.
function buildPlacementOptions(trackData, levelKey, t) {
  const level = IMPORT_LEVELS.find((l) => l.key === levelKey);
  const depth = level ? level.depth : 6;
  const options = [];
  trackData.ultimateObjectives.forEach((uo) => {
    const uoLabel = `${uo.label}${uo.text ? " — " + truncate(uo.text, 40) : ""}`;
    if (depth === 1) {
      options.push({ crumb: uoLabel, path: { uoId: uo.id, outcomeId: "", outputId: "", waveId: "", quarterId: "", monthId: "" } });
      return;
    }
    uo.krOutcomes.forEach((o) => {
      const oLabel = `${uoLabel} › ${o.label}${o.text ? " — " + truncate(o.text, 40) : ""}`;
      if (depth === 2) {
        options.push({ crumb: oLabel, path: { uoId: uo.id, outcomeId: o.id, outputId: "", waveId: "", quarterId: "", monthId: "" } });
        return;
      }
      o.krOutputs.forEach((ko) => {
        const koLabel = `${oLabel} › ${ko.label}`;
        ko.waves.forEach((w, wi) => {
          const wLabel = `${koLabel} › ${t("Волна")} ${wi + 1}${w.objective6mo ? " — " + truncate(w.objective6mo, 40) : ""}`;
          if (depth === 4) {
            options.push({ crumb: wLabel, path: { uoId: uo.id, outcomeId: o.id, outputId: ko.id, waveId: w.id, quarterId: "", monthId: "" } });
            return;
          }
          w.quarters.forEach((q) => {
            const qLabel = `${wLabel} › ${t("Квартал")} ${q.label}${q.initiative3mo ? " — " + truncate(q.initiative3mo, 40) : ""}`;
            if (depth === 5) {
              options.push({ crumb: qLabel, path: { uoId: uo.id, outcomeId: o.id, outputId: ko.id, waveId: w.id, quarterId: q.id, monthId: "" } });
              return;
            }
            q.monthlyKRs.forEach((m) => {
              const mLabel = `${qLabel} › ${m.label}${m.text ? " — " + truncate(m.text, 30) : ""}`;
              options.push({ crumb: mLabel, path: { uoId: uo.id, outcomeId: o.id, outputId: ko.id, waveId: w.id, quarterId: q.id, monthId: m.id } });
            });
          });
        });
      });
    });
  });
  return options;
}

function HierarchyBrowserModal({ trackData, levelKey, onPick, onClose }) {
  const t = useT();
  const [query, setQuery] = useState("");
  const options = buildPlacementOptions(trackData, levelKey, t);
  const q = query.trim().toLowerCase();
  const filtered = q ? options.filter((o) => o.crumb.toLowerCase().includes(q)) : options;
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white rounded-2xl p-4 w-full max-w-xl max-h-[80vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="text-sm font-semibold text-neutral-800 mb-2">{t("Выберите место в дереве трека")}</div>
        <input
          autoFocus value={query} onChange={(e) => setQuery(e.target.value)}
          placeholder={t("Поиск по дереву…")}
          className="text-sm border border-neutral-200 rounded-md px-2.5 py-1.5 mb-2 outline-none focus:border-neutral-400"
        />
        <div className="overflow-y-auto space-y-1 flex-1">
          {filtered.length === 0 && <div className="text-xs text-neutral-400 py-4 text-center">{t("Ничего не найдено")}</div>}
          {filtered.map((o, i) => (
            <button
              key={i}
              onClick={() => onPick(o.path)}
              className="w-full text-left text-xs rounded-lg px-2.5 py-1.5 border border-neutral-200 hover:border-neutral-400 hover:bg-neutral-50"
            >
              {o.crumb}
            </button>
          ))}
        </div>
        <button onClick={onClose} className="w-full text-xs text-neutral-400 hover:text-neutral-600 pt-2 shrink-0">{t("Отмена")}</button>
      </div>
    </div>
  );
}

function TargetPicker({ trackData, levelKey, path, onChange }) {
  const t = useT();
  const level = IMPORT_LEVELS.find((l) => l.key === levelKey);
  const depth = level ? level.depth : 6;
  const uo = trackData.ultimateObjectives.find((u) => u.id === path.uoId);
  const outcome = uo && uo.krOutcomes.find((o) => o.id === path.outcomeId);
  const output = outcome && outcome.krOutputs.find((k) => k.id === path.outputId);
  const wave = output && output.waves.find((w) => w.id === path.waveId);
  const quarter = wave && wave.quarters.find((q) => q.id === path.quarterId);
  const set = (patch) => onChange({ ...path, ...patch });
  const selCls = "text-xs border border-neutral-200 rounded-md px-1.5 py-1 bg-white";

  return (
    <div className="flex flex-wrap gap-1.5">
      <select
        value={path.uoId} className={selCls}
        onChange={(e) => set({ uoId: e.target.value, outcomeId: "", outputId: "", waveId: "", quarterId: "", monthId: "" })}
      >
        <option value="">{t("Главная цель…")}</option>
        {trackData.ultimateObjectives.map((u) => (
          <option key={u.id} value={u.id}>{u.label}{u.text ? " — " + truncate(u.text, 30) : ""}</option>
        ))}
        {trackData.ultimateObjectives.length < MAX_ULTIMATE_OBJECTIVES && <option value="__new__">+ {t("новая Главная цель")}</option>}
      </select>

      {depth >= 2 && path.uoId && path.uoId !== "__new__" && (
        <select
          value={path.outcomeId} className={selCls}
          onChange={(e) => set({ outcomeId: e.target.value, outputId: "", waveId: "", quarterId: "", monthId: "" })}
        >
          <option value="">{t("KR Outcome…")}</option>
          {uo && uo.krOutcomes.map((o) => <option key={o.id} value={o.id}>{o.label}{o.text ? " — " + truncate(o.text, 30) : ""}</option>)}
          {uo && uo.krOutcomes.length < MAX_KR_OUTCOMES && <option value="__new__">+ {t("новый KR Outcome")}</option>}
        </select>
      )}
      {depth >= 2 && path.uoId === "__new__" && (
        <span className="t11 text-neutral-400 self-center">{t("KR Outcome появится после создания цели")}</span>
      )}

      {depth >= 4 && path.outcomeId && path.outcomeId !== "__new__" && (
        <select value={path.outputId} className={selCls} onChange={(e) => set({ outputId: e.target.value, waveId: "", quarterId: "", monthId: "" })}>
          <option value="">{t("KR Output…")}</option>
          {outcome && outcome.krOutputs.map((k) => <option key={k.id} value={k.id}>{k.label}</option>)}
        </select>
      )}

      {depth >= 4 && path.outputId && (
        <select value={path.waveId} className={selCls} onChange={(e) => set({ waveId: e.target.value, quarterId: "", monthId: "" })}>
          <option value="">{t("Волна…")}</option>
          {output && output.waves.map((w, i) => <option key={w.id} value={w.id}>{t("Волна")} {i + 1}</option>)}
        </select>
      )}

      {depth >= 5 && path.waveId && (
        <select value={path.quarterId} className={selCls} onChange={(e) => set({ quarterId: e.target.value, monthId: "" })}>
          <option value="">{t("Квартал…")}</option>
          {wave && wave.quarters.map((q) => <option key={q.id} value={q.id}>{t("Квартал")} {q.label}</option>)}
        </select>
      )}

      {depth >= 6 && path.quarterId && (
        <select value={path.monthId} className={selCls} onChange={(e) => set({ monthId: e.target.value })}>
          <option value="">{t("Месяц…")}</option>
          {quarter && quarter.monthlyKRs.map((m) => <option key={m.id} value={m.id}>{m.label}</option>)}
        </select>
      )}
    </div>
  );
}

function BitrixImportModule() {
  const t = useT();
  const fileInputRef = useRef(null);
  const [rows, setRows] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [tracksCache, setTracksCache] = useState({});
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("okr-bitrix-import:draft", false);
        const draft = res ? JSON.parse(res.value) : [];
        setRows(Array.isArray(draft) ? draft : []);
      } catch { setRows([]); }
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    const h = setTimeout(() => {
      window.storage.set("okr-bitrix-import:draft", JSON.stringify(rows), false).catch(() => {});
    }, 600);
    return () => clearTimeout(h);
  }, [rows, loaded]);

  useEffect(() => {
    const missing = [...new Set(rows.map((r) => r.trackId).filter(Boolean))].filter((tid) => !tracksCache[tid]);
    if (missing.length === 0) return;
    (async () => {
      const updates = {};
      for (const tid of missing) {
        try {
          const res = await window.storage.get(`okr-track:${tid}`, false);
          updates[tid] = res ? migrateTrackData(JSON.parse(res.value)) : createTrackData();
        } catch {
          updates[tid] = createTrackData();
        }
      }
      setTracksCache((c) => ({ ...c, ...updates }));
    })();
  }, [rows, tracksCache]);

  // Once a row's track finishes loading, check whether this exact Bitrix task was already
  // imported before (matched by its Bitrix ID, not by guessing) — if so, pre-fill its known
  // location and mark it so the review list shows "already imported" instead of prompting
  // the person to place it again from scratch.
  useEffect(() => {
    setRows((rs) => {
      let changed = false;
      const next = rs.map((r) => {
        if (r.idChecked || !r.bitrixTaskId || !r.trackId) return r;
        const trackData = tracksCache[r.trackId];
        if (!trackData) return r;
        changed = true;
        const match = findTaskByBitrixId(trackData, r.bitrixTaskId);
        if (!match) return { ...r, idChecked: true };
        return { ...r, idChecked: true, path: match.path, alreadyImportedText: match.task.text, include: false };
      });
      return changed ? next : rs;
    });
  }, [tracksCache]);

  const handleFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setBusy(true); setMessage("");
    try {
      const text = await file.text();
      const parsed = parseBitrixExport(text);
      if (parsed.length === 0) throw new Error(t("не нашёл таблицу с колонкой «Название» в файле"));
      const newRows = parsed.map((p) => {
        const trackId = guessTrackFromTags(p.tags);
        const guess = guessImportLevel(p.rawText);
        return {
          rowId: p.rowId, rawText: p.rawText, editedText: p.rawText, deadline: p.deadline,
          bitrixTaskId: p.bitrixTaskId || "",
          trackId, levelKey: guess === "skip" ? "task" : guess,
          include: guess !== "skip",
          path: emptyTargetPath(),
          forceApply: false,
          idChecked: false,
          alreadyImportedText: "",
        };
      });
      setRows(newRows);
      setMessage(t("Загружено строк: ") + newRows.length);
    } catch (err) {
      setMessage(t("Не удалось прочитать файл: ") + err.message);
    }
    setBusy(false);
  };

  const updateRow = (rowId, patch) => setRows((rs) => rs.map((r) => (r.rowId === rowId ? { ...r, ...patch } : r)));
  const removeRow = (rowId) => setRows((rs) => rs.filter((r) => r.rowId !== rowId));
  const acceptAll = () => setRows((rs) => rs.map((r) => ({ ...r, include: true })));

  // A row counts as "conflicting" the same way the Gantt preview and the warning badge below
  // decide it — if it doesn't conflict, or the person explicitly forced it, it's eligible for
  // the bulk apply. Conflicting rows stay staged (visible in the Gantt as a draft) and sit out
  // of "Применить отмеченное" until cleaned up or force-applied one at a time.
  const rowConflict = (r) => {
    const trackData = tracksCache[r.trackId];
    if (!trackData || !pathIsComplete(r.levelKey, r.path)) return null;
    if (r.levelKey === "task") return detectTaskDuplicate(trackData, r.path, r.editedText || r.rawText) ? "dup" : null;
    return detectFieldConflict(trackData, r.levelKey, r.path);
  };

  const readyRows = rows.filter((r) => {
    if (!r.include || !r.trackId || !pathIsComplete(r.levelKey, r.path)) return false;
    const conflict = rowConflict(r);
    return !conflict || r.forceApply;
  });
  const conflictRows = rows.filter((r) => {
    if (!r.include || !r.trackId || !pathIsComplete(r.levelKey, r.path)) return false;
    return !!rowConflict(r) && !r.forceApply;
  });

  const applyAll = async () => {
    if (readyRows.length === 0) return;
    setBusy(true); setMessage("");
    try {
      const byTrack = {};
      const appliedIds = [];
      readyRows.forEach((r) => {
        const base = byTrack[r.trackId] || tracksCache[r.trackId];
        if (!base) return;
        const next = applyImportRow(base, r.levelKey, r.path, r.editedText || r.rawText, r.bitrixTaskId);
        if (next) { byTrack[r.trackId] = next; appliedIds.push(r.rowId); }
      });
      const trackIds = Object.keys(byTrack);
      if (trackIds.length === 0) {
        setMessage(t("Ничего не применено — проверьте лимиты (например, максимум Главных целей/KR Outcome/задач в месяце)."));
        setBusy(false);
        return;
      }
      await Promise.all(trackIds.map((tid) => window.storage.set(`okr-track:${tid}`, JSON.stringify(byTrack[tid]), false)));
      const appliedSet = new Set(appliedIds);
      setRows((rs) => rs.filter((r) => !appliedSet.has(r.rowId)));
      setTracksCache((c) => ({ ...c, ...byTrack }));
      setMessage(t("Применено: ") + appliedIds.length + (conflictRows.length ? " · " + t("конфликтных строк оставлено на доработку: ") + conflictRows.length : ""));
    } catch (err) {
      setMessage(t("Ошибка применения: ") + err.message);
    }
    setBusy(false);
  };

  if (!loaded) return <div className="text-sm text-neutral-400 py-12 text-center">{t("Загрузка…")}</div>;

  return (
    <div className="space-y-3">
      <div className="border border-neutral-200 rounded-xl p-3 flex items-center justify-between flex-wrap gap-3">
        <div>
          <div className="text-sm font-medium text-neutral-800">{t("Импорт задач из Bitrix24")}</div>
          <div className="text-xs text-neutral-400 mt-0.5">
            {t("Файл экспорта задач из Bitrix24 (.xls) — колонки «Название», «Крайний срок», «Теги», «№ заявки».")}
          </div>
          <div className="t11 text-sky-700 mt-1">
            {t("Совет: добавьте в выгрузку колонку с номером задачи Bitrix (ID / № заявки) — после того как вы один раз разместите задачу вручную, её номер сохранится, и при повторном импорте того же файла приложение само узнает эту задачу и не предложит разместить её ещё раз.")}
          </div>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => fileInputRef.current && fileInputRef.current.click()}
            disabled={busy}
            className="flex items-center gap-1.5 text-xs border border-neutral-200 rounded-md px-2.5 py-1.5 hover:bg-neutral-50 disabled:opacity-50"
          >
            <FileUp size={13} /> {t("Загрузить файл Bitrix")}
          </button>
          <input ref={fileInputRef} type="file" accept=".xls,.xlsx,.html,.htm" className="hidden" onChange={handleFile} />
        </div>
      </div>
      {message && <div className="text-xs text-neutral-500">{message}</div>}

      {rows.length === 0 ? (
        <div className="text-sm text-neutral-400 text-center py-10 border border-dashed border-neutral-200 rounded-xl">
          {t("Загрузите файл выгрузки задач из Bitrix24, чтобы начать проверку.")}
        </div>
      ) : (
        <>
          <div className="flex items-center justify-between t11 text-neutral-400 flex-wrap gap-2">
            <span>
              {t("Готовы к применению:")} {readyRows.length}
              {conflictRows.length > 0 && <> · <span className="text-orange-600">{t("на конфликте (остаются черновиком в Ганте):")} {conflictRows.length}</span></>}
              {" "}{t("из")} {rows.length}
            </span>
            <div className="flex gap-1.5">
              <button onClick={acceptAll} className="text-xs border border-neutral-200 rounded-md px-2.5 py-1.5 hover:bg-neutral-50">
                {t("Принять все")}
              </button>
              <button
                onClick={applyAll} disabled={busy || readyRows.length === 0}
                className="flex items-center gap-1.5 text-xs bg-neutral-900 text-white rounded-md px-2.5 py-1.5 disabled:opacity-40"
              >
                <CheckCircle2 size={13} /> {t("Применить отмеченное")}
              </button>
            </div>
          </div>

          <div className="space-y-2">
            {rows.map((r) => {
              const trackData = tracksCache[r.trackId];
              const complete = trackData && pathIsComplete(r.levelKey, r.path);
              const fieldConflict = complete && r.levelKey !== "task" ? detectFieldConflict(trackData, r.levelKey, r.path) : null;
              const taskDup = complete && r.levelKey === "task" ? detectTaskDuplicate(trackData, r.path, r.editedText || r.rawText) : false;
              const hasConflict = !!fieldConflict || taskDup;
              const willApply = complete && (!hasConflict || r.forceApply);
              return (
                <div key={r.rowId} className={`rounded-xl border p-3 ${!complete ? "border-amber-200" : hasConflict && !r.forceApply ? "border-orange-300" : "border-neutral-200"}`}>
                  <div className="flex items-start gap-2.5">
                    <input
                      type="checkbox" checked={r.include} className="mt-1"
                      onChange={(e) => updateRow(r.rowId, { include: e.target.checked })}
                    />
                    <div className="flex-1 min-w-0 space-y-2">
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <select
                          value={r.trackId}
                          onChange={(e) => updateRow(r.rowId, { trackId: e.target.value, path: emptyTargetPath(), forceApply: false })}
                          className="t11 font-medium rounded px-1.5 py-0.5 border"
                          style={
                            r.trackId
                              ? { background: TRACK_COLORS[r.trackId], borderColor: TRACK_COLORS[r.trackId], color: "#fff" }
                              : { borderColor: "#e5e5e5", color: "#a3a3a3" }
                          }
                        >
                          <option value="">{t("Трек…")}</option>
                          {TRACKS.map((tr) => <option key={tr.id} value={tr.id}>{tr.name}</option>)}
                        </select>
                        {r.bitrixTaskId && <span className="t11 text-neutral-400">№ {r.bitrixTaskId}</span>}
                        {r.deadline && <span className="t11 text-neutral-400">{t("срок")} {r.deadline}</span>}
                        {r.alreadyImportedText && (
                          <span className="t11 text-sky-700 bg-sky-50 border border-sky-200 rounded px-1.5 py-0.5">
                            {t("уже импортирована по ID")}
                          </span>
                        )}
                        {!complete && (
                          <span className="t11 text-amber-700 bg-amber-50 border border-amber-200 rounded px-1.5 py-0.5">
                            {t("не размещено")}
                          </span>
                        )}
                        {complete && hasConflict && !r.forceApply && (
                          <span className="t11 text-orange-700 bg-orange-50 border border-orange-200 rounded px-1.5 py-0.5">
                            ⚠ {t("конфликт — сначала виден только в Ганте")}
                          </span>
                        )}
                        {complete && hasConflict && r.forceApply && (
                          <span className="t11 text-neutral-500 bg-neutral-100 border border-neutral-200 rounded px-1.5 py-0.5">
                            {t("будет объединено рядом с существующим при применении")}
                          </span>
                        )}
                        <button onClick={() => removeRow(r.rowId)} className="ml-auto text-neutral-300 hover:text-red-500 p-0.5 shrink-0">
                          <Trash2 size={12} />
                        </button>
                      </div>
                      <Field small value={r.editedText} onChange={(v) => updateRow(r.rowId, { editedText: v })} placeholder={t("Текст")} />
                      {r.alreadyImportedText && (
                        <div className="t11 text-sky-700 bg-sky-50 border border-sky-200 rounded-md px-2 py-1">
                          {r.alreadyImportedText.trim() === (r.editedText || r.rawText).trim()
                            ? t("Уже в дереве без изменений — можно пропустить.")
                            : <>{t("В дереве сейчас другой текст:")} «<Truncated text={r.alreadyImportedText} limit={90} />» — {t("если в Bitrix задачу переименовали, отметьте чекбокс, чтобы дописать новый текст рядом.")}</>}
                        </div>
                      )}
                      {fieldConflict && (
                        <div className="t11 text-orange-700 bg-orange-50 border border-orange-200 rounded-md px-2 py-1 flex items-center justify-between gap-2 flex-wrap">
                          <span>{t("Сейчас там:")} «<Truncated text={fieldConflict} limit={90} />»</span>
                          {!r.forceApply && (
                            <button onClick={() => updateRow(r.rowId, { forceApply: true })} className="shrink-0 text-orange-800 underline">
                              {t("применить всё равно (дописать рядом)")}
                            </button>
                          )}
                        </div>
                      )}
                      {taskDup && !fieldConflict && (
                        <div className="t11 text-orange-700 bg-orange-50 border border-orange-200 rounded-md px-2 py-1 flex items-center justify-between gap-2 flex-wrap">
                          <span>{t("Похожая задача уже есть в этом месяце")}</span>
                          {!r.forceApply && (
                            <button onClick={() => updateRow(r.rowId, { forceApply: true })} className="shrink-0 text-orange-800 underline">
                              {t("применить всё равно (добавить рядом)")}
                            </button>
                          )}
                        </div>
                      )}
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <select
                          value={r.levelKey} className="text-xs border border-neutral-200 rounded-md px-1.5 py-1 bg-white"
                          onChange={(e) => updateRow(r.rowId, { levelKey: e.target.value, path: emptyTargetPath(), forceApply: false })}
                        >
                          {IMPORT_LEVELS.map((l) => <option key={l.key} value={l.key}>{t(l.label)}</option>)}
                        </select>
                        {r.trackId && trackData && (
                          <>
                            <button
                              type="button"
                              onClick={() => updateRow(r.rowId, { browsing: true })}
                              className="flex items-center gap-1 text-xs border border-neutral-200 rounded-md px-2 py-1 bg-white hover:bg-neutral-50"
                            >
                              <Eye size={12} /> {t("Обзор дерева")}
                            </button>
                            <TargetPicker
                              trackData={trackData} levelKey={r.levelKey} path={r.path}
                              onChange={(p) => updateRow(r.rowId, { path: p, forceApply: false })}
                            />
                            {r.browsing && (
                              <HierarchyBrowserModal
                                trackData={trackData} levelKey={r.levelKey}
                                onPick={(p) => updateRow(r.rowId, { path: p, forceApply: false, browsing: false })}
                                onClose={() => updateRow(r.rowId, { browsing: false })}
                              />
                            )}
                          </>
                        )}
                        {r.trackId && !trackData && <span className="t11 text-neutral-400">{t("Загрузка трека…")}</span>}
                      </div>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
}

const MODULES = [
  { key: "tracks_all", label: "Треки (все)" },
  { key: "track_health-os", label: "Трек Health OS" },
  { key: "track_growth-model", label: "Трек GrowthModel" },
  { key: "track_prime-growth", label: "Трек PrimeGrowth" },
  { key: "track_coral-evo", label: "Трек CoralEVO" },
  { key: "track_it-model", label: "Трек ITModel" },
  { key: "combined_tree", label: "Общее дерево OKR" },
  { key: "combined_tree_links", label: "Общее дерево OKR (связи)" },
  { key: "projects", label: "Проекты" },
  { key: "projects_dashboard", label: "Дашборд проектов" },
  { key: "tracking_dashboard", label: "Дашборд трекинга" },
  { key: "gantt", label: "Гант" },
  { key: "directory", label: "Справочник" },
  { key: "bitrix_import", label: "Импорт из Bitrix24" },
  { key: "privileges", label: "Управление привилегиями" },
];

function moduleKeyForDataKey(key) {
  if (key.startsWith("okr-track:")) return "track_" + key.slice("okr-track:".length);
  if (key === "okr-projects" || key === "okr-project-statuses") return "projects";
  if (key === "okr-directory") return "directory";
  if (key === "okr-links") return "combined_tree_links";
  if (key.startsWith("okr-bitrix-import")) return "bitrix_import";
  return "privileges";
}
function effectivePermission(permissions, moduleKey) {
  const level = (permissions && permissions[moduleKey]) || "none";
  if (moduleKey.startsWith("track_")) {
    const all = (permissions && permissions["tracks_all"]) || "none";
    if (all === "edit") return "edit";
    if (all === "view" && level === "none") return "view";
  }
  return level;
}
function canEditModule(permissions, moduleKey) { return effectivePermission(permissions, moduleKey) === "edit"; }
function canViewModule(permissions, moduleKey) {
  const l = effectivePermission(permissions, moduleKey);
  return l === "view" || l === "edit";
}

function tSync(s) {
  let lang = "ru";
  try { lang = localStorage.getItem("okr-lang") || "ru"; } catch {}
  return (lang === "en" && EN_DICT[s]) ? EN_DICT[s] : s;
}

function useRoleAuth() {
  const t = useT();
  const [roles, setRoles] = useState([]);
  const [rolesLoaded, setRolesLoaded] = useState(false);
  const [session, setSession] = useState(() => {
    try { return JSON.parse(localStorage.getItem(ROLE_SESSION_KEY) || "null"); } catch { return null; }
  });
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    (async () => {
      const { data, error } = await supabase.rpc("okr_list_roles");
      if (!error && data) setRoles(data);
      setRolesLoaded(true);
    })();
  }, []);

  const signIn = useCallback(async (roleId, password) => {
    setError("");
    setBusy(true);
    try {
      const { data, error } = await supabase.rpc("okr_login", { p_role_id: roleId, p_password: password });
      if (error || !data || data.length === 0) {
        setError(t("Неверная роль или пароль."));
        return false;
      }
      const roleData = data[0];
      const s = { roleId: roleData.id, roleName: roleData.name, permissions: roleData.permissions, password };
      setSession(s);
      try { localStorage.setItem(ROLE_SESSION_KEY, JSON.stringify(s)); } catch {}
      return true;
    } catch (e) {
      setError(e.message || t("Не удалось войти."));
      return false;
    } finally {
      setBusy(false);
    }
  }, [t]);

  const signOut = useCallback(() => {
    setSession(null);
    try { localStorage.removeItem(ROLE_SESSION_KEY); } catch {}
  }, []);

  return { roles, rolesLoaded, session, signIn, signOut, signOutError: error, busy, signedIn: !!session };
}

// Drop-in replacement for the local window.storage API — backed by Supabase.
// Both reads AND writes now go through password-checked RPCs (okr_load_data / okr_save_data) —
// the table itself has no direct SELECT grant for anon/authenticated, so an unauthenticated
// visitor (or anyone with just the public anon key) can no longer read the data directly.
function makeSupabaseStorage(session) {
  return {
    async get(key) {
      const { data, error } = await supabase.rpc("okr_load_data", {
        p_role_id: session.roleId, p_password: session.password, p_key: key,
      });
      if (error || data == null) return null;
      return { key, value: JSON.stringify(data), shared: true };
    },
    async set(key, value) {
      if (!session) throw new Error(tSync("Войдите, чтобы сохранять изменения."));
      const moduleKey = moduleKeyForDataKey(key);
      let parsed;
      try { parsed = JSON.parse(value); } catch { parsed = value; }
      const { error } = await supabase.rpc("okr_save_data", {
        p_role_id: session.roleId, p_password: session.password,
        p_module_key: moduleKey, p_data_key: key, p_value: parsed,
      });
      if (error) throw new Error(error.message || tSync("Не удалось сохранить данные."));
      return { key, value, shared: true };
    },
    async delete(key) {
      return this.set(key, "null");
    },
    async list(prefix) {
      const { data, error } = await supabase.rpc("okr_list_data_keys", {
        p_role_id: session.roleId, p_password: session.password, p_prefix: prefix || null,
      });
      if (error) return { keys: [], prefix, shared: true };
      return { keys: data || [], prefix, shared: true };
    },
  };
}

// Fallback used while signed out — no session means no password to authenticate reads with,
// so every call resolves to "nothing here" rather than touching Supabase at all. This is what
// makes "open the site without logging in" show an empty state instead of any real data.
function makeLocalDemoStorage() {
  return {
    async get() { return null; },
    async set() { throw new Error(tSync("Войдите, чтобы сохранять изменения.")); },
    async delete() { throw new Error(tSync("Войдите, чтобы сохранять изменения.")); },
    async list() { return { keys: [], shared: false }; },
  };
}

const AuthCtx = createContext({ signedIn: false, roleId: null, roleName: null, permissions: {} });
function useAuthCtx() { return useContext(AuthCtx); }

function LockOverlay({ locked, reason, children }) {
  if (!locked) return children;
  return (
    <div className="relative">
      <div style={{ pointerEvents: "none", opacity: 0.55 }}>{children}</div>
      <div className="absolute inset-0 flex items-start justify-center pt-10 pointer-events-none">
        <div className="bg-neutral-900 text-white text-xs px-3 py-1.5 rounded-full shadow-lg">🔒 {reason}</div>
      </div>
    </div>
  );
}

function useDirectory(reloadToken) {
  const [directory, setDirectory] = useState([]);
  const [loaded, setLoaded] = useState(false);
  useEffect(() => {
    setLoaded(false);
    (async () => {
      try {
        const res = await window.storage.get("okr-directory", false);
        setDirectory(res ? JSON.parse(res.value) : []);
      } catch {
        setDirectory([]);
      }
      setLoaded(true);
    })();
  }, [reloadToken]);
  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => {
      window.storage.set("okr-directory", JSON.stringify(directory), false).catch(() => {});
    }, 400);
    return () => clearTimeout(t);
  }, [directory, loaded]);
  return [directory, setDirectory];
}

const TYPE_LABEL = { company: "Компания", department: "Отдел", function: "Функция" };
const TYPE_COLOR = {
  company: "bg-purple-100 text-purple-700",
  department: "bg-sky-100 text-sky-700",
  function: "bg-teal-100 text-teal-700",
};

function DirectoryManager({ directory, setDirectory }) {
  const t = useT();
  const [name, setName] = useState("");
  const [type, setType] = useState("function");
  const [open, setOpen] = useState(false);

  const add = () => {
    if (!name.trim()) return;
    setDirectory((prev) => [...prev, { id: newId(), name: name.trim(), type }]);
    setName("");
  };
  const remove = (id) => setDirectory((prev) => prev.filter((d) => d.id !== id));

  return (
    <div className="mb-4 border border-neutral-200 rounded-xl overflow-hidden">
      <button onClick={() => setOpen((o) => !o)} className="w-full flex items-center gap-2 px-3 py-2 bg-neutral-50 hover:bg-neutral-100">
        {open ? <ChevronDown size={15} className="text-neutral-500" /> : <ChevronRight size={15} className="text-neutral-500" />}
        <Users size={14} className="text-neutral-500" />
        <span className="text-xs font-medium uppercase tracking-wide text-neutral-600">{t("Справочник: компания / отделы / функции")}</span>
        <span className="text-xs text-neutral-400 ml-auto">{directory.length}</span>
      </button>
      {open && (
        <div className="p-3 space-y-2">
          <div className="flex gap-1.5">
            <input
              className="flex-1 text-sm border border-neutral-200 rounded-md px-2 py-1 outline-none"
              placeholder={t("Название (например: Отдел маркетинга)")}
              value={name}
              onChange={(e) => setName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && add()}
            />
            <select value={type} onChange={(e) => setType(e.target.value)} className="text-sm border border-neutral-200 rounded-md px-1.5">
              <option value="company">{t("Компания")}</option>
              <option value="department">{t("Отдел")}</option>
              <option value="function">{t("Функция")}</option>
            </select>
            <button onClick={add} className="px-2.5 rounded-md bg-neutral-900 text-white text-sm hover:bg-neutral-800">
              <Plus size={14} />
            </button>
          </div>
          <div className="space-y-1 max-h-56 overflow-y-auto">
            {directory.length === 0 && (
              <div className="text-xs text-neutral-400 py-2 text-center">{t("Справочник пуст — добавьте функции, отделы или компанию")}</div>
            )}
            {directory.map((d) => (
              <div key={d.id} className="flex items-center gap-2 px-2 py-1 rounded-md hover:bg-neutral-50">
                <span className={`t10 px-1.5 py-0.5 rounded ${TYPE_COLOR[d.type]}`}>{t(TYPE_LABEL[d.type])}</span>
                <span className="text-sm flex-1 min-w-0 break-words">{d.name}</span>
                <button onClick={() => remove(d.id)} className="text-neutral-300 hover:text-red-500 p-0.5">
                  <Trash2 size={12} />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function RoleAuthBar({ auth, onSignedInChange }) {
  const t = useT();
  const { roles, rolesLoaded, session, signIn, signOut, signOutError, busy, signedIn } = auth;
  const [roleId, setRoleId] = useState("");
  const [password, setPassword] = useState("");
  const [open, setOpen] = useState(false);
  const prevSignedIn = useRef(signedIn);

  useEffect(() => {
    if (signedIn !== prevSignedIn.current) {
      prevSignedIn.current = signedIn;
      onSignedInChange();
    }
  }, [signedIn, onSignedInChange]);

  const handleSignIn = async () => {
    if (!roleId || !password) return;
    await signIn(roleId, password);
    setPassword("");
  };

  return (
    <div className="mb-4 border border-neutral-200 rounded-xl overflow-hidden">
      <button onClick={() => setOpen((o) => !o)} className="w-full flex items-center gap-2 flex-wrap px-3 py-2 bg-neutral-50 hover:bg-neutral-100 text-left">
        {open ? <ChevronDown size={15} className="text-neutral-500 shrink-0" /> : <ChevronRight size={15} className="text-neutral-500 shrink-0" />}
        <span className="text-xs font-medium uppercase tracking-wide text-neutral-600 mr-1">{t("Вход")}</span>
        {!signedIn ? (
          <span className="text-xs text-neutral-400">{t("не подключено")}</span>
        ) : (
          <span className="text-xs px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700">{session.roleName}</span>
        )}
      </button>

      {open && (
        <div className="p-3 space-y-2">
          {!signedIn ? (
            <div className="flex items-center gap-2 flex-wrap">
              <select
                value={roleId} onChange={(e) => setRoleId(e.target.value)} disabled={!rolesLoaded}
                className="text-xs border border-neutral-200 rounded-md px-2 py-1.5 bg-neutral-50"
                style={{ minWidth: 220 }}
              >
                <option value="">{t("Выберите роль")}</option>
                {roles.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
              </select>
              <input
                type="password" value={password} onChange={(e) => setPassword(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleSignIn()}
                placeholder={t("Пароль")}
                className="text-xs border border-neutral-200 rounded-md px-2 py-1.5 bg-neutral-50"
              />
              <button
                onClick={handleSignIn} disabled={busy || !roleId || !password}
                className="flex items-center gap-1.5 text-xs bg-neutral-900 text-white rounded-md px-2.5 py-1.5 disabled:opacity-50"
              >
                <Users size={12} /> {busy ? t("Входим…") : t("Войти")}
              </button>
            </div>
          ) : (
            <button onClick={signOut} className="text-xs text-neutral-400 hover:text-red-500 underline">{t("выйти")}</button>
          )}

          {signOutError && <div className="text-xs text-red-600">{signOutError}</div>}
          <div className="text-xs text-neutral-400">
            {t("Без входа приложение доступно только на просмотр. Роль и пароль вам выдаёт администратор.")}
          </div>
        </div>
      )}
    </div>
  );
}

// ---- Signature illustration for the empty/logged-out state — a rolling-wave tree of goals ----
// Trunk = mission (grounded, singular). It splits into 5 branches, one per track, each carrying
// the track's own brand color. Branches thin out into small leaf-clusters — the nearer a leaf
// sits to the trunk, the larger it is (an 18-month Outcome); the further out, the smaller and more
// numerous (a week's task). A dashed ring behind it echoes the app's own empty-state border and
// stands for the 18-month horizon the whole tree grows inside.
function GoalsTreeIllustration() {
  const branches = [
    { color: "#16a34a", d: "M150,232 C120,190 78,168 40,152", leaves: [[40, 152, 5], [58, 138, 3.5], [30, 133, 3]] },
    { color: "#ca8a04", d: "M150,232 C132,180 118,140 108,96", leaves: [[108, 96, 5], [92, 82, 3.5], [122, 76, 3]] },
    { color: "#b45309", d: "M150,232 C150,178 150,132 150,84", leaves: [[150, 84, 5.5], [134, 66, 3.5], [166, 66, 3.5]] },
    { color: "#8b5cf6", d: "M150,232 C168,180 182,140 192,96", leaves: [[192, 96, 5], [178, 82, 3], [208, 76, 3.5]] },
    { color: "#f87171", d: "M150,232 C180,190 222,168 260,152", leaves: [[260, 152, 5], [242, 138, 3], [270, 133, 3.5]] },
  ];
  return (
    <svg viewBox="0 0 300 250" className="w-full h-auto" role="img" aria-hidden="true">
      <ellipse cx="150" cy="128" rx="132" ry="112" fill="none" stroke="#e5e5e5" strokeWidth="1" strokeDasharray="3 5" />
      <line x1="70" y1="234" x2="230" y2="234" stroke="#e5e5e5" strokeWidth="1" />
      {branches.map((b, i) => (
        <g key={i}>
          <path d={b.d} fill="none" stroke={b.color} strokeWidth="2.5" strokeLinecap="round" opacity="0.85" />
          {b.leaves.map(([x, y, r], j) => (
            <circle key={j} cx={x} cy={y} r={r} fill={b.color} opacity={j === 0 ? 0.9 : 0.55} />
          ))}
        </g>
      ))}
      <path d="M150,232 C150,214 150,200 150,190" fill="none" stroke="#1F4E3D" strokeWidth="5" strokeLinecap="round" />
      <path d="M150,234 C138,240 122,242 108,240" fill="none" stroke="#1F4E3D" strokeWidth="2" strokeLinecap="round" opacity="0.5" />
      <path d="M150,234 C162,240 178,242 192,240" fill="none" stroke="#1F4E3D" strokeWidth="2" strokeLinecap="round" opacity="0.5" />
    </svg>
  );
}

function EmptyStateInvite({ signedIn }) {
  const t = useT();
  return (
    <div className="flex flex-col items-center text-center py-10 px-6 border border-dashed border-neutral-200 rounded-xl">
      <div className="w-48 sm:w-56">
        <GoalsTreeIllustration />
      </div>
      <div className="mt-2 text-sm font-medium text-neutral-700">
        {signedIn ? t("Этот раздел не по вашей роли") : t("Дерево ещё не открыто")}
      </div>
      <div className="mt-1 text-sm text-neutral-400 max-w-sm">
        {signedIn
          ? t("У вашей роли нет доступа сюда — переключитесь на раздел, который вам доступен, или попросите администратора расширить права.")
          : t("Войдите под своей ролью — и увидите свою ветку: от цели на 18 месяцев до задач этой недели.")}
      </div>
    </div>
  );
}

export default function App() {
  const [trackId, setTrackId] = useState(TRACKS[0].id);
  const [refreshKey, setRefreshKey] = useState(0);
  const [directory, setDirectory] = useDirectory(refreshKey);
  const roleAuth = useRoleAuth();
  const [lang, setLang] = useState(() => {
    try { return localStorage.getItem("okr-lang") || "ru"; } catch { return "ru"; }
  });
  const toggleLang = () => {
    setLang((l) => {
      const next = l === "ru" ? "en" : "ru";
      try { localStorage.setItem("okr-lang", next); } catch {}
      return next;
    });
  };
  const t = (s) => (lang === "en" && EN_DICT[s]) ? EN_DICT[s] : s;

  const permissions = roleAuth.session ? roleAuth.session.permissions : {};
  const authCtxValue = {
    signedIn: roleAuth.signedIn,
    roleId: roleAuth.session ? roleAuth.session.roleId : null,
    roleName: roleAuth.session ? roleAuth.session.roleName : null,
    permissions,
    isAdmin: canEditModule(permissions, "privileges"),
    canEdit: (moduleKey) => canEditModule(permissions, moduleKey),
    canView: (moduleKey) => canViewModule(permissions, moduleKey),
  };

  useEffect(() => {
    window.storage = roleAuth.session ? makeSupabaseStorage(roleAuth.session) : makeLocalDemoStorage();
  }, [roleAuth.session]);

  return (
    <LanguageCtx.Provider value={lang}>
    <div className="w-full bg-white text-neutral-900">
      <style>{`.t11{font-size:11px;line-height:1.35}.t10{font-size:10px;line-height:1.3}`}</style>
      <div className="w-full px-3 sm:px-6 lg:px-10 py-6">
        <div className="mb-4 flex items-start justify-between gap-3">
          <div>
            <h1 className="text-lg font-medium">{t("OKR-конструктор трека")}</h1>
          </div>
          <button
            onClick={toggleLang}
            className="shrink-0 flex items-center gap-1 text-xs border border-neutral-200 rounded-full px-3 py-1.5 hover:bg-neutral-50"
            title="Switch language / Переключить язык"
          >
            🌐 {lang === "ru" ? "RU / EN" : "EN / RU"}
          </button>
        </div>

        <RoleAuthBar auth={roleAuth} onSignedInChange={() => setRefreshKey((k) => k + 1)} />

        {authCtxValue.isAdmin && (
          <ExportImportBar directory={directory} setDirectory={setDirectory} onImported={() => setRefreshKey((k) => k + 1)} />
        )}

        {authCtxValue.canView("directory") && (
          <LockOverlay locked={!authCtxValue.canEdit("directory")} reason={t("Справочник редактирует администратор")}>
            <DirectoryManager directory={directory} setDirectory={setDirectory} />
          </LockOverlay>
        )}

        <div className="flex gap-1 mb-4 flex-wrap">
          {TRACKS.map((tr) => {
            if (!authCtxValue.canView(`track_${tr.id}`)) return null;
            const color = TRACK_COLORS[tr.id];
            const active = trackId === tr.id;
            return (
              <button
                key={tr.id}
                onClick={() => setTrackId(tr.id)}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm border"
                style={
                  active
                    ? { background: color, borderColor: color, color: "#fff" }
                    : { background: color + "14", borderColor: color + "40", color }
                }
              >
                <span className="w-2 h-2 rounded-full shrink-0" style={{ background: active ? "#fff" : color }} />
                {tr.name}
              </button>
            );
          })}
          {authCtxValue.canView("combined_tree") && (
            <button
              onClick={() => setTrackId("combined")}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm ${
                trackId === "combined" ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-600 hover:bg-neutral-200"
              }`}
            >
              <Network size={13} /> {t("Общее дерево OKR")}
            </button>
          )}
          {authCtxValue.canView("projects") && (
            <button
              onClick={() => setTrackId("projects")}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm ${
                trackId === "projects" ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-600 hover:bg-neutral-200"
              }`}
            >
              <Briefcase size={13} /> {t("Проекты")}
            </button>
          )}
          {authCtxValue.canView("projects_dashboard") && (
            <button
              onClick={() => setTrackId("projects-dashboard")}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm ${
                trackId === "projects-dashboard" ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-600 hover:bg-neutral-200"
              }`}
            >
              <PieChart size={13} /> {t("Дашборд проектов")}
            </button>
          )}
          {authCtxValue.canView("tracking_dashboard") && (
            <button
              onClick={() => setTrackId("dashboard")}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm ${
                trackId === "dashboard" ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-600 hover:bg-neutral-200"
              }`}
            >
              <BarChart3 size={13} /> {t("Дашборд трекинга")}
            </button>
          )}
          {authCtxValue.canView("gantt") && (
            <button
              onClick={() => setTrackId("gantt")}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm ${
                trackId === "gantt" ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-600 hover:bg-neutral-200"
              }`}
            >
              <CalendarRange size={13} /> {t("Гант")}
            </button>
          )}
          {authCtxValue.canView("bitrix_import") && (
            <button
              onClick={() => setTrackId("bitrix-import")}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm ${
                trackId === "bitrix-import" ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-600 hover:bg-neutral-200"
              }`}
            >
              <FileUp size={13} /> {t("Импорт из Bitrix24")}
            </button>
          )}
          {authCtxValue.canView("privileges") && (
            <button
              onClick={() => setTrackId("privileges")}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm ${
                trackId === "privileges" ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-600 hover:bg-neutral-200"
              }`}
            >
              <ShieldCheck size={13} /> {t("Управление привилегиями")}
            </button>
          )}
        </div>

        <AuthCtx.Provider value={authCtxValue}>
        <DirectoryCtx.Provider value={directory}>
          {trackId === "combined" && authCtxValue.canView("combined_tree") ? (
            <CombinedTree key={`combined-${refreshKey}`} />
          ) : trackId === "projects" && authCtxValue.canView("projects") ? (
            <LockOverlay locked={!authCtxValue.canEdit("projects")} reason={t("Проекты редактирует администратор")}>
              <ProjectsModule key={`projects-${refreshKey}`} />
            </LockOverlay>
          ) : trackId === "projects-dashboard" && authCtxValue.canView("projects_dashboard") ? (
            <ProjectsDashboard key={`projects-dashboard-${refreshKey}`} />
          ) : trackId === "dashboard" && authCtxValue.canView("tracking_dashboard") ? (
            <TrackingDashboard key={`dashboard-${refreshKey}`} />
          ) : trackId === "gantt" && authCtxValue.canView("gantt") ? (
            <GanttChart key={`gantt-${refreshKey}`} />
          ) : trackId === "bitrix-import" && authCtxValue.canView("bitrix_import") ? (
            <LockOverlay locked={!authCtxValue.canEdit("bitrix_import")} reason={t("Импорт доступен администратору")}>
              <BitrixImportModule key={`bitrix-import-${refreshKey}`} />
            </LockOverlay>
          ) : trackId === "privileges" && authCtxValue.canView("privileges") ? (
            <PrivilegesManager key={`privileges-${refreshKey}`} auth={roleAuth} />
          ) : TRACKS.some((tr) => tr.id === trackId) && authCtxValue.canView(`track_${trackId}`) ? (
            <LockOverlay
              locked={!authCtxValue.canEdit(`track_${trackId}`)}
              reason={
                !authCtxValue.signedIn
                  ? t("Войдите, чтобы редактировать")
                  : t("Доступен только просмотр")
              }
            >
              <TrackEditor key={`${trackId}-${refreshKey}`} trackId={trackId} />
            </LockOverlay>
          ) : (
            <EmptyStateInvite signedIn={authCtxValue.signedIn} />
          )}
        </DirectoryCtx.Provider>
        </AuthCtx.Provider>
      </div>
    </div>
    </LanguageCtx.Provider>
  );
}
