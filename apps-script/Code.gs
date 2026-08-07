/**
 * 電子閱讀護照系統 v3.0 - Google Apps Script 後端
 * 架構：GitHub Pages + Google Apps Script + Google 試算表
 *
 * 使用方式：
 * 1. 將本檔完整貼到「綁定在資料庫試算表」的 Apps Script 專案。
 * 2. 執行 setupSpreadsheet() 一次。
 * 3. 到 Settings 工作表查看 adminKey；到 Teachers 工作表新增教師與班級金鑰。
 * 4. 部署為 Web App：執行身分「自己」、存取權「任何人」。
 * 5. 將 /exec 網址貼到前端 config.js 的 API_URL。
 */

const RP_CONFIG = {
  SPREADSHEET_ID: '', // 綁定試算表時留空；若是獨立 Apps Script，可填入試算表 ID。
  TIME_ZONE: 'Asia/Taipei',
  MAX_REVIEW_LENGTH: 20000,
  MIN_REVIEW_LENGTH: 30,
  VIEW_DEDUP_HOURS: 24
};

const SHEETS = {
  RECORDS: 'ReadingRecords',
  STUDENTS: 'Students',
  VIEWS: 'ArticleViews',
  SETTINGS: 'Settings',
  TEACHERS: 'Teachers',
  BOOKS: 'Books',
  TASKS: 'ReadingTasks'
};

const HEADERS = {
  ReadingRecords: [
    'recordId','taskId','studentId','studentName','grade','className',
    'bookTitle','isbn','author','category','pages','title',
    'reviewContent','reviewSummary','status','isPublic','publishedAt',
    'viewCount','teacherFeedback','source','createdAt','updatedAt',
    'approvedAt','approvedBy'
  ],
  Students: [
    'studentId','studentName','grade','className','email','seatNo','active','createdAt','updatedAt'
  ],
  ArticleViews: [
    'viewId','recordId','viewerKey','viewDate','viewMonth','timestamp'
  ],
  Settings: ['key','value'],
  Teachers: [
    'teacherId','teacherName','className','apiKey','active','createdAt','updatedAt'
  ],
  Books: [
    'bookId','isbn','bookTitle','author','publisher','category','pages','coverUrl','active','createdAt','updatedAt'
  ],
  ReadingTasks: [
    'taskId','title','scopeType','scopeValue','scopeLabel','bookId','bookTitle','isbn',
    'description','pages','dueDate','targetCount','status',
    'createdByRole','createdByClass','createdAt','updatedAt'
  ]
};

/* =========================
 * Web App entry points
 * ========================= */

function doGet(e) {
  try {
    return output_(routeGet_((e && e.parameter) || {}));
  } catch (err) {
    return output_(fail_(safeError_(err)));
  }
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const body = JSON.parse(raw);
    return output_(routePost_(body || {}));
  } catch (err) {
    return output_(fail_(safeError_(err)));
  }
}

function routeGet_(p) {
  switch (String(p.action || '')) {
    case 'ping':
      return ok_({ version: '3.0', now: nowIso_() });
    case 'getPublicDashboard':
      return ok_(getPublicDashboard_(p.period || 'semester'));
    case 'getPublicRanking':
      return ok_({ ranking: getRanking_(p.period || 'semester', p.metric || 'books') });
    case 'getPublicReview':
      return ok_({ review: getPublicReview_(p.recordId) });
    default:
      return fail_('Unknown action');
  }
}

function routePost_(b) {
  switch (String(b.action || '')) {
    // 公開前台
    case 'addReviewView':
      return ok_(withLock_(function () { return addReviewView_(b.recordId, b.viewerKey); }));

    // 管理者 / 教師後台
    case 'login':
      return ok_(login_(b));
    case 'getManagementData':
      return ok_(getManagementData_(b));
    case 'reviewRecord':
      return ok_(withLock_(function () { return reviewRecord_(b); }));
    case 'saveEntity':
      return ok_(withLock_(function () { return saveEntity_(b); }));
    case 'deleteEntity':
      return ok_(withLock_(function () { return deleteEntity_(b); }));

    // 學生指定閱讀任務
    case 'getStudentTasks':
      return ok_(getStudentTasks_(b));
    case 'submitTaskReview':
      return ok_(withLock_(function () { return submitTaskReview_(b); }));

    default:
      return fail_('Unknown action');
  }
}

/* =========================
 * 初始化
 * ========================= */

function setupSpreadsheet() {
  const ss = getSpreadsheet_();
  try { ss.setSpreadsheetTimeZone(RP_CONFIG.TIME_ZONE); } catch (e) {}

  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    ensureHeaders_(sh, HEADERS[name]);
    sh.setFrozenRows(1);
    if (sh.getLastColumn() > 0) sh.autoResizeColumns(1, sh.getLastColumn());
  });

  ensureSetting_('schoolName', '興嘉國小');
  ensureSetting_('quoteText', '閱讀，是隨身攜帶的避難所。');
  ensureSetting_('quoteAuthor', '毛姆');

  if (!getSetting_('adminKey')) {
    setSetting_('adminKey', makeKey_('ADM'));
  }

  ensureSetting_('systemVersion', '3.0');
  ensureSetting_('semesterStartMonth', '8'); // 0-based: 8 = 9 月

  Logger.log('初始化完成。管理者金鑰請到 Settings 工作表的 adminKey 查看。');
  Logger.log('教師請到 Teachers 工作表新增 teacherId / teacherName / className / apiKey / active。');
  return 'OK';
}

function ensureHeaders_(sh, requiredHeaders) {
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const missing = requiredHeaders.filter(function (h) { return current.indexOf(h) === -1; });
  if (missing.length) {
    sh.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
  }
}

/* =========================
 * 權限 / 登入
 * ========================= */

function login_(b) {
  const auth = authorize_(b);
  return {
    role: auth.role,
    className: auth.className || '',
    teacherName: auth.teacherName || '',
    schoolName: getSetting_('schoolName') || ''
  };
}

function authorize_(b) {
  const role = String(b.role || '').trim().toLowerCase();
  const apiKey = String(b.apiKey || '').trim();
  if (!apiKey) throw new Error('請輸入管理金鑰');

  if (role === 'admin') {
    const expected = String(getSetting_('adminKey') || '').trim();
    if (!expected || !secureEquals_(apiKey, expected)) throw new Error('管理者金鑰錯誤');
    return { role: 'admin', className: '', teacherName: '系統管理者' };
  }

  if (role === 'teacher') {
    const className = normalizeClassName_(b.className);
    if (!className) throw new Error('請輸入管理班級');

    const teacher = sheetObjects_(SHEETS.TEACHERS).find(function (t) {
      return truthyDefaultTrue_(t.active) &&
        normalizeClassName_(t.className) === className &&
        secureEquals_(String(t.apiKey || '').trim(), apiKey);
    });

    if (!teacher) throw new Error('班級或教師金鑰錯誤');
    return {
      role: 'teacher',
      className: className,
      teacherName: String(teacher.teacherName || teacher.teacherId || '班級教師')
    };
  }

  throw new Error('登入身分錯誤');
}

function requireAdmin_(auth) {
  if (!auth || auth.role !== 'admin') throw new Error('此操作僅限系統管理者');
}

function canManageRecord_(auth, record) {
  if (auth.role === 'admin') return true;
  return auth.role === 'teacher' && normalizeClassName_(record.className) === auth.className;
}

function canManageTask_(auth, task) {
  if (auth.role === 'admin') return true;
  if (auth.role !== 'teacher') return false;
  return String(task.scopeType) === 'class' &&
    normalizeClassName_(task.scopeValue) === auth.className &&
    normalizeClassName_(task.createdByClass || task.scopeValue) === auth.className;
}

/* =========================
 * 管理中心資料
 * ========================= */

function getManagementData_(b) {
  const auth = authorize_(b);
  let records = sheetObjects_(SHEETS.RECORDS);
  let tasks = sheetObjects_(SHEETS.TASKS);
  let students = sheetObjects_(SHEETS.STUDENTS);
  let books = sheetObjects_(SHEETS.BOOKS);

  if (auth.role === 'teacher') {
    records = records.filter(function (r) { return normalizeClassName_(r.className) === auth.className; });
    tasks = tasks.filter(function (t) {
      return String(t.scopeType) === 'class' && normalizeClassName_(t.scopeValue) === auth.className;
    });
    students = students.filter(function (s) { return normalizeClassName_(s.className) === auth.className; });
    // 教師建立任務時仍需要看到書目，因此 books 保留啟用書籍。
  }

  books = books.filter(function (x) { return truthyDefaultTrue_(x.active); });

  const taskCounts = approvedTaskCounts_();
  tasks = tasks.map(function (t) {
    const x = clone_(t);
    x.scopeLabel = taskScopeLabel_(x);
    x.completedCount = taskCounts[String(x.taskId)] || 0;
    return normalizeTaskForJson_(x);
  });

  records = records.map(normalizeRecordForJson_);
  students = students.map(normalizeStudentForJson_);
  books = books.map(normalizeBookForJson_);

  const approved = records.filter(function (r) { return r.status === 'approved'; });
  const pending = records.filter(function (r) { return r.status === 'pending'; });

  return {
    summary: {
      approved: approved.length,
      pending: pending.length,
      rejected: records.filter(function (r) { return r.status === 'rejected'; }).length,
      activeTasks: tasks.filter(function (t) { return t.status === 'active'; }).length,
      totalPages: approved.reduce(function (sum, r) { return sum + number_(r.pages); }, 0),
      studentCount: unique_(records.map(function (r) { return String(r.studentId || ''); }).filter(Boolean)).length
    },
    records: records,
    students: auth.role === 'admin' ? students : [],
    books: auth.role === 'admin' ? books : books,
    tasks: tasks,
    user: {
      role: auth.role,
      className: auth.className || '',
      teacherName: auth.teacherName || ''
    }
  };
}

function approvedTaskCounts_() {
  const map = {};
  sheetObjects_(SHEETS.RECORDS).forEach(function (r) {
    if (String(r.status).toLowerCase() !== 'approved' || !r.taskId) return;
    const key = String(r.taskId);
    map[key] = (map[key] || 0) + 1;
  });
  return map;
}

/* =========================
 * 心得審核
 * ========================= */

function reviewRecord_(b) {
  const auth = authorize_(b);
  const recordId = requiredText_(b.recordId, '缺少 recordId');
  const status = String(b.status || '').toLowerCase();
  if (['approved', 'rejected'].indexOf(status) === -1) throw new Error('審核狀態只能是 approved 或 rejected');

  const found = findRowById_(SHEETS.RECORDS, 'recordId', recordId);
  if (!found) throw new Error('找不到閱讀紀錄');
  if (!canManageRecord_(auth, found.object)) throw new Error('無權審核其他班級的閱讀心得');

  const patch = {
    status: status,
    teacherFeedback: cleanText_(b.teacherFeedback, 5000),
    updatedAt: new Date(),
    approvedBy: auth.role === 'admin' ? 'admin' : (auth.teacherName || auth.className)
  };

  if (status === 'approved') {
    patch.approvedAt = new Date();
    if (truthy_(found.object.isPublic) && !found.object.publishedAt) patch.publishedAt = new Date();
  } else {
    patch.approvedAt = '';
    patch.approvedBy = '';
    patch.publishedAt = '';
  }

  updateRowByNumber_(SHEETS.RECORDS, found.rowNumber, patch);
  return { recordId: recordId, status: status, counted: status === 'approved' };
}

/* =========================
 * 通用編修 / 刪除
 * ========================= */

function saveEntity_(b) {
  const auth = authorize_(b);
  const entity = String(b.entity || '').toLowerCase();
  const id = String(b.id || '').trim();
  const data = b.data || {};

  if (entity === 'task') return saveTask_(auth, id, data);

  requireAdmin_(auth);
  if (entity === 'record') return saveRecordByAdmin_(id, data);
  if (entity === 'student') return saveStudent_(id, data);
  if (entity === 'book') return saveBook_(id, data);

  throw new Error('不支援的 entity');
}

function saveTask_(auth, id, data) {
  const existing = id ? findRowById_(SHEETS.TASKS, 'taskId', id) : null;
  if (existing && !canManageTask_(auth, existing.object)) throw new Error('無權修改此閱讀任務');

  let scopeType = String(data.scopeType || 'grade').toLowerCase();
  let scopeValue = String(data.scopeValue || '').trim();
  if (['grade', 'class'].indexOf(scopeType) === -1) throw new Error('指定範圍錯誤');

  if (auth.role === 'teacher') {
    scopeType = 'class';
    scopeValue = auth.className;
  }

  if (scopeType === 'grade') {
    const g = normalizeGrade_(scopeValue);
    if (!g) throw new Error('請輸入正確年級，例如 5 或 五年級');
    scopeValue = String(g);
  } else {
    scopeValue = normalizeClassName_(scopeValue);
    if (!scopeValue) throw new Error('請輸入班級');
  }

  const bookTitle = requiredText_(data.bookTitle, '請輸入指定書籍');
  const status = String(data.status || 'active').toLowerCase();
  if (['active', 'closed'].indexOf(status) === -1) throw new Error('任務狀態錯誤');

  const now = new Date();
  const obj = {
    taskId: existing ? existing.object.taskId : Utilities.getUuid(),
    title: cleanText_(data.title || (bookTitle + '閱讀任務'), 120),
    scopeType: scopeType,
    scopeValue: scopeValue,
    scopeLabel: scopeType === 'grade' ? gradeLabel_(scopeValue) : scopeValue,
    bookId: cleanText_(data.bookId, 100),
    bookTitle: cleanText_(bookTitle, 300),
    isbn: cleanText_(data.isbn, 30),
    description: cleanText_(data.description || '完成閱讀並撰寫心得。', 5000),
    pages: clampInt_(data.pages, 0, 99999),
    dueDate: normalizeDateValue_(data.dueDate),
    targetCount: clampInt_(data.targetCount, 0, 99999),
    status: status,
    createdByRole: existing ? existing.object.createdByRole : auth.role,
    createdByClass: existing ? existing.object.createdByClass : (auth.role === 'teacher' ? auth.className : ''),
    createdAt: existing ? existing.object.createdAt : now,
    updatedAt: now
  };

  if (existing) updateRowByNumber_(SHEETS.TASKS, existing.rowNumber, obj);
  else appendObject_(SHEETS.TASKS, obj);

  return { taskId: obj.taskId };
}

function saveRecordByAdmin_(id, data) {
  const existing = id ? findRowById_(SHEETS.RECORDS, 'recordId', id) : null;
  const now = new Date();
  const status = String(data.status || (existing ? existing.object.status : 'pending')).toLowerCase();
  if (['pending', 'approved', 'rejected'].indexOf(status) === -1) throw new Error('閱讀紀錄狀態錯誤');

  const student = data.studentId ? findStudent_(data.studentId) : null;
  const isPublic = truthy_(data.isPublic);
  const obj = {
    recordId: existing ? existing.object.recordId : Utilities.getUuid(),
    taskId: cleanText_(data.taskId || (existing && existing.object.taskId), 100),
    studentId: cleanText_(data.studentId, 100),
    studentName: cleanText_(data.studentName || (student && student.studentName), 100),
    grade: cleanText_(data.grade || (student && student.grade), 20),
    className: normalizeClassName_(data.className || (student && student.className)),
    bookTitle: cleanText_(requiredText_(data.bookTitle, '請輸入書名'), 300),
    isbn: cleanText_(data.isbn, 30),
    author: cleanText_(data.author, 200),
    category: cleanText_(data.category, 100),
    pages: clampInt_(data.pages, 0, 99999),
    title: cleanText_(data.title || ((data.bookTitle || '') + '閱讀心得'), 120),
    reviewContent: cleanReview_(data.reviewContent || ''),
    reviewSummary: cleanText_(data.reviewSummary || String(data.reviewContent || '').slice(0, 120), 300),
    status: status,
    isPublic: isPublic,
    publishedAt: status === 'approved' && isPublic ? (existing && existing.object.publishedAt ? existing.object.publishedAt : now) : '',
    viewCount: existing ? number_(existing.object.viewCount) : 0,
    teacherFeedback: cleanText_(data.teacherFeedback || (existing && existing.object.teacherFeedback), 5000),
    source: cleanText_(data.source || (existing && existing.object.source) || 'manual', 50),
    createdAt: existing ? existing.object.createdAt : now,
    updatedAt: now,
    approvedAt: status === 'approved' ? (existing && existing.object.approvedAt ? existing.object.approvedAt : now) : '',
    approvedBy: status === 'approved' ? 'admin' : ''
  };

  if (existing) updateRowByNumber_(SHEETS.RECORDS, existing.rowNumber, obj);
  else appendObject_(SHEETS.RECORDS, obj);
  return { recordId: obj.recordId };
}

function saveStudent_(id, data) {
  const studentId = requiredText_(data.studentId || id, '請輸入學號');
  const existing = findRowById_(SHEETS.STUDENTS, 'studentId', id || studentId);
  if (!existing) {
    const dup = findRowById_(SHEETS.STUDENTS, 'studentId', studentId);
    if (dup) throw new Error('學號已存在');
  } else if (studentId !== String(existing.object.studentId)) {
    const dup2 = findRowById_(SHEETS.STUDENTS, 'studentId', studentId);
    if (dup2) throw new Error('新學號已存在');
  }

  const now = new Date();
  const obj = {
    studentId: cleanText_(studentId, 100),
    studentName: cleanText_(requiredText_(data.studentName, '請輸入學生姓名'), 100),
    grade: cleanText_(data.grade, 20),
    className: normalizeClassName_(data.className),
    email: cleanText_(data.email, 200),
    seatNo: cleanText_(data.seatNo, 20),
    active: data.active === undefined ? true : truthy_(data.active),
    createdAt: existing ? existing.object.createdAt : now,
    updatedAt: now
  };

  if (existing) updateRowByNumber_(SHEETS.STUDENTS, existing.rowNumber, obj);
  else appendObject_(SHEETS.STUDENTS, obj);
  return { studentId: studentId };
}

function saveBook_(id, data) {
  const existing = id ? findRowById_(SHEETS.BOOKS, 'bookId', id) : null;
  const now = new Date();
  const obj = {
    bookId: existing ? existing.object.bookId : (cleanText_(data.bookId, 100) || Utilities.getUuid()),
    isbn: cleanText_(data.isbn, 30),
    bookTitle: cleanText_(requiredText_(data.bookTitle, '請輸入書名'), 300),
    author: cleanText_(data.author, 200),
    publisher: cleanText_(data.publisher, 200),
    category: cleanText_(data.category, 100),
    pages: clampInt_(data.pages, 0, 99999),
    coverUrl: cleanText_(data.coverUrl, 1000),
    active: data.active === undefined ? true : truthy_(data.active),
    createdAt: existing ? existing.object.createdAt : now,
    updatedAt: now
  };
  if (existing) updateRowByNumber_(SHEETS.BOOKS, existing.rowNumber, obj);
  else appendObject_(SHEETS.BOOKS, obj);
  return { bookId: obj.bookId };
}

function deleteEntity_(b) {
  const auth = authorize_(b);
  const entity = String(b.entity || '').toLowerCase();
  const id = requiredText_(b.id, '缺少 id');

  const map = {
    task: [SHEETS.TASKS, 'taskId'],
    record: [SHEETS.RECORDS, 'recordId'],
    student: [SHEETS.STUDENTS, 'studentId'],
    book: [SHEETS.BOOKS, 'bookId']
  };
  if (!map[entity]) throw new Error('不支援的 entity');

  const found = findRowById_(map[entity][0], map[entity][1], id);
  if (!found) throw new Error('找不到要刪除的資料');

  if (entity === 'task') {
    if (!canManageTask_(auth, found.object)) throw new Error('無權刪除此任務');
    const used = sheetObjects_(SHEETS.RECORDS).some(function (r) { return String(r.taskId) === id; });
    if (used) throw new Error('此任務已有學生繳交紀錄，請改為「已結束」，不可刪除');
  } else {
    requireAdmin_(auth);
  }

  sheet_(map[entity][0]).deleteRow(found.rowNumber);
  return { deleted: true, entity: entity, id: id };
}

/* =========================
 * 學生任務 / 繳交心得
 * ========================= */

function getStudentTasks_(b) {
  const studentId = requiredText_(b.studentId, '請輸入學號');
  const className = normalizeClassName_(requiredText_(b.className, '請輸入班級'));
  const student = findStudent_(studentId);
  if (!student || !truthyDefaultTrue_(student.active)) throw new Error('找不到啟用中的學生資料');
  if (normalizeClassName_(student.className) !== className) throw new Error('學號與班級不符');

  const grade = normalizeGrade_(student.grade || className);
  const submissions = sheetObjects_(SHEETS.RECORDS).filter(function (r) {
    return String(r.studentId) === String(studentId) && !!r.taskId;
  });
  const byTask = {};
  submissions.forEach(function (r) { byTask[String(r.taskId)] = r; });

  const tasks = sheetObjects_(SHEETS.TASKS)
    .filter(function (t) {
      if (String(t.status || '').toLowerCase() !== 'active') return false;
      const type = String(t.scopeType || '').toLowerCase();
      if (type === 'class') return normalizeClassName_(t.scopeValue) === className;
      if (type === 'grade') return normalizeGrade_(t.scopeValue) === grade;
      return false;
    })
    .sort(function (a, b2) {
      return dateMs_(a.dueDate, 8640000000000000) - dateMs_(b2.dueDate, 8640000000000000);
    })
    .map(function (t) {
      const r = byTask[String(t.taskId)] || {};
      return {
        taskId: String(t.taskId || ''),
        title: String(t.title || ''),
        scopeType: String(t.scopeType || ''),
        scopeValue: String(t.scopeValue || ''),
        scopeLabel: taskScopeLabel_(t),
        bookTitle: String(t.bookTitle || ''),
        isbn: String(t.isbn || ''),
        description: String(t.description || ''),
        pages: number_(t.pages),
        dueDate: toIso_(t.dueDate),
        status: String(t.status || ''),
        submissionStatus: String(r.status || ''),
        recordId: String(r.recordId || ''),
        reviewTitle: String(r.title || ''),
        submittedPages: number_(r.pages),
        reviewContent: String(r.reviewContent || ''),
        isPublic: truthy_(r.isPublic),
        teacherFeedback: String(r.teacherFeedback || '')
      };
    });

  return {
    student: normalizeStudentForJson_(student),
    tasks: tasks
  };
}

function submitTaskReview_(b) {
  const studentId = requiredText_(b.studentId, '缺少學號');
  const className = normalizeClassName_(requiredText_(b.className, '缺少班級'));
  const taskId = requiredText_(b.taskId, '缺少 taskId');

  const student = findStudent_(studentId);
  if (!student || !truthyDefaultTrue_(student.active)) throw new Error('找不到啟用中的學生資料');
  if (normalizeClassName_(student.className) !== className) throw new Error('學號與班級不符');

  const taskFound = findRowById_(SHEETS.TASKS, 'taskId', taskId);
  if (!taskFound) throw new Error('找不到指定閱讀任務');
  const task = taskFound.object;
  if (String(task.status || '').toLowerCase() !== 'active') throw new Error('此閱讀任務已結束');
  if (!taskAppliesToStudent_(task, student)) throw new Error('此任務不屬於你的年級或班級');

  const reviewContent = cleanReview_(b.reviewContent);
  if (reviewContent.length < RP_CONFIG.MIN_REVIEW_LENGTH) {
    throw new Error('閱讀心得至少需 ' + RP_CONFIG.MIN_REVIEW_LENGTH + ' 字');
  }

  const pages = clampInt_(b.pages, 0, 99999);
  const reviewTitle = cleanText_(b.title || (task.bookTitle + '閱讀心得'), 120);
  const now = new Date();

  const existing = sheetObjectsWithRows_(SHEETS.RECORDS).find(function (x) {
    return String(x.object.studentId) === String(studentId) && String(x.object.taskId) === String(taskId);
  });

  if (existing && String(existing.object.status).toLowerCase() === 'approved') {
    throw new Error('此任務已審核通過並完成，不能重複繳交');
  }

  const book = findBookForTask_(task);
  const obj = {
    recordId: existing ? existing.object.recordId : Utilities.getUuid(),
    taskId: taskId,
    studentId: String(student.studentId),
    studentName: String(student.studentName || ''),
    grade: String(student.grade || normalizeGrade_(student.className) || ''),
    className: normalizeClassName_(student.className),
    bookTitle: String(task.bookTitle || ''),
    isbn: String(task.isbn || (book && book.isbn) || ''),
    author: String((book && book.author) || ''),
    category: String((book && book.category) || ''),
    pages: pages,
    title: reviewTitle,
    reviewContent: reviewContent,
    reviewSummary: reviewContent.slice(0, 120),
    status: 'pending',
    isPublic: !!b.isPublic,
    publishedAt: '',
    viewCount: existing ? number_(existing.object.viewCount) : 0,
    teacherFeedback: '',
    source: 'task',
    createdAt: existing ? existing.object.createdAt : now,
    updatedAt: now,
    approvedAt: '',
    approvedBy: ''
  };

  if (existing) updateRowByNumber_(SHEETS.RECORDS, existing.rowNumber, obj);
  else appendObject_(SHEETS.RECORDS, obj);

  return { recordId: obj.recordId, status: 'pending' };
}

function taskAppliesToStudent_(task, student) {
  const type = String(task.scopeType || '').toLowerCase();
  if (type === 'class') return normalizeClassName_(task.scopeValue) === normalizeClassName_(student.className);
  if (type === 'grade') return normalizeGrade_(task.scopeValue) === normalizeGrade_(student.grade || student.className);
  return false;
}

function findBookForTask_(task) {
  const books = sheetObjects_(SHEETS.BOOKS);
  if (task.bookId) {
    const byId = books.find(function (b) { return String(b.bookId) === String(task.bookId); });
    if (byId) return byId;
  }
  if (task.isbn) {
    const byIsbn = books.find(function (b) { return String(b.isbn) === String(task.isbn); });
    if (byIsbn) return byIsbn;
  }
  return books.find(function (b) { return String(b.bookTitle) === String(task.bookTitle); }) || null;
}

/* =========================
 * 公開前台
 * ========================= */

function getPublicDashboard_(period) {
  const rows = publicRecords_();
  return {
    generatedAt: nowIso_(),
    summary: summary_(rows),
    studentRanking: getRankingFromRows_(rows, period, 'books'),
    gradeStatistics: gradeStats_(rows),
    monthlyStatistics: monthlyStats_(rows),
    categoryStatistics: categoryStats_(rows),
    latestReviews: reviews_(rows, 'latest', 10),
    popularReviews: reviews_(rows, 'month', 10),
    popularReviewsAll: reviews_(rows, 'all', 10),
    quote: getQuote_()
  };
}

function publicRecords_() {
  return sheetObjects_(SHEETS.RECORDS).filter(function (r) {
    return String(r.status || '').toLowerCase() === 'approved' && truthy_(r.isPublic);
  });
}

function summary_(rows) {
  const books = unique_(rows.map(function (r) { return String(r.isbn || r.bookTitle || '').trim(); }).filter(Boolean));
  const students = unique_(rows.map(function (r) { return String(r.studentId || (r.grade + '-' + r.className + '-' + r.studentName)); }));
  return {
    articleCount: rows.length,
    bookCount: books.length,
    totalPages: rows.reduce(function (s, r) { return s + number_(r.pages); }, 0),
    studentCount: students.length
  };
}

function getRanking_(period, metric) {
  return getRankingFromRows_(publicRecords_(), period, metric);
}

function getRankingFromRows_(rows, period, metric) {
  const filtered = filterPeriod_(rows, period);
  const map = {};
  filtered.forEach(function (r) {
    const key = String(r.studentId || (r.grade + '-' + r.className + '-' + r.studentName));
    if (!map[key]) {
      map[key] = {
        studentName: String(r.studentName || ''),
        className: normalizeClassName_(r.className),
        books: 0,
        pages: 0,
        articles: 0,
        _books: {}
      };
    }
    const bookKey = String(r.isbn || r.bookTitle || '');
    if (bookKey) map[key]._books[bookKey] = true;
    map[key].pages += number_(r.pages);
    map[key].articles += 1;
  });

  const list = Object.keys(map).map(function (k) {
    const x = map[k];
    return {
      studentName: x.studentName,
      className: x.className,
      books: Object.keys(x._books).length,
      pages: x.pages,
      articles: x.articles
    };
  });

  if (['books', 'pages', 'articles'].indexOf(metric) === -1) metric = 'books';
  return list.sort(function (a, b) {
    return number_(b[metric]) - number_(a[metric]) || b.pages - a.pages || b.articles - a.articles;
  }).slice(0, 20);
}

function gradeStats_(rows) {
  return [1,2,3,4,5,6].map(function (g) {
    const data = rows.filter(function (r) { return normalizeGrade_(r.grade || r.className) === g; });
    return {
      grade: g + '年級',
      books: unique_(data.map(function (r) { return String(r.isbn || r.bookTitle || ''); }).filter(Boolean)).length,
      pages: data.reduce(function (s, r) { return s + number_(r.pages); }, 0),
      articles: data.length
    };
  });
}

function monthlyStats_(rows) {
  const now = new Date();
  const startYear = now.getMonth() >= 8 ? now.getFullYear() : now.getFullYear() - 1;
  const result = [];
  for (let i = 0; i < 12; i++) {
    const d = new Date(startYear, 8 + i, 1);
    const next = new Date(startYear, 8 + i + 1, 1);
    if (d > now && i > 9) break;
    result.push({
      month: (d.getMonth() + 1) + '月',
      articles: rows.filter(function (r) {
        const x = parseDate_(r.publishedAt || r.approvedAt || r.createdAt);
        return x && x >= d && x < next;
      }).length
    });
  }
  return result;
}

function categoryStats_(rows) {
  const map = {};
  rows.forEach(function (r) {
    const key = String(r.category || '其他').trim() || '其他';
    map[key] = (map[key] || 0) + 1;
  });
  return Object.keys(map).map(function (k) { return { category: k, count: map[k] }; })
    .sort(function (a, b) { return b.count - a.count; }).slice(0, 8);
}

function reviews_(rows, type, limit) {
  let data = rows.slice();
  if (type === 'month') {
    const start = new Date();
    start.setDate(1);
    start.setHours(0, 0, 0, 0);
    data = data.filter(function (r) {
      const d = parseDate_(r.publishedAt || r.approvedAt || r.createdAt);
      return d && d >= start;
    }).sort(function (a, b) { return number_(b.viewCount) - number_(a.viewCount); });
  } else if (type === 'all') {
    data.sort(function (a, b) { return number_(b.viewCount) - number_(a.viewCount); });
  } else {
    data.sort(function (a, b) {
      return dateMs_(b.publishedAt || b.approvedAt || b.createdAt, 0) - dateMs_(a.publishedAt || a.approvedAt || a.createdAt, 0);
    });
  }
  return data.slice(0, limit).map(publicReviewObject_);
}

function getPublicReview_(recordId) {
  const row = publicRecords_().find(function (r) { return String(r.recordId) === String(recordId); });
  if (!row) throw new Error('找不到公開心得');
  return publicReviewObject_(row);
}

function publicReviewObject_(r) {
  return {
    recordId: String(r.recordId || ''),
    studentName: maskStudentName_(String(r.studentName || '')),
    className: normalizeClassName_(r.className),
    bookTitle: String(r.bookTitle || ''),
    title: String(r.title || ((r.bookTitle || '') + '閱讀心得')),
    summary: String(r.reviewSummary || String(r.reviewContent || '').slice(0, 120)),
    content: String(r.reviewContent || ''),
    publishedAt: toIso_(r.publishedAt || r.approvedAt || r.createdAt),
    viewCount: number_(r.viewCount)
  };
}

function addReviewView_(recordId, viewerKey) {
  recordId = requiredText_(recordId, '缺少 recordId');
  viewerKey = cleanText_(requiredText_(viewerKey, '缺少 viewerKey'), 200);

  const record = findRowById_(SHEETS.RECORDS, 'recordId', recordId);
  if (!record || String(record.object.status).toLowerCase() !== 'approved' || !truthy_(record.object.isPublic)) {
    throw new Error('找不到公開心得');
  }

  const now = new Date();
  const cutoff = new Date(now.getTime() - RP_CONFIG.VIEW_DEDUP_HOURS * 60 * 60 * 1000);
  const duplicate = sheetObjects_(SHEETS.VIEWS).some(function (v) {
    return String(v.recordId) === String(recordId) &&
      String(v.viewerKey) === String(viewerKey) &&
      parseDate_(v.timestamp) && parseDate_(v.timestamp) >= cutoff;
  });
  if (duplicate) return { counted: false, viewCount: number_(record.object.viewCount) };

  const day = Utilities.formatDate(now, RP_CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  appendObject_(SHEETS.VIEWS, {
    viewId: Utilities.getUuid(),
    recordId: recordId,
    viewerKey: viewerKey,
    viewDate: day,
    viewMonth: day.slice(0, 7),
    timestamp: now
  });

  const count = number_(record.object.viewCount) + 1;
  updateRowByNumber_(SHEETS.RECORDS, record.rowNumber, { viewCount: count });
  return { counted: true, viewCount: count };
}

function filterPeriod_(rows, period) {
  const now = new Date();
  let start;
  if (period === 'week') {
    start = new Date(now);
    start.setDate(now.getDate() - 6);
    start.setHours(0, 0, 0, 0);
  } else if (period === 'month') {
    start = new Date(now.getFullYear(), now.getMonth(), 1);
  } else {
    start = now.getMonth() >= 8 ? new Date(now.getFullYear(), 8, 1) : new Date(now.getFullYear() - 1, 8, 1);
  }
  return rows.filter(function (r) {
    const d = parseDate_(r.publishedAt || r.approvedAt || r.createdAt);
    return d && d >= start;
  });
}

/* =========================
 * 試算表 helpers
 * ========================= */

function getSpreadsheet_() {
  if (RP_CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(RP_CONFIG.SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到綁定的 Google 試算表，請設定 RP_CONFIG.SPREADSHEET_ID');
  return ss;
}

function sheet_(name) {
  const sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) throw new Error('找不到工作表：' + name + '，請先執行 setupSpreadsheet()');
  return sh;
}

function sheetObjects_(name) {
  return sheetObjectsWithRows_(name).map(function (x) { return x.object; });
}

function sheetObjectsWithRows_(name) {
  const sh = sheet_(name);
  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!row.some(function (v) { return v !== '' && v !== null; })) continue;
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    out.push({ object: obj, rowNumber: r + 1 });
  }
  return out;
}

function appendObject_(name, obj) {
  const sh = sheet_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  sh.appendRow(headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; }));
}

function updateRowByNumber_(name, rowNumber, patch) {
  const sh = sheet_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (h, i) {
    if (Object.prototype.hasOwnProperty.call(patch, h)) current[i] = patch[h];
  });
  sh.getRange(rowNumber, 1, 1, headers.length).setValues([current]);
}

function findRowById_(name, idHeader, id) {
  const target = String(id || '');
  return sheetObjectsWithRows_(name).find(function (x) {
    return String(x.object[idHeader] || '') === target;
  }) || null;
}

function findStudent_(studentId) {
  const found = findRowById_(SHEETS.STUDENTS, 'studentId', studentId);
  return found ? found.object : null;
}

/* =========================
 * Settings helpers
 * ========================= */

function getSetting_(key) {
  const row = sheetObjects_(SHEETS.SETTINGS).find(function (x) { return String(x.key) === String(key); });
  return row ? row.value : '';
}

function setSetting_(key, value) {
  const found = findRowById_(SHEETS.SETTINGS, 'key', key);
  if (found) updateRowByNumber_(SHEETS.SETTINGS, found.rowNumber, { value: value });
  else appendObject_(SHEETS.SETTINGS, { key: key, value: value });
}

function ensureSetting_(key, value) {
  if (!String(getSetting_(key) || '')) setSetting_(key, value);
}

/* =========================
 * 格式 / 驗證 helpers
 * ========================= */

function normalizeRecordForJson_(r) {
  return {
    recordId: String(r.recordId || ''),
    taskId: String(r.taskId || ''),
    studentId: String(r.studentId || ''),
    studentName: String(r.studentName || ''),
    grade: String(r.grade || ''),
    className: normalizeClassName_(r.className),
    bookTitle: String(r.bookTitle || ''),
    isbn: String(r.isbn || ''),
    author: String(r.author || ''),
    category: String(r.category || ''),
    pages: number_(r.pages),
    title: String(r.title || ''),
    reviewContent: String(r.reviewContent || ''),
    reviewSummary: String(r.reviewSummary || ''),
    status: String(r.status || 'pending').toLowerCase(),
    isPublic: truthy_(r.isPublic),
    publishedAt: toIso_(r.publishedAt),
    viewCount: number_(r.viewCount),
    teacherFeedback: String(r.teacherFeedback || ''),
    source: String(r.source || ''),
    createdAt: toIso_(r.createdAt),
    updatedAt: toIso_(r.updatedAt),
    approvedAt: toIso_(r.approvedAt),
    approvedBy: String(r.approvedBy || '')
  };
}

function normalizeStudentForJson_(s) {
  return {
    studentId: String(s.studentId || ''),
    studentName: String(s.studentName || ''),
    grade: String(s.grade || ''),
    className: normalizeClassName_(s.className),
    email: String(s.email || ''),
    seatNo: String(s.seatNo || ''),
    active: truthyDefaultTrue_(s.active)
  };
}

function normalizeBookForJson_(b) {
  return {
    bookId: String(b.bookId || ''),
    isbn: String(b.isbn || ''),
    bookTitle: String(b.bookTitle || ''),
    author: String(b.author || ''),
    publisher: String(b.publisher || ''),
    category: String(b.category || ''),
    pages: number_(b.pages),
    coverUrl: String(b.coverUrl || ''),
    active: truthyDefaultTrue_(b.active)
  };
}

function normalizeTaskForJson_(t) {
  const x = clone_(t);
  x.taskId = String(x.taskId || '');
  x.title = String(x.title || '');
  x.scopeType = String(x.scopeType || '');
  x.scopeValue = String(x.scopeValue || '');
  x.scopeLabel = String(x.scopeLabel || taskScopeLabel_(x));
  x.bookId = String(x.bookId || '');
  x.bookTitle = String(x.bookTitle || '');
  x.isbn = String(x.isbn || '');
  x.description = String(x.description || '');
  x.pages = number_(x.pages);
  x.dueDate = toIso_(x.dueDate);
  x.targetCount = number_(x.targetCount);
  x.completedCount = number_(x.completedCount);
  x.status = String(x.status || 'active').toLowerCase();
  x.createdAt = toIso_(x.createdAt);
  x.updatedAt = toIso_(x.updatedAt);
  return x;
}

function normalizeClassName_(value) {
  let s = String(value || '').trim().replace(/\s+/g, '');
  if (!s) return '';
  s = s.replace('年级', '年級').replace('班級', '');
  // 若輸入「三年一班」等格式，直接保留；若是全形數字則轉半形。
  s = toHalfWidthDigits_(s);
  return s;
}

function normalizeGrade_(value) {
  const s = toHalfWidthDigits_(String(value || '').trim());
  const digit = s.match(/[1-6]/);
  if (digit) return Number(digit[0]);
  const map = { '一':1, '二':2, '三':3, '四':4, '五':5, '六':6 };
  for (const k in map) if (s.indexOf(k) !== -1) return map[k];
  return 0;
}

function gradeLabel_(value) {
  const g = normalizeGrade_(value);
  return g ? g + '年級' : String(value || '');
}

function taskScopeLabel_(t) {
  return String(t.scopeType) === 'grade' ? gradeLabel_(t.scopeValue) : normalizeClassName_(t.scopeValue);
}

function maskStudentName_(name) {
  const s = String(name || '').trim();
  if (!s) return '匿名同學';
  if (s.length === 1) return s;
  if (s.length === 2) return s.charAt(0) + '○';
  return s.charAt(0) + '○' + s.charAt(s.length - 1);
}

function cleanReview_(value) {
  const s = String(value || '').trim();
  if (s.length > RP_CONFIG.MAX_REVIEW_LENGTH) throw new Error('心得內容過長');
  return s;
}

function cleanText_(value, maxLen) {
  let s = String(value === undefined || value === null ? '' : value).trim();
  if (maxLen && s.length > maxLen) s = s.slice(0, maxLen);
  return s;
}

function requiredText_(value, message) {
  const s = String(value || '').trim();
  if (!s) throw new Error(message || '缺少必要欄位');
  return s;
}

function clampInt_(value, min, max) {
  let n = Math.round(number_(value));
  if (n < min) n = min;
  if (n > max) n = max;
  return n;
}

function truthy_(v) {
  if (v === true) return true;
  const s = String(v || '').trim().toLowerCase();
  return ['true','1','yes','y','公開','是','啟用','active'].indexOf(s) !== -1;
}

function truthyDefaultTrue_(v) {
  if (v === '' || v === null || v === undefined) return true;
  return truthy_(v);
}

function number_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function unique_(arr) {
  const seen = {};
  return arr.filter(function (x) {
    const k = String(x);
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
}

function parseDate_(v) {
  if (!v) return null;
  const d = v instanceof Date ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function dateMs_(v, fallback) {
  const d = parseDate_(v);
  return d ? d.getTime() : fallback;
}

function toIso_(v) {
  const d = parseDate_(v);
  return d ? d.toISOString() : '';
}

function normalizeDateValue_(v) {
  if (!v) return '';
  const d = parseDate_(v);
  if (!d) throw new Error('日期格式錯誤');
  return d;
}

function nowIso_() {
  return new Date().toISOString();
}

function toHalfWidthDigits_(s) {
  return String(s).replace(/[０-９]/g, function (c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });
}

function clone_(o) {
  const x = {};
  Object.keys(o || {}).forEach(function (k) { x[k] = o[k]; });
  return x;
}

function secureEquals_(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function makeKey_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); }
  finally { lock.releaseLock(); }
}

function getQuote_() {
  return {
    text: String(getSetting_('quoteText') || '閱讀，是隨身攜帶的避難所。'),
    author: String(getSetting_('quoteAuthor') || '毛姆')
  };
}

function safeError_(err) {
  return err && err.message ? String(err.message) : '系統發生錯誤';
}

function ok_(data) {
  return { success: true, data: data };
}

function fail_(message) {
  return { success: false, message: String(message || '操作失敗') };
}

function output_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
