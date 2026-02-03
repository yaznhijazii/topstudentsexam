// ==================== Global State ====================
const state = {
    files: [],
    config: {
        timestampCol: 'طابع زمني',
        nameCol: 'الاسم الكامل',
        schoolCol: 'المدرسة',
        scoreCol: 'النتيجة',
        minNameParts: 2,
        minExams: 1
    },
    results: null
};

// ==================== DOM Elements ====================
const elements = {
    uploadZone: document.getElementById('uploadZone'),
    fileInput: document.getElementById('fileInput'),
    selectFilesBtn: document.getElementById('selectFilesBtn'),
    filesList: document.getElementById('filesList'),
    resetBtn: document.getElementById('resetBtn'),
    processBtn: document.getElementById('processBtn'),
    step1: document.getElementById('step1'),
    step2: document.getElementById('step2'),
    step3: document.getElementById('step3'),
    step4: document.getElementById('step4'),
    processingStatus: document.getElementById('processingStatus'),
    progressFill: document.getElementById('progressFill'),
    toastContainer: document.getElementById('toastContainer')
};

// ==================== Utility Functions ====================
function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast-msg ${type}`;

    const icons = {
        success: '✅',
        error: '❌',
        warning: '⚠️',
        info: 'ℹ️'
    };

    toast.innerHTML = `
        <span class="toast-icon">${icons[type]}</span>
        <div class="toast-text">${message}</div>
    `;

    elements.toastContainer.appendChild(toast);

    setTimeout(() => {
        toast.style.animation = 'slideInLeft 0.5s cubic-bezier(0.23, 1, 0.32, 1) reverse forwards';
        setTimeout(() => toast.remove(), 500);
    }, 4000);
}

function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function updateProgress(percent, message) {
    elements.progressFill.style.width = percent + '%';
    if (message) {
        elements.processingStatus.innerHTML = `<p>${message}</p>`;
    }
}

// ==================== File Upload Handlers ====================
elements.selectFilesBtn.addEventListener('click', () => {
    elements.fileInput.click();
});

elements.uploadZone.addEventListener('click', (e) => {
    // If we didn't click the button itself (which has its own listener), click the input
    if (e.target.id !== 'selectFilesBtn') {
        elements.fileInput.click();
    }
});

elements.uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    elements.uploadZone.classList.add('drag-active');
});

elements.uploadZone.addEventListener('dragleave', () => {
    elements.uploadZone.classList.remove('drag-active');
});

elements.uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    elements.uploadZone.classList.remove('drag-active');
    handleFiles(e.dataTransfer.files);
});

elements.fileInput.addEventListener('change', (e) => {
    handleFiles(e.target.files);
});

function handleFiles(files) {
    const excelFiles = Array.from(files).filter(file =>
        file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );

    if (excelFiles.length === 0) {
        showToast('الرجاء اختيار ملفات Excel فقط (.xlsx)', 'error');
        return;
    }

    state.files.push(...excelFiles);
    renderFilesList();
    showToast(`تم إضافة ${excelFiles.length} ملف`, 'success');

    // Show step 2
    elements.step2.style.display = 'block';
    elements.step2.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderFilesList() {
    if (state.files.length === 0) {
        elements.filesList.innerHTML = '';
        elements.step2.style.display = 'none';
        return;
    }

    elements.filesList.innerHTML = state.files.map((file, index) => `
        <div class="file-card">
            <div class="file-box">📄</div>
            <div class="file-info">
                <div class="file-name">${file.name}</div>
                <div class="file-size">${formatBytes(file.size)}</div>
            </div>
            <div class="remove-file" onclick="removeFile(${index})">✕</div>
        </div>
    `).join('');
}

function removeFile(index) {
    state.files.splice(index, 1);
    renderFilesList();
    showToast('تم حذف الملف', 'info');
}

// ==================== Configuration ====================
document.getElementById('timestampCol').addEventListener('input', (e) => {
    state.config.timestampCol = e.target.value;
});

document.getElementById('nameCol').addEventListener('input', (e) => {
    state.config.nameCol = e.target.value;
});

document.getElementById('schoolCol').addEventListener('input', (e) => {
    state.config.schoolCol = e.target.value;
});

document.getElementById('scoreCol').addEventListener('input', (e) => {
    state.config.scoreCol = e.target.value;
});

document.getElementById('minNameParts').addEventListener('input', (e) => {
    state.config.minNameParts = parseInt(e.target.value) || 2;
});

document.getElementById('minExams').addEventListener('input', (e) => {
    state.config.minExams = parseInt(e.target.value) || 1;
});

// ==================== Processing ====================
elements.processBtn.addEventListener('click', async () => {
    if (state.files.length === 0) {
        showToast('الرجاء رفع ملفات أولاً', 'error');
        return;
    }

    elements.step3.style.display = 'block';
    elements.step3.scrollIntoView({ behavior: 'smooth', block: 'start' });

    try {
        await processExams();
        elements.step3.style.display = 'none';
        elements.step4.style.display = 'block';
        elements.step4.scrollIntoView({ behavior: 'smooth', block: 'start' });
        showToast('تمت المعالجة بنجاح!', 'success');
    } catch (error) {
        showToast('حدث خطأ أثناء المعالجة: ' + error.message, 'error');
        elements.step3.style.display = 'none';
    }
});

async function processExams() {
    updateProgress(0, 'جاري تحميل الملفات...');

    const allData = [];
    const examStats = [];

    // Process each file
    for (let i = 0; i < state.files.length; i++) {
        const file = state.files[i];
        updateProgress((i / state.files.length) * 40, `معالجة: ${file.name}`);

        try {
            const data = await readExcelFile(file);
            if (data && data.length > 0) {
                console.log(`✅ ${file.name}: ${data.length} students`);
                allData.push(...data);
                examStats.push({
                    name: file.name.replace('.xlsx', '').replace('.xls', ''),
                    count: data.length,
                    avgScore: data.reduce((sum, row) => sum + (row.نسبة || 0), 0) / data.length
                });
            } else {
                console.warn(`⚠️ ${file.name}: No data`);
            }
        } catch (error) {
            console.error(`❌ Error processing ${file.name}:`, error);
            showToast(`خطأ في ${file.name}`, 'error');
        }
    }

    console.log(`📊 Total records before deduplication: ${allData.length}`);

    updateProgress(50, 'حذف التكرارات...');

    // Remove duplicates across all exams (keep best attempt per student per exam)
    const examStudentMap = new Map();

    allData.forEach(row => {
        const key = `${row['الاسم الكامل']}|||${row['امتحان']}`;
        const existing = examStudentMap.get(key);

        if (!existing) {
            examStudentMap.set(key, row);
        } else {
            // Keep the one with higher percentage
            if (row.نسبة > existing.نسبة) {
                examStudentMap.set(key, row);
            } else if (row.نسبة === existing.نسبة) {
                // If same score, keep the faster one (better speed rank)
                if (row.ترتيب_السرعة < existing.ترتيب_السرعة) {
                    examStudentMap.set(key, row);
                }
            }
        }
    });

    const dedupedData = Array.from(examStudentMap.values());
    console.log(`🔥 After deduplication: ${dedupedData.length} records`);
    console.log(`✂️ Removed ${allData.length - dedupedData.length} duplicates`);

    updateProgress(65, 'حساب الإحصائيات...');

    // Calculate student profiles
    const studentMap = new Map();

    dedupedData.forEach(row => {
        const name = row['الاسم الكامل'];
        if (!studentMap.has(name)) {
            studentMap.set(name, {
                name: name,
                exams: [],
                totalScore: 0,
                totalSpeed: 0,
                totalSpeedRank: 0,
                count: 0
            });
        }

        const student = studentMap.get(name);
        student.exams.push(row);
        student.totalScore += row.نسبة || 0;
        student.totalSpeed += row.نسبة_السرعة || 0;
        student.totalSpeedRank += row.ترتيب_السرعة || 0;
        student.count++;
    });

    console.log(`👥 Total unique students: ${studentMap.size}`);

    updateProgress(80, 'ترتيب النتائج...');

    // Convert to array and calculate averages
    const students = Array.from(studentMap.values())
        .filter(s => s.count >= state.config.minExams)
        .map(s => ({
            name: s.name,
            examCount: s.count,
            avgScore: s.totalScore / s.count,
            avgSpeed: s.totalSpeed / s.count,
            avgSpeedRank: s.totalSpeedRank / s.count,
            score: (s.count * 100) + (s.totalScore / s.count) + (s.totalSpeed / s.count * 0.1)
        }))
        .sort((a, b) => b.score - a.score);

    console.log(`✅ Qualified students (min ${state.config.minExams} exams): ${students.length}`);

    if (students.length > 0) {
        console.log('🏆 Winner:', students[0]);
    }

    // Exam distribution
    const examDist = {};
    students.forEach(s => {
        examDist[s.examCount] = (examDist[s.examCount] || 0) + 1;
    });
    console.log('📈 Exam distribution:', examDist);

    updateProgress(100, 'اكتمل!');

    // Store results
    state.results = {
        students,
        allData: dedupedData,
        examStats,
        totalStudents: students.length,
        totalExams: state.files.length,
        totalRecords: dedupedData.length
    };

    // Render results
    renderResults();
}

async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {
                    type: 'array',
                    cellDates: true,
                    cellNF: true,
                    cellStyles: true
                });

                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const range = XLSX.utils.decode_range(firstSheet['!ref']);

                // Get headers
                const headers = [];
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellAddress = XLSX.utils.encode_cell({ r: 0, c: C });
                    const cell = firstSheet[cellAddress];
                    headers[C] = cell ? String(cell.v).trim() : '';
                }

                // Find column indices
                const timestampColIdx = headers.findIndex(h => h.includes(state.config.timestampCol) || h.includes('طابع'));
                const nameColIdx = headers.findIndex(h => h.includes(state.config.nameCol) || h.includes('الاسم'));
                const scoreColIdx = headers.findIndex(h => h.includes(state.config.scoreCol) || h.includes('النتيجة'));

                console.log('📋 Column indices:', { timestampColIdx, nameColIdx, scoreColIdx });
                console.log('📋 Headers:', headers);

                if (nameColIdx === -1 || scoreColIdx === -1) {
                    console.warn('⚠️ Required columns not found');
                    showToast(`تحذير: لم يتم العثور على الأعمدة المطلوبة في ${file.name}`, 'warning');
                }

                // Process rows
                const rows = [];
                for (let R = range.s.r + 1; R <= range.e.r; ++R) {
                    const row = {};

                    // Get name
                    if (nameColIdx !== -1) {
                        const nameCell = firstSheet[XLSX.utils.encode_cell({ r: R, c: nameColIdx })];
                        row.name = nameCell ? String(nameCell.v).trim() : '';
                    }

                    // Get timestamp
                    if (timestampColIdx !== -1) {
                        const tsCell = firstSheet[XLSX.utils.encode_cell({ r: R, c: timestampColIdx })];
                        if (tsCell) {
                            if (tsCell.t === 'd') {
                                row.timestamp = tsCell.v;
                            } else if (tsCell.v) {
                                row.timestamp = new Date(tsCell.v);
                            }
                        }
                    }

                    // Get score and max score
                    if (scoreColIdx !== -1) {
                        const scoreCell = firstSheet[XLSX.utils.encode_cell({ r: R, c: scoreColIdx })];
                        if (scoreCell) {
                            // Get the actual score value
                            row.score = parseFloat(scoreCell.v) || 0;

                            // Try to extract max score from number format
                            row.maxScore = extractMaxScoreFromFormat(scoreCell.z || scoreCell.w);

                            // If no format, try to parse from display value
                            if (!row.maxScore && scoreCell.w) {
                                row.maxScore = extractMaxScoreFromString(scoreCell.w);
                            }

                            // Default to 100 if still not found
                            if (!row.maxScore) {
                                row.maxScore = 100;
                            }
                        }
                    }

                    if (row.name) {
                        rows.push(row);
                    }
                }

                console.log(`📊 Read ${rows.length} rows from ${file.name}`);

                // Filter by name parts
                const filtered = rows.filter(row => {
                    const parts = row.name.split(' ').filter(p => p.length > 0);
                    return parts.length >= state.config.minNameParts;
                });

                console.log(`✅ ${filtered.length} rows after filtering (min ${state.config.minNameParts} name parts)`);

                // Sort by timestamp (earliest first = fastest)
                filtered.sort((a, b) => {
                    if (!a.timestamp || !b.timestamp) return 0;
                    return new Date(a.timestamp) - new Date(b.timestamp);
                });

                // Remove duplicates - keep highest score for each name
                const nameMap = new Map();
                filtered.forEach(row => {
                    const existing = nameMap.get(row.name);
                    if (!existing) {
                        nameMap.set(row.name, row);
                    } else {
                        // Keep the one with higher percentage
                        const existingPct = (existing.score / existing.maxScore) * 100;
                        const currentPct = (row.score / row.maxScore) * 100;

                        if (currentPct > existingPct) {
                            nameMap.set(row.name, row);
                        } else if (currentPct === existingPct) {
                            // If same score, keep the faster one (earlier timestamp)
                            if (row.timestamp && existing.timestamp &&
                                new Date(row.timestamp) < new Date(existing.timestamp)) {
                                nameMap.set(row.name, row);
                            }
                        }
                    }
                });

                const unique = Array.from(nameMap.values());
                console.log(`🔥 ${unique.length} unique students after deduplication`);

                // Re-sort by timestamp for speed ranking
                unique.sort((a, b) => {
                    if (!a.timestamp || !b.timestamp) return 0;
                    return new Date(a.timestamp) - new Date(b.timestamp);
                });

                // Calculate speed rankings
                const processed = unique.map((row, index) => {
                    const percentage = (row.score / row.maxScore) * 100;
                    const speedRank = index + 1;
                    const speedPct = unique.length > 1
                        ? ((unique.length - index - 1) / (unique.length - 1)) * 100
                        : 100;

                    return {
                        'الاسم الكامل': row.name,
                        'طابع زمني': row.timestamp,
                        'امتحان': file.name.replace('.xlsx', '').replace('.xls', ''),
                        'درجة': row.score,
                        'درجة قصوى': row.maxScore,
                        'نسبة': percentage,
                        'ترتيب_السرعة': speedRank,
                        'نسبة_السرعة': speedPct
                    };
                });

                console.log(`✨ Processed ${processed.length} final records`);
                if (processed.length > 0) {
                    console.log('📝 Sample:', processed[0]);
                }

                resolve(processed);
            } catch (error) {
                console.error(`❌ Error reading ${file.name}:`, error);
                showToast(`خطأ في قراءة ${file.name}: ${error.message}`, 'error');
                reject(error);
            }
        };

        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function extractMaxScoreFromFormat(format) {
    if (!format || format === 'General') return null;

    // Pattern: '0" / 40"' or '0 / 40' or similar
    const patterns = [
        /\/\s*"?\s*(\d+\.?\d*)/,  // / 40 or / "40"
        /من\s*(\d+\.?\d*)/,        // من 40
        /out of\s*(\d+\.?\d*)/i    // out of 40
    ];

    for (const pattern of patterns) {
        const match = format.match(pattern);
        if (match) {
            const value = parseFloat(match[1]);
            if (value > 0) return value;
        }
    }

    return null;
}

function extractMaxScoreFromString(str) {
    if (!str) return null;

    // Try to parse "15 / 40" or "15/40"
    const match = str.match(/(\d+\.?\d*)\s*\/\s*(\d+\.?\d*)/);
    if (match) {
        const maxScore = parseFloat(match[2]);
        if (maxScore > 0) return maxScore;
    }

    return null;
}

// ==================== Render Results ====================
function renderResults() {
    const { students, totalStudents, totalExams } = state.results;

    // Winner
    if (students.length > 0) {
        const winner = students[0];
        document.getElementById('winnerInfo').innerHTML = `
            <div style="font-size: 1.5rem; font-weight: 800; color: #fff; margin-bottom: 1rem;">${winner.name}</div>
            <div style="display: flex; justify-content: center; gap: 2rem; flex-wrap: wrap;">
                <div>
                    <div style="color: var(--text-dim); font-size: 0.8rem;">متوسط النسبة</div>
                    <div style="font-size: 1.2rem; font-weight: 700; color: var(--primary);">${winner.avgScore.toFixed(2)}%</div>
                </div>
                <div>
                    <div style="color: var(--text-dim); font-size: 0.8rem;">الامتحانات</div>
                    <div style="font-size: 1.2rem; font-weight: 700; color: var(--secondary);">${winner.examCount}</div>
                </div>
                <div>
                    <div style="color: var(--text-dim); font-size: 0.8rem;">النقاط الكلية</div>
                    <div style="font-size: 1.2rem; font-weight: 700; color: var(--success);">${winner.score.toFixed(2)}</div>
                </div>
            </div>
        `;
    }

    // Statistics
    const avgExamsPerStudent = students.length > 0 ? (students.reduce((sum, s) => sum + s.examCount, 0) / students.length) : 0;
    const avgScore = students.length > 0 ? (students.reduce((sum, s) => sum + s.avgScore, 0) / students.length) : 0;

    document.getElementById('statsGrid').innerHTML = `
        <div class="stat-item">
            <span class="stat-val">${totalStudents}</span>
            <span class="stat-lbl">إجمالي الطلاب</span>
        </div>
        <div class="stat-item">
            <span class="stat-val">${totalExams}</span>
            <span class="stat-lbl">عدد الامتحانات</span>
        </div>
        <div class="stat-item">
            <span class="stat-val">${avgExamsPerStudent.toFixed(1)}</span>
            <span class="stat-lbl">متوسط الامتحانات/طالب</span>
        </div>
        <div class="stat-item">
            <span class="stat-val">${avgScore.toFixed(1)}%</span>
            <span class="stat-lbl">متوسط الدرجات</span>
        </div>
    `;

    // Top 15 Students
    const top15 = students.slice(0, 15);
    document.getElementById('topStudentsTable').innerHTML = `
        <thead>
            <tr>
                <th style="width: 60px;">#</th>
                <th>اسم الطالب</th>
                <th>الامتحانات</th>
                <th>النسبة</th>
                <th>السرعة</th>
                <th>النقاط</th>
            </tr>
        </thead>
        <tbody>
            ${top15.map((student, index) => `
                <tr>
                    <td>
                        <span class="rank-dot ${index < 3 ? 'rank-' + (index + 1) : ''}">
                            ${index + 1}
                        </span>
                    </td>
                    <td style="font-weight: 700; color: #fff;">${student.name}</td>
                    <td>${student.examCount}</td>
                    <td>${student.avgScore.toFixed(2)}%</td>
                    <td>${student.avgSpeed.toFixed(2)}%</td>
                    <td style="color: var(--primary); font-weight: 800;">${student.score.toFixed(2)}</td>
                </tr>
            `).join('')}
        </tbody>
    `;

    // Fastest 10 Students
    const fastest10 = [...students]
        .sort((a, b) => b.avgSpeed - a.avgSpeed)
        .slice(0, 10);

    document.getElementById('fastestStudentsTable').innerHTML = `
        <thead>
            <tr>
                <th style="width: 60px;">#</th>
                <th>الاسم</th>
                <th>نسبة السرعة</th>
                <th>متوسط النسبة</th>
                <th>الامتحانات</th>
            </tr>
        </thead>
        <tbody>
            ${fastest10.map((student, index) => `
                <tr>
                    <td><span class="rank-dot">${index + 1}</span></td>
                    <td style="font-weight: 700; color: #fff;">${student.name}</td>
                    <td style="color: var(--secondary); font-weight: 800;">${student.avgSpeed.toFixed(2)}%</td>
                    <td>${student.avgScore.toFixed(2)}%</td>
                    <td>${student.examCount}</td>
                </tr>
            `).join('')}
        </tbody>
    `;

    // Export buttons
    document.getElementById('exportButtons').innerHTML = `
        <button class="btn btn-main" onclick="exportToExcel('students')">
            <span>📄</span> بروفايل الطلاب
        </button>
        <button class="btn btn-main" onclick="exportToExcel('top15')">
            <span>🏅</span> توب 15
        </button>
        <button class="btn btn-main" onclick="exportToExcel('fastest')">
            <span>⚡</span> أسرع الطلاب
        </button>
        <button class="btn btn-main" onclick="exportToExcel('all')">
            <span>📊</span> البيانات الكاملة
        </button>
    `;
}

// ==================== Export Functions ====================
function exportToExcel(type) {
    const { students, allData } = state.results;
    let data, filename;

    switch (type) {
        case 'students':
            data = students.map(s => ({
                'الاسم الكامل': s.name,
                'عدد الامتحانات': s.examCount,
                'متوسط النسبة': s.avgScore.toFixed(2),
                'متوسط نسبة السرعة': s.avgSpeed.toFixed(2),
                'النقاط الكلية': s.score.toFixed(2)
            }));
            filename = 'students_profile.xlsx';
            break;

        case 'top15':
            data = students.slice(0, 15).map((s, i) => ({
                'الترتيب': i + 1,
                'الاسم الكامل': s.name,
                'عدد الامتحانات': s.examCount,
                'متوسط النسبة': s.avgScore.toFixed(2),
                'متوسط نسبة السرعة': s.avgSpeed.toFixed(2),
                'النقاط الكلية': s.score.toFixed(2)
            }));
            filename = 'top_15_students.xlsx';
            break;

        case 'fastest':
            data = [...students]
                .sort((a, b) => b.avgSpeed - a.avgSpeed)
                .slice(0, 10)
                .map((s, i) => ({
                    'الترتيب': i + 1,
                    'الاسم الكامل': s.name,
                    'نسبة السرعة': s.avgSpeed.toFixed(2),
                    'متوسط النسبة': s.avgScore.toFixed(2),
                    'عدد الامتحانات': s.examCount
                }));
            filename = 'fastest_students.xlsx';
            break;

        case 'all':
            data = allData;
            filename = 'all_exam_data.xlsx';
            break;
    }

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, filename);

    showToast(`تم تحميل ${filename}`, 'success');
}

// ==================== Reset ====================
elements.resetBtn.addEventListener('click', () => {
    if (confirm('هل أنت متأكد من إعادة تعيين كل شيء؟')) {
        state.files = [];
        state.results = null;
        elements.fileInput.value = '';
        renderFilesList();
        elements.step2.style.display = 'none';
        elements.step3.style.display = 'none';
        elements.step4.style.display = 'none';
        elements.step1.scrollIntoView({ behavior: 'smooth' });
        showToast('تم إعادة التعيين', 'info');
    }
});

// ==================== Initialize ====================
console.log('🚀 Exam Processing Tool Initialized');
showToast('مرحباً! قم برفع ملفات Excel للبدء', 'info');
