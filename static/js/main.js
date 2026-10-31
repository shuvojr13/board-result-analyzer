// main.js - Frontend Logic for Board Result Analyzer (Bulk, Individual & Comprehensive Analysis)

document.addEventListener("DOMContentLoaded", () => {
    // Shared Controls
    const examSelect = document.getElementById("exam");
    const yearSelect = document.getElementById("year");
    const boardSelect = document.getElementById("board");
    const demoModeCheck = document.getElementById("demoMode");

    // Bulk Mode Elements
    const bulkForm = document.getElementById("bulkForm");
    const fileInput = document.getElementById("fileInput");
    const dropZone = document.getElementById("dropZone");
    const fileInfo = document.getElementById("fileInfo");
    const fileName = document.getElementById("fileName");
    const submitBulkBtn = document.getElementById("submitBulkBtn");

    // Individual Mode Elements
    const individualForm = document.getElementById("individualForm");
    const indRollInput = document.getElementById("indRoll");
    const indRegInput = document.getElementById("indReg");
    const submitIndBtn = document.getElementById("submitIndBtn");

    // Marksheet Card Elements
    const individualResultCard = document.getElementById("individualResultCard");
    const marksheetName = document.getElementById("marksheetName");
    const marksheetExam = document.getElementById("marksheetExam");
    const marksheetYear = document.getElementById("marksheetYear");
    const marksheetBoard = document.getElementById("marksheetBoard");
    const marksheetGpa = document.getElementById("marksheetGpa");
    const marksheetResultBadge = document.getElementById("marksheetResultBadge");
    const marksheetRoll = document.getElementById("marksheetRoll");
    const marksheetReg = document.getElementById("marksheetReg");
    const marksheetGroup = document.getElementById("marksheetGroup");
    const marksheetType = document.getElementById("marksheetType");
    const marksheetFather = document.getElementById("marksheetFather");
    const marksheetMother = document.getElementById("marksheetMother");
    const marksheetInstitute = document.getElementById("marksheetInstitute");
    const marksheetInstituteRow = document.getElementById("marksheetInstituteRow");
    const marksheetGradesBody = document.getElementById("marksheetGradesBody");
    const cumillaHscAlert = document.getElementById("cumillaHscAlert");

    // Progress Elements (Bulk Scraping)
    const progressCard = document.getElementById("progressCard");
    const progressTitle = document.getElementById("progressTitle");
    const progressBar = document.getElementById("progressBar");
    const progressPercentBadge = document.getElementById("progressPercentBadge");
    const statusMessage = document.getElementById("statusMessage");
    const studentCounter = document.getElementById("studentCounter");

    // Download Elements
    const downloadSection = document.getElementById("downloadSection");
    const downloadXlsxBtn = document.getElementById("downloadXlsxBtn");
    const downloadCsvBtn = document.getElementById("downloadCsvBtn");

    // Results Table Elements
    const resultsBody = document.getElementById("resultsBody");
    const resultsCount = document.getElementById("resultsCount");

    // Result Analysis Buttons & Elements
    const openAnalysisBtn = document.getElementById("openAnalysisBtn");
    const headerAnalysisBtn = document.getElementById("headerAnalysisBtn");
    const bottomAnalysisBtn = document.getElementById("bottomAnalysisBtn");
    const uploadFileForAnalysisBtn = document.getElementById("uploadFileForAnalysisBtn");
    const analysisFileInput = document.getElementById("analysisFileInput");
    const exportAnalysisExcelBtn = document.getElementById("exportAnalysisExcelBtn");
    const printAnalysisBtn = document.getElementById("printAnalysisBtn");

    // In-memory data store for live analysis
    let eventSource = null;
    let extractedCount = 0;
    let currentRecords = [];
    let latestAnalysisData = null;
    let latestJobId = null;
    let analysisModalInstance = null;

    // Bootstrap Modal initialization
    const analysisModalEl = document.getElementById("resultAnalysisModal");
    if (analysisModalEl && typeof bootstrap !== "undefined") {
        analysisModalInstance = new bootstrap.Modal(analysisModalEl);
    }

    // Dynamic Server Notice for Cumilla HSC
    function updateCumillaServerNotice() {
        const isCumilla = (boardSelect.value || "").toLowerCase() === "comilla";
        const isHsc = (examSelect.value || "").toLowerCase().includes("hsc");
        if (isCumilla && isHsc) {
            if (cumillaHscAlert) cumillaHscAlert.style.display = "flex";
            indRollInput.placeholder = "e.g. 160000";
            indRegInput.placeholder = "Optional (Not needed for Cumilla HSC)";
        } else {
            if (cumillaHscAlert) cumillaHscAlert.style.display = "none";
            indRollInput.placeholder = "e.g. 100101";
            indRegInput.placeholder = "e.g. 1512345678 (optional)";
        }
    }

    examSelect.addEventListener("change", updateCumillaServerNotice);
    boardSelect.addEventListener("change", updateCumillaServerNotice);
    updateCumillaServerNotice();

    // --- Drag and Drop Handlers for Bulk File ---
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('dragover');
        });
    });

    dropZone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            fileInput.files = files;
            displaySelectedFile(files[0].name);
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            displaySelectedFile(fileInput.files[0].name);
        }
    });

    function displaySelectedFile(name) {
        fileName.textContent = name;
        fileInfo.classList.remove('d-none');
    }

    // ==========================================
    // 1. INDIVIDUAL STUDENT RESULT CHECK
    // ==========================================
    individualForm.addEventListener("submit", async (e) => {
        e.preventDefault();

        const rollVal = indRollInput.value.trim();
        const regVal = indRegInput.value.trim();

        if (!rollVal) {
            alert("Please enter a valid Roll Number.");
            return;
        }

        // UI Loading State
        submitIndBtn.disabled = true;
        submitIndBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-2" role="status"></span>Fetching Result...`;
        individualResultCard.style.display = "none";

        const formData = new FormData();
        formData.append("exam", examSelect.value);
        formData.append("year", yearSelect.value);
        formData.append("board", boardSelect.value);
        formData.append("roll", rollVal);
        formData.append("reg", regVal);
        formData.append("demo_mode", demoModeCheck.checked ? "true" : "false");

        try {
            const response = await fetch("/api/check-individual", {
                method: "POST",
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || "Failed to fetch student result.");
            }

            const resJson = await response.json();
            const student = resJson.data;

            // Save in currentRecords
            currentRecords.push(student);
            latestAnalysisData = null; // Invalidate to recompute on next click

            // Render detailed Marksheet Card
            renderIndividualMarksheet(student);

            // Also append to the preview log table
            appendResultRow(student);

        } catch (error) {
            alert(`Error: ${error.message}`);
        } finally {
            submitIndBtn.disabled = false;
            submitIndBtn.innerHTML = `<i class="bi bi-search me-1"></i> Check Single Result`;
        }
    });

    function renderIndividualMarksheet(student) {
        marksheetName.textContent = student.Name || "N/A";
        marksheetExam.textContent = student.Exam || examSelect.options[examSelect.selectedIndex].text;
        marksheetYear.textContent = student.Year || yearSelect.value;
        marksheetBoard.textContent = student.Board || boardSelect.options[boardSelect.selectedIndex].text;
        marksheetGpa.textContent = student.GPA || "0.00";
        marksheetRoll.textContent = student.Roll || "-";
        marksheetReg.textContent = student.Registration || "N/A";
        marksheetGroup.textContent = student.Group || "-";
        marksheetType.textContent = student.Type || "-";
        marksheetFather.textContent = student["Father Name"] || "-";
        marksheetMother.textContent = student["Mother Name"] || "-";

        if (marksheetInstitute && marksheetInstituteRow) {
            if (student.Institute) {
                marksheetInstitute.textContent = student.Institute;
                marksheetInstituteRow.style.display = "block";
            } else {
                marksheetInstituteRow.style.display = "none";
            }
        }

        // Result badge color
        const isNotPublished = (student.Result || "").includes("Not Published");
        const isError = (student.Result || "").includes("Fail") || (student.Result || "").includes("Not Found") || (student.Result || "").includes("Error");

        marksheetResultBadge.textContent = student.Result || "N/A";
        if (isNotPublished) {
            marksheetResultBadge.className = "badge bg-warning-subtle text-dark fs-7 px-3 py-2 border border-warning-subtle d-inline-block mb-1";
        } else if (isError) {
            marksheetResultBadge.className = "badge bg-danger-subtle text-danger fs-7 px-3 py-2 border border-danger-subtle d-inline-block mb-1";
        } else {
            marksheetResultBadge.className = "badge bg-success-subtle text-success fs-7 px-3 py-2 border border-success-subtle d-inline-block mb-1";
        }

        // Populate Subject Grades Table
        marksheetGradesBody.innerHTML = "";
        const grades = student.Subject_Grades || {};
        const subjectKeys = Object.keys(grades);

        if (subjectKeys.length === 0) {
            marksheetGradesBody.innerHTML = `
                <tr>
                    <td colspan="2" class="text-center text-muted py-3">
                        ${student.Subject_Grades_Summary || "No subject-wise breakdown available."}
                    </td>
                </tr>
            `;
        } else {
            subjectKeys.forEach(subj => {
                const grade = grades[subj];
                let badgeClass = "badge bg-secondary";
                if (grade === "A+") badgeClass = "badge-gpa-gold";
                else if (grade === "A" || grade === "A-") badgeClass = "badge-gpa-green";
                else if (grade === "B" || grade === "C" || grade === "D") badgeClass = "badge-gpa-blue";
                else if (grade === "F") badgeClass = "badge-gpa-red";

                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td class="ps-3 fw-medium text-dark">${subj}</td>
                    <td class="text-center"><span class="badge ${badgeClass} px-3 py-1">${grade}</span></td>
                `;
                marksheetGradesBody.appendChild(tr);
            });
        }

        // Show card and scroll into view smoothly
        individualResultCard.style.display = "block";
        individualResultCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }

    // ==========================================
    // 2. BULK EXTRACTION (CSV / EXCEL UPLOAD)
    // ==========================================
    bulkForm.addEventListener("submit", async (e) => {
        e.preventDefault();

        if (!fileInput.files || fileInput.files.length === 0) {
            alert("Please select or drop a CSV or Excel file containing Roll numbers.");
            return;
        }

        // Reset UI State
        extractedCount = 0;
        currentRecords = [];
        latestAnalysisData = null;
        latestJobId = null;

        resultsCount.textContent = "0 Extracted";
        resultsBody.innerHTML = "";
        downloadSection.style.setProperty("display", "none", "important");
        downloadXlsxBtn.classList.remove("pulse-success");
        individualResultCard.style.display = "none";

        submitBulkBtn.disabled = true;
        submitBulkBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-2" role="status"></span>Starting Scraping...`;

        progressCard.style.display = "block";
        progressTitle.textContent = "Scraping in progress...";
        progressBar.style.width = "0%";
        progressBar.className = "progress-bar progress-bar-striped progress-bar-animated bg-primary";
        progressPercentBadge.textContent = "0%";
        statusMessage.innerHTML = `<i class="bi bi-arrow-repeat spin me-1"></i> Reading uploaded student file...`;
        studentCounter.textContent = "0 / 0";

        const formData = new FormData();
        formData.append("file", fileInput.files[0]);
        formData.append("exam", examSelect.value);
        formData.append("year", yearSelect.value);
        formData.append("board", boardSelect.value);
        formData.append("demo_mode", demoModeCheck.checked ? "true" : "false");

        try {
            const response = await fetch("/api/start-extraction", {
                method: "POST",
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || "Failed to start extraction task.");
            }

            const data = await response.json();
            const jobId = data.job_id;
            const totalStudents = data.total_students;
            latestJobId = jobId;

            studentCounter.textContent = `0 / ${totalStudents}`;
            statusMessage.innerHTML = `<i class="bi bi-arrow-repeat spin me-1"></i> Connecting to board result portal...`;

            // Connect to Server-Sent Events (SSE) Stream
            connectProgressStream(jobId, totalStudents);

        } catch (error) {
            alert(`Error: ${error.message}`);
            submitBulkBtn.disabled = false;
            submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Start Bulk Extraction`;
            progressCard.style.display = "none";
        }
    });

    function connectProgressStream(jobId, total) {
        if (eventSource) {
            eventSource.close();
        }

        eventSource = new EventSource(`/api/progress-stream/${jobId}`);

        eventSource.onmessage = (event) => {
            const msg = JSON.parse(event.data);

            if (msg.type === "progress") {
                const cur = msg.current;
                const tot = msg.total;
                const pct = msg.percent;
                const res = msg.data;

                // Store in memory
                currentRecords.push(res);

                // Update Progress bar
                progressBar.style.width = `${pct}%`;
                progressPercentBadge.textContent = `${pct}%`;
                studentCounter.textContent = `${cur} / ${tot}`;
                statusMessage.innerHTML = `<i class="bi bi-gear spin me-1"></i> Processed Roll: <strong>${res.Roll}</strong> (${res.Name})`;

                // Append student row to live results table
                appendResultRow(res);

            } else if (msg.type === "completed") {
                progressBar.style.width = "100%";
                progressBar.className = "progress-bar bg-success";
                progressPercentBadge.textContent = "100%";
                progressPercentBadge.className = "badge bg-success rounded-pill px-3 py-2";
                progressTitle.textContent = "Extraction Completed!";
                statusMessage.innerHTML = `<i class="bi bi-check-circle-fill text-success me-1"></i> Successfully finished fetching ${msg.total} student records!`;

                // Set Download Links
                downloadXlsxBtn.href = msg.download_xlsx;
                downloadCsvBtn.href = msg.download_csv;
                downloadSection.style.setProperty("display", "flex", "important");
                downloadXlsxBtn.classList.add("pulse-success");

                // Capture analysis from backend if returned
                if (msg.analysis) {
                    latestAnalysisData = msg.analysis;
                }
                if (msg.download_analysis_xlsx && exportAnalysisExcelBtn) {
                    exportAnalysisExcelBtn.href = msg.download_analysis_xlsx;
                }

                // Reset Submit Button
                submitBulkBtn.disabled = false;
                submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Start Bulk Extraction`;

                eventSource.close();

            } else if (msg.type === "error") {
                progressBar.className = "progress-bar bg-danger";
                statusMessage.innerHTML = `<i class="bi bi-exclamation-triangle-fill text-danger me-1"></i> ${msg.message}`;
                alert(msg.message);

                submitBulkBtn.disabled = false;
                submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Start Bulk Extraction`;

                eventSource.close();
            }
        };

        eventSource.onerror = (err) => {
            console.error("SSE connection error:", err);
        };
    }

    // ==========================================
    // 3. TABLE LOG HELPER
    // ==========================================
    function appendResultRow(data) {
        // Remove empty placeholder row if present
        const emptyRow = document.getElementById("emptyRow");
        if (emptyRow) {
            emptyRow.remove();
        }

        extractedCount++;
        resultsCount.textContent = `${extractedCount} Recorded`;

        const row = document.createElement("tr");

        // Format GPA Badge
        let gpaBadgeClass = "badge-gpa-blue";
        const gpaNum = parseFloat(data.GPA);
        if (gpaNum >= 5.0) {
            gpaBadgeClass = "badge-gpa-gold";
        } else if (gpaNum >= 4.0) {
            gpaBadgeClass = "badge-gpa-green";
        } else if (gpaNum === 0 || isNaN(gpaNum) || (data.Result && (data.Result.includes("Fail") || data.Result.includes("Not Found")))) {
            gpaBadgeClass = "badge-gpa-red";
        }

        // Format Result Status Badge
        let resultBadge = `<span class="badge bg-success-subtle text-success border border-success-subtle">${data.Result || "PASSED"}</span>`;
        if (data.Result && (data.Result.includes("Fail") || data.Result.includes("Not Found"))) {
            resultBadge = `<span class="badge bg-danger-subtle text-danger border border-danger-subtle">${data.Result}</span>`;
        }

        // Truncate long subject grades for table display
        const gradesSummary = data.Subject_Grades_Summary || "-";
        const displayGrades = gradesSummary.length > 50 
            ? `${gradesSummary.substring(0, 48)}...` 
            : gradesSummary;

        row.innerHTML = `
            <td class="ps-4 fw-bold text-dark">${data.Roll || "-"}</td>
            <td class="text-secondary">${data.Registration || "-"}</td>
            <td class="fw-medium">${data.Name || "N/A"}</td>
            <td><span class="badge ${gpaBadgeClass} px-2 py-1 fs-7">${data.GPA || "0.00"}</span></td>
            <td>${resultBadge}</td>
            <td class="pe-4 text-muted small" title="${gradesSummary.replace(/"/g, '&quot;')}">${displayGrades}</td>
        `;

        resultsBody.prepend(row);
    }

    // ==========================================
    // 4. COMPREHENSIVE RESULT ANALYSIS LOGIC
    // ==========================================

    // Open Analysis modal trigger handlers
    function openResultAnalysis() {
        const meta = {
            exam: examSelect ? examSelect.options[examSelect.selectedIndex].text : "HSC",
            year: yearSelect ? yearSelect.value : "2026",
            board: boardSelect ? boardSelect.options[boardSelect.selectedIndex].text : "Board"
        };

        if (latestAnalysisData) {
            displayResultAnalysis(latestAnalysisData, meta);
            return;
        }

        if (currentRecords.length > 0) {
            const analysis = computeClientAnalysis(currentRecords);
            latestAnalysisData = analysis;
            displayResultAnalysis(analysis, meta);
            return;
        }

        // If no records are present, offer user to analyze an existing file
        const confirmUpload = confirm(
            "No extracted results are currently loaded on this page.\n\n" +
            "Would you like to select an existing Board Results CSV or Excel file to analyze right now?"
        );
        if (confirmUpload && analysisFileInput) {
            analysisFileInput.click();
        }
    }

    if (openAnalysisBtn) openAnalysisBtn.addEventListener("click", openResultAnalysis);
    if (headerAnalysisBtn) headerAnalysisBtn.addEventListener("click", openResultAnalysis);
    if (bottomAnalysisBtn) bottomAnalysisBtn.addEventListener("click", openResultAnalysis);

    if (uploadFileForAnalysisBtn) {
        uploadFileForAnalysisBtn.addEventListener("click", () => {
            if (analysisFileInput) analysisFileInput.click();
        });
    }

    // 4th Subject Mode Selector State and Listeners
    let currentFourthSubjectMode = "with_4th";
    const btnModeWith4th = document.getElementById("btnModeWith4th");
    const btnModeWithout4th = document.getElementById("btnModeWithout4th");

    if (btnModeWith4th) {
        btnModeWith4th.addEventListener("click", () => {
            currentFourthSubjectMode = "with_4th";
            updateGpaDisplayMode(latestAnalysisData);
        });
    }
    if (btnModeWithout4th) {
        btnModeWithout4th.addEventListener("click", () => {
            currentFourthSubjectMode = "without_4th";
            updateGpaDisplayMode(latestAnalysisData);
        });
    }

    // Upload an existing result file to analyze directly
    if (analysisFileInput) {
        analysisFileInput.addEventListener("change", async () => {
            if (!analysisFileInput.files || analysisFileInput.files.length === 0) return;
            const file = analysisFileInput.files[0];

            const originalBtnHtml = bottomAnalysisBtn.innerHTML;
            bottomAnalysisBtn.disabled = true;
            bottomAnalysisBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-2"></span>Analyzing file...`;

            const formData = new FormData();
            formData.append("file", file);

            try {
                const response = await fetch("/api/analyze-file", {
                    method: "POST",
                    body: formData
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.detail || "Failed to analyze result file.");
                }

                const resJson = await response.json();
                latestAnalysisData = resJson.analysis;

                // Setup Excel Export via direct download stream
                if (exportAnalysisExcelBtn) {
                    exportAnalysisExcelBtn.onclick = async (e) => {
                        e.preventDefault();
                        const exportForm = new FormData();
                        exportForm.append("file", file);
                        const expResp = await fetch("/api/export-analysis-file", {
                            method: "POST",
                            body: exportForm
                        });
                        const blob = await expResp.blob();
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement("a");
                        a.href = url;
                        a.download = `Board_Result_Analysis_${file.name.replace(/\.[^/.]+$/, "")}.xlsx`;
                        document.body.appendChild(a);
                        a.click();
                        a.remove();
                    };
                }

                displayResultAnalysis(latestAnalysisData, {
                    exam: "Analyzed File",
                    year: "",
                    board: file.name
                });

            } catch (err) {
                alert(`Analysis Error: ${err.message}`);
            } finally {
                bottomAnalysisBtn.disabled = false;
                bottomAnalysisBtn.innerHTML = originalBtnHtml;
                analysisFileInput.value = "";
            }
        });
    }

    // Display Analysis Modal with All Aggregations
    function displayResultAnalysis(analysisData, metaInfo) {
        if (!analysisData) return;

        const summary = analysisData.summary || {};

        // 1. Overall KPI Cards
        const kpiAppeared = document.getElementById("kpiAppeared");
        const kpiTotalStudents = document.getElementById("kpiTotalStudents");
        const kpiAbsent = document.getElementById("kpiAbsent");
        const kpiAbsentRate = document.getElementById("kpiAbsentRate");
        const kpiPassed = document.getElementById("kpiPassed");
        const kpiPassRate = document.getElementById("kpiPassRate");
        const kpiFailed = document.getElementById("kpiFailed");
        const kpiFailRate = document.getElementById("kpiFailRate");
        const kpiGpa5 = document.getElementById("kpiGpa5");
        const kpiGpa5Rate = document.getElementById("kpiGpa5Rate");
        const kpi1Fail = document.getElementById("kpi1Fail");
        const kpi2Fail = document.getElementById("kpi2Fail");
        const analysisMetaText = document.getElementById("analysisMetaText");

        if (kpiAppeared) kpiAppeared.textContent = summary.appeared || 0;
        if (kpiTotalStudents) kpiTotalStudents.textContent = `${summary.total_students || 0} Processed`;
        if (kpiAbsent) kpiAbsent.textContent = summary.absent || 0;
        if (kpiAbsentRate) kpiAbsentRate.textContent = `${summary.absent_rate || 0}% Absent`;
        if (kpiPassed) kpiPassed.textContent = summary.passed || 0;
        if (kpiPassRate) kpiPassRate.textContent = `${summary.pass_rate || 0}% Pass`;
        if (kpiFailed) kpiFailed.textContent = summary.failed || 0;
        if (kpiFailRate) kpiFailRate.textContent = `${summary.fail_rate || 0}% Fail`;
        if (kpi1Fail) kpi1Fail.textContent = summary.one_subject_fail_count || 0;
        if (kpi2Fail) kpi2Fail.textContent = summary.two_subject_fail_count || 0;

        // 4th Subject Bar Real-time Stats Badges
        const kpiGpa5ValWith = document.getElementById("kpiGpa5ValWith");
        const kpiGpa5RateWith = document.getElementById("kpiGpa5RateWith");
        const kpiGpa5ValWithout = document.getElementById("kpiGpa5ValWithout");
        const kpiGpa5RateWithout = document.getElementById("kpiGpa5RateWithout");

        const countWith = (summary.gpa_5_count_with_4th !== undefined) ? summary.gpa_5_count_with_4th : (summary.gpa_5_count || 0);
        const rateWith = (summary.gpa_5_rate_with_4th !== undefined) ? summary.gpa_5_rate_with_4th : (summary.gpa_5_rate || 0);
        const countWithout = (summary.gpa_5_count_without_4th !== undefined) ? summary.gpa_5_count_without_4th : 0;
        const rateWithout = (summary.gpa_5_rate_without_4th !== undefined) ? summary.gpa_5_rate_without_4th : 0;

        if (kpiGpa5ValWith) kpiGpa5ValWith.textContent = countWith;
        if (kpiGpa5RateWith) kpiGpa5RateWith.textContent = `${rateWith}%`;
        if (kpiGpa5ValWithout) kpiGpa5ValWithout.textContent = countWithout;
        if (kpiGpa5RateWithout) kpiGpa5RateWithout.textContent = `${rateWithout}%`;

        if (analysisMetaText) {
            analysisMetaText.textContent = `${metaInfo.exam || "Examination"} ${metaInfo.year || ""} | ${metaInfo.board || "Board"} | ${summary.total_students || 0} Records Analyzed`;
        }

        // 2. Regular vs Irregular KPI Quick Banner
        const reg = analysisData.regular_summary || {};
        const irreg = analysisData.irregular_summary || {};

        const kpiRegularBadge = document.getElementById("kpiRegularBadge");
        const kpiRegularRate = document.getElementById("kpiRegularRate");
        const kpiRegularPassed = document.getElementById("kpiRegularPassed");
        const kpiRegularFailed = document.getElementById("kpiRegularFailed");
        const kpiRegularAbsent = document.getElementById("kpiRegularAbsent");
        const kpiRegularGpa5 = document.getElementById("kpiRegularGpa5");
        const kpiRegular1Fail = document.getElementById("kpiRegular1Fail");

        if (kpiRegularBadge) kpiRegularBadge.textContent = `${reg.total || 0} Students`;
        if (kpiRegularRate) kpiRegularRate.textContent = `${reg.pass_rate || 0}% Pass`;
        if (kpiRegularPassed) kpiRegularPassed.textContent = reg.passed || 0;
        if (kpiRegularFailed) kpiRegularFailed.textContent = reg.failed || 0;
        if (kpiRegularAbsent) kpiRegularAbsent.textContent = reg.absent || 0;
        if (kpiRegularGpa5) kpiRegularGpa5.textContent = reg.gpa5_count || 0;
        if (kpiRegular1Fail) kpiRegular1Fail.textContent = reg.one_fail_count || 0;

        const kpiIrregularBadge = document.getElementById("kpiIrregularBadge");
        const kpiIrregularRate = document.getElementById("kpiIrregularRate");
        const kpiIrregularPassed = document.getElementById("kpiIrregularPassed");
        const kpiIrregularFailed = document.getElementById("kpiIrregularFailed");
        const kpiIrregularAbsent = document.getElementById("kpiIrregularAbsent");
        const kpiIrregularGpa5 = document.getElementById("kpiIrregularGpa5");
        const kpiIrregular1Fail = document.getElementById("kpiIrregular1Fail");

        if (kpiIrregularBadge) kpiIrregularBadge.textContent = `${irreg.total || 0} Students`;
        if (kpiIrregularRate) kpiIrregularRate.textContent = `${irreg.pass_rate || 0}% Pass`;
        if (kpiIrregularPassed) kpiIrregularPassed.textContent = irreg.passed || 0;
        if (kpiIrregularFailed) kpiIrregularFailed.textContent = irreg.failed || 0;
        if (kpiIrregularAbsent) kpiIrregularAbsent.textContent = irreg.absent || 0;
        if (kpiIrregularGpa5) kpiIrregularGpa5.textContent = irreg.gpa5_count || 0;
        if (kpiIrregular1Fail) kpiIrregular1Fail.textContent = irreg.one_fail_count || 0;

        // Tab badge counters
        const badge1FailCount = document.getElementById("badge1FailCount");
        const badge2FailCount = document.getElementById("badge2FailCount");
        const badgeMFailCount = document.getElementById("badgeMFailCount");
        const badgeAbsentCount = document.getElementById("badgeAbsentCount");

        if (badge1FailCount) badge1FailCount.textContent = (analysisData.one_subject_fail && analysisData.one_subject_fail.total) || 0;
        if (badge2FailCount) badge2FailCount.textContent = (analysisData.two_subject_fail && analysisData.two_subject_fail.total) || 0;
        if (badgeMFailCount) badgeMFailCount.textContent = (analysisData.multi_subject_fail && analysisData.multi_subject_fail.total) || 0;
        const absentTotal = (analysisData.absent_analysis && analysisData.absent_analysis.total_absent !== undefined) ? analysisData.absent_analysis.total_absent : (summary.absent || 0);
        if (badgeAbsentCount) badgeAbsentCount.textContent = absentTotal;

        // Render Tabs
        renderSubjectGradeTable(analysisData.subject_analysis || [], analysisData.available_groups || []);
        renderRegularVsIrregular(analysisData.regular_summary || {}, analysisData.irregular_summary || {}, analysisData.types_summary || {}, analysisData.group_type_comparison || []);
        renderOneSubjectFail(analysisData.one_subject_fail || {});
        renderTwoSubjectFail(analysisData.two_subject_fail || {});
        renderMultiSubjectFail(analysisData.multi_subject_fail || {});
        renderGroupAndGpaSummary(analysisData.groups_summary || {}, analysisData.gpa_distribution || {});
        renderAbsentStudents(analysisData.absent_analysis || {});

        // Synchronize 4th Subject display mode
        updateGpaDisplayMode(analysisData);

        // Setup Print Button
        if (printAnalysisBtn) {
            printAnalysisBtn.onclick = () => printAnalysisReport(analysisData, metaInfo);
        }

        // Show Modal
        if (analysisModalInstance) {
            analysisModalInstance.show();
        } else if (analysisModalEl) {
            const bsModal = new bootstrap.Modal(analysisModalEl);
            bsModal.show();
        }
    }

    // TAB 1: Subject Wise Grade Table
    function renderSubjectGradeTable(subjectAnalysis, availableGroups) {
        const groupFilter = document.getElementById("analysisGroupFilter");
        const typeFilter = document.getElementById("analysisTypeFilter");
        const searchInput = document.getElementById("analysisSubjectSearch");
        const tableBody = document.getElementById("subjectGradeBody");

        if (groupFilter) {
            const currentVal = groupFilter.value;
            groupFilter.innerHTML = '<option value="ALL">All Groups Combined</option>';
            (availableGroups || []).forEach(grp => {
                const opt = document.createElement("option");
                opt.value = grp;
                opt.textContent = grp;
                if (grp === currentVal) opt.selected = true;
                groupFilter.appendChild(opt);
            });
        }

        function filterAndRender() {
            const selectedGroup = groupFilter ? groupFilter.value : "ALL";
            const selectedType = typeFilter ? typeFilter.value : "ALL";
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            tableBody.innerHTML = "";

            const filtered = subjectAnalysis.filter(item => {
                const matchGroup = (selectedGroup === "ALL") ? (item.group === "ALL") : (item.group === selectedGroup);
                if (!matchGroup) return false;
                const matchType = (selectedType === "ALL") ? (item.type === "ALL") : (item.type === selectedType);
                if (!matchType) return false;
                if (query && !item.subject.toLowerCase().includes(query)) return false;
                return true;
            });

            if (filtered.length === 0) {
                tableBody.innerHTML = `<tr><td colspan="14" class="text-center py-4 text-muted">No subject records found matching the filter.</td></tr>`;
                return;
            }

            filtered.forEach(item => {
                const tr = document.createElement("tr");
                const g = item.grades || {};
                const passPct = item.pass_rate || 0;
                const progressColor = passPct >= 80 ? "bg-success" : (passPct >= 50 ? "bg-primary" : "bg-danger");

                let typeBadge = `<span class="badge bg-secondary-subtle text-secondary border">${item.type}</span>`;
                if (item.type === "REGULAR") {
                    typeBadge = `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`;
                } else if (item.type === "IRREGULAR") {
                    typeBadge = `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`;
                } else if (item.type === "ALL") {
                    typeBadge = `<span class="badge bg-light text-muted border">ALL</span>`;
                }

                tr.innerHTML = `
                    <td class="text-start ps-3 fw-bold text-dark">${item.subject}</td>
                    <td><span class="badge bg-light text-secondary border">${item.group}</span></td>
                    <td>${typeBadge}</td>
                    <td class="fw-semibold">${item.appeared}</td>
                    <td class="text-success fw-bold">${item.passed}</td>
                    <td class="${item.failed > 0 ? 'text-danger fw-bold' : 'text-muted'}">${item.failed}</td>
                    <td>
                        <div class="fw-bold small mb-1">${passPct}%</div>
                        <div class="progress" style="height: 5px;">
                            <div class="progress-bar ${progressColor}" style="width: ${passPct}%"></div>
                        </div>
                    </td>
                    <td class="badge-gpa-gold">${g["A+"] || 0}</td>
                    <td class="badge-gpa-green">${g["A"] || 0}</td>
                    <td class="badge-gpa-green">${g["A-"] || 0}</td>
                    <td class="badge-gpa-blue">${g["B"] || 0}</td>
                    <td class="badge-gpa-blue">${g["C"] || 0}</td>
                    <td>${g["D"] || 0}</td>
                    <td class="${g['F'] > 0 ? 'badge-gpa-red fw-bold text-danger' : 'text-muted'}">${g["F"] || 0}</td>
                `;
                tableBody.appendChild(tr);
            });
        }

        if (groupFilter) groupFilter.onchange = filterAndRender;
        if (typeFilter) typeFilter.onchange = filterAndRender;
        if (searchInput) searchInput.oninput = filterAndRender;
        filterAndRender();
    }

    // TAB: Regular vs Irregular Performance
    function renderRegularVsIrregular(regularSummary, irregularSummary, typesSummary, groupTypeComparison) {
        const cardRegCount = document.getElementById("cardRegCount");
        const cardRegAppeared = document.getElementById("cardRegAppeared");
        const cardRegPassed = document.getElementById("cardRegPassed");
        const cardRegPassRate = document.getElementById("cardRegPassRate");
        const cardRegFailed = document.getElementById("cardRegFailed");
        const cardRegGpa5 = document.getElementById("cardRegGpa5");
        const cardReg1Fail = document.getElementById("cardReg1Fail");
        const cardReg2Fail = document.getElementById("cardReg2Fail");

        if (cardRegCount) cardRegCount.textContent = `${regularSummary.total || 0} Students`;
        if (cardRegAppeared) cardRegAppeared.textContent = regularSummary.appeared || 0;
        if (cardRegPassed) cardRegPassed.textContent = regularSummary.passed || 0;
        if (cardRegPassRate) cardRegPassRate.textContent = `${regularSummary.pass_rate || 0}%`;
        if (cardRegFailed) cardRegFailed.textContent = regularSummary.failed || 0;
        if (cardRegGpa5) cardRegGpa5.textContent = regularSummary.gpa5_count || 0;
        if (cardReg1Fail) cardReg1Fail.textContent = regularSummary.one_fail_count || 0;
        if (cardReg2Fail) cardReg2Fail.textContent = regularSummary.two_fail_count || 0;

        const cardIrregCount = document.getElementById("cardIrregCount");
        const cardIrregAppeared = document.getElementById("cardIrregAppeared");
        const cardIrregPassed = document.getElementById("cardIrregPassed");
        const cardIrregPassRate = document.getElementById("cardIrregPassRate");
        const cardIrregFailed = document.getElementById("cardIrregFailed");
        const cardIrregGpa5 = document.getElementById("cardIrregGpa5");
        const cardIrreg1Fail = document.getElementById("cardIrreg1Fail");
        const cardIrreg2Fail = document.getElementById("cardIrreg2Fail");

        if (cardIrregCount) cardIrregCount.textContent = `${irregularSummary.total || 0} Students`;
        if (cardIrregAppeared) cardIrregAppeared.textContent = irregularSummary.appeared || 0;
        if (cardIrregPassed) cardIrregPassed.textContent = irregularSummary.passed || 0;
        if (cardIrregPassRate) cardIrregPassRate.textContent = `${irregularSummary.pass_rate || 0}%`;
        if (cardIrregFailed) cardIrregFailed.textContent = irregularSummary.failed || 0;
        if (cardIrregGpa5) cardIrregGpa5.textContent = irregularSummary.gpa5_count || 0;
        if (cardIrreg1Fail) cardIrreg1Fail.textContent = irregularSummary.one_fail_count || 0;
        if (cardIrreg2Fail) cardIrreg2Fail.textContent = irregularSummary.two_fail_count || 0;

        // Render Overall Type Table
        const typeTableBody = document.getElementById("typeSummaryTableBody");
        if (typeTableBody) {
            typeTableBody.innerHTML = "";
            const typeList = Object.values(typesSummary || {});
            if (typeList.length === 0) {
                typeTableBody.innerHTML = `<tr><td colspan="10" class="text-center py-3 text-muted">No candidate type data recorded.</td></tr>`;
            } else {
                typeList.forEach(t => {
                    const badgeClass = t.type === "REGULAR" ? "bg-primary" : (t.type === "IRREGULAR" ? "bg-warning text-dark" : "bg-secondary");
                    const tr = document.createElement("tr");
                    tr.innerHTML = `
                        <td class="text-start ps-3"><span class="badge ${badgeClass} px-2.5 py-1 fw-bold">${t.type}</span></td>
                        <td class="fw-semibold">${t.total}</td>
                        <td>${t.appeared}</td>
                        <td class="text-success fw-bold">${t.passed}</td>
                        <td class="text-danger fw-bold">${t.failed}</td>
                        <td><span class="badge bg-success-subtle text-success border border-success-subtle px-2 py-1">${t.pass_rate}%</span></td>
                        <td class="badge-gpa-gold">${t.gpa5_count}</td>
                        <td><span class="badge bg-primary-subtle text-primary border border-primary-subtle px-2 py-1">${t.one_fail_count}</span></td>
                        <td><span class="badge bg-danger-subtle text-danger border border-danger-subtle px-2 py-1">${t.two_fail_count}</span></td>
                        <td><span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle px-2 py-1">${t.multi_fail_count}</span></td>
                    `;
                    typeTableBody.appendChild(tr);
                });
            }
        }

        // Render Group x Candidate Type Matrix Table
        const groupTypeTableBody = document.getElementById("groupTypeComparisonTableBody");
        if (groupTypeTableBody) {
            groupTypeTableBody.innerHTML = "";
            const matrix = groupTypeComparison || [];
            if (matrix.length === 0) {
                groupTypeTableBody.innerHTML = `<tr><td colspan="11" class="text-center py-3 text-muted">No group x type data recorded.</td></tr>`;
            } else {
                matrix.forEach(row => {
                    const typeBadge = row.type === "REGULAR" 
                        ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle px-2 py-1">REGULAR</span>`
                        : (row.type === "IRREGULAR" 
                            ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle px-2 py-1">IRREGULAR</span>`
                            : `<span class="badge bg-light text-secondary border px-2 py-1">${row.type}</span>`);

                    const tr = document.createElement("tr");
                    tr.innerHTML = `
                        <td class="text-start ps-3 fw-bold text-dark">${row.group}</td>
                        <td>${typeBadge}</td>
                        <td class="fw-semibold">${row.total}</td>
                        <td>${row.appeared}</td>
                        <td class="text-success fw-bold">${row.passed}</td>
                        <td class="text-danger fw-bold">${row.failed}</td>
                        <td><span class="badge bg-success-subtle text-success border border-success-subtle px-2 py-1">${row.pass_rate}%</span></td>
                        <td class="badge-gpa-gold">${row.gpa5_count}</td>
                        <td><span class="badge bg-primary-subtle text-primary border border-primary-subtle px-2 py-1">${row.one_fail_count}</span></td>
                        <td><span class="badge bg-danger-subtle text-danger border border-danger-subtle px-2 py-1">${row.two_fail_count}</span></td>
                        <td><span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle px-2 py-1">${row.multi_fail_count}</span></td>
                    `;
                    groupTypeTableBody.appendChild(tr);
                });
            }
        }
    }

    // TAB 2: One Subject Fail with Subject
    function renderOneSubjectFail(oneFailData) {
        const pillsContainer = document.getElementById("oneSubPillsContainer");
        const tableBody = document.getElementById("oneSubTableBody");
        const summaryBadge = document.getElementById("oneSubTotalSummary");
        const noticeEl = document.getElementById("oneSubFilteredNotice");
        const searchInput = document.getElementById("oneSubSearch");
        const typeButtonsContainer = document.getElementById("oneSubTypeButtons");

        const total = oneFailData.total || 0;
        if (summaryBadge) summaryBadge.textContent = `${total} Total Students`;

        let activeSubject = "ALL";
        let activeType = "ALL";
        const students = oneFailData.students || [];

        // Build interactive filter pills
        if (pillsContainer) {
            pillsContainer.innerHTML = "";
            const allBtn = document.createElement("button");
            allBtn.type = "button";
            allBtn.className = "btn btn-sm btn-dark filter-pill active";
            allBtn.innerHTML = `All Subjects (${total})`;
            allBtn.onclick = () => {
                activeSubject = "ALL";
                updatePills();
                filterAndRender();
            };
            pillsContainer.appendChild(allBtn);

            Object.entries(oneFailData.subjects_summary || {}).forEach(([subj, count]) => {
                const btn = document.createElement("button");
                btn.type = "button";
                btn.className = "btn btn-sm btn-outline-danger filter-pill";
                btn.innerHTML = `${subj} <span class="badge bg-danger ms-1">${count}</span>`;
                btn.onclick = () => {
                    activeSubject = subj;
                    updatePills();
                    filterAndRender();
                };
                pillsContainer.appendChild(btn);
            });
        }

        function updatePills() {
            if (!pillsContainer) return;
            pillsContainer.querySelectorAll(".filter-pill").forEach(btn => {
                if (activeSubject === "ALL" && btn.textContent.startsWith("All Subjects")) {
                    btn.className = "btn btn-sm btn-dark filter-pill active";
                } else if (btn.textContent.startsWith(activeSubject)) {
                    btn.className = "btn btn-sm btn-danger filter-pill active";
                } else {
                    btn.className = btn.textContent.startsWith("All Subjects") 
                        ? "btn btn-sm btn-outline-dark filter-pill" 
                        : "btn btn-sm btn-outline-danger filter-pill";
                }
            });
        }

        // Setup candidate type toggle buttons
        if (typeButtonsContainer) {
            typeButtonsContainer.querySelectorAll("button").forEach(btn => {
                btn.onclick = () => {
                    activeType = btn.getAttribute("data-type") || "ALL";
                    typeButtonsContainer.querySelectorAll("button").forEach(b => {
                        b.classList.remove("active", "btn-dark", "btn-primary", "btn-warning");
                        const t = b.getAttribute("data-type");
                        if (t === "ALL") b.className = "btn btn-outline-dark";
                        else if (t === "REGULAR") b.className = "btn btn-outline-primary";
                        else if (t === "IRREGULAR") b.className = "btn btn-outline-warning text-dark";
                    });

                    btn.classList.add("active");
                    if (activeType === "ALL") btn.className = "btn btn-dark active";
                    else if (activeType === "REGULAR") btn.className = "btn btn-primary active";
                    else if (activeType === "IRREGULAR") btn.className = "btn btn-warning text-dark active";

                    filterAndRender();
                };
            });
        }

        function filterAndRender() {
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            const filtered = students.filter(st => {
                if (activeSubject !== "ALL" && st.Failed_Subject !== activeSubject) return false;
                if (activeType !== "ALL" && st.Type !== activeType) return false;
                if (query) {
                    const text = `${st.Roll} ${st.Registration || ""} ${st.Name} ${st.Group} ${st.Type || ""} ${st.Failed_Subject}`.toLowerCase();
                    if (!text.includes(query)) return false;
                }
                return true;
            });

            if (noticeEl) {
                const subText = activeSubject === "ALL" ? "all subjects" : activeSubject;
                const typeText = activeType === "ALL" ? "all types" : activeType;
                noticeEl.textContent = `Showing ${filtered.length} students (${subText}, ${typeText})`;
            }

            tableBody.innerHTML = "";
            if (filtered.length === 0) {
                tableBody.innerHTML = `<tr><td colspan="8" class="text-center py-4 text-muted">No students found matching criteria.</td></tr>`;
                return;
            }

            filtered.forEach(st => {
                const typeBadge = st.Type === "REGULAR" 
                    ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`
                    : (st.Type === "IRREGULAR" 
                        ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`
                        : `<span class="badge bg-light text-secondary border">${st.Type || "OTHER"}</span>`);

                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td class="ps-3 fw-bold text-dark">${st.Roll}</td>
                    <td class="text-secondary small">${st.Registration || "-"}</td>
                    <td class="fw-semibold">${st.Name}</td>
                    <td><span class="badge bg-light text-dark border">${st.Group}</span></td>
                    <td>${typeBadge}</td>
                    <td><span class="badge bg-danger px-2 py-1 fs-7"><i class="bi bi-x-circle me-1"></i>${st.Failed_Subject} (F)</span></td>
                    <td class="small text-muted" title="${st.Passed_Grades_Summary}">${st.Passed_Grades_Summary}</td>
                    <td class="text-center"><span class="badge bg-success-subtle text-success border border-success-subtle">Eligible for Improvement</span></td>
                `;
                tableBody.appendChild(tr);
            });
        }

        if (searchInput) searchInput.oninput = filterAndRender;
        filterAndRender();
    }

    // TAB 3: Two Subject Fail with Subject
    function renderTwoSubjectFail(twoFailData) {
        const pillsContainer = document.getElementById("twoSubPillsContainer");
        const tableBody = document.getElementById("twoSubTableBody");
        const summaryBadge = document.getElementById("twoSubTotalSummary");
        const noticeEl = document.getElementById("twoSubFilteredNotice");
        const searchInput = document.getElementById("twoSubSearch");
        const typeButtonsContainer = document.getElementById("twoSubTypeButtons");

        const total = twoFailData.total || 0;
        if (summaryBadge) summaryBadge.textContent = `${total} Total Students`;

        let activeSubject = "ALL";
        let activeType = "ALL";
        const students = twoFailData.students || [];

        // Build interactive pills
        if (pillsContainer) {
            pillsContainer.innerHTML = "";
            const allBtn = document.createElement("button");
            allBtn.type = "button";
            allBtn.className = "btn btn-sm btn-dark filter-pill active";
            allBtn.innerHTML = `All Two-Subject Fails (${total})`;
            allBtn.onclick = () => {
                activeSubject = "ALL";
                updatePills();
                filterAndRender();
            };
            pillsContainer.appendChild(allBtn);

            Object.entries(twoFailData.subjects_summary || {}).forEach(([subj, count]) => {
                const btn = document.createElement("button");
                btn.type = "button";
                btn.className = "btn btn-sm btn-outline-danger filter-pill";
                btn.innerHTML = `${subj} <span class="badge bg-danger ms-1">${count}</span>`;
                btn.onclick = () => {
                    activeSubject = subj;
                    updatePills();
                    filterAndRender();
                };
                pillsContainer.appendChild(btn);
            });
        }

        function updatePills() {
            if (!pillsContainer) return;
            pillsContainer.querySelectorAll(".filter-pill").forEach(btn => {
                if (activeSubject === "ALL" && btn.textContent.startsWith("All Two-Subject")) {
                    btn.className = "btn btn-sm btn-dark filter-pill active";
                } else if (btn.textContent.startsWith(activeSubject)) {
                    btn.className = "btn btn-sm btn-danger filter-pill active";
                } else {
                    btn.className = btn.textContent.startsWith("All Two-Subject") 
                        ? "btn btn-sm btn-outline-dark filter-pill" 
                        : "btn btn-sm btn-outline-danger filter-pill";
                }
            });
        }

        // Setup candidate type toggle buttons
        if (typeButtonsContainer) {
            typeButtonsContainer.querySelectorAll("button").forEach(btn => {
                btn.onclick = () => {
                    activeType = btn.getAttribute("data-type") || "ALL";
                    typeButtonsContainer.querySelectorAll("button").forEach(b => {
                        b.classList.remove("active", "btn-dark", "btn-primary", "btn-warning");
                        const t = b.getAttribute("data-type");
                        if (t === "ALL") b.className = "btn btn-outline-dark";
                        else if (t === "REGULAR") b.className = "btn btn-outline-primary";
                        else if (t === "IRREGULAR") b.className = "btn btn-outline-warning text-dark";
                    });

                    btn.classList.add("active");
                    if (activeType === "ALL") btn.className = "btn btn-dark active";
                    else if (activeType === "REGULAR") btn.className = "btn btn-primary active";
                    else if (activeType === "IRREGULAR") btn.className = "btn btn-warning text-dark active";

                    filterAndRender();
                };
            });
        }

        function filterAndRender() {
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            const filtered = students.filter(st => {
                const fails = st.Failed_Subjects || [];
                if (activeSubject !== "ALL" && !fails.includes(activeSubject)) return false;
                if (activeType !== "ALL" && st.Type !== activeType) return false;
                if (query) {
                    const text = `${st.Roll} ${st.Registration || ""} ${st.Name} ${st.Group} ${st.Type || ""} ${fails.join(" ")}`.toLowerCase();
                    if (!text.includes(query)) return false;
                }
                return true;
            });

            if (noticeEl) {
                const subText = activeSubject === "ALL" ? "all subjects" : activeSubject;
                const typeText = activeType === "ALL" ? "all types" : activeType;
                noticeEl.textContent = `Showing ${filtered.length} students (${subText}, ${typeText})`;
            }

            tableBody.innerHTML = "";
            if (filtered.length === 0) {
                tableBody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-muted">No students found matching criteria.</td></tr>`;
                return;
            }

            filtered.forEach(st => {
                const fails = st.Failed_Subjects || [];
                const f1 = fails[0] || "-";
                const f2 = fails[1] || "-";

                const typeBadge = st.Type === "REGULAR" 
                    ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`
                    : (st.Type === "IRREGULAR" 
                        ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`
                        : `<span class="badge bg-light text-secondary border">${st.Type || "OTHER"}</span>`);

                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td class="ps-3 fw-bold text-dark">${st.Roll}</td>
                    <td class="text-secondary small">${st.Registration || "-"}</td>
                    <td class="fw-semibold">${st.Name}</td>
                    <td><span class="badge bg-light text-dark border">${st.Group}</span></td>
                    <td>${typeBadge}</td>
                    <td>
                        <span class="badge bg-danger px-2 py-1 fs-7 me-1"><i class="bi bi-x-circle me-1"></i>${f1}</span>
                        <span class="badge bg-danger px-2 py-1 fs-7"><i class="bi bi-x-circle me-1"></i>${f2}</span>
                    </td>
                    <td class="small text-muted" title="${st.Passed_Grades_Summary}">${st.Passed_Grades_Summary}</td>
                `;
                tableBody.appendChild(tr);
            });
        }

        if (searchInput) searchInput.oninput = filterAndRender;
        filterAndRender();
    }

    // TAB 4: 3+ Subject Fail
    function renderMultiSubjectFail(multiFailData) {
        const tableBody = document.getElementById("multiSubTableBody");
        const summaryBadge = document.getElementById("multiSubTotalSummary");
        const students = multiFailData.students || [];

        if (summaryBadge) summaryBadge.textContent = `${students.length} Students`;
        if (!tableBody) return;
        tableBody.innerHTML = "";

        if (students.length === 0) {
            tableBody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-muted">No students failed in 3 or more subjects.</td></tr>`;
            return;
        }

        students.forEach(st => {
            const tr = document.createElement("tr");
            const fails = (st.Failed_Subjects || []).join(", ");
            const typeBadge = st.Type === "REGULAR" 
                ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`
                : (st.Type === "IRREGULAR" 
                    ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`
                    : `<span class="badge bg-light text-secondary border">${st.Type || "OTHER"}</span>`);

            tr.innerHTML = `
                <td class="ps-3 fw-bold text-dark">${st.Roll}</td>
                <td class="fw-semibold">${st.Name}</td>
                <td><span class="badge bg-light text-dark border">${st.Group}</span></td>
                <td>${typeBadge}</td>
                <td class="text-center"><span class="badge bg-danger px-2 py-1">${st.Failed_Count} Subjects</span></td>
                <td class="text-danger small fw-semibold">${fails}</td>
                <td class="small text-muted">${st.Passed_Grades_Summary}</td>
            `;
            tableBody.appendChild(tr);
        });
    }

    // TAB 5: Group Comparison & GPA Breakdown
    function renderGroupAndGpaSummary(groupsSummary, gpaDistribution) {
        const groupTableBody = document.getElementById("groupSummaryTableBody");

        if (groupTableBody) {
            groupTableBody.innerHTML = "";
            Object.values(groupsSummary || {}).forEach(grp => {
                const tr = document.createElement("tr");
                const gpa5With = (grp.gpa5_count_with_4th !== undefined) ? grp.gpa5_count_with_4th : (grp.gpa5_count || 0);
                const gpa5Without = (grp.gpa5_count_without_4th !== undefined) ? grp.gpa5_count_without_4th : 0;
                tr.innerHTML = `
                    <td class="text-start ps-3 fw-bold text-dark">${grp.group}</td>
                    <td class="fw-semibold">${grp.total}</td>
                    <td>${grp.appeared}</td>
                    <td class="text-secondary fw-bold">${grp.absent || 0}</td>
                    <td class="text-success fw-bold">${grp.passed}</td>
                    <td class="text-danger fw-bold">${grp.failed}</td>
                    <td><span class="badge bg-success-subtle text-success border border-success-subtle px-2 py-1">${grp.pass_rate}%</span></td>
                    <td class="badge-gpa-gold fw-bold">${gpa5With}</td>
                    <td class="text-primary fw-bold">${gpa5Without}</td>
                    <td><span class="badge bg-primary-subtle text-primary border border-primary-subtle px-2 py-1">${grp.one_fail_count}</span></td>
                    <td><span class="badge bg-danger-subtle text-danger border border-danger-subtle px-2 py-1">${grp.two_fail_count}</span></td>
                    <td><span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle px-2 py-1">${grp.multi_fail_count}</span></td>
                `;
                groupTableBody.appendChild(tr);
            });
        }

        renderGpaDistributionTiers(gpaDistribution);
    }

    // Render GPA Distribution Tier Cards
    function renderGpaDistributionTiers(gpaDistribution) {
        const gpaRow = document.getElementById("gpaDistributionRow");
        if (!gpaRow) return;
        gpaRow.innerHTML = "";
        Object.entries(gpaDistribution || {}).forEach(([tier, count]) => {
            const col = document.createElement("div");
            col.className = "col-6 col-md-3 col-lg-auto flex-fill";
            let colorClass = "text-primary";
            if (tier.includes("5.00")) colorClass = "text-warning";
            else if (tier.includes("0.00") || tier.includes("Fail")) colorClass = "text-danger";
            else if (tier.includes("4.00") || tier.includes("3.50")) colorClass = "text-success";

            col.innerHTML = `
                <div class="card border rounded-3 p-3 bg-light text-center h-100">
                    <span class="text-muted fs-8 fw-semibold">${tier}</span>
                    <h4 class="fw-bold ${colorClass} mt-1 mb-0">${count}</h4>
                </div>
            `;
            gpaRow.appendChild(col);
        });
    }

    // Toggle 4th Subject display mode on KPI and GPA cards
    function updateGpaDisplayMode(analysisData) {
        const data = analysisData || latestAnalysisData;
        if (!data) return;

        const summary = data.summary || {};
        const kpiGpa5 = document.getElementById("kpiGpa5");
        const kpiGpa5Rate = document.getElementById("kpiGpa5Rate");
        const kpiGpa5Title = document.getElementById("kpiGpa5Title");
        const gpaDistModeNotice = document.getElementById("gpaDistModeNotice");
        const btnWith = document.getElementById("btnModeWith4th");
        const btnWithout = document.getElementById("btnModeWithout4th");

        if (currentFourthSubjectMode === "with_4th") {
            if (btnWith) {
                btnWith.classList.add("active", "btn-primary");
                btnWith.classList.remove("btn-outline-primary");
            }
            if (btnWithout) {
                btnWithout.classList.remove("active", "btn-primary");
                btnWithout.classList.add("btn-outline-primary");
            }
            if (kpiGpa5Title) kpiGpa5Title.textContent = "GPA 5.0 (With 4th Sub)";
            const countWith = (summary.gpa_5_count_with_4th !== undefined) ? summary.gpa_5_count_with_4th : (summary.gpa_5_count || 0);
            const rateWith = (summary.gpa_5_rate_with_4th !== undefined) ? summary.gpa_5_rate_with_4th : (summary.gpa_5_rate || 0);
            if (kpiGpa5) kpiGpa5.textContent = countWith;
            if (kpiGpa5Rate) kpiGpa5Rate.textContent = `${rateWith}% of Appeared`;
            if (gpaDistModeNotice) gpaDistModeNotice.innerHTML = `<i class="bi bi-info-circle me-1"></i>Showing: With 4th Subject Mode`;
            renderGpaDistributionTiers(data.gpa_distribution_with_4th || data.gpa_distribution || {});
        } else {
            if (btnWith) {
                btnWith.classList.remove("active", "btn-primary");
                btnWith.classList.add("btn-outline-primary");
            }
            if (btnWithout) {
                btnWithout.classList.add("active", "btn-primary");
                btnWithout.classList.remove("btn-outline-primary");
            }
            if (kpiGpa5Title) kpiGpa5Title.textContent = "Golden A+ (Without 4th Sub)";
            const countWithout = (summary.gpa_5_count_without_4th !== undefined) ? summary.gpa_5_count_without_4th : 0;
            const rateWithout = (summary.gpa_5_rate_without_4th !== undefined) ? summary.gpa_5_rate_without_4th : 0;
            if (kpiGpa5) kpiGpa5.textContent = countWithout;
            if (kpiGpa5Rate) kpiGpa5Rate.textContent = `${rateWithout}% of Appeared`;
            if (gpaDistModeNotice) gpaDistModeNotice.innerHTML = `<i class="bi bi-star-fill text-warning me-1"></i>Showing: Without 4th Subject (Golden / Pure A+)`;
            renderGpaDistributionTiers(data.gpa_distribution_without_4th || {});
        }
    }

    // TAB 6: Absent Students Analysis & Roster
    function renderAbsentStudents(absentData) {
        if (!absentData) return;

        // 1. KPI Summary Cards in Absent Pane
        const absentPaneTotal = document.getElementById("absentPaneTotal");
        const absentPaneRate = document.getElementById("absentPaneRate");
        const absentPaneRegular = document.getElementById("absentPaneRegular");
        const absentPaneRegularRate = document.getElementById("absentPaneRegularRate");
        const absentPaneIrregular = document.getElementById("absentPaneIrregular");
        const absentPaneIrregularRate = document.getElementById("absentPaneIrregularRate");
        const absentPaneAppeared = document.getElementById("absentPaneAppeared");

        if (absentPaneTotal) absentPaneTotal.textContent = absentData.total_absent || 0;
        if (absentPaneRate) absentPaneRate.textContent = `${absentData.absent_rate || 0}% of Total`;
        if (absentPaneRegular) absentPaneRegular.textContent = absentData.regular_absent || 0;
        if (absentPaneRegularRate) absentPaneRegularRate.textContent = `${absentData.regular_absent_rate || 0}%`;
        if (absentPaneIrregular) absentPaneIrregular.textContent = absentData.irregular_absent || 0;
        if (absentPaneIrregularRate) absentPaneIrregularRate.textContent = `${absentData.irregular_absent_rate || 0}%`;
        if (absentPaneAppeared) absentPaneAppeared.textContent = absentData.appeared || 0;

        // 2. Group-wise Absent Breakdown Table
        const groupTbody = document.getElementById("absentGroupTableBody");
        if (groupTbody) {
            groupTbody.innerHTML = "";
            const groups = Object.values(absentData.groups_absent || {});
            if (groups.length === 0) {
                groupTbody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">No group absent records available.</td></tr>`;
            } else {
                groups.forEach(g => {
                    const tr = document.createElement("tr");
                    const rateColor = g.absent_rate > 5 ? "text-danger fw-bold" : "text-muted";
                    tr.innerHTML = `
                        <td class="text-start ps-3 fw-bold text-dark">${g.group}</td>
                        <td class="fw-semibold">${g.total}</td>
                        <td>${g.appeared}</td>
                        <td class="fw-bold text-danger">${g.absent}</td>
                        <td class="${rateColor}">${g.absent_rate}%</td>
                    `;
                    groupTbody.appendChild(tr);
                });
            }
        }

        // 3. Absent Students Roster Table with filter and search
        const rosterTbody = document.getElementById("absentTableBody");
        const noticeEl = document.getElementById("absentFilteredNotice");
        const searchInput = document.getElementById("absentSearch");
        const typeButtonsContainer = document.getElementById("absentTypeButtons");

        let currentFilterType = "ALL";
        const students = absentData.students || [];

        function filterAndRenderRoster() {
            if (!rosterTbody) return;
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            const filtered = students.filter(st => {
                const stType = String(st.Type || "REGULAR").toUpperCase();
                if (currentFilterType !== "ALL" && !stType.includes(currentFilterType)) {
                    return false;
                }
                if (query) {
                    const roll = String(st.Roll || "").toLowerCase();
                    const reg = String(st.Registration || "").toLowerCase();
                    const name = String(st.Name || "").toLowerCase();
                    const father = String(st["Father Name"] || st.Father_Name || "").toLowerCase();
                    const inst = String(st.Institute || "").toLowerCase();
                    const grp = String(st.Group || "").toLowerCase();
                    if (!roll.includes(query) && !reg.includes(query) && !name.includes(query) && !father.includes(query) && !inst.includes(query) && !grp.includes(query)) {
                        return false;
                    }
                }
                return true;
            });

            if (noticeEl) {
                noticeEl.textContent = `Showing ${filtered.length} of ${students.length} absent students`;
            }

            rosterTbody.innerHTML = "";
            if (filtered.length === 0) {
                rosterTbody.innerHTML = `<tr><td colspan="9" class="text-center py-4 text-muted">No absent students match the criteria.</td></tr>`;
                return;
            }

            filtered.forEach(st => {
                const tr = document.createElement("tr");
                let typeBadge = `<span class="badge bg-secondary-subtle text-secondary border">${st.Type || "REGULAR"}</span>`;
                if (st.Type === "REGULAR") {
                    typeBadge = `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`;
                } else if (st.Type === "IRREGULAR") {
                    typeBadge = `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`;
                }

                tr.innerHTML = `
                    <td class="ps-3 fw-bold text-dark">${st.Roll || "-"}</td>
                    <td class="font-monospace text-muted small">${st.Registration || "-"}</td>
                    <td class="fw-semibold text-dark">${st.Name || "N/A"}</td>
                    <td class="text-muted small">${st["Father Name"] || st.Father_Name || "-"}</td>
                    <td><span class="badge bg-light text-dark border">${st.Group || "GENERAL"}</span></td>
                    <td>${typeBadge}</td>
                    <td class="small text-truncate" style="max-width: 180px;" title="${st.Institute || ''}">${st.Institute || "-"}</td>
                    <td class="text-center"><span class="badge bg-secondary px-2.5 py-1">ABSENT</span></td>
                    <td class="small text-muted">${st.Remarks || "Did not sit for examination"}</td>
                `;
                rosterTbody.appendChild(tr);
            });
        }

        if (typeButtonsContainer) {
            const btns = typeButtonsContainer.querySelectorAll("button");
            btns.forEach(btn => {
                btn.onclick = () => {
                    btns.forEach(b => {
                        b.classList.remove("active", "btn-dark", "btn-primary", "btn-warning");
                        if (b.dataset.type === "REGULAR") b.classList.add("btn-outline-primary");
                        else if (b.dataset.type === "IRREGULAR") b.classList.add("btn-outline-warning");
                        else b.classList.add("btn-outline-dark");
                    });
                    btn.classList.add("active");
                    if (btn.dataset.type === "REGULAR") {
                        btn.classList.remove("btn-outline-primary");
                        btn.classList.add("btn-primary");
                    } else if (btn.dataset.type === "IRREGULAR") {
                        btn.classList.remove("btn-outline-warning");
                        btn.classList.add("btn-warning");
                    } else {
                        btn.classList.remove("btn-outline-dark");
                        btn.classList.add("btn-dark");
                    }
                    currentFilterType = btn.dataset.type;
                    filterAndRenderRoster();
                };
            });
        }

        if (searchInput) {
            searchInput.oninput = filterAndRenderRoster;
        }

        filterAndRenderRoster();
    }

    // Client-side fallback analyzer (Used if analyzing directly from active memory)
    function computeClientAnalysis(records) {
        const validGrades = ["A+", "A", "A-", "B", "C", "D", "F"];
        const gradePoints = { "A+": 5.0, "A": 4.0, "A-": 3.5, "B": 3.0, "C": 2.0, "D": 1.0, "F": 0.0 };
        const oneFailStudents = [];
        const twoFailStudents = [];
        const multiFailStudents = [];
        const absentStudents = [];
        const oneSubFreq = {};
        const twoSubFreq = {};
        const subjectStats = {};
        const groupStudents = {};
        const typeStudents = {};

        const gpaDistWith4th = {
            "5.00 (A+)": 0, "4.00 - 4.99 (A)": 0, "3.50 - 3.99 (A-)": 0,
            "3.00 - 3.49 (B)": 0, "2.00 - 2.99 (C)": 0, "1.00 - 1.99 (D)": 0,
            "0.00 (F / Fail)": 0
        };
        const gpaDistWithout4th = {
            "5.00 (A+)": 0, "4.00 - 4.99 (A)": 0, "3.50 - 3.99 (A-)": 0,
            "3.00 - 3.49 (B)": 0, "2.00 - 2.99 (C)": 0, "1.00 - 1.99 (D)": 0,
            "0.00 (F / Fail)": 0
        };

        let passedCount = 0;
        let failedCount = 0;
        let absentCount = 0;

        function identifyFourthSubjectClient(subjectGrades, group) {
            const keys = Object.keys(subjectGrades);
            if (keys.length <= 1) return null;
            const fourthKeywords = [
                "4TH", "FOURTH", "OPTIONAL", "AGRICULTURE", "HIGHER MATH", "BIOLOGY",
                "STATISTICS", "HOME SCIENCE", "PSYCHOLOGY", "GEOGRAPHY", "ISLAMIC STUDIES",
                "CIVICS", "ECONOMICS", "SOCIOLOGY", "SOCIAL WORK", "FINANCE"
            ];
            for (let i = keys.length - 1; i >= 0; i--) {
                const sUpper = keys[i].toUpperCase();
                for (const kw of fourthKeywords) {
                    if (sUpper.includes(kw)) return keys[i];
                }
            }
            if (keys.length >= 7) return keys[keys.length - 1];
            return null;
        }

        records.forEach(raw => {
            const roll = String(raw.Roll || "").trim();
            const reg = String(raw.Registration || "").trim();
            const name = String(raw.Name || "N/A").trim();
            const father = String(raw["Father Name"] || raw.Father_Name || "-").trim();
            const institute = String(raw.Institute || "-").trim();
            let group = String(raw.Group || "GENERAL").trim().toUpperCase();
            if (!group || group === "NAN" || group === "-") group = "GENERAL";

            let resStatus = String(raw["Result Status"] || raw.Result || "").trim().toUpperCase();
            let gpaNum = parseFloat(raw.GPA || 0) || 0.0;

            const subjectGrades = {};
            if (raw.Subject_Grades && typeof raw.Subject_Grades === "object") {
                Object.entries(raw.Subject_Grades).forEach(([k, v]) => {
                    const g = String(v).trim().toUpperCase();
                    if (validGrades.includes(g)) subjectGrades[k.trim()] = g;
                });
            } else if (raw.Subject_Grades_Summary || raw["Subject Grades Summary"]) {
                const sumStr = String(raw.Subject_Grades_Summary || raw["Subject Grades Summary"]);
                sumStr.split(",").forEach(part => {
                    if (part.includes(":")) {
                        const [sName, sGrade] = part.split(":");
                        const g = sGrade.trim().toUpperCase();
                        if (validGrades.includes(g)) subjectGrades[sName.trim()] = g;
                    }
                });
            }

            // Detect candidate type
            let cType = String(raw.Type || raw["Candidate Type"] || raw["Student Type"] || "").trim().toUpperCase();
            if (!cType || cType === "NAN" || cType === "-") {
                cType = (Object.keys(subjectGrades).length > 0 && Object.keys(subjectGrades).length <= 4) ? "IRREGULAR" : "REGULAR";
            } else if (cType.includes("IRREG")) {
                cType = "IRREGULAR";
            } else if (cType.includes("REG")) {
                cType = "REGULAR";
            } else if (cType.includes("IMP")) {
                cType = "IMPROVEMENT";
            } else if (cType.includes("PRIV")) {
                cType = "PRIVATE";
            }

            const isAbsent = resStatus.includes("ABSENT") || resStatus.includes("ABS");

            if (isAbsent) {
                absentCount++;
                const absentObj = {
                    Roll: roll,
                    Registration: reg,
                    Name: name,
                    "Father Name": father,
                    Group: group,
                    Type: cType,
                    Institute: institute,
                    Result: "ABSENT",
                    Remarks: "Did not sit for examination",
                    isAbsent: true,
                    isPassed: false,
                    isFail: false
                };
                absentStudents.push(absentObj);

                if (!groupStudents[group]) groupStudents[group] = [];
                groupStudents[group].push(absentObj);

                if (!typeStudents[cType]) typeStudents[cType] = [];
                typeStudents[cType].push(absentObj);
                return;
            }

            // Fourth subject analysis
            const fourthSubj = identifyFourthSubjectClient(subjectGrades, group);
            const mainSubjects = {};
            let fourthGrade = null;
            Object.entries(subjectGrades).forEach(([s, g]) => {
                if (s === fourthSubj) {
                    fourthGrade = g;
                } else {
                    mainSubjects[s] = g;
                }
            });

            const mainFailed = Object.entries(mainSubjects).filter(([_, g]) => g === "F").map(([s]) => s);
            const passedSubjs = Object.entries(subjectGrades).filter(([_, g]) => g !== "F").map(([s, g]) => `${s} (${g})`);
            const failCnt = mainFailed.length; // Failing 4th sub does not fail candidate

            const isFail = resStatus.includes("FAIL") || failCnt > 0;
            const isPassed = !isFail && (gpaNum > 0 || resStatus.includes("PASSED"));

            // Compute GPA without 4th subject
            let gpaWithout4th = 0.0;
            const mainKeys = Object.keys(mainSubjects);
            if (mainKeys.length > 0) {
                const sumMainGp = mainKeys.reduce((acc, s) => acc + (gradePoints[mainSubjects[s]] || 0.0), 0);
                gpaWithout4th = Math.round((sumMainGp / mainKeys.length) * 100) / 100;
            } else {
                gpaWithout4th = gpaNum;
            }

            // Compute GPA with 4th subject
            let computedW4 = gpaNum;
            if (isPassed && mainKeys.length > 0) {
                let fourthBonus = 0.0;
                if (fourthGrade && gradePoints[fourthGrade] !== undefined) {
                    fourthBonus = Math.max(0.0, gradePoints[fourthGrade] - 2.0);
                }
                const sumMainGp = mainKeys.reduce((acc, s) => acc + (gradePoints[mainSubjects[s]] || 0.0), 0);
                computedW4 = Math.min(5.0, Math.round(((sumMainGp + fourthBonus) / mainKeys.length) * 100) / 100);
            }
            const gpaWith4th = (gpaNum > 0) ? Math.max(gpaNum, computedW4) : computedW4;

            const isGpa5With = isPassed && gpaWith4th >= 5.0;
            const isGpa5Without = isPassed && gpaWithout4th >= 5.0;

            if (isPassed) passedCount++;
            else if (isFail) failedCount++;

            const stInfo = {
                Roll: roll,
                Registration: reg,
                Name: name,
                "Father Name": father,
                Institute: institute,
                Group: group,
                Type: cType,
                isPassed,
                isFail,
                isAbsent: false,
                gpaNum: gpaWith4th,
                gpaWith4th: gpaWith4th,
                gpaWithout4th: gpaWithout4th,
                isGpa5With: isGpa5With,
                isGpa5Without: isGpa5Without,
                failCnt,
                Result: isFail ? "FAIL" : (isPassed ? "PASSED" : "UNKNOWN")
            };

            if (!groupStudents[group]) groupStudents[group] = [];
            groupStudents[group].push(stInfo);

            if (!typeStudents[cType]) typeStudents[cType] = [];
            typeStudents[cType].push(stInfo);

            // Fail categories
            if (failCnt === 1) {
                oneFailStudents.push({
                    Roll: roll, Registration: reg, Name: name, Group: group, Type: cType,
                    GPA_With_4th: gpaWith4th.toFixed(2),
                    GPA_Without_4th: gpaWithout4th.toFixed(2),
                    Failed_Subject: mainFailed[0],
                    Passed_Grades_Summary: passedSubjs.join(", ") || "-"
                });
                oneSubFreq[mainFailed[0]] = (oneSubFreq[mainFailed[0]] || 0) + 1;
            } else if (failCnt === 2) {
                twoFailStudents.push({
                    Roll: roll, Registration: reg, Name: name, Group: group, Type: cType,
                    GPA_With_4th: gpaWith4th.toFixed(2),
                    GPA_Without_4th: gpaWithout4th.toFixed(2),
                    Failed_Subjects: mainFailed,
                    Passed_Grades_Summary: passedSubjs.join(", ") || "-"
                });
                mainFailed.forEach(s => twoSubFreq[s] = (twoSubFreq[s] || 0) + 1);
            } else if (failCnt >= 3) {
                multiFailStudents.push({
                    Roll: roll, Name: name, Group: group, Type: cType,
                    GPA_With_4th: gpaWith4th.toFixed(2),
                    GPA_Without_4th: gpaWithout4th.toFixed(2),
                    Failed_Count: failCnt, Failed_Subjects: mainFailed,
                    Passed_Grades_Summary: passedSubjs.join(", ") || "-"
                });
            }

            // GPA Distribution With 4th
            if (isFail || !isPassed) gpaDistWith4th["0.00 (F / Fail)"]++;
            else if (gpaWith4th >= 5.0) gpaDistWith4th["5.00 (A+)"]++;
            else if (gpaWith4th >= 4.0) gpaDistWith4th["4.00 - 4.99 (A)"]++;
            else if (gpaWith4th >= 3.5) gpaDistWith4th["3.50 - 3.99 (A-)"]++;
            else if (gpaWith4th >= 3.0) gpaDistWith4th["3.00 - 3.49 (B)"]++;
            else if (gpaWith4th >= 2.0) gpaDistWith4th["2.00 - 2.99 (C)"]++;
            else if (gpaWith4th >= 1.0) gpaDistWith4th["1.00 - 1.99 (D)"]++;
            else gpaDistWith4th["0.00 (F / Fail)"]++;

            // GPA Distribution Without 4th
            if (isFail || !isPassed) gpaDistWithout4th["0.00 (F / Fail)"]++;
            else if (gpaWithout4th >= 5.0) gpaDistWithout4th["5.00 (A+)"]++;
            else if (gpaWithout4th >= 4.0) gpaDistWithout4th["4.00 - 4.99 (A)"]++;
            else if (gpaWithout4th >= 3.5) gpaDistWithout4th["3.50 - 3.99 (A-)"]++;
            else if (gpaWithout4th >= 3.0) gpaDistWithout4th["3.00 - 3.49 (B)"]++;
            else if (gpaWithout4th >= 2.0) gpaDistWithout4th["2.00 - 2.99 (C)"]++;
            else if (gpaWithout4th >= 1.0) gpaDistWithout4th["1.00 - 1.99 (D)"]++;
            else gpaDistWithout4th["0.00 (F / Fail)"]++;

            // Subject Grade distribution per (group, type, subject)
            Object.entries(subjectGrades).forEach(([subj, g]) => {
                const keys = [
                    `${group}:::${cType}:::${subj}`,
                    `ALL:::${cType}:::${subj}`,
                    `${group}:::ALL:::${subj}`,
                    `ALL:::ALL:::${subj}`
                ];
                keys.forEach(key => {
                    const [grpPart, typePart, subjPart] = key.split(":::");
                    if (!subjectStats[key]) {
                        subjectStats[key] = { group: grpPart, type: typePart, subject: subjPart, grades: {} };
                        validGrades.forEach(vg => subjectStats[key].grades[vg] = 0);
                    }
                    subjectStats[key].grades[g] = (subjectStats[key].grades[g] || 0) + 1;
                });
            });
        });

        const totalStudents = records.length;
        const appeared = totalStudents - absentCount;
        const passRate = appeared > 0 ? ((passedCount / appeared) * 100).toFixed(2) : "0.00";
        const failRate = appeared > 0 ? ((failedCount / appeared) * 100).toFixed(2) : "0.00";
        const absentRate = totalStudents > 0 ? ((absentCount / totalStudents) * 100).toFixed(2) : "0.00";
        const gpa5CountWith = gpaDistWith4th["5.00 (A+)"];
        const gpa5RateWith = appeared > 0 ? ((gpa5CountWith / appeared) * 100).toFixed(2) : "0.00";
        const gpa5CountWithout = gpaDistWithout4th["5.00 (A+)"];
        const gpa5RateWithout = appeared > 0 ? ((gpa5CountWithout / appeared) * 100).toFixed(2) : "0.00";

        // Build groups summary
        const groupsSummary = {};
        const groupsAbsent = {};
        Object.entries(groupStudents).forEach(([grp, list]) => {
            const grpTot = list.length;
            const grpPass = list.filter(s => s.isPassed).length;
            const grpFail = list.filter(s => s.isFail).length;
            const grpAbs = list.filter(s => s.isAbsent).length;
            const grpApp = grpTot - grpAbs;
            const grpPassRate = grpApp > 0 ? ((grpPass / grpApp) * 100).toFixed(2) : "0.00";
            const grpAbsRate = grpTot > 0 ? ((grpAbs / grpTot) * 100).toFixed(2) : "0.00";
            const grpGpa5With = list.filter(s => s.isGpa5With).length;
            const grpGpa5Without = list.filter(s => s.isGpa5Without).length;
            groupsSummary[grp] = {
                group: grp, total: grpTot, appeared: grpApp, passed: grpPass,
                failed: grpFail, absent: grpAbs, pass_rate: grpPassRate,
                gpa5_count: grpGpa5With,
                gpa5_count_with_4th: grpGpa5With,
                gpa5_count_without_4th: grpGpa5Without,
                one_fail_count: list.filter(s => s.failCnt === 1).length,
                two_fail_count: list.filter(s => s.failCnt === 2).length,
                multi_fail_count: list.filter(s => s.failCnt >= 3).length
            };
            groupsAbsent[grp] = {
                group: grp, total: grpTot, appeared: grpApp, absent: grpAbs, absent_rate: grpAbsRate
            };
        });

        // Build Candidate Types Summary
        const typesSummary = {};
        const typesAbsent = {};
        Object.entries(typeStudents).forEach(([cType, list]) => {
            const tTot = list.length;
            const tPass = list.filter(s => s.isPassed).length;
            const tFail = list.filter(s => s.isFail).length;
            const tAbs = list.filter(s => s.isAbsent).length;
            const tApp = tTot - tAbs;
            const tPassRate = tApp > 0 ? ((tPass / tApp) * 100).toFixed(2) : "0.00";
            const tFailRate = tApp > 0 ? ((tFail / tApp) * 100).toFixed(2) : "0.00";
            const tAbsRate = tTot > 0 ? ((tAbs / tTot) * 100).toFixed(2) : "0.00";
            const tGpa5With = list.filter(s => s.isGpa5With).length;
            const tGpa5Without = list.filter(s => s.isGpa5Without).length;
            typesSummary[cType] = {
                type: cType,
                total: tTot,
                appeared: tApp,
                passed: tPass,
                failed: tFail,
                absent: tAbs,
                pass_rate: tPassRate,
                fail_rate: tFailRate,
                absent_rate: tAbsRate,
                gpa5_count: tGpa5With,
                gpa5_count_with_4th: tGpa5With,
                gpa5_count_without_4th: tGpa5Without,
                gpa5_rate: tApp > 0 ? ((tGpa5With / tApp) * 100).toFixed(2) : "0.00",
                one_fail_count: list.filter(s => s.failCnt === 1).length,
                two_fail_count: list.filter(s => s.failCnt === 2).length,
                multi_fail_count: list.filter(s => s.failCnt >= 3).length
            };
            typesAbsent[cType] = {
                type: cType, total: tTot, appeared: tApp, absent: tAbs, absent_rate: tAbsRate
            };
        });

        // Build Group x Type Matrix
        const groupTypeComparison = [];
        Object.keys(groupStudents).sort().forEach(grp => {
            ["REGULAR", "IRREGULAR", "IMPROVEMENT", "PRIVATE"].forEach(cType => {
                const sub = (groupStudents[grp] || []).filter(s => s.Type === cType);
                if (sub.length > 0) {
                    const subTot = sub.length;
                    const subPass = sub.filter(s => s.isPassed).length;
                    const subFail = sub.filter(s => s.isFail).length;
                    const subAbs = sub.filter(s => s.isAbsent).length;
                    const subApp = subTot - subAbs;
                    const subRate = subApp > 0 ? ((subPass / subApp) * 100).toFixed(2) : "0.00";
                    const subGpa5With = sub.filter(s => s.isGpa5With).length;
                    const subGpa5Without = sub.filter(s => s.isGpa5Without).length;
                    groupTypeComparison.push({
                        group: grp,
                        type: cType,
                        total: subTot,
                        appeared: subApp,
                        passed: subPass,
                        failed: subFail,
                        absent: subAbs,
                        pass_rate: subRate,
                        gpa5_count: subGpa5With,
                        gpa5_count_with_4th: subGpa5With,
                        gpa5_count_without_4th: subGpa5Without,
                        one_fail_count: sub.filter(s => s.failCnt === 1).length,
                        two_fail_count: sub.filter(s => s.failCnt === 2).length,
                        multi_fail_count: sub.filter(s => s.failCnt >= 3).length
                    });
                }
            });
        });

        const regularSummary = typesSummary["REGULAR"] || {
            type: "REGULAR", total: 0, appeared: 0, passed: 0, failed: 0, absent: 0, pass_rate: "0.00",
            gpa5_count: 0, gpa5_count_with_4th: 0, gpa5_count_without_4th: 0,
            one_fail_count: 0, two_fail_count: 0, multi_fail_count: 0
        };
        const irregularSummary = typesSummary["IRREGULAR"] || {
            type: "IRREGULAR", total: 0, appeared: 0, passed: 0, failed: 0, absent: 0, pass_rate: "0.00",
            gpa5_count: 0, gpa5_count_with_4th: 0, gpa5_count_without_4th: 0,
            one_fail_count: 0, two_fail_count: 0, multi_fail_count: 0
        };

        // Build subject analysis list
        const subjectAnalysis = [];
        Object.values(subjectStats).forEach(entry => {
            const tot = Object.values(entry.grades).reduce((a, b) => a + b, 0);
            if (tot === 0) return;
            const fCount = entry.grades["F"] || 0;
            const pCount = tot - fCount;
            const pRate = ((pCount / tot) * 100).toFixed(2);
            subjectAnalysis.push({
                group: entry.group,
                type: entry.type,
                subject: entry.subject,
                appeared: tot,
                passed: pCount,
                failed: fCount,
                pass_rate: pRate,
                grades: entry.grades
            });
        });

        subjectAnalysis.sort((a, b) => a.group.localeCompare(b.group) || a.type.localeCompare(b.type) || a.subject.localeCompare(b.subject));

        const regAbsent = absentStudents.filter(s => s.Type === "REGULAR").length;
        const irregAbsent = absentStudents.filter(s => s.Type === "IRREGULAR").length;

        const absentAnalysis = {
            total_absent: absentCount,
            absent_rate: absentRate,
            regular_absent: regAbsent,
            regular_absent_rate: (regularSummary.total > 0 ? ((regAbsent / regularSummary.total) * 100).toFixed(2) : "0.00"),
            irregular_absent: irregAbsent,
            irregular_absent_rate: (irregularSummary.total > 0 ? ((irregAbsent / irregularSummary.total) * 100).toFixed(2) : "0.00"),
            appeared: appeared,
            groups_absent: groupsAbsent,
            types_absent: typesAbsent,
            students: absentStudents
        };

        return {
            total_students: totalStudents,
            summary: {
                total_students: totalStudents,
                appeared: appeared,
                passed: passedCount,
                failed: failedCount,
                absent: absentCount,
                pass_rate: passRate,
                fail_rate: failRate,
                absent_rate: absentRate,
                gpa_5_count: gpa5CountWith,
                gpa_5_rate: gpa5RateWith,
                gpa_5_count_with_4th: gpa5CountWith,
                gpa_5_rate_with_4th: gpa5RateWith,
                gpa_5_count_without_4th: gpa5CountWithout,
                gpa_5_rate_without_4th: gpa5RateWithout,
                one_subject_fail_count: oneFailStudents.length,
                two_subject_fail_count: twoFailStudents.length,
                multi_subject_fail_count: multiFailStudents.length
            },
            groups_summary: groupsSummary,
            types_summary: typesSummary,
            regular_summary: regularSummary,
            irregular_summary: irregularSummary,
            group_type_comparison: groupTypeComparison,
            subject_analysis: subjectAnalysis,
            one_subject_fail: {
                total: oneFailStudents.length,
                subjects_summary: Object.fromEntries(Object.entries(oneSubFreq).sort((a, b) => b[1] - a[1])),
                students: oneFailStudents
            },
            two_subject_fail: {
                total: twoFailStudents.length,
                subjects_summary: Object.fromEntries(Object.entries(twoSubFreq).sort((a, b) => b[1] - a[1])),
                students: twoFailStudents
            },
            multi_subject_fail: {
                total: multiFailStudents.length,
                students: multiFailStudents
            },
            absent_analysis: absentAnalysis,
            gpa_distribution: gpaDistWith4th,
            gpa_distribution_with_4th: gpaDistWith4th,
            gpa_distribution_without_4th: gpaDistWithout4th,
            available_groups: Object.keys(groupStudents).sort(),
            available_types: Object.keys(typeStudents).sort()
        };
    }

    // Print Report in a formatted new window
    function printAnalysisReport(analysisData, metaInfo) {
        const printWin = window.open("", "_blank");
        if (!printWin) {
            alert("Please allow popups to open the print layout.");
            return;
        }

        const summary = analysisData.summary || {};
        const reg = analysisData.regular_summary || {};
        const irreg = analysisData.irregular_summary || {};
        const groupsSummary = analysisData.groups_summary || {};
        const absentData = analysisData.absent_analysis || {};
        const oneFail = analysisData.one_subject_fail || {};
        const twoFail = analysisData.two_subject_fail || {};
        const mFail = analysisData.multi_subject_fail || {};
        const groupType = analysisData.group_type_comparison || [];

        const gpa5WithCount = (summary.gpa_5_count_with_4th !== undefined) ? summary.gpa_5_count_with_4th : (summary.gpa_5_count || 0);
        const gpa5WithRate = (summary.gpa_5_rate_with_4th !== undefined) ? summary.gpa_5_rate_with_4th : (summary.gpa_5_rate || 0);
        const gpa5WithoutCount = (summary.gpa_5_count_without_4th !== undefined) ? summary.gpa_5_count_without_4th : 0;
        const gpa5WithoutRate = (summary.gpa_5_rate_without_4th !== undefined) ? summary.gpa_5_rate_without_4th : 0;

        const html = `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <title>Board Result Comprehensive Analysis Report - ${metaInfo.exam || "HSC"} ${metaInfo.year || ""}</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
            <style>
                @page { size: A4; margin: 10mm 8mm; }
                body {
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                    font-size: 11px;
                    color: #1e293b;
                    background-color: #f8fafc;
                    padding: 20px;
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                }
                .report-container {
                    max-width: 1100px;
                    margin: 0 auto;
                    background: #ffffff;
                    padding: 25px 30px;
                    border-radius: 12px;
                    box-shadow: 0 4px 15px rgba(0,0,0,0.06);
                }
                .report-header {
                    border-bottom: 2px solid #0f172a;
                    padding-bottom: 15px;
                    margin-bottom: 20px;
                }
                .section-title {
                    font-size: 13px;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    color: #0f172a;
                    border-bottom: 1.5px solid #cbd5e1;
                    padding-bottom: 6px;
                    margin-top: 24px;
                    margin-bottom: 12px;
                    display: flex;
                    align-items: center;
                    gap: 6px;
                }
                .section-desc {
                    font-size: 10.5px;
                    color: #64748b;
                    margin-top: -8px;
                    margin-bottom: 12px;
                }
                .kpi-card {
                    background: #ffffff;
                    border: 1px solid #e2e8f0;
                    border-radius: 8px;
                    padding: 10px 12px;
                    text-align: center;
                    height: 100%;
                }
                .kpi-card .kpi-num {
                    font-size: 20px;
                    font-weight: 800;
                    line-height: 1.2;
                    margin-top: 4px;
                    margin-bottom: 2px;
                }
                .kpi-card .kpi-label {
                    font-size: 9.5px;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    color: #64748b;
                }
                .kpi-card .kpi-sub {
                    font-size: 9px;
                    font-weight: 600;
                }
                .kpi-slate { border-top: 3px solid #475569; }
                .kpi-slate .kpi-num { color: #1e293b; }
                .kpi-blue { border-top: 3px solid #2563eb; }
                .kpi-blue .kpi-num { color: #1d4ed8; }
                .kpi-green { border-top: 3px solid #16a34a; }
                .kpi-green .kpi-num { color: #15803d; }
                .kpi-red { border-top: 3px solid #dc2626; }
                .kpi-red .kpi-num { color: #b91c1c; }
                .kpi-gold { border-top: 3px solid #d97706; }
                .kpi-gold .kpi-num { color: #b45309; }
                .kpi-purple { border-top: 3px solid #7c3aed; }
                .kpi-purple .kpi-num { color: #6d28d9; }
                .kpi-orange { border-top: 3px solid #ea580c; }
                .kpi-orange .kpi-num { color: #c2410c; }

                table {
                    width: 100%;
                    margin-bottom: 15px;
                    font-size: 10px;
                    border-collapse: collapse;
                }
                th {
                    background-color: #0f172a !important;
                    color: #ffffff !important;
                    font-weight: 600;
                    text-transform: uppercase;
                    font-size: 9px;
                    letter-spacing: 0.3px;
                    padding: 6px 6px;
                    border: 1px solid #334155 !important;
                    text-align: center;
                    vertical-align: middle;
                }
                td {
                    padding: 5px 6px;
                    border: 1px solid #e2e8f0;
                    text-align: center;
                    vertical-align: middle;
                }
                tbody tr:nth-child(even) { background-color: #f8fafc; }
                .table-highlight-gold { background-color: #fef3c7 !important; font-weight: 700; color: #92400e; }
                .table-highlight-red { background-color: #fee2e2 !important; font-weight: 700; color: #b91c1c; }
                .table-highlight-green { background-color: #dcfce7 !important; font-weight: 700; color: #15803d; }
                .guide-box {
                    background-color: #f0fdf4;
                    border: 1px solid #bbf7d0;
                    border-radius: 8px;
                    padding: 12px 16px;
                    margin-top: 15px;
                    margin-bottom: 20px;
                }
                .guide-item {
                    font-size: 10px;
                    margin-bottom: 4px;
                    color: #1e3a24;
                }
                .badge-pill-custom {
                    font-size: 9px;
                    padding: 3px 8px;
                    border-radius: 12px;
                    font-weight: 600;
                    display: inline-block;
                }
                @media print {
                    body { background-color: #ffffff !important; padding: 0 !important; }
                    .report-container { box-shadow: none !important; padding: 0 !important; max-width: 100% !important; }
                    .no-print { display: none !important; }
                    .report-section { page-break-inside: avoid; }
                    tr { page-break-inside: avoid; }
                }
            </style>
        </head>
        <body>
            <!-- Interactive Top Action Toolbar (Hidden on paper) -->
            <div class="no-print d-flex justify-content-between align-items-center mb-4 p-3 bg-white rounded-3 shadow-sm border">
                <div class="d-flex align-items-center gap-2">
                    <span class="badge bg-primary fs-7 px-3 py-2"><i class="bi bi-file-earmark-pdf-fill me-1"></i> Print / PDF Layout Ready</span>
                    <span class="text-muted small">Optimized for standard A4 paper and clean reading</span>
                </div>
                <div class="d-flex gap-2">
                    <button class="btn btn-primary btn-sm fw-bold px-4 py-2 rounded-2" onclick="window.print()">
                        <i class="bi bi-printer-fill me-1"></i> Print / Save as PDF
                    </button>
                    <button class="btn btn-outline-secondary btn-sm px-3 py-2 rounded-2" onclick="window.close()">
                        <i class="bi bi-x-lg me-1"></i> Close
                    </button>
                </div>
            </div>

            <div class="report-container">
                <!-- Document Header -->
                <div class="report-header d-flex justify-content-between align-items-center flex-wrap gap-2">
                    <div>
                        <div class="badge-pill-custom bg-dark text-white mb-1">
                            <i class="bi bi-shield-check text-warning me-1"></i> OFFICIAL EDUCATION BOARD RESULT AUDIT
                        </div>
                        <h3 class="fw-bold text-dark mb-0" style="letter-spacing: -0.5px;">INSTITUTION RESULT PERFORMANCE REPORT</h3>
                        <div class="text-muted small mt-1">
                            <strong>Exam:</strong> ${metaInfo.exam || "Examination"} ${metaInfo.year || ""} &bull; 
                            <strong>Board:</strong> ${metaInfo.board || "Cumilla (Fast Server)"} &bull; 
                            <strong>Generated:</strong> ${new Date().toLocaleDateString('en-US', { day: 'numeric', month: 'short', year: 'numeric' })} at ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                        </div>
                    </div>
                    <div class="text-end">
                        <span class="badge bg-primary-subtle text-primary border border-primary-subtle px-3 py-2 fs-7 fw-bold d-block mb-1">
                            Total Students: ${summary.total_students || 0}
                        </span>
                        <span class="badge bg-success-subtle text-success border border-success-subtle px-2 py-1 fs-8">
                            Pass Rate: ${summary.pass_rate || 0}%
                        </span>
                    </div>
                </div>

                <!-- Executive Visual KPI Dashboard -->
                <div class="row g-2 mb-3">
                    <div class="col-3">
                        <div class="kpi-card kpi-slate">
                            <div class="kpi-label">Total Enrolled</div>
                            <div class="kpi-num">${summary.total_students || 0}</div>
                            <div class="kpi-sub text-muted">Registered Candidates</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-blue">
                            <div class="kpi-label">Exam Attended</div>
                            <div class="kpi-num">${summary.appeared || 0}</div>
                            <div class="kpi-sub text-primary">Sat for Examination</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-green">
                            <div class="kpi-label">Total Passed</div>
                            <div class="kpi-num">${summary.passed || 0}</div>
                            <div class="kpi-sub text-success">${summary.pass_rate || 0}% of Attended</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-red">
                            <div class="kpi-label">Total Failed</div>
                            <div class="kpi-num">${summary.failed || 0}</div>
                            <div class="kpi-sub text-danger">${summary.fail_rate || 0}% of Attended</div>
                        </div>
                    </div>
                </div>

                <div class="row g-2 mb-3">
                    <div class="col-3">
                        <div class="kpi-card kpi-gold">
                            <div class="kpi-label">GPA 5.0 (With 4th)</div>
                            <div class="kpi-num">${gpa5WithCount}</div>
                            <div class="kpi-sub text-warning">${gpa5WithRate}% &bull; Official Board GPA</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-purple">
                            <div class="kpi-label">Golden A+ (Without 4th)</div>
                            <div class="kpi-num">${gpa5WithoutCount}</div>
                            <div class="kpi-sub" style="color: #6d28d9;">${gpa5WithoutRate}% &bull; Pure Main Subjects</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-orange">
                            <div class="kpi-label">1-Subject Fail</div>
                            <div class="kpi-num">${summary.one_subject_fail_count || 0}</div>
                            <div class="kpi-sub text-danger">Supplementary Eligible</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-slate">
                            <div class="kpi-label">Absent Students</div>
                            <div class="kpi-num">${summary.absent || 0}</div>
                            <div class="kpi-sub text-secondary">${summary.absent_rate || 0}% Did Not Sit</div>
                        </div>
                    </div>
                </div>

                <!-- Plain-English Explanatory Guide Box for easy understanding -->
                <div class="guide-box">
                    <div class="fw-bold text-success mb-1" style="font-size: 11px;">
                        <i class="bi bi-info-circle-fill me-1"></i> How to Easily Understand These Metrics:
                    </div>
                    <div class="row g-2">
                        <div class="col-md-6">
                            <div class="guide-item">
                                &bull; <strong>GPA 5.0 (With 4th Subject):</strong> The official Board GPA including bonus points above 2.0 earned in the optional 4th subject.
                            </div>
                            <div class="guide-item">
                                &bull; <strong>Golden A+ (Without 4th Subject):</strong> Pure curriculum mastery where student achieved 5.00 GPA in compulsory & group subjects without needing optional bonus points.
                            </div>
                        </div>
                        <div class="col-md-6">
                            <div class="guide-item">
                                &bull; <strong>4th Subject Safety Rule:</strong> Failing in the 4th optional subject does NOT fail a student who passed all main subjects (Official Board Rule).
                            </div>
                            <div class="guide-item">
                                &bull; <strong>One Subject Fail:</strong> Students who failed in only 1 subject and are eligible for single-subject supplementary re-examination.
                            </div>
                        </div>
                    </div>
                </div>

                <!-- SECTION 1: OVERALL INSTITUTION SUMMARY -->
                <div class="report-section">
                    <div class="section-title">
                        <i class="bi bi-buildings-fill text-primary"></i> 1. Overall Institution Performance Summary
                    </div>
                    <div class="section-desc">Consolidated executive overview of all registered students, attendance, pass rates, and grade tiers.</div>
                    <table class="table table-bordered table-sm text-center">
                        <thead>
                            <tr>
                                <th>Total Enrolled</th>
                                <th>Attended Exam</th>
                                <th>Absent</th>
                                <th>Passed</th>
                                <th>Failed</th>
                                <th>Pass Rate (%)</th>
                                <th>GPA 5.0 (With 4th)</th>
                                <th>Golden A+ (Without 4th)</th>
                                <th>1-Sub Fail</th>
                                <th>2-Sub Fail</th>
                                <th>3+ Sub Fail</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td class="fw-bold">${summary.total_students || 0}</td>
                                <td class="fw-semibold text-primary">${summary.appeared || 0}</td>
                                <td class="fw-bold text-secondary">${summary.absent || 0}</td>
                                <td class="fw-bold text-success">${summary.passed || 0}</td>
                                <td class="fw-bold text-danger">${summary.failed || 0}</td>
                                <td class="fw-bold table-highlight-green">${summary.pass_rate || 0}%</td>
                                <td class="fw-bold table-highlight-gold">${gpa5WithCount}</td>
                                <td class="fw-bold" style="color: #6d28d9; background-color: #f3e8ff;">${gpa5WithoutCount}</td>
                                <td class="fw-bold text-primary">${summary.one_subject_fail_count || 0}</td>
                                <td class="fw-bold text-danger">${summary.two_subject_fail_count || 0}</td>
                                <td class="text-muted">${summary.multi_subject_fail_count || 0}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <!-- SECTION 2: REGULAR VS IRREGULAR CANDIDATE PERFORMANCE -->
                <div class="report-section">
                    <div class="section-title">
                        <i class="bi bi-people-fill text-primary"></i> 2. Regular vs. Irregular Candidates Breakdown
                    </div>
                    <div class="section-desc">Compares fresh regular batch candidates against candidates appearing for improvement or irregular examination.</div>
                    <table class="table table-bordered table-sm text-center">
                        <thead>
                            <tr>
                                <th class="text-start ps-3">Candidate Category</th>
                                <th>Total Enrolled</th>
                                <th>Attended</th>
                                <th>Absent</th>
                                <th>Passed</th>
                                <th>Failed</th>
                                <th>Pass Rate (%)</th>
                                <th>GPA 5.0 (With 4th)</th>
                                <th>Golden A+ (Without 4th)</th>
                                <th>1-Sub Fail</th>
                                <th>2-Sub Fail</th>
                                <th>3+ Sub Fail</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td class="text-start ps-3 fw-bold text-primary">
                                    <span class="badge bg-primary-subtle text-primary border border-primary-subtle me-1">REGULAR</span> Fresh Batch
                                </td>
                                <td>${reg.total || 0}</td>
                                <td class="fw-semibold">${reg.appeared || 0}</td>
                                <td class="text-secondary fw-bold">${reg.absent || 0}</td>
                                <td class="fw-bold text-success">${reg.passed || 0}</td>
                                <td class="fw-bold text-danger">${reg.failed || 0}</td>
                                <td class="fw-bold text-success">${reg.pass_rate || 0}%</td>
                                <td class="fw-bold table-highlight-gold">${reg.gpa5_count_with_4th ?? reg.gpa5_count ?? 0}</td>
                                <td class="fw-bold" style="color: #6d28d9;">${reg.gpa5_count_without_4th ?? 0}</td>
                                <td class="fw-semibold">${reg.one_fail_count || 0}</td>
                                <td class="text-danger fw-semibold">${reg.two_fail_count || 0}</td>
                                <td class="text-muted">${reg.multi_fail_count || 0}</td>
                            </tr>
                            <tr>
                                <td class="text-start ps-3 fw-bold text-warning">
                                    <span class="badge bg-warning-subtle text-dark border border-warning-subtle me-1">IRREGULAR</span> Improvement / Retake
                                </td>
                                <td>${irreg.total || 0}</td>
                                <td class="fw-semibold">${irreg.appeared || 0}</td>
                                <td class="text-secondary fw-bold">${irreg.absent || 0}</td>
                                <td class="fw-bold text-success">${irreg.passed || 0}</td>
                                <td class="fw-bold text-danger">${irreg.failed || 0}</td>
                                <td class="fw-bold text-success">${irreg.pass_rate || 0}%</td>
                                <td class="fw-bold table-highlight-gold">${irreg.gpa5_count_with_4th ?? irreg.gpa5_count ?? 0}</td>
                                <td class="fw-bold" style="color: #6d28d9;">${irreg.gpa5_count_without_4th ?? 0}</td>
                                <td class="fw-semibold">${irreg.one_fail_count || 0}</td>
                                <td class="text-danger fw-semibold">${irreg.two_fail_count || 0}</td>
                                <td class="text-muted">${irreg.multi_fail_count || 0}</td>
                            </tr>
                        </tbody>
                    </table>

                    ${groupType.length > 0 ? `
                    <div class="fw-bold text-secondary mb-1" style="font-size: 10px;">Group-Wise Regular vs. Irregular Detailed Matrix</div>
                    <table class="table table-bordered table-sm text-center mb-3">
                        <thead>
                            <tr>
                                <th class="text-start ps-2">Academic Group</th>
                                <th>Category</th>
                                <th>Enrolled</th>
                                <th>Attended</th>
                                <th>Absent</th>
                                <th>Passed</th>
                                <th>Failed</th>
                                <th>Pass Rate (%)</th>
                                <th>GPA 5 (With 4th)</th>
                                <th>Golden A+ (Without 4th)</th>
                                <th>1-Sub Fail</th>
                                <th>2-Sub Fail</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${groupType.map(r => `
                                <tr>
                                    <td class="text-start ps-2 fw-bold">${r.group}</td>
                                    <td><span class="badge ${r.type === 'REGULAR' ? 'bg-primary-subtle text-primary border border-primary-subtle' : 'bg-warning-subtle text-dark border border-warning-subtle'}">${r.type}</span></td>
                                    <td>${r.total}</td>
                                    <td>${r.appeared}</td>
                                    <td class="text-secondary">${r.absent || 0}</td>
                                    <td class="text-success fw-bold">${r.passed}</td>
                                    <td class="text-danger fw-bold">${r.failed}</td>
                                    <td class="fw-bold">${r.pass_rate}%</td>
                                    <td class="table-highlight-gold">${r.gpa5_count_with_4th ?? r.gpa5_count ?? 0}</td>
                                    <td style="color: #6d28d9; font-weight: 700;">${r.gpa5_count_without_4th ?? 0}</td>
                                    <td>${r.one_fail_count}</td>
                                    <td>${r.two_fail_count}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>` : ""}
                </div>

                <!-- SECTION 3: GROUP PERFORMANCE COMPARISON -->
                <div class="report-section">
                    <div class="section-title">
                        <i class="bi bi-grid-3x3-gap-fill text-primary"></i> 3. Group Performance Comparison
                    </div>
                    <div class="section-desc">Academic achievement comparison across Science, Humanities, and Business Studies groups.</div>
                    <table class="table table-bordered table-sm text-center">
                        <thead>
                            <tr>
                                <th class="text-start ps-3">Academic Group</th>
                                <th>Total Enrolled</th>
                                <th>Attended</th>
                                <th>Absent</th>
                                <th>Passed</th>
                                <th>Failed</th>
                                <th>Pass Rate (%)</th>
                                <th>GPA 5.0 (With 4th)</th>
                                <th>Golden A+ (Without 4th)</th>
                                <th>1-Sub Fail</th>
                                <th>2-Sub Fail</th>
                                <th>3+ Sub Fail</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${Object.values(groupsSummary).map(grp => `
                                <tr>
                                    <td class="text-start ps-3 fw-bold text-dark">${grp.group}</td>
                                    <td>${grp.total}</td>
                                    <td class="fw-semibold">${grp.appeared}</td>
                                    <td class="text-secondary fw-bold">${grp.absent || 0}</td>
                                    <td class="text-success fw-bold">${grp.passed}</td>
                                    <td class="text-danger fw-bold">${grp.failed}</td>
                                    <td class="fw-bold text-success">${grp.pass_rate}%</td>
                                    <td class="fw-bold table-highlight-gold">${grp.gpa5_count_with_4th ?? grp.gpa5_count ?? 0}</td>
                                    <td class="fw-bold" style="color: #6d28d9;">${grp.gpa5_count_without_4th ?? 0}</td>
                                    <td>${grp.one_fail_count}</td>
                                    <td>${grp.two_fail_count}</td>
                                    <td class="text-muted">${grp.multi_fail_count}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>

                <!-- SECTION 4: SUBJECT WISE GRADE ANALYSIS -->
                <div class="report-section">
                    <div class="section-title">
                        <i class="bi bi-book-half text-primary"></i> 4. SUBJECT WISE GRADE ANALYSIS
                    </div>
                    <div class="section-desc">Complete breakdown of every subject showing student attendance, pass rate, and full grade distribution (A+ down to F).</div>
                    <table class="table table-bordered table-sm text-center">
                        <thead>
                            <tr>
                                <th class="text-start ps-2">Subject Name</th>
                                <th>Group</th>
                                <th>Type</th>
                                <th>Attended</th>
                                <th>Passed</th>
                                <th>Failed</th>
                                <th>Pass %</th>
                                <th class="table-highlight-gold">A+</th>
                                <th>A</th>
                                <th>A-</th>
                                <th>B</th>
                                <th>C</th>
                                <th>D</th>
                                <th class="table-highlight-red">F (Fail)</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(analysisData.subject_analysis || []).map(item => `
                                <tr>
                                    <td class="text-start ps-2 fw-bold text-dark">${item.subject}</td>
                                    <td><span class="badge bg-light text-secondary border">${item.group}</span></td>
                                    <td><span class="badge bg-light text-dark border">${item.type || "ALL"}</span></td>
                                    <td>${item.appeared}</td>
                                    <td class="text-success fw-bold">${item.passed}</td>
                                    <td class="${item.failed > 0 ? 'text-danger fw-bold' : 'text-muted'}">${item.failed}</td>
                                    <td class="fw-bold ${item.pass_rate >= 80 ? 'text-success' : (item.pass_rate < 50 ? 'text-danger' : 'text-primary')}">${item.pass_rate}%</td>
                                    <td class="fw-bold text-warning">${item.grades["A+"] || 0}</td>
                                    <td>${item.grades["A"] || 0}</td>
                                    <td>${item.grades["A-"] || 0}</td>
                                    <td>${item.grades["B"] || 0}</td>
                                    <td>${item.grades["C"] || 0}</td>
                                    <td>${item.grades["D"] || 0}</td>
                                    <td class="${item.grades["F"] > 0 ? 'table-highlight-red' : 'text-muted'}">${item.grades["F"] || 0}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>

                <!-- SECTION 5: ONE SUBJECT FAIL STUDENTS -->
                <div class="report-section">
                    <div class="section-title">
                        <i class="bi bi-exclamation-octagon-fill text-warning"></i> 5. One Subject Fail - Supplementary Eligible Students (Total: ${oneFail.total || 0})
                    </div>
                    <div class="section-desc">These students failed in only ONE subject. Remedial coaching and supplementary examination will allow them to pass easily.</div>
                    ${(oneFail.students || []).length > 0 ? `
                    <table class="table table-bordered table-sm">
                        <thead>
                            <tr>
                                <th style="width: 80px;">Roll No</th>
                                <th style="width: 90px;">Reg No</th>
                                <th>Student Name</th>
                                <th style="width: 90px;">Group</th>
                                <th style="width: 80px;">Type</th>
                                <th style="width: 140px;">Failed Subject</th>
                                <th>Passed Subject Grades</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(oneFail.students || []).map(st => `
                                <tr>
                                    <td class="fw-bold text-center">${st.Roll}</td>
                                    <td class="text-center font-monospace">${st.Registration || "-"}</td>
                                    <td class="fw-semibold text-dark">${st.Name}</td>
                                    <td class="text-center">${st.Group}</td>
                                    <td class="text-center"><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border border-primary-subtle' : 'bg-warning-subtle text-dark border border-warning-subtle'}">${st.Type || "REGULAR"}</span></td>
                                    <td class="fw-bold text-danger table-highlight-red text-center">${st.Failed_Subject}</td>
                                    <td class="small text-muted">${st.Passed_Grades_Summary}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>` : `<p class="text-muted fst-italic">No students failed in only one subject.</p>`}
                </div>

                <!-- SECTION 6: TWO SUBJECT FAIL STUDENTS -->
                <div class="report-section">
                    <div class="section-title">
                        <i class="bi bi-exclamation-triangle-fill text-danger"></i> 6. Two Subject Fail Students (Total: ${twoFail.total || 0})
                    </div>
                    <div class="section-desc">Students who failed in exactly two subjects and require intervention in these specific subjects.</div>
                    ${(twoFail.students || []).length > 0 ? `
                    <table class="table table-bordered table-sm">
                        <thead>
                            <tr>
                                <th style="width: 80px;">Roll No</th>
                                <th style="width: 90px;">Reg No</th>
                                <th>Student Name</th>
                                <th style="width: 90px;">Group</th>
                                <th style="width: 80px;">Type</th>
                                <th style="width: 180px;">Failed Subjects</th>
                                <th>Passed Subject Grades</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(twoFail.students || []).map(st => `
                                <tr>
                                    <td class="fw-bold text-center">${st.Roll}</td>
                                    <td class="text-center font-monospace">${st.Registration || "-"}</td>
                                    <td class="fw-semibold text-dark">${st.Name}</td>
                                    <td class="text-center">${st.Group}</td>
                                    <td class="text-center"><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border border-primary-subtle' : 'bg-warning-subtle text-dark border border-warning-subtle'}">${st.Type || "REGULAR"}</span></td>
                                    <td class="fw-bold text-danger table-highlight-red text-center">${(st.Failed_Subjects || []).join(" & ")}</td>
                                    <td class="small text-muted">${st.Passed_Grades_Summary}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>` : `<p class="text-muted fst-italic">No students failed in two subjects.</p>`}
                </div>

                <!-- SECTION 7: 3+ SUBJECT FAIL STUDENTS (IF ANY) -->
                ${(mFail.students || []).length > 0 ? `
                <div class="report-section">
                    <div class="section-title">
                        <i class="bi bi-x-circle-fill text-secondary"></i> 7. Multi-Subject Fail Students (3+ Subjects) - Total: ${mFail.total || 0}
                    </div>
                    <table class="table table-bordered table-sm">
                        <thead>
                            <tr>
                                <th style="width: 80px;">Roll No</th>
                                <th>Student Name</th>
                                <th style="width: 90px;">Group</th>
                                <th style="width: 80px;">Type</th>
                                <th style="width: 80px;">Fails Count</th>
                                <th>Failed Subjects</th>
                                <th>Passed Subject Grades</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(mFail.students || []).map(st => `
                                <tr>
                                    <td class="fw-bold text-center">${st.Roll}</td>
                                    <td class="fw-semibold">${st.Name}</td>
                                    <td class="text-center">${st.Group}</td>
                                    <td class="text-center">${st.Type || "REGULAR"}</td>
                                    <td class="fw-bold text-danger text-center">${st.Failed_Count}</td>
                                    <td class="text-danger fw-semibold">${(st.Failed_Subjects || []).join(", ")}</td>
                                    <td class="small text-muted">${st.Passed_Grades_Summary}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>` : ""}

                <!-- SECTION 8: ABSENT CANDIDATES ANALYSIS & ROSTER -->
                <div class="report-section">
                    <div class="section-title">
                        <i class="bi bi-person-x-fill text-secondary"></i> 8. Absent Candidates Roster (Did Not Sit for Exam) - Total: ${absentData.total_absent || summary.absent || 0}
                    </div>
                    <div class="section-desc">Registered candidates who were absent during the examination.</div>
                    
                    <div class="p-2 mb-3 bg-light border rounded d-flex justify-content-between align-items-center flex-wrap gap-2" style="font-size: 10.5px;">
                        <div>
                            <strong>Total Absent:</strong> <span class="badge bg-secondary">${absentData.total_absent || summary.absent || 0}</span> (${absentData.absent_rate || summary.absent_rate || 0}% of Enrolled) &bull; 
                            <strong>Regular Absent:</strong> ${absentData.regular_absent || 0} &bull; 
                            <strong>Irregular Absent:</strong> ${absentData.irregular_absent || 0}
                        </div>
                        <div>
                            <strong>Total Appeared (Attended Exam):</strong> <span class="badge bg-success">${absentData.appeared || summary.appeared || 0}</span>
                        </div>
                    </div>

                    ${(absentData.students || []).length > 0 ? `
                    <table class="table table-bordered table-sm">
                        <thead>
                            <tr>
                                <th style="width: 80px;">Roll No</th>
                                <th style="width: 90px;">Reg No</th>
                                <th>Student Name</th>
                                <th>Father's Name</th>
                                <th style="width: 90px;">Group</th>
                                <th style="width: 80px;">Type</th>
                                <th>Institute</th>
                                <th style="width: 75px;">Status</th>
                                <th>Remarks</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(absentData.students || []).map(st => `
                                <tr>
                                    <td class="fw-bold text-center">${st.Roll || "-"}</td>
                                    <td class="text-center font-monospace">${st.Registration || "-"}</td>
                                    <td class="fw-semibold text-dark">${st.Name || "N/A"}</td>
                                    <td class="text-muted small">${st["Father Name"] || st.Father_Name || "-"}</td>
                                    <td class="text-center">${st.Group || "GENERAL"}</td>
                                    <td class="text-center"><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border border-primary-subtle' : 'bg-warning-subtle text-dark border border-warning-subtle'}">${st.Type || "REGULAR"}</span></td>
                                    <td class="small text-truncate" style="max-width: 140px;">${st.Institute || "-"}</td>
                                    <td class="fw-bold text-center text-secondary"><span class="badge bg-secondary">ABSENT</span></td>
                                    <td class="small text-muted">${st.Remarks || "Did not sit for examination"}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>` : `<p class="text-muted fst-italic">No absent candidates recorded in this examination result.</p>`}
                </div>

                <!-- Footer Signoff Section -->
                <div class="row pt-4 mt-4 border-top text-center text-muted" style="font-size: 10px;">
                    <div class="col-4">
                        <div class="border-top border-dark mx-4 pt-1 fw-bold">Prepared by (Staff)</div>
                    </div>
                    <div class="col-4">
                        <div class="border-top border-dark mx-4 pt-1 fw-bold">Checked by (Tabulator)</div>
                    </div>
                    <div class="col-4">
                        <div class="border-top border-dark mx-4 pt-1 fw-bold">Principal / Head of Institution</div>
                    </div>
                </div>

                <!-- Bottom Print Button (Hidden on paper) -->
                <div class="no-print text-center my-4 pt-3 border-top">
                    <button class="btn btn-primary btn-lg px-5 shadow-sm fw-bold" onclick="window.print()">
                        <i class="bi bi-printer-fill me-2"></i> Print / Save as PDF
                    </button>
                    <button class="btn btn-outline-secondary btn-lg px-4 ms-2" onclick="window.close()">
                        Close Window
                    </button>
                </div>
            </div>
        </body>
        </html>
        `;
        printWin.document.write(html);
        printWin.document.close();
    }
});
