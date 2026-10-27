const time = [2004, 2026];

const irs_url = "tabels/irs/Tabelas_RF_Continente_${}.xlsx";

const sidebarDiv = document.querySelector("#sidebar");
const viewer = document.querySelector("#xlsx-viewer");
const titlebar = document.querySelector("#window-titlebar");
const windowTitle = document.querySelector("#window-title");
const closeButton = document.querySelector("#close-button");
const headerDiv = document.querySelector("#header");
const sheetContainer = document.querySelector("#sheet-container");

let currentRequest = 0;
let selectedYear = -1;

// Create a button for every year
for (let i = time[1]; i >= time[0]; --i) {
    const button = document.createElement("div");

    button.classList.add("year-button");
    button.id = `year-button-${(i - time[1]) * -1}`
    button.textContent = i;
    button.tabIndex = 0;

    button.onclick = () => openxlsxFile(
        irs_url.replace("${}", i),
        i
    );

    sidebarDiv.appendChild(button);
}

// Keep sidebar and window visibility in sync.
function setWindowOpen(isOpen) {
    viewer.classList.toggle("open", isOpen);
    document.body.classList.toggle("window-open", isOpen);
    sidebarDiv.setAttribute("aria-hidden", String(isOpen));
}

        // Open the window and load the selected year.
        async function openxlsxFile(file, year) {
            setWindowOpen(true);

            windowTitle.textContent = `IRS - ${year}`;
            const requestId = ++currentRequest;
            headerDiv.textContent = `Loading ${year}...`;
            sheetContainer.innerHTML = "";

            try {
                const response = await fetch(file);
                if (!response.ok) throw new Error(`HTTP error: ${response.status}`);
                const arrayBuffer = await response.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: "array" });
                if (requestId !== currentRequest) return;

                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                sheetContainer.innerHTML = XLSX.utils.sheet_to_html(worksheet, { editable: false });
                headerDiv.textContent = `${year} — ${sheetName}`;
                applySheetOrganizer();
                updateCalculatorTableInfo();
                calculateSalary();
            } catch (error) {
                if (requestId !== currentRequest) return;
                console.error("Failed to load Excel file:", error);
                headerDiv.textContent = `Unable to load ${year}`;
                sheetContainer.textContent =
                    "The file could not be loaded. Check that it exists and that the server is serving it correctly.";
            }
        }

        closeButton.addEventListener("click", (event) => {
            event.stopPropagation();
            closeWindow();
        });

        function closeWindow() {
            setWindowOpen(false);
            ++currentRequest;
        }

        // The viewer is fixed fullscreen and cannot be dragged or resized.
        let scrolling = false;
        let scrollDirection = 0;
        let scrollIsHorizontal = false;
        let scrollAnimation = null;
        let scrollPlace = [0, 0]

        document.addEventListener('keydown', function(e) {
            if (e.key.toLowerCase() === "b" && !e.repeat &&
                !e.target.closest('input, textarea, select, [contenteditable="true"]')) {
                e.preventDefault();
                scrollPlace = [0, 0];
                if (viewer.classList.contains("open")) {
                    closeWindow();
                    if (selectedYear >= 0) moveSidebar(null);
                } else {
                    const selectedButton = document.querySelector(`#year-button-${selectedYear >= 0 ? selectedYear : 0}`);
                    if (selectedButton) selectedButton.click();
                }
            } else if (e.key === "PageUp") {
                e.preventDefault();
                if (viewer.classList.contains("open")) {
                    if (scrolling && scrollDirection === -1) {
                        stopScrolling();
                        scrollDirection = 0;
                    } else {
                        // Start scrolling, or reverse direction.
                        stopScrolling();
                        scrollDirection = -1;
                        startScrolling();
                    }
                } else {
                    moveSidebar(false)
                }
            } else if (e.key === "PageDown") {
                e.preventDefault();
                if (viewer.classList.contains("open")) {
                    if (scrolling && scrollDirection === 1) {
                        stopScrolling();
                        scrollDirection = 0;
                    } else {
                        // Start scrolling, or reverse direction.
                        stopScrolling();
                        scrollDirection = 1;
                        startScrolling();
                    }

                } else {
                    moveSidebar(true)
                }
            } else if (e.key === "Shift") {
                scrollIsHorizontal = true;
                scrollDirection = 0;

            } else if (e.key === "Escape") {
                scrollIsHorizontal = false;
                scrollDirection = 0;
            }
        });

        function startScrolling() {
            scrolling = true;
            let lastTime = null;
            const speed = 200; // pixels per second
            const sheet = document.querySelector("#sheet-container");

            function frame(timestamp) {
                if (!scrolling) return;

                if (lastTime !== null) {
                    const distance = speed * (timestamp - lastTime) / 1000;

                    if (scrollIsHorizontal) {
                        scrollPlace[0] += distance * scrollDirection;

                        const maxScrollX = sheet.scrollWidth - sheet.clientWidth;
                        scrollPlace[0] = Math.max(0, Math.min(scrollPlace[0], maxScrollX));

                        if (
                            (scrollDirection > 0 && scrollPlace[0] >= maxScrollX) ||
                            (scrollDirection < 0 && scrollPlace[0] <= 0)
                        ) {
                            sheet.scrollTo(scrollPlace[0], scrollPlace[1]);
                            stopScrolling();
                            return;
                        }
                    } else {
                        scrollPlace[1] += distance * scrollDirection;

                        const maxScrollY = sheet.scrollHeight - sheet.clientHeight;
                        scrollPlace[1] = Math.max(0, Math.min(scrollPlace[1], maxScrollY));

                        if (
                            (scrollDirection > 0 && scrollPlace[1] >= maxScrollY) ||
                            (scrollDirection < 0 && scrollPlace[1] <= 0)
                        ) {
                            sheet.scrollTo(scrollPlace[0], scrollPlace[1]);
                            stopScrolling();
                            return;
                        }
                    }
                    sheet.scrollTo(scrollPlace[0], scrollPlace[1]);
                }

                lastTime = timestamp;
                scrollAnimation = requestAnimationFrame(frame);
            }

    scrollAnimation = requestAnimationFrame(frame);
}
        function stopScrolling() {
            scrolling = false;

            if (scrollAnimation !== null) {
                cancelAnimationFrame(scrollAnimation);
                scrollAnimation = null;
            }
        }

        function moveSidebar(up = true) {
            if (up !== null) {
                let totalYears = document.querySelectorAll(".year-button").length;
                selectedYear = up ? (selectedYear + 1) % totalYears : (selectedYear - 1 + totalYears) % totalYears;
            }

            document.querySelector(`#year-button-${selectedYear}`).focus();

            console.log("Selected year:", selectedYear);
        }

        // ==========================================
        // Sheet Organizer Logic
        // ==========================================
        const organizerState = {
            marital: null, // "nao_casado" | "casado"
            deps: null,    // "sem_dependentes" | "com_dependentes"
            holders: null, // "2_titulares" | "unico_titular"
            disab: null    // true | false
        };

        const organizerStatusEl = document.querySelector("#organizer-status");
        const organizerResetBtn = document.querySelector("#organizer-reset");
        const organizerButtons = document.querySelectorAll(".organizer-btn");

        organizerButtons.forEach((btn) => {
            btn.addEventListener("click", () => {
                const group = btn.closest(".organizer-group");
                const question = group.dataset.question;
                const val = btn.dataset.value;

                // Toggle behavior: clicking an active option unselects it
                if (btn.classList.contains("active")) {
                    btn.classList.remove("active");
                    if (question === "marital") organizerState.marital = null;
                    else if (question === "deps") organizerState.deps = null;
                    else if (question === "holders") organizerState.holders = null;
                    else if (question === "disab") organizerState.disab = null;
                } else {
                    group.querySelectorAll(".organizer-btn").forEach((b) => b.classList.remove("active"));
                    btn.classList.add("active");

                    if (question === "marital") organizerState.marital = val;
                    else if (question === "deps") {
                        organizerState.deps = val;
                        if (val === "sem_dependentes" && calcDepsInput) {
                            calcDepsInput.value = 0;
                        } else if (val === "com_dependentes" && calcDepsInput && parseInt(calcDepsInput.value, 10) === 0) {
                            calcDepsInput.value = 1;
                        }
                    }
                    else if (question === "holders") organizerState.holders = val;
                    else if (question === "disab") {
                        organizerState.disab = (val === "com_deficiencia");
                    }
                }

                applySheetOrganizer();
            });
        });

        if (organizerResetBtn) {
            organizerResetBtn.addEventListener("click", () => {
                organizerState.marital = null;
                organizerState.deps = null;
                organizerState.holders = null;
                organizerState.disab = null;

                organizerButtons.forEach((b) => b.classList.remove("active"));
                applySheetOrganizer();
            });
        }

        // Text checking in lower caps; ensures "casado" has no preceding "não" / "nao"
        function hasNaoCasado(lowerText) {
            return /(?:n[ãa]o)[\s,.-]+casado/i.test(lowerText);
        }

        function hasCasadoWithoutNao(lowerText) {
            // Check if there is an occurrence of "casado" NOT preceded by "não" / "nao"
            // For example:
            // "não casado ou casado dois titulares" -> true (2nd occurrence is not preceded by não)
            // "não casado com um ou mais dependentes" -> false (only preceded by não)
            // "casado, único titular" -> true
            const regex = /(?:(n[ãa]o)[\s,.-]+)?casado/gi;
            let match;
            while ((match = regex.exec(lowerText)) !== null) {
                if (!match[1]) {
                    return true;
                }
            }
            return false;
        }

        function hasSemDependentes(lowerText) {
            return /sem\s+dependentes?/i.test(lowerText);
        }

        function hasComDependentes(lowerText) {
            return /(?:com\s+(?:um\s+ou\s+mais\s+)?|um\s+ou\s+mais\s+)dependentes?/i.test(lowerText);
        }

        function hasDoisTitulares(lowerText) {
            return /(?:2|dois)\s*titulares?/i.test(lowerText);
        }

        function hasUnicoTitular(lowerText) {
            return /[úu]nico\s*titula(?:r)?/i.test(lowerText);
        }

        function hasDeficiencia(lowerText) {
            return /pessoa\s+com\s+defici[êe]ncia|deficiente/i.test(lowerText);
        }

        // Separate the search by "ou"
        // e.g., "Não casado sem dependentes ou casado 2 titulares" -> ["Não casado sem dependentes", "casado 2 titulares"]
        function splitByOu(text) {
            // Match " ou " where "ou" is not part of "um ou mais" or digits like "5 ou mais"
            const pattern = /(?<!\bum)(?<!\b\d)\s+\bou\b\s+(?!mais\b)/gi;
            const parts = text.split(pattern);
            if (parts.length <= 1) {
                return [text.trim()];
            }

            const part1 = parts[0].trim();
            const part2 = parts[1].trim();

            // When part2 has trailing qualifiers (e.g. "sem dependentes - Pessoa com deficiência")
            // that apply to both sides (e.g. "Não casado ou casado dois titulares sem dependentes - Pessoa com deficiência"),
            // propagate those qualifiers to part1 as well.
            const qualifierMatch = part2.match(/(sem\s+dependentes?.*|com\s+defici.*|deficiente.*|-.*)/i);
            if (qualifierMatch && !/(?:sem\s+dependentes?|defici)/i.test(part1)) {
                const augmentedPart1 = `${part1} ${qualifierMatch[0]}`.trim();
                return [augmentedPart1, part2];
            }

            return [part1, part2];
        }

        function applySheetOrganizer() {
            // Clear any existing highlights in the sheet
            sheetContainer.querySelectorAll(".organizer-highlight").forEach((el) => {
                el.classList.remove("organizer-highlight");
            });

            const rows = Array.from(sheetContainer.querySelectorAll("tr"));
            if (rows.length === 0) {
                if (organizerStatusEl) {
                    organizerStatusEl.className = "organizer-status";
                    organizerStatusEl.textContent = "Selecione as opções para focar na tabela.";
                }
                return;
            }

            const { marital, deps, holders, disab } = organizerState;
            const anySelected = marital !== null || deps !== null || holders !== null || disab !== null;

            if (!anySelected) {
                if (organizerStatusEl) {
                    organizerStatusEl.className = "organizer-status";
                    organizerStatusEl.textContent = "Selecione as opções para focar na tabela.";
                }
                return;
            }

            let bestRow = null;
            let bestBranch = "";
            let bestScore = -1;

            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const rawText = (row.textContent || "").trim();
                if (!rawText) continue;

                // Check in lower caps
                const fullLower = rawText.toLowerCase();

                // Target header-like rows
                if (!/(?:casado|titula|defici|tabela)/i.test(fullLower)) continue;
                if (fullLower.length > 250) continue;
                if (/f[óo]rmula:|remunera[çc][ãa]o\s+mensal/i.test(fullLower)) continue;

                // Separate the search by "ou"
                const branches = splitByOu(rawText);

                let rowBestScore = -1;
                let rowBestBranch = "";

                for (const branch of branches) {
                    const lower = branch.toLowerCase();

                    const branchHasNaoCasado = hasNaoCasado(lower);
                    const branchHasCasado = hasCasadoWithoutNao(lower);
                    const branchHas2Tit = hasDoisTitulares(lower);
                    const branchHas1Tit = hasUnicoTitular(lower);
                    const branchHasSemDep = hasSemDependentes(lower);
                    const branchHasComDep = hasComDependentes(lower);
                    const branchHasDisab = hasDeficiencia(lower);

                    let score = 0;

                    // Disability check
                    if (disab === true) {
                        if (!branchHasDisab) continue;
                        score += 20;
                    } else if (disab === false) {
                        if (branchHasDisab) continue;
                        score += 5;
                    }

                    // Marital status check
                    if (marital === "nao_casado") {
                        if (!branchHasNaoCasado) continue;
                        score += 20;
                    } else if (marital === "casado") {
                        if (!branchHasCasado) continue;
                        score += 20;
                    }

                    // Holders check
                    if (holders === "2_titulares") {
                        if (branchHas2Tit) {
                            score += 15;
                        } else if (branchHas1Tit) {
                            continue;
                        }
                    } else if (holders === "unico_titular") {
                        if (branchHas1Tit) {
                            score += 15;
                        } else if (branchHas2Tit) {
                            continue;
                        }
                    }

                    // Dependents check
                    if (deps === "sem_dependentes") {
                        if (branchHasSemDep) {
                            score += 15;
                        } else if (branchHasComDep) {
                            continue;
                        }
                    } else if (deps === "com_dependentes") {
                        if (branchHasComDep) {
                            score += 15;
                        } else if (branchHasSemDep) {
                            continue;
                        }
                    }

                    if (score > rowBestScore) {
                        rowBestScore = score;
                        rowBestBranch = branch;
                    }
                }

                if (rowBestScore > bestScore) {
                    bestScore = rowBestScore;
                    bestRow = row;
                    bestBranch = rowBestBranch;
                }
            }

            if (bestRow && bestScore > 0) {
                let scrollTarget = bestRow;
                // If row is preceded by "Tabela I - ...", include that row in highlight and scroll
                const prevRow = bestRow.previousElementSibling;
                if (prevRow && /tabela\s+[ivx]+/i.test(prevRow.textContent || "")) {
                    scrollTarget = prevRow;
                    prevRow.classList.add("organizer-highlight");
                }
                bestRow.classList.add("organizer-highlight");

                // Focus and scroll to that part of the sheet
                const containerRect = sheetContainer.getBoundingClientRect();
                const targetRect = scrollTarget.getBoundingClientRect();
                sheetContainer.scrollTo({
                    top: sheetContainer.scrollTop + (targetRect.top - containerRect.top) - 10,
                    behavior: "smooth"
                });

                if (organizerStatusEl) {
                    organizerStatusEl.className = "organizer-status has-match";
                    const displayMatch = (bestBranch || bestRow.textContent).replace(/\s+/g, " ").trim();
                    organizerStatusEl.textContent = `Focado em: ${displayMatch.slice(0, 70)}`;
                }

                currentActiveTable = {
                    row: bestRow,
                    name: (bestBranch || bestRow.textContent).replace(/\s+/g, " ").trim()
                };
                updateCalculatorTableInfo();
                calculateSalary();
            } else {
                if (organizerStatusEl) {
                    organizerStatusEl.className = "organizer-status";
                    organizerStatusEl.textContent = "Nenhuma tabela encontrada para os critérios selecionados.";
                }

                currentActiveTable = null;
                updateCalculatorTableInfo();
                calculateSalary();
            }
        }

        // ==========================================
        // View Changer (Tabela vs Calculadora)
        // ==========================================
        const btnShowOrganizer = document.querySelector("#btn-show-organizer");
        const btnShowCalculator = document.querySelector("#btn-show-calculator");
        const tableOrganizerView = document.querySelector("#table-organizer");
        const tableCalculatorView = document.querySelector("#table-calculator");

        if (btnShowOrganizer && btnShowCalculator) {
            btnShowOrganizer.addEventListener("click", () => {
                btnShowOrganizer.classList.add("active");
                btnShowCalculator.classList.remove("active");
                if (tableOrganizerView) tableOrganizerView.style.display = "flex";
                if (tableCalculatorView) tableCalculatorView.style.display = "none";
            });

            btnShowCalculator.addEventListener("click", () => {
                btnShowCalculator.classList.add("active");
                btnShowOrganizer.classList.remove("active");
                if (tableOrganizerView) tableOrganizerView.style.display = "none";
                if (tableCalculatorView) tableCalculatorView.style.display = "flex";
                updateCalculatorTableInfo();
                calculateSalary();
            });
        }

        // ==========================================
        // Salary & Deductions Calculator Logic
        // ==========================================
        let currentActiveTable = null; // { row: Element, name: string }

        const calcGrossInput = document.querySelector("#calc-gross");
        const calcDepsInput = document.querySelector("#calc-deps");
        const calcDepMinusBtn = document.querySelector("#calc-dep-minus");
        const calcDepPlusBtn = document.querySelector("#calc-dep-plus");
        const calcUnionCheckbox = document.querySelector("#calc-sindicato");
        const calcTableNameEl = document.querySelector("#calc-table-name");

        const resGrossEl = document.querySelector("#res-gross");
        const resTsuEl = document.querySelector("#res-tsu");
        const resIrsEl = document.querySelector("#res-irs");
        const resIrsFormulaTextEl = document.querySelector("#res-irs-formula-text");
        const resUnionEl = document.querySelector("#res-union");
        const calcRowUnionEl = document.querySelector("#calc-row-union");
        const resTotalDeductionsEl = document.querySelector("#res-total-deductions");
        const resNetEl = document.querySelector("#res-net");

        function formatEUR(val) {
            return val.toLocaleString("pt-PT", {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }) + " €";
        }

        function getDefaultTable() {
            const rows = Array.from(sheetContainer.querySelectorAll("tr"));
            for (const r of rows) {
                const t = (r.textContent || "").trim();
                if (/(?:tabela\s+[ivx]+|n[ãa]o\s+casado|casado)/i.test(t) &&
                    !/(?:at[ée]|superior|remunera[çc][ãa]o)/i.test(t)) {
                    return { row: r, name: t.replace(/\s+/g, " ").trim() };
                }
            }
            return null;
        }

        function updateCalculatorTableInfo() {
            const active = currentActiveTable || getDefaultTable();
            if (calcTableNameEl) {
                if (active && active.name) {
                    calcTableNameEl.textContent = active.name.slice(0, 65);
                } else {
                    calcTableNameEl.textContent = "Nenhuma tabela carregada";
                }
            }
        }

        function parseNum(str) {
            if (str === null || str === undefined || str === "") return 0;

            // SheetJS renders the source workbook using its number formats.  That
            // means a rate can be rendered as "47.17%" and a monetary value as
            // either "1,272.31" or "1.272,31".  These must not be parsed with
            // parseFloat directly: 47.17% is 0.4717, and both currency examples
            // represent 1272.31.
            let value = String(str).trim();
            const isPercent = value.includes("%");
            value = value.replace(/[^0-9,.-]/g, "");

            const lastComma = value.lastIndexOf(",");
            const lastDot = value.lastIndexOf(".");
            const decimalIndex = Math.max(lastComma, lastDot);

            if (decimalIndex !== -1) {
                const hasBothSeparators = lastComma !== -1 && lastDot !== -1;
                const fraction = value.slice(decimalIndex + 1);
                // Excel uses a thousands separator without decimals for the
                // bracket limits (for example, "1,819").
                const isThousandsOnly = !hasBothSeparators && fraction.length === 3 &&
                    !/^[-+]?0[.,]/.test(value);

                if (isThousandsOnly) {
                    value = value.replace(/[.,]/g, "");
                } else {
                    const integerPart = value.slice(0, decimalIndex).replace(/[.,]/g, "");
                    const decimalPart = fraction.replace(/[.,]/g, "");
                    value = `${integerPart}.${decimalPart}`;
                }
            }

            const parsed = parseFloat(value) || 0;
            return isPercent ? parsed / 100 : parsed;
        }

        function isBracketRow(rowEl) {
            if (!rowEl) return false;
            const text = (rowEl.textContent || "").trim().toLowerCase();
            if (!text) return false;
            return text.includes("até") || text.includes("ate") || text.includes("superior");
        }

        function getBracketRowsForTable(tableRow) {
            if (!tableRow) return [];
            const brackets = [];
            let curr = tableRow;

            // Step 1: Advance to the first bracket row belonging to this table
            while (curr) {
                if (isBracketRow(curr)) {
                    break;
                }
                // If we hit another table header before finding any bracket row, stop
                if (curr !== tableRow) {
                    const text = (curr.textContent || "").trim();
                    if (/tabela\s+[ivx]+/i.test(text)) {
                        break;
                    }
                }
                curr = curr.nextElementSibling;
            }

            // Step 2: Collect all bracket rows for this table
            while (curr) {
                if (isBracketRow(curr)) {
                    const cells = Array.from(curr.querySelectorAll("td, th"))
                        .map((c) => c.textContent.trim())
                        .filter(Boolean);
                    if (cells.length > 0) {
                        brackets.push(cells);
                    }
                } else {
                    const text = (curr.textContent || "").trim();
                    // Stop when reaching a non-empty, non-bracket row (e.g. formula row or next table)
                    if (text.length > 0) {
                        break;
                    }
                }
                curr = curr.nextElementSibling;
            }
            return brackets;
        }

        function computeIRS(salary, deps) {
            const active = currentActiveTable || getDefaultTable();
            if (!active || !active.row) {
                return { irs: 0, rate: 0 };
            }

            const bracketRows = getBracketRowsForTable(active.row);
            if (bracketRows.length === 0) {
                return { irs: 0, rate: 0 };
            }

            // Find the bracket that matches salary
            let matched = null;
            for (const cells of bracketRows) {
                const firstCell = (cells[0] || "").toLowerCase();
                const isSup = firstCell.includes("superior");
                let limit = 0;
                for (let i = 1; i < Math.min(3, cells.length); i++) {
                    const num = parseNum(cells[i]);
                    if (num > 0) {
                        limit = num;
                        break;
                    }
                }
                if (isSup) {
                    if (salary > limit) {
                        matched = cells;
                        break;
                    }
                } else {
                    if (salary <= limit) {
                        matched = cells;
                        break;
                    }
                }
            }

            if (!matched && bracketRows.length > 0) {
                matched = bracketRows[bracketRows.length - 1];
            }
            if (!matched) return { irs: 0, rate: 0 };

            const line = matched.join(" ");

            // Modern marginal rate system (post-2023)
            // Fórmula a aplicar: Remuneração x Taxa - Parcela a abater - Parcela adicional a abater x nº dependentes
            const isModern = /f[óo]rmula|parcela\s+a\s+abater|taxa\s+marginal/i.test(sheetContainer.textContent || "") ||
                line.includes("- R") || matched.length > 8;

            if (isModern) {
                const rate = parseNum(matched[2]);
                let parcela = 0;
                let parcelaAdicional = parseNum(matched[matched.length - 2]);

                if (line.includes("- R")) {
                    let factor = 1.0;
                    let upper = salary;
                    for (let i = 0; i < matched.length; i++) {
                        if (matched[i] === "x" && i + 1 < matched.length) {
                            const f = parseNum(matched[i + 1]);
                            if (f > 0) factor = f;
                        }
                        if (matched[i].includes("- R") && i > 0) {
                            const u = parseNum(matched[i - 1]);
                            if (u > 0) upper = u;
                        }
                    }
                    parcela = rate * factor * Math.max(0, upper - salary);
                } else {
                    parcela = parseNum(matched[3]);
                }

                // Fórmula a aplicar: Remuneração x Taxa - Parcela a abater - Parcela adicional a abater x nº dependentes
                const irs = Math.max(0, (salary * rate) - parcela - (parcelaAdicional * deps));
                const formulaText = `(${formatEUR(salary)} × ${(rate * 100).toFixed(2)}%) - ${formatEUR(parcela)} - (${formatEUR(parcelaAdicional)} × ${deps})`;
                return { irs, rate, parcela, parcelaAdicional, isModern: true, formulaText };
            } else {
                // Classic effective rate system (pre-2024)
                const rates = [];
                for (let i = 2; i < matched.length; i++) {
                    const v = parseNum(matched[i]);
                    if (!isNaN(v)) rates.push(v);
                }
                const depIdx = Math.min(deps, rates.length - 1);
                const rate = rates[depIdx] || 0;
                const irs = salary * rate;
                const formulaText = `${formatEUR(salary)} × ${(rate * 100).toFixed(2)}% (${deps} dependentes)`;
                return { irs, rate, isModern: false, formulaText };
            }
        }

        function calculateSalary() {
            if (!calcGrossInput) return;

            const grossSalary = Math.max(0, parseFloat(calcGrossInput.value) || 0);
            const dependents = Math.max(0, parseInt(calcDepsInput ? calcDepsInput.value : 0, 10) || 0);
            const hasUnion = calcUnionCheckbox ? calcUnionCheckbox.checked : false;

            // 1. TSU (11%)
            const tsu = grossSalary * 0.11;

            // 2. Sindicato (1% opcional)
            const union = hasUnion ? grossSalary * 0.01 : 0;

            // 3. IRS: Fórmula a aplicar: Remuneração x Taxa - Parcela a abater - Parcela adicional a abater x nº dependentes
            const irsResult = computeIRS(grossSalary, dependents);
            const irs = irsResult.irs;
            const effectiveRate = grossSalary > 0 ? (irs / grossSalary) * 100 : 0;

            // 4. Totals
            const totalDeductions = tsu + irs + union;
            const netSalary = Math.max(0, grossSalary - totalDeductions);

            // Update DOM
            if (resGrossEl) resGrossEl.textContent = formatEUR(grossSalary);
            if (resTsuEl) resTsuEl.textContent = `- ${formatEUR(tsu)}`;
            if (resIrsEl) resIrsEl.textContent = `- ${formatEUR(irs)}`;
            if (resIrsFormulaTextEl) {
                resIrsFormulaTextEl.textContent = `Fórmula: ${irsResult.formulaText || ""}`;
            }

            if (calcRowUnionEl) {
                if (hasUnion) {
                    calcRowUnionEl.style.display = "flex";
                    if (resUnionEl) resUnionEl.textContent = `- ${formatEUR(union)}`;
                } else {
                    calcRowUnionEl.style.display = "none";
                }
            }

            if (resTotalDeductionsEl) resTotalDeductionsEl.textContent = `- ${formatEUR(totalDeductions)}`;
            if (resNetEl) resNetEl.textContent = formatEUR(netSalary);
        }

        // Calculator Event Listeners
        if (calcGrossInput) {
            calcGrossInput.addEventListener("input", calculateSalary);
        }
        if (calcDepsInput) {
            calcDepsInput.addEventListener("input", calculateSalary);
        }
        if (calcDepMinusBtn && calcDepsInput) {
            calcDepMinusBtn.addEventListener("click", () => {
                let val = parseInt(calcDepsInput.value, 10) || 0;
                if (val > 0) {
                    calcDepsInput.value = val - 1;
                    calculateSalary();
                }
            });
        }
        if (calcDepPlusBtn && calcDepsInput) {
            calcDepPlusBtn.addEventListener("click", () => {
                let val = parseInt(calcDepsInput.value, 10) || 0;
                calcDepsInput.value = val + 1;
                calculateSalary();
            });
        }
        if (calcUnionCheckbox) {
            calcUnionCheckbox.addEventListener("change", calculateSalary);
        }

