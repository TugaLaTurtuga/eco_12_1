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
                    else if (question === "deps") organizerState.deps = val;
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
            } else {
                if (organizerStatusEl) {
                    organizerStatusEl.className = "organizer-status";
                    organizerStatusEl.textContent = "Nenhuma tabela encontrada para os critérios selecionados.";
                }
            }
        }

