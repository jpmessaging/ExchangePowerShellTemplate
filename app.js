const ON_PREM_PREPARE_GUIDANCE_TEXT = 
`任意の Exchange サーバー上の Exchange 管理シェルで以下の情報を取得のうえ、後述のアップロードサイトよりご提供ください。

`;

const EXO_PREPARE_GUIDANCE_TEXT = 
`PowerShell を使用して Exchange Online に接続し、以下の情報を取得のうえ、後述のアップロードサイトよりご提供ください。
Exchange Online に PowerShell で接続する方法は、以下の技術情報をご参照ください。

Title: Exchange Online PowerShell に接続する
URL: https://learn.microsoft.com/ja-jp/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps

`;

const GRAPH_PREPARE_GUIDANCE_TEXT = 
`Microsoft Graph PowerShell SDK を使用して Microsoft 365 テナントに接続し、以下の情報を取得のうえ、後述のアップロードサイトよりご提供ください。

Microsoft Graph PowerShell SDK がインストールされた端末をご用意ください。 手順については下記にてご案内しておりますので必要に応じてご確認いただけますと幸いです

Title: Install the Microsoft Graph PowerShell SDK
URL: https://learn.microsoft.com/ja-jp/powershell/microsoftgraph/installation?view=graph-powershell-1.0

`;

const commandDefinition = window.commandDefinition || [];

function normalizeTagArray(tagValue) {
    if (Array.isArray(tagValue)) {
        return tagValue.filter(tag => !!tag).map(tag => String(tag));
    }

    if (typeof tagValue === "string" && tagValue) {
        return [tagValue];
    }

    return [];
}

// Normalize tag shape to string-array. All commands are expected to define tag.
commandDefinition.forEach(cmd => {
    const normalizedTags = normalizeTagArray(cmd.tag);
    cmd.tag = normalizedTags;
});

// Defines display order for Tag filter chips.
const TAG_FILTER_ORDER = [
    "Mailbox",
    "Group",
    "Other recipient",
    "Organization",
    "Calendar",
    "Transport",
];

const tagOrderMap = new Map(
    TAG_FILTER_ORDER.map((tag, index) => [tag, index])
);

const tagCatalog = Array.from(
    new Set(commandDefinition.flatMap(cmd => normalizeTagArray(cmd.tag)))
).sort((a, b) => {
    const rankA = tagOrderMap.has(a) ? tagOrderMap.get(a) : Number.MAX_SAFE_INTEGER;
    const rankB = tagOrderMap.has(b) ? tagOrderMap.get(b) : Number.MAX_SAFE_INTEGER;

    if (rankA !== rankB) {
        return rankA - rankB;
    }

    return a.localeCompare(b, "ja");
});

function getTagFilterKey(tag) {
    return `tag:${String(tag).toLowerCase()}`;
}

const commandUiState = {
    searchText: "",
    filters: {
        exo: false,
        onprem: false,
        graph: false,
        orgLevel: false,
        hasParams: false
    },
    collapsedGroups: {}
};

tagCatalog.forEach(tag => {
    commandUiState.filters[getTagFilterKey(tag)] = false;
});

function renderTagFilterChips() {
    const container = document.getElementById("tagFilterChips");
    if (!container) {
        return;
    }

    container.innerHTML = "";

    tagCatalog.forEach(tag => {
        const filterKey = getTagFilterKey(tag);
        const button = document.createElement("button");
        button.type = "button";
        button.className = "filter-chip";
        button.id = `filterChip_${filterKey}`;
        button.textContent = tag;
        button.setAttribute("aria-pressed", "false");
        button.addEventListener("click", () => toggleCommandFilter(filterKey));
        container.appendChild(button);
    });
}

function buildCommandSearchText(commandDef) {
    const paramKeys = (commandDef.parameters || []).map(param => param.key || "").join(" ");
    const optionsText = (commandDef.parameters || [])
        .flatMap(param => (param.options || []))
        .map(option => (typeof option === "string" ? option : `${option.value || ""} ${option.label || ""}`))
        .join(" ");
    const applicableText = (commandDef.applicable || []).join(" ");

    return `${commandDef.commandName} ${paramKeys} ${optionsText} ${applicableText}`.toLowerCase();
}

function handleCommandSearchInput(inputEl) {
    commandUiState.searchText = (inputEl?.value || "").trim().toLowerCase();
    applyCommandFilters();
}

function clearCommandSearch() {
    const searchInput = document.getElementById("commandSearch");
    if (searchInput) {
        searchInput.value = "";
    }

    commandUiState.searchText = "";
    applyCommandFilters();
}

function toggleCommandFilter(filterKey) {
    if (!Object.prototype.hasOwnProperty.call(commandUiState.filters, filterKey)) {
        return;
    }

    commandUiState.filters[filterKey] = !commandUiState.filters[filterKey];
    applyCommandFilters();
}

function updateCommandFilterChipStates() {
    Object.entries(commandUiState.filters).forEach(([key, active]) => {
        const chip = document.getElementById(`filterChip_${key}`);
        if (!chip) {
            return;
        }

        chip.classList.toggle("active", active);
        chip.setAttribute("aria-pressed", String(active));
    });
}

function commandMatchesActiveFilters(commandDef) {
    const filters = commandUiState.filters;
    const activeApplicableFilters = ["exo", "onprem", "graph"].filter(key => filters[key]);
    const activeTagFilters = tagCatalog.filter(tag => filters[getTagFilterKey(tag)]);

    if (activeApplicableFilters.length > 0) {
        const applicable = commandDef.applicable || [];
        const hasAnyApplicableMatch = activeApplicableFilters.some(key => applicable.includes(key));
        if (!hasAnyApplicableMatch) {
            return false;
        }
    }

    if (filters.orgLevel && !commandDef.orgLevel) {
        return false;
    }

    if (filters.hasParams && (!Array.isArray(commandDef.parameters) || commandDef.parameters.length === 0)) {
        return false;
    }

    if (activeTagFilters.length > 0) {
        const commandTags = normalizeTagArray(commandDef.tag);
        const hasTagMatch = activeTagFilters.some(tag => commandTags.includes(tag));
        if (!hasTagMatch) {
            return false;
        }
    }

    const query = commandUiState.searchText;
    if (!query) {
        return true;
    }

    const searchable = buildCommandSearchText(commandDef);
    return searchable.includes(query);
}

function applyCommandFilters() {
    const totalCount = commandDefinition.length;
    let visibleCount = 0;

    commandDefinition.forEach(cmd => {
        const block = document.querySelector(`.cmd-block[data-cmd="${cmd.commandName}"]`);
        if (!block) {
            return;
        }

        const visible = commandMatchesActiveFilters(cmd);
        block.hidden = !visible;
        if (visible) {
            visibleCount += 1;
        }
    });

    const summary = document.getElementById("commandSummary");
    if (summary) {
        summary.textContent = `${visibleCount} / ${totalCount}`;
    }

    const noResults = document.getElementById("commandNoResults");
    if (noResults) {
        noResults.hidden = visibleCount > 0;
    }

    updateCommandFilterChipStates();
}

function generateGuidanceSteps(varDefsText, commandText) {
    const guidanceTypeValue = document.getElementById("guidanceType")?.value ?? "exo";
    const includeTranscript = document.getElementById("includeTranscript");
    const transcriptEnabled = !!(includeTranscript && includeTranscript.checked);
    const exportToDesktop = document.getElementById("exportToDesktop");
    const exportFolderPath = getExportFolderPath();
    const varDefsGuidanceText = varDefsText ? indentNonEmptyLines(varDefsText.trimEnd()) + "\n" : "";
    const commandGuidanceText = commandText ? indentNonEmptyLines(commandText.trimEnd()) + "\n" : "";
    let guidanceText = "";
    let step = 1;

    if (guidanceTypeValue === "exoSimple" || guidanceTypeValue === "onpremSimple") {
        switch (guidanceTypeValue) {
            case "exoSimple":
                guidanceText += "Exchange Online に接続した PowerShell にて以下のコマンドを実行し、デスクトップに出力された XML ファイルをお寄せください。\n\n"
                break;
            case "onpremSimple":
                guidanceText += "任意のメールボックス サーバーで起動した Exchange 管理シェルにて以下のコマンドを実行し、デスクトップに出力された XML ファイルをお寄せください。\n\n"
                break;
        }
        guidanceText += `${varDefsGuidanceText}\n`
        guidanceText += `   cd ${exportFolderPath}\n`
        // guidanceText += "XML ファイルの出力先は適宜変更ください。\n\n"
        guidanceText += `${commandGuidanceText}`;
        return guidanceText;
    }

    switch (guidanceTypeValue) {
        case "onprem":
            guidanceText += ON_PREM_PREPARE_GUIDANCE_TEXT;
            break;
        case "graph":
            guidanceText += GRAPH_PREPARE_GUIDANCE_TEXT;
            break;
        case "exo":
        default:
            guidanceText += EXO_PREPARE_GUIDANCE_TEXT;
            break;
    }

    if (guidanceTypeValue === "onprem") {
        guidanceText += `${step++}. 任意のメールボックス サーバーで Exchange 管理シェルを起動します。\n`;
        guidanceText += `${step++}. 以下の cd コマンドにて情報採取結果を出力するフォルダーに移動します。\n`;
    }
    else {
        guidanceText += `${step++}. PowerShell を起動し、以下の cd コマンドにて情報採取結果を出力するフォルダーに移動します。\n`;
    }

    if (exportToDesktop && exportToDesktop.checked) {
        guidanceText += "   以下の実行例ではデスクトップに移動しますが、任意のフォルダーに変更いただいて構いません。\n\n";
    }
    else {
        guidanceText += "   移動先は任意のフォルダーに変更いただいて構いません。\n\n";
    }

    guidanceText += `   cd ${exportFolderPath}\n\n`;

    if (transcriptEnabled) {
        guidanceText += `${step++}. 以下のコマンドを実行し、トランスクリプトを開始します。\n\n`;
        guidanceText += `   Start-Transcript -Path "PowerShell_transcript.$(Get-Date -UFormat %Y%m%d%H%M%S).txt"\n\n`;
    }

    if (guidanceTypeValue === "exo") {
        guidanceText += `${step++}. 以下のコマンドを実行し、管理者アカウントで Exchange Online に接続します。\n\n`;
        guidanceText += `   Connect-ExchangeOnline\n\n`;
    } else if (guidanceTypeValue === "graph") {
        guidanceText += `${step++}. 以下のコマンドを実行し、管理者アカウントで Microsoft Graph に接続します。\n\n`;
        guidanceText += `   Connect-MgGraph -Scopes "User.Read.All"\n\n`;
    }

    guidanceText += `${step++}. 以下のコマンドを順に実行し、出力される XML ファイルを弊社までお寄せください。\n\n`;
    guidanceText += "   コマンド)\n";
    guidanceText += `${varDefsGuidanceText}\n`;
    guidanceText += `${commandGuidanceText}`;

    if (transcriptEnabled) {
        guidanceText += `\n${step++}. 以下のコマンドを実行し、トランスクリプトを終了します。出力された PowerShell_transcript.xxxxxxxxxxxxxx.txt ファイルをお寄せください。\n\n`;
        guidanceText += "   Stop-Transcript\n";
    }

    return guidanceText;
}

// Normalizes and sanitizes a folder path for safe Windows-style output.
function sanitizeFolderPath(raw) {
    if (!raw) return "";

    let value = raw.replace(/\//g, "\\");
    value = value.replace(/[\u0000-\u001F<>"|?*]/g, "");

    const hasDrivePrefix = /^[A-Za-z]:/.test(value);

    if (hasDrivePrefix) {
        value = value.slice(0, 2) + value.slice(2).replace(/:/g, "");
    } else {
        value = value.replace(/:/g, "");
    }

    if (value.startsWith("\\\\")) {
        value = "\\\\" + value.slice(2).replace(/\\{2,}/g, "\\");
    } else {
        value = value.replace(/\\{2,}/g, "\\");
    }

    return value;
}

// Converts input text into a valid PowerShell variable name.
function sanitizeVarName(raw) {
    const noSpaces = raw.replace(/\s+/g, "");
    const allowedOnly = noSpaces.replace(/[^A-Za-z0-9_]/g, "");

    return allowedOnly.replace(/^[^A-Za-z_]+/, "");
}

// Updates a command-level checkbox state based on its target checkbox states.
function updateParentCheckbox(commandName) {
    const cmdToggle = document.querySelector(`.cmd-toggle[data-cmd="${commandName}"]`);
    const commandDef = commandDefinition.find(cmd => cmd.commandName === commandName);

    if (!cmdToggle) {
        return;
    }

    // orgLevel commands do not have per-target children, so keep a binary checked state.
    if (commandDef && commandDef.orgLevel) {
        cmdToggle.indeterminate = false;
        return;
    }

    const children = document.querySelectorAll(`.${commandName}_target`);

    const total = children.length;
    const checked = Array.from(children).filter(cb => cb.checked).length;

    if (checked === 0) {
        cmdToggle.checked = false;
        cmdToggle.indeterminate = false;
    } else if (checked === total) {
        cmdToggle.checked = true;
        cmdToggle.indeterminate = false;
    } else {
        cmdToggle.checked = false;
        cmdToggle.indeterminate = true;
    }
}

// Recalculates parent checkbox states for all commands.
function updateAllParentCheckboxes() {
    commandDefinition.forEach(cmd => {
        updateParentCheckbox(cmd.commandName);
    });
}

// Applies a command-level checkbox state to all target checkboxes for that command.
function toggleAllTargetsForCommand(commandName, parentCheckbox) {
    const commandDef = commandDefinition.find(cmd => cmd.commandName === commandName);

    if (commandDef && commandDef.orgLevel) {
        parentCheckbox.indeterminate = false;
        updateGuidanceTypeIfNeeded();
        return;
    }

    const children = document.querySelectorAll(`.${commandName}_target`);

    children.forEach(cb => {
        cb.checked = parentCheckbox.checked;
    });

    parentCheckbox.indeterminate = false;
    updateGuidanceTypeIfNeeded();
}

// Checks if all checked commands support common guidance types and updates accordingly with priority: exo > onprem > graph.
function updateGuidanceTypeIfNeeded() {
    const checkedCommands = commandDefinition.filter(cmd => {
        const cmdToggle = document.querySelector(`.cmd-toggle[data-cmd="${cmd.commandName}"]`);
        return cmdToggle && (cmdToggle.checked || cmdToggle.indeterminate);
    });

    if (checkedCommands.length === 0) {
        return;
    }

    const guidanceTypeSelect = document.getElementById("guidanceType");
    if (!guidanceTypeSelect) {
        return;
    }

    // Find common applicable types across all checked commands
    const commonApplicable = checkedCommands.reduce((common, cmd) => {
        const cmdSet = new Set(cmd.applicable || []);
        return common.filter(type => cmdSet.has(type));
    }, [...(checkedCommands[0]?.applicable || [])]);

    // Apply priority: exo > onprem > graph
    const priorityOrder = ["exo", "onprem", "graph"];
    for (const guidanceType of priorityOrder) {
        if (commonApplicable.includes(guidanceType)) {
            guidanceTypeSelect.value = guidanceType;
            guidanceTypeSelect.dispatchEvent(new Event('change'));
            return;
        }
    }
}

// Applies expanded or collapsed UI state to a specific command section.
function setGroupExpanded(commandName, group, expanded) {
    const el = document.getElementById(`${group === "targets" ? "targetGroup" : "paramGroup"}_${commandName}`);
    const collapse = document.querySelector(`.collapse-toggle[data-cmd="${commandName}"][data-group="${group}"]`);
    const chevron = document.getElementById(`chevron_${group}_${commandName}`);

    if (!el || !collapse || !chevron) return;

    if (expanded) {
        el.classList.add("open");
        collapse.setAttribute("aria-expanded", "true");
        chevron.textContent = "▼";
    } else {
        el.classList.remove("open");
        collapse.setAttribute("aria-expanded", "false");
        chevron.textContent = "▶";
    }
}

// Closes all collapsible command sections except the specified one.
function closeAllGroupsExcept(openCommandName, openGroup) {
    commandDefinition.forEach(cmd => {
        ["targets", "params"].forEach(group => {
            if (cmd.commandName === openCommandName && group === openGroup) {
                return;
            }

            setGroupExpanded(cmd.commandName, group, false);
        });
    });
}

// Toggles one collapsible section and closes all other command sections.
function toggleGroup(commandName, group) {
    const collapse = document.querySelector(`.collapse-toggle[data-cmd="${commandName}"][data-group="${group}"]`);
    const expanded = collapse && collapse.getAttribute("aria-expanded") === "true";

    if (expanded) {
        setGroupExpanded(commandName, group, false);
        return;
    }

    if (group === "targets") {
        commandDefinition.forEach(cmd => {
            setGroupExpanded(cmd.commandName, "params", false);
        });
    } else {
        closeAllGroupsExcept(commandName, group);
    }
    setGroupExpanded(commandName, group, true);
}

// Captures current command UI selections and collapse state for re-rendering.
function saveState() {
    const state = {};

    commandDefinition.forEach(cmd => {
        const commandName = cmd.commandName;
        const cmdToggle = document.querySelector(`.cmd-toggle[data-cmd="${commandName}"]`);
        const targetsCollapse = document.querySelector(`.collapse-toggle[data-cmd="${commandName}"][data-group="targets"]`);
        const paramsCollapse = document.querySelector(`.collapse-toggle[data-cmd="${commandName}"][data-group="params"]`);

        state[commandName] = {
            enabled: !!(cmdToggle && cmdToggle.checked),
            targets: [],
            params: [],
            parameterValues: {},
            expandedTargets: !!(targetsCollapse && targetsCollapse.getAttribute("aria-expanded") === "true"),
            expandedParams: !!(paramsCollapse && paramsCollapse.getAttribute("aria-expanded") === "true")
        };

        document.querySelectorAll(`.${commandName}_target:checked`).forEach(cb => {
            state[commandName].targets.push(cb.value);
        });

        document.querySelectorAll(`.${commandName}_param:checked`).forEach(cb => {
            state[commandName].params.push(cb.value);
        });

        if (cmd.parameters && cmd.parameters.length > 0) {
            cmd.parameters.forEach(param => {
                if (param.type === "multi-checkbox-split" || param.type === "multi-checkbox") {
                    const selectedValues = Array.from(
                        document.querySelectorAll(`.${commandName}_value_${param.key}:checked`)
                    ).map(cb => cb.value);
                    const hasRenderedInputs = !!document.querySelector(`.${commandName}_value_${param.key}`);
                    const defaultSelectedOptions = Array.isArray(param.defaultSelectedOptions)
                        ? param.defaultSelectedOptions
                        : [];

                    state[commandName].parameterValues[param.key] =
                        selectedValues.length > 0
                            ? selectedValues
                            : (!hasRenderedInputs ? defaultSelectedOptions : []);
                }
            });
        }
    });

    return state;
}

// Restores saved command UI selections and expansion state.
function restoreState(state) {
    commandDefinition.forEach(cmd => {
        const commandName = cmd.commandName;

        const cmdToggle = document.querySelector(`.cmd-toggle[data-cmd="${commandName}"]`);
        if (cmdToggle) {
            cmdToggle.checked = !!state[commandName].enabled;
        }

        document.querySelectorAll(`.${commandName}_target`).forEach(cb => {
            if (state[commandName].targets.includes(cb.value)) {
                cb.checked = true;
            }
        });

        document.querySelectorAll(`.${commandName}_param`).forEach(cb => {
            if (state[commandName].params.includes(cb.value)) {
                cb.checked = true;
            }
        });

        if (cmd.parameters && cmd.parameters.length > 0) {
            cmd.parameters.forEach(param => {
                if (param.type === "multi-checkbox-split" || param.type === "multi-checkbox") {
                    const selectedValues = (state[commandName].parameterValues && state[commandName].parameterValues[param.key]) || [];
                    document.querySelectorAll(`.${commandName}_value_${param.key}`).forEach(cb => {
                        cb.checked = selectedValues.includes(cb.value);
                    });
                    updateDefaultOptionPresetState(commandName, param.key);
                }
            });
        }

        setGroupExpanded(commandName, "targets", !!state[commandName].expandedTargets);
        setGroupExpanded(commandName, "params", !!state[commandName].expandedParams);
    });

    updateAllParentCheckboxes();
}

// Rebuilds the command selection UI and restores the previous interactive state.
function renderCommands() {
    const state = saveState();
    const container = document.getElementById("commands");
    container.innerHTML = "";
    const targetNameInputs = document.querySelectorAll(".target-var-name");

    commandDefinition.forEach(cmd => {
        const commandName = cmd.commandName;
        const hasParameters = Array.isArray(cmd.parameters) && cmd.parameters.length > 0;
        const hasTargetScope = !cmd.orgLevel;
        const div = document.createElement("div");
        div.className = "cmd-block";
        div.setAttribute("data-cmd", commandName);

        let badgeHtml = "";
        if (cmd.applicable && Array.isArray(cmd.applicable)) {
            badgeHtml = cmd.applicable
                .map(type => `<span class="cmd-badge cmd-badge-${type}">${type.toUpperCase()}</span>`)
                .join("");
        }

        let html = `
            <div class="cmd-header">
                <label>
                    <input type="checkbox"
                           class="cmd-toggle"
                           data-cmd="${commandName}"
                           onchange="toggleAllTargetsForCommand('${commandName}', this)">
                    ${commandName}${badgeHtml}
                </label>

                <button type="button"
                        class="collapse-toggle"
                        aria-expanded="false"
                        onclick="toggleGroup('${commandName}', 'targets')"
                        data-cmd="${commandName}"
                        data-group="targets"
                        ${hasTargetScope ? "" : "disabled aria-disabled=\"true\""}>
                    <span id="chevron_targets_${commandName}" class="chevron">▶</span>
                    対象
                </button>

                <button type="button"
                        class="collapse-toggle"
                        aria-expanded="false"
                        onclick="toggleGroup('${commandName}', 'params')"
                        data-cmd="${commandName}"
                        data-group="params"
                        ${hasParameters ? "" : "disabled aria-disabled=\"true\""}>
                    <span id="chevron_params_${commandName}" class="chevron">▶</span>
                    パラメーター
                </button>
            </div>

            <div id="targetGroup_${commandName}" class="target-group">
        `;

        if (hasTargetScope) {
            targetNameInputs.forEach((input, i) => {
                const name = input.value.trim();

                if (name) {
                    html += `
                        <label>
                            <input type="checkbox"
                                   class="${commandName}_target"
                                   value="${i}"
                                   onchange="updateParentCheckbox('${commandName}')">
                            ${name}
                        </label>
                    `;
                }
            });
        }

        html += `
            </div>
            <div id="paramGroup_${commandName}" class="param-group">
        `;

        const switchParams = (cmd.parameters || []).filter(
            param => param.type === "switch" || param.type === "switch-split"
        );
        if (switchParams.length > 0) {
            switchParams.forEach(param => {
                html += `
                    <label>
                        <input type="checkbox"
                               class="${commandName}_param"
                               value="${param.key}">
                        ${param.key}
                    </label>
                `;
            });
        }

        if (cmd.parameters && cmd.parameters.length > 0) {
            cmd.parameters.forEach(param => {
                if (param.type === "multi-checkbox-split" || param.type === "multi-checkbox") {
                    const hasDefaultPreset = Array.isArray(param.defaultSelectedOptions) && param.defaultSelectedOptions.length > 0;
                    const defaultPresetOptionList = hasDefaultPreset
                        ? param.defaultSelectedOptions.join(", ")
                        : "";
                    const defaultPresetHtml = hasDefaultPreset

                        ? `
                            <div class="option-sub param-default-preset-row">
                                <label>
                                    <input type="checkbox"
                                           class="param-default-preset"
                                           data-cmd="${commandName}"
                                           data-param="${param.key}"
                                           onchange="toggleDefaultOptionPreset('${commandName}', '${param.key}', this)">
                                    <span>
                                        <span>Default Preset</span>
                                        <span class="param-default-preset-caption">(${defaultPresetOptionList})</span>
                                    </span>
                                </label>
                            </div>
                        `
                        : "";
                    html += `
                        <div class="param-input-block">
                            <span class="param-input-title">${param.key}</span>
                            ${defaultPresetHtml}
                            <div class="param-options-grid scrollable">
                    `;

                    param.options.forEach(option => {
                        if (typeof option === "string") {
                            html += `
                                <label class="param-option">
                                    <input type="checkbox"
                                           class="${commandName}_value_${param.key}"
                                           value="${option}"
                                           onchange="updateDefaultOptionPresetState('${commandName}', '${param.key}')">
                                    <span class="param-option-text" title="${option}">${option}</span>
                                </label>
                            `;
                        } else {
                            html += `
                                <label class="param-option">
                                    <input type="checkbox"
                                           class="${commandName}_value_${param.key}"
                                           value="${option.value}"
                                           onchange="updateDefaultOptionPresetState('${commandName}', '${param.key}')">
                                    <span class="param-option-text" title="${option.label || option.value}">${option.label || option.value}</span>
                                </label>
                            `;
                        }
                    });

                    html += `
                            </div>
                        </div>
                    `;
                }
            });
        }

        html += `
            </div>
        `;

        div.innerHTML = html;
        container.appendChild(div);
    });

    restoreState(state);
    applyCommandFilters();
}

// Clears all command-related checkboxes in the command selection area.
function resetCommandSelections() {
    document.querySelectorAll("#commands input[type='checkbox']").forEach(input => {
        input.checked = false;
        input.indeterminate = false;
    });

    updateGuidanceTypeIfNeeded();
}

// Sanitizes variable names and optionally re-renders command target checkboxes.
function handleVarNameInput(inputEl = null, shouldRender = inputEl !== null) {
    const targetNameInputs = inputEl ? [inputEl] : Array.from(document.querySelectorAll(".target-var-name"));

    targetNameInputs.forEach(input => {
        const sanitized = sanitizeVarName(input.value);
        if (input.value !== sanitized) {
            input.value = sanitized;
        }
    });

    if (shouldRender) {
        renderCommands();
    }
}

// Sanitizes the export path input field as the user types.
function handleExportPathInput(inputEl) {
    const sanitized = sanitizeFolderPath(inputEl.value);
    if (inputEl.value !== sanitized) {
        inputEl.value = sanitized;
    }
}

// Applies path sanitization to the export path input during initialization.
function normalizeExportPathInput() {
    const inputEl = document.getElementById("exportPath");
    if (!inputEl) return;
    inputEl.value = sanitizeFolderPath(inputEl.value);
}

// Returns a parameter definition by command and parameter key.
function getParameterDefinition(commandName, parameterKey) {
    const commandDef = commandDefinition.find(cmd => cmd.commandName === commandName);
    if (!commandDef || !Array.isArray(commandDef.parameters)) return null;

    return commandDef.parameters.find(param => param.key === parameterKey) || null;
}

// Syncs preset checkbox state to whether all default options are selected.
function updateDefaultOptionPresetState(commandName, parameterKey) {
    const paramDef = getParameterDefinition(commandName, parameterKey);
    if (!paramDef || !Array.isArray(paramDef.defaultSelectedOptions) || paramDef.defaultSelectedOptions.length === 0) {
        return;
    }

    const presetCheckbox = document.querySelector(
        `.param-default-preset[data-cmd="${commandName}"][data-param="${parameterKey}"]`
    );
    if (!presetCheckbox) return;

    const selectedSet = new Set(
        Array.from(document.querySelectorAll(`.${commandName}_value_${parameterKey}:checked`)).map(cb => cb.value)
    );
    const allDefaultsSelected = paramDef.defaultSelectedOptions.every(value => selectedSet.has(value));
    presetCheckbox.checked = allDefaultsSelected;
}

// Toggles only default-selected options for a parameter via preset checkbox.
function toggleDefaultOptionPreset(commandName, parameterKey, presetCheckbox) {
    const paramDef = getParameterDefinition(commandName, parameterKey);
    if (!paramDef || !Array.isArray(paramDef.defaultSelectedOptions) || paramDef.defaultSelectedOptions.length === 0) {
        return;
    }

    const defaultSet = new Set(paramDef.defaultSelectedOptions);
    document.querySelectorAll(`.${commandName}_value_${parameterKey}`).forEach(cb => {
        if (defaultSet.has(cb.value)) {
            cb.checked = presetCheckbox.checked;
        }
    });

    updateDefaultOptionPresetState(commandName, parameterKey);
}

// Enables or disables expandVarNameInFileName based on keepVarNameInFileName.
function toggleVarNameInFileNameOptions() {
    const keepVarNameInFileName = document.getElementById("keepVarNameInFileName");
    const expandVarNameInFileName = document.getElementById("expandVarNameInFileName");

    if (!keepVarNameInFileName || !expandVarNameInFileName) return;

    expandVarNameInFileName.disabled = !keepVarNameInFileName.checked;
}

// Enables or disables the export path input based on the desktop output option.
function toggleExportPathInput() {
    const desktopCheckbox = document.getElementById("exportToDesktop");
    const exportPathInput = document.getElementById("exportPath");
    const includeGuidanceText = document.getElementById("includeGuidanceText");
    const guidanceTypeValue = document.getElementById("guidanceType")?.value ?? "exo";

    if (!desktopCheckbox || !exportPathInput) return;

    const guidanceDisabled = !!(includeGuidanceText && !includeGuidanceText.checked);
    const simpleGuidanceSelected = guidanceTypeValue === "exoSimple" || guidanceTypeValue === "onpremSimple";

    desktopCheckbox.disabled = guidanceDisabled || simpleGuidanceSelected;
    exportPathInput.disabled = guidanceDisabled || simpleGuidanceSelected || desktopCheckbox.checked;
}

// Enables or disables guidance-type radio options based on guidance-text option.
function toggleGuidanceTypeOptions() {
    const includeGuidanceText = document.getElementById("includeGuidanceText");
    const guidanceTypeElement = document.getElementById("guidanceType");
    const includeTranscript = document.getElementById("includeTranscript");
    const premiseSuffix = document.getElementById("premiseSuffix");
    const selectedGuidanceTypeValue = guidanceTypeElement?.value ?? "exo";

    if (!includeGuidanceText || !guidanceTypeElement || !includeTranscript || !premiseSuffix) return;

    const disabled = !includeGuidanceText.checked;
    const simpleGuidanceSelected = selectedGuidanceTypeValue === "exoSimple" || selectedGuidanceTypeValue === "onpremSimple";
    const graphGuidanceSelected = selectedGuidanceTypeValue === "graph";
    guidanceTypeElement.disabled = disabled;
    includeTranscript.disabled = disabled || simpleGuidanceSelected;
    toggleExportPathInput();
}

// Adds a new target row with default variable name and mailbox value.
function addTarget() {
    const targetTableBody = document.querySelector("#targetTable tbody");
    const index = targetTableBody.rows.length + 1;

    const row = targetTableBody.insertRow();

    row.innerHTML = `
        <td><input class="target-var-name" value="User${index}" onchange="handleVarNameInput(this)"></td>
        <td><input class="target-var-value" value="事象発生ユーザーのメール アドレス"></td>
        <td><button type="button" onclick="removeTarget(this)">削除</button></td>
    `;

    renderCommands();
}

// Removes a target row while ensuring at least one target remains.
function removeTarget(buttonEl) {
    const targetTableBody = document.querySelector("#targetTable tbody");

    if (targetTableBody.rows.length <= 1) {
        alert("少なくとも 1 ユーザーは必要です。");
        return;
    }

    buttonEl.closest("tr").remove();
    renderCommands();
}

// Converts a command name into a file-name prefix by removing the verb part.
function getfileNamePrefixFromCommand(commandName) {
    const parts = commandName.split("-");

    if (parts.length <= 1) {
        return commandName;
    }

    return parts.slice(1).join("");
}

// Returns a normalized output directory path from the current option settings.
function getExportDirectoryPath() {
    const inputEl = document.getElementById("exportPath");
    const raw = inputEl ? inputEl.value : "";
    const sanitized = sanitizeFolderPath(raw || "C:\\temp");
    const base = sanitized.replace(/\s+$/g, "");

    if (inputEl && inputEl.value !== base) {
        inputEl.value = base;
    }

    if (/^[A-Za-z]:\\$/.test(base)) {
        return base;
    }

    return base.replace(/\\+$/, "");
}

// Returns the actual output folder path based on desktop-output option.
function getExportFolderPath() {
    const exportToDesktop = document.getElementById("exportToDesktop");

    if (exportToDesktop && exportToDesktop.checked) {
        return "~\\Desktop";
    }

    return getExportDirectoryPath();
}

// Returns filename suffix for selected guidance type when premiseSuffix is enabled.
function getPremiseSuffixForFileName() {
    const premiseSuffix = document.getElementById("premiseSuffix");
    const guidanceTypeValue = document.getElementById("guidanceType")?.value ?? "exo";

    if (!premiseSuffix || !premiseSuffix.checked) {
        return "";
    }

    if (guidanceTypeValue === "exo" || guidanceTypeValue === "exoSimple") {
        return "_ExO";
    }

    if (guidanceTypeValue === "onprem" || guidanceTypeValue === "onpremSimple") {
        return "_OnPrem";
    }

    return "";
}

// Prepends indentation to each non-empty line in the generated script body.
function indentNonEmptyLines(text, indent = "   ") {
    return text
        .replace(/\r\n/g, "\n")
        .split("\n")
        .map(line => (line ? `${indent}${line}` : line))
        .join("\n");
}

// Only for Get-MailboxFolderPermission
// Finds an English folder name paired with the selected localized folder label.
function getPairedEnglishFolderName(commandDef, folderName) {
    const folderParam = (commandDef.parameters || []).find(param => param.key === "Folder");
    if (!folderParam || !folderParam.options) return "";

    for (const option of folderParam.options) {
        if (typeof option === "string") {
            if (option === folderName) {
                return option;
            }
        } else if (option.value === folderName) {
            return option.english || option.value;
        }
    }

    return "";
}

// Collects selected switch parameters and parameter values for a command.
function collectSelectedParams(commandDef) {
    const commandName = commandDef.commandName;
    const selectedSwitchKeys = Array.from(document.querySelectorAll(`.${commandName}_param:checked`))
        .map(cb => cb.value);
    const selectedParams = {
        switch: [],
        switchSplit: [],
        multiCheckbox: {},
        multiCheckboxSplit: {}
    };
    const parameters = commandDef.parameters || [];

    for (const param of parameters) {
        if (param.type === "switch") {
            if (selectedSwitchKeys.includes(param.key)) {
                selectedParams.switch.push(param.key);
            }
            continue;
        }

        if (param.type === "switch-split") {
            if (selectedSwitchKeys.includes(param.key)) {
                selectedParams.switchSplit.push(param.key);
            }
            continue;
        }

        if (param.type !== "multi-checkbox-split" && param.type !== "multi-checkbox") {
            continue;
        }

        const selectedValues = Array.from(document.querySelectorAll(`.${commandName}_value_${param.key}:checked`))
            .map(cb => cb.value);

        if (selectedValues.length > 0) {
            if (param.type === "multi-checkbox") {
                selectedParams.multiCheckbox[param.key] = selectedValues;
            } else {
                selectedParams.multiCheckboxSplit[param.key] = selectedValues;
            }
        }
    }

    return selectedParams;
}

// Builds one formatted command line for a command, target, and parameter combination.
function buildCommandLine(commandDef, targetVarName, selectedParamsByType) {
    const EXPORT_CLIXML_ENCODING = "Export-CliXml -Encoding UTF8";
    
    const keepVarNameInFileName = document.getElementById("keepVarNameInFileName")?.checked ?? false;
    const expandVarNameInFileName = document.getElementById("expandVarNameInFileName")?.checked ?? false;
    const premiseSuffix = getPremiseSuffixForFileName();
    const selectedAllSwitchParams = [
        ...(selectedParamsByType.switch || []),
        ...(selectedParamsByType.switchSplit || [])
    ];
    const selectedSwitchSplitParams = selectedParamsByType.switchSplit || [];
    const selectedMultiCheckboxParams = selectedParamsByType.multiCheckbox || {};
    const selectedMultiCheckboxSplitParams = selectedParamsByType.multiCheckboxSplit || {};

    const commandName = commandDef.commandName || commandDef.CommandName;
    const fileNamePrefix = getfileNamePrefixFromCommand(commandDef.commandName);
    const bracedTargetVarName = "${" + targetVarName + "}";
    const switchSplitSuffix = selectedSwitchSplitParams.length > 0 ? `_${selectedSwitchSplitParams.map(param => param.replace(/^-+/, "")).join("_")}` : "";
    const varNameInFileName = !keepVarNameInFileName
        ? ""
        : (expandVarNameInFileName
            ? `${bracedTargetVarName}${premiseSuffix}`
            : `${targetVarName}${premiseSuffix}`);
    const varNameSegment = varNameInFileName ? `_${varNameInFileName}` : "";
    const switchParamText = selectedAllSwitchParams.length > 0
        ? ` ${selectedAllSwitchParams.join(" ")}`
        : "";
    const multiCheckboxParamText = Object.entries(selectedMultiCheckboxParams)
        .filter(([, values]) => Array.isArray(values) && values.length > 0)
        .map(([key, values]) => ` ${key} ${values.join(",")}`)
        .join("");
    const getVariantValue = key => selectedMultiCheckboxSplitParams[key] || "";

    if (commandDef.orgLevel) {
        const outputFile = `${fileNamePrefix}${premiseSuffix}${switchSplitSuffix}.xml`;
        return `${commandName}${switchParamText}${multiCheckboxParamText} | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
    }

    if (commandDef.commandName === "Get-Mailbox") {
        const mailboxSwitchParams = selectedAllSwitchParams.includes("-IncludeInactiveMailbox")
            && (
                selectedAllSwitchParams.includes("-SoftDeletedMailbox")
                || selectedAllSwitchParams.includes("-InactiveMailboxOnly")
            )
            ? selectedAllSwitchParams.filter(param => param !== "-IncludeInactiveMailbox")
            : selectedAllSwitchParams;

        if (mailboxSwitchParams.includes("-SoftDeletedMailbox") && mailboxSwitchParams.includes("-InactiveMailboxOnly")) {
            return "";
        }

        const mailboxSuffixParts = mailboxSwitchParams
            .filter(param => selectedSwitchSplitParams.includes(param) || param === "-IncludeInactiveMailbox")
            .map(param => {
                if (param === "-SoftDeletedMailbox") return "SoftDeleted";
                if (param === "-InactiveMailboxOnly") return "Inactive";
                if (param === "-IncludeInactiveMailbox") return "IncludeInactive";
                return param.replace(/^-+/, "");
            });
        const mailboxSwitchSplitSuffix = mailboxSuffixParts.length > 0
            ? `_${mailboxSuffixParts.join("_")}`
            : "";
        const mailboxSwitchParamText = mailboxSwitchParams.length > 0
            ? ` ${mailboxSwitchParams.join(" ")}`
            : "";
        const outputFile = `${fileNamePrefix}${varNameSegment}${mailboxSwitchSplitSuffix}.xml`;

        return `${commandName} $${targetVarName}${mailboxSwitchParamText} | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
    }

    if (commandDef.commandName === "Get-MailboxFolderStatistics") {
        const folderScopeValue = getVariantValue("-FolderScope");
        const folderScopeParam = folderScopeValue ? ` -FolderScope ${folderScopeValue}` : "";
        const folderScopeSuffix = folderScopeValue ? `_${folderScopeValue}` : "";
        const outputFile = `${fileNamePrefix}${varNameSegment}${folderScopeSuffix}${switchSplitSuffix}.xml`;

        return `${commandName} $${targetVarName}${folderScopeParam}${switchParamText}${multiCheckboxParamText} | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
    }

    if (commandDef.commandName === "Export-MailboxDiagnosticLogs") {
        const componentName = getVariantValue("-ComponentName");
        const componentParam = componentName ? ` -ComponentName ${componentName}` : "";
        const fileSuffix = componentName ? `_${componentName}` : "";
        const outputFile = `${fileNamePrefix}${varNameSegment}${fileSuffix}${switchSplitSuffix}.xml`;

        return `${commandName} $${targetVarName}${componentParam}${switchParamText}${multiCheckboxParamText} | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
    }

    if (commandDef.commandName === "Get-MailboxFolderPermission") {
        const folder = getVariantValue("Folder");
        const pairedEnglish = getPairedEnglishFolderName(commandDef, folder);
        const englishFolderName = pairedEnglish || folder;
        const selectedFolders = Array.from(
            document.querySelectorAll(`.${commandName}_value_Folder:checked`)
        ).map(cb => cb.value);
        const selectedTargetVarNames = Array.from(
            document.querySelectorAll(`.${commandName}_target:checked`)
        )
            .map(cb => {
                const index = parseInt(cb.value, 10);
                const input = document.querySelectorAll(".target-var-name")[index];
                return input ? input.value.trim() : "";
            })
            .filter(name => !!name);
        const isLastTarget = selectedTargetVarNames.length === 0
            ? true
            : targetVarName === selectedTargetVarNames[selectedTargetVarNames.length - 1];
        const isLastSelectedFolder = selectedFolders.length === 0
            ? true
            : folder === selectedFolders[selectedFolders.length - 1];
        const outputFile = `${fileNamePrefix}${varNameSegment}_${englishFolderName}${switchSplitSuffix}.xml`;
        const commandLine = folder === "ルート"
            ? `${commandName} ${bracedTargetVarName} | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`
            : `${commandName} "${bracedTargetVarName}:\\${folder}" | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;

        if (!(isLastTarget && isLastSelectedFolder)) {
            return commandLine;
        }

        const noteLines = selectedFolders
            .map(selectedFolder => {
                const selectedPairedEnglish = getPairedEnglishFolderName(commandDef, selectedFolder);
                if (!selectedPairedEnglish || selectedPairedEnglish === selectedFolder || selectedFolder === "ルート") {
                    return "";
                }

                return `# エラーになる場合は、":\\${selectedFolder}" の代わりに ":\\${selectedPairedEnglish}" をお試しください。`;
            })
            .filter(line => !!line);

        if (noteLines.length === 0) {
            return commandLine;
        }

        return `${commandLine}\n${noteLines.join("\n")}`;
    }

    if (commandDef.commandName === "Get-InboxRule") {
        const outputFile = `${fileNamePrefix}${varNameSegment}${switchSplitSuffix}.xml`;

        return `${commandName} -Mailbox $${targetVarName}${switchParamText}${multiCheckboxParamText} | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
    }

    if (commandDef.commandName === "Get-MgUser" || commandDef.commandName === "Get-MgUserLicenseDetail") {
        const outputFile = `${fileNamePrefix}${varNameSegment}${switchSplitSuffix}.xml`;

        return `${commandName} -UserId $${targetVarName}${switchParamText}${multiCheckboxParamText} | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
    }

    if (commandDef.commandName === "Get-UnifiedGroupLinks") {
        const linkTypeValue = getVariantValue("-LinkType");
        const linkTypeParam = linkTypeValue ? ` -LinkType ${linkTypeValue}` : "";
        const linkTypeSuffix = linkTypeValue ? `_${linkTypeValue}` : "";
        const outputFile = `${fileNamePrefix}${varNameSegment}${linkTypeSuffix}${switchSplitSuffix}.xml`;

        return `${commandName} $${targetVarName}${linkTypeParam}${switchParamText}${multiCheckboxParamText} | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
    }

    if (commandDef.commandName.startsWith("Get-AD") && commandDef.commandName !== "Get-ADPermission") {
        const ldapFilter = `"(|(UserPrincipalName=${bracedTargetVarName})(ProxyAddresses=*${bracedTargetVarName}*)(Name=${bracedTargetVarName})(DisplayName=${bracedTargetVarName})(mailNickName=${bracedTargetVarName}))"`
        const outputFile = `${fileNamePrefix}${varNameSegment}${switchSplitSuffix}.xml`;

        return `${commandName} -LDAPFilter ${ldapFilter}${switchParamText}${multiCheckboxParamText} -Properties * | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
    }
    if (commandDef.commandName === "Get-ADPermission") {
        const outputFile = `${fileNamePrefix}${varNameSegment}${switchSplitSuffix}.xml`;

        return `${commandName} (Get-Mailbox $${targetVarName})[0].DistinguishedName | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
    }
    
    const outputFile = `${fileNamePrefix}${varNameSegment}${switchSplitSuffix}.xml`;

    return `${commandName} $${targetVarName}${switchParamText}${multiCheckboxParamText} | ${EXPORT_CLIXML_ENCODING} "${outputFile}"`;
}

// Builds all command lines for a target, including archive and split-value variants.
function buildCommandLinesForTarget(commandDef, targetVarName, selectedParams) {
    const selectedSwitchParams = selectedParams.switch || [];
    const selectedSwitchSplitParams = selectedParams.switchSplit || [];
    const selectedMultiCheckboxParams = selectedParams.multiCheckbox || {};
    const selectedMultiCheckboxSplitParams = selectedParams.multiCheckboxSplit || {};
    
    let switchSplitVariants = [[]];
    const lines = [];

    // Expand switch-split keys into with/without variants.
    selectedSwitchSplitParams.forEach(key => {
        const expandedSets = [];
        switchSplitVariants.forEach(set => {
            expandedSets.push(set); // with
            expandedSets.push([...set, key]); // without
        });
        switchSplitVariants = expandedSets;
    });

    // Expand multi-checkbox-split keys into one command per selected value (cartesian product across keys).
    let multiCheckboxSplitVariants = [{}];
    Object.entries(selectedMultiCheckboxSplitParams).forEach(([key, values]) => {
        if (!Array.isArray(values)) {
            return;
        }

        const expandedVariants = [];
        values.forEach(value => {
            multiCheckboxSplitVariants.forEach(variant => {
                expandedVariants.push({
                    ...variant,
                    [key]: value
                });
            });
        });

        multiCheckboxSplitVariants = expandedVariants;
    });

    switchSplitVariants.forEach(switchSplitVariant => {
        multiCheckboxSplitVariants.forEach(multiCheckboxSplitVariant => {
            const line = buildCommandLine(commandDef, targetVarName, {
                switch: selectedSwitchParams,
                switchSplit: switchSplitVariant,
                multiCheckbox: selectedMultiCheckboxParams,
                multiCheckboxSplit: multiCheckboxSplitVariant
            });
            if (line) {
                lines.push(line);
            }
        });
    });

    return lines;
}

// Generates the output chunk for one command across all selected targets.
function buildCommandOutputChunk(selectedCommand, targetVarNames) {
    const commandName = selectedCommand.commandName;
    const cmdToggle = document.querySelector(`.cmd-toggle[data-cmd="${commandName}"]`);

    if (!cmdToggle || (!cmdToggle.checked && !cmdToggle.indeterminate)) {
        return "";
    }

    const selectedParams = collectSelectedParams(selectedCommand);
    if (!selectedParams) {
        return null;
    }

    const lines = [];

    if (selectedCommand.orgLevel) {
        lines.push(...buildCommandLinesForTarget(selectedCommand, "", selectedParams));
        return lines.length > 0 ? lines.join("\n") + "\n" : "";
    }

    document.querySelectorAll(`.${commandName}_target:checked`).forEach(cb => {
        const index = parseInt(cb.value, 10);
        const targetVarName = targetVarNames[index].value.trim();
        lines.push(...buildCommandLinesForTarget(selectedCommand, targetVarName, selectedParams));
    });

    if (lines.length === 0) {
        return "";
    }

    return lines.join("\n") + "\n";
}

// Validates inputs before generation and returns normalized context for output creation.
function validateGenerateInputs() {
    const psVarNamePattern = /^[A-Za-z_][A-Za-z0-9_]*$/;
    let varDefsText = "";
    const targetNameInputs = document.querySelectorAll(".target-var-name");
    const targetValueInputs = document.querySelectorAll(".target-var-value");

    for (const [index, targetNameInputEl] of targetNameInputs.entries()) {
        const targetValueInputEl = targetValueInputs[index];
        const targetNameText = targetNameInputEl.value.trim();
        const targetValueText = targetValueInputEl.value.trim();

        if (!targetValueText) {
            alert("値 (メール アドレス) が未入力です。");
            targetValueInputEl.focus();
            return null;
        }

        if (targetNameText && !psVarNamePattern.test(targetNameText)) {
            alert(`無効な変数名です: ${targetNameText}`);
            targetNameInputEl.focus();
            return null;
        }

        if (targetNameText && targetValueText) {
            varDefsText += `$${targetNameText} = "${targetValueText}"\n`;
        }
    }

    const selectedCommands = commandDefinition.filter(cmd => {
        const cmdToggle = document.querySelector(`.cmd-toggle[data-cmd="${cmd.commandName}"]`);
        return !!(cmdToggle && (cmdToggle.checked || cmdToggle.indeterminate));
    });

    for (const selectedCommand of selectedCommands) {
        for (const selectedParameter of (selectedCommand.parameters || [])) {
            if (selectedParameter.type !== "multi-checkbox" && selectedParameter.type !== "multi-checkbox-split") {
                continue;
            }

            if (selectedParameter.isMandatory === false) {
                continue;
            }

            const selectedValues = Array.from(
                document.querySelectorAll(`.${selectedCommand.commandName}_value_${selectedParameter.key}:checked`)
            );
            if (selectedValues.length > 0) {
                continue;
            }

            alert(`${selectedCommand.commandName} は ${selectedParameter.key} の選択が必須です。`);
            const firstOption = document.querySelector(`.${selectedCommand.commandName}_value_${selectedParameter.key}`);
            if (firstOption) {
                firstOption.focus();
            }
            return null;
        }
    }

    const warningMessages = [];
    const isKeepVarNameInFileNameEnabled = document.getElementById("keepVarNameInFileName")?.checked ?? false;
    if (targetNameInputs.length > 1 && !isKeepVarNameInFileNameEnabled) {
        warningMessages.push("複数の変数が定義されています。\n[ファイル名に変数を含める] が無効の場合、ファイル名が重複する可能性があります。");
    }

    const guidanceTypeValue = document.getElementById("guidanceType")?.value ?? "exo";
    const applicableGuidanceType = guidanceTypeValue === "exoSimple"
        ? "exo"
        : guidanceTypeValue === "onpremSimple"
            ? "onprem"
            : guidanceTypeValue;
    const guidanceTypeMap = {
        exo: "Exchange Online",
        exoSimple: "Exchange Online",
        onprem: "オンプレミス",
        onpremSimple: "オンプレミス",
        graph: "Microsoft Graph"
    };

    const includesExoAddressListCmd =
        (guidanceTypeValue === "exo" || guidanceTypeValue === "exoSimple")
        && selectedCommands.some(cmd =>
            cmd.commandName === "Get-GlobalAddressList"
            || cmd.commandName === "Get-OfflineAddressBook"
            || cmd.commandName === "Get-AddressList"
        );

    if (includesExoAddressListCmd) {
        warningMessages.push("以下のコマンドは Exchange Online では Address Lists の役割が必要です。\n- Get-GlobalAddressList\n- Get-OfflineAddressBook\n- Get-AddressList");
    }

    const unavailableCommands = selectedCommands
        .filter(cmd => cmd.applicable && !cmd.applicable.includes(applicableGuidanceType))
        .map(cmd => cmd.commandName);

    if (unavailableCommands.length > 0) {
        warningMessages.push(
            `${guidanceTypeMap[guidanceTypeValue]} で利用不可のコマンドがあります。\n- ${unavailableCommands.join("\n- ")}`
        );
    }

    for (const warningMessage of warningMessages) {
        alert(warningMessage);
    }

    const shouldIncludeVarDefsText = !(selectedCommands.length > 0 && selectedCommands.every(cmd => !!cmd.orgLevel));
    const includeGuidanceTextEnabled = !!document.getElementById("includeGuidanceText")?.checked;

    return {
        targetVarNames: targetNameInputs,
        varDefsText,
        shouldIncludeVarDefsText,
        selectedCommands,
        includeGuidanceTextEnabled
    };
}

// Validates inputs and composes the final PowerShell script output.
function generate() {
    handleVarNameInput(null, false);
    updateAllParentCheckboxes();
    let commandsText = "";

    const validationResult = validateGenerateInputs();
    if (!validationResult) {
        return;
    }

    const targetVarNames = validationResult.targetVarNames;
    const varDefsText = validationResult.varDefsText;
    const shouldIncludeVarDefsText = validationResult.shouldIncludeVarDefsText;
    const includeGuidanceTextEnabled = validationResult.includeGuidanceTextEnabled;

    for (const selectedCommand of validationResult.selectedCommands) {
        const chunk = buildCommandOutputChunk(selectedCommand, targetVarNames);
        if (chunk === null) {
            return;
        }

        commandsText += chunk;
    }

    if (includeGuidanceTextEnabled) {
        commandsText = generateGuidanceSteps(shouldIncludeVarDefsText ? varDefsText : "", commandsText);
    } else {
        commandsText = shouldIncludeVarDefsText
            ? varDefsText + "\n" + commandsText
            : commandsText;
    }

    const output = document.getElementById("output");
    if (output) {
        output.value = commandsText;
    }
}

// Copies the generated script text to the clipboard.
function copy() {
    const textarea = document.getElementById("output");
    textarea.select();
    document.execCommand("copy");
}

// Returns all saved presets from localStorage.
function getAllPresets() {
    try {
        return JSON.parse(localStorage.getItem("exchangeTemplatePresets") || "{}");
    } catch {
        return {};
    }
}

// Collects current UI state into a plain object for serialization.
function collectAllSettings() {
    const targets = [];
    document.querySelectorAll("#targetTable tbody tr").forEach(row => {
        const nameInput = row.querySelector(".target-var-name");
        const valueInput = row.querySelector(".target-var-value");
        targets.push({
            name: nameInput ? nameInput.value : "",
            value: valueInput ? valueInput.value : ""
        });
    });

    return {
        targets,
        commandState: saveState(),
        keepVarNameInFileName: document.getElementById("keepVarNameInFileName")?.checked ?? false,
        expandVarNameInFileName: document.getElementById("expandVarNameInFileName")?.checked ?? false,
        includeGuidanceText: document.getElementById("includeGuidanceText")?.checked ?? false,
        guidanceType: document.getElementById("guidanceType")?.value ?? "exo",
        includeTranscript: document.getElementById("includeTranscript")?.checked ?? false,
        exportToDesktop: document.getElementById("exportToDesktop")?.checked ?? false,
        exportPath: document.getElementById("exportPath")?.value ?? "",
        premiseSuffix: document.getElementById("premiseSuffix")?.checked ?? false
    };
}

// Saves current settings to localStorage under the given name.
function saveSettingsToStorage(name) {
    const presets = getAllPresets();
    presets[name] = collectAllSettings();
    presets[name].name = name;
    localStorage.setItem("exchangeTemplatePresets", JSON.stringify(presets));
    updatePresetSelect();
}

// Applies a settings object to the current UI.
function applySettings(settings, presetName = "") {
    document.getElementById("presetName").value = presetName;

    const targetTableBody = document.querySelector("#targetTable tbody");
    if (targetTableBody && Array.isArray(settings.targets) && settings.targets.length > 0) {
        targetTableBody.innerHTML = "";
        settings.targets.forEach(target => {
            const row = targetTableBody.insertRow();
            row.innerHTML = `
                <td><input class="target-var-name" onchange="handleVarNameInput(this)"></td>
                <td><input class="target-var-value"></td>
                <td><button type="button" onclick="removeTarget(this)">削除</button></td>
            `;
            row.querySelector(".target-var-name").value = target.name || "";
            row.querySelector(".target-var-value").value = target.value || "";
        });
    }

    const setChecked = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.checked = !!val;
    };
    const setValue = (id, val) => {
        const el = document.getElementById(id);
        if (el && val !== undefined) el.value = val;
    };

    setChecked("keepVarNameInFileName", settings.keepVarNameInFileName);
    setChecked("expandVarNameInFileName", settings.expandVarNameInFileName);
    setChecked("includeGuidanceText", settings.includeGuidanceText);
    setValue("guidanceType", settings.guidanceType);
    setChecked("includeTranscript", settings.includeTranscript);
    setChecked("exportToDesktop", settings.exportToDesktop);
    setValue("exportPath", settings.exportPath);
    setChecked("premiseSuffix", settings.premiseSuffix);

    handleVarNameInput(null, false);
    normalizeExportPathInput();
    toggleVarNameInFileNameOptions();
    toggleGuidanceTypeOptions();

    // Build a default state for all current commands, then overlay with saved state
    // to safely handle commands added after the preset was saved.
    const defaultCommandState = {};
    commandDefinition.forEach(cmd => {
        defaultCommandState[cmd.commandName] = {
            enabled: false, targets: [], params: [], parameterValues: {},
            expandedTargets: false, expandedParams: false
        };
    });
    const mergedCommandState = Object.assign(defaultCommandState, settings.commandState || {});

    renderCommands();
    restoreState(mergedCommandState);
}

// Loads settings from localStorage by name and applies them to the UI.
function loadSettingsFromStorage(name) {
    const presets = getAllPresets();
    const settings = presets[name];
    if (!settings) return;
    applySettings(settings, name);
}

// Updates the preset dropdown to reflect current localStorage contents.
function updatePresetSelect() {
    const select = document.getElementById("presetSelect");
    if (!select) return;
    const presets = getAllPresets();
    const currentValue = select.value;
    select.innerHTML = '<option value="">-- 設定を選択 --</option>';
    Object.keys(presets).sort().forEach(name => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        select.appendChild(option);
    });
    if (currentValue && presets[currentValue]) {
        select.value = currentValue;
    }
}

// Saves the current settings under the name entered in the preset name input.
function savePreset() {
    const nameInput = document.getElementById("presetName");
    const name = nameInput ? nameInput.value.trim() : "";
    if (!name) {
        alert("設定名を入力してください。");
        if (nameInput) nameInput.focus();
        return;
    }
    const presets = getAllPresets();
    if (presets[name] && !confirm(`"${name}" は既に存在します。上書きしますか？`)) {
        return;
    }
    saveSettingsToStorage(name);
    const select = document.getElementById("presetSelect");
    if (select) select.value = name;
}

// Loads the preset selected in the dropdown.
function loadPreset() {
    const select = document.getElementById("presetSelect");
    const name = select ? select.value : "";
    if (!name) {
        alert("読み込む設定を選択してください。");
        return;
    }
    loadSettingsFromStorage(name);
}

// Sets the current settings as the default preset.
function setDefault() {
    const select = document.getElementById("presetSelect");
    const name = select ? select.value : "";
    if (!name) {
        alert("読み込む設定を選択してください。");
        return;
    }
    const presets = getAllPresets();
    const settings = presets[name];
    if (!settings) return;
    Object.keys(presets).forEach(key => {
        presets[key].isDefault = key === name;
    });
    localStorage.setItem("exchangeTemplatePresets", JSON.stringify(presets));
    alert(`"${name}" をデフォルト設定にしました。`);
}

// Deletes the preset selected in the dropdown.
function deleteSelectedPreset() {
    const select = document.getElementById("presetSelect");
    const name = select ? select.value : "";
    if (!name) {
        alert("削除する設定を選択してください。");
        return;
    }
    if (!confirm(`"${name}" を削除しますか？`)) return;
    const presets = getAllPresets();
    delete presets[name];
    localStorage.setItem("exchangeTemplatePresets", JSON.stringify(presets));
    updatePresetSelect();
}

window.onload = function () {
    handleVarNameInput(null, false);
    normalizeExportPathInput();
    toggleVarNameInFileNameOptions();
    toggleExportPathInput();
    toggleGuidanceTypeOptions();
    renderTagFilterChips();
    const keepVarNameInFileName = document.getElementById("keepVarNameInFileName");
    if (keepVarNameInFileName) {
        keepVarNameInFileName.addEventListener("change", toggleVarNameInFileNameOptions);
    }
    const includeGuidanceText = document.getElementById("includeGuidanceText");
    if (includeGuidanceText) {
        includeGuidanceText.addEventListener("change", toggleGuidanceTypeOptions);
    }
    const guidanceTypeElement = document.getElementById("guidanceType");
    if (guidanceTypeElement) {
        guidanceTypeElement.addEventListener("change", toggleGuidanceTypeOptions);
    }

    const commandSearchInput = document.getElementById("commandSearch");
    if (commandSearchInput) {
        commandSearchInput.value = commandUiState.searchText;
    }

    updatePresetSelect();

    const defaultPresetEntry = Object.entries(getAllPresets()).find(([, preset]) => preset.isDefault);
    if (defaultPresetEntry) {
        const [defaultName, defaultSettings] = defaultPresetEntry;
        applySettings(defaultSettings, defaultName);
        const select = document.getElementById("presetSelect");
        if (select) select.value = defaultName;
    } else {
        renderCommands();
    }
};