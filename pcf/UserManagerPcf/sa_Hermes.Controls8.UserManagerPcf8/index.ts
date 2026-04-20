import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface UserRecordViewModel {
    id: string;
    datasetId: string;
    userAccount: string;
    typeUser: string;
    titleUser: string;
    role: string;
    lieux: string[];
    createdOn?: Date;
}

type SortDir = "asc" | "desc";
type SortKey = "userAccount" | "typeUser" | "titleUser" | "role" | "createdOn" | "lieux";

interface ColFilter {
    userAccount: string;
    typeUser: string;
    titleUser: string;
    role: string;
    lieux: string;
}

export class UserManagerPcf15 implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private root!: HTMLDivElement;
    private notifyOutputChanged!: () => void;
    private context!: ComponentFramework.Context<IInputs>;

    private selectedIds = new Set<string>();

    private sortKey: SortKey = "userAccount";
    private sortDir: SortDir = "asc";
    private colFilters: ColFilter = { userAccount: "", typeUser: "", titleUser: "", role: "", lieux: "" };

    private currentPage = 1;
    private readonly PAGE_SIZE = 10;
    private activeFilterKey: keyof ColFilter | null = null;

    private outputAction = "";
    private outputSelectedIds = "[]";
    private outputFormData = "";
    private outputEventToken = "";
    private outputSaveSignal = "";
    private outputSearchTerm = "";
    private outputSelectedId = ""; // New: stores SharePoint ID of selected element
    private outputSequence = 0;

    private modalOverlay: HTMLDivElement | null = null;
    private searchTermTimer: ReturnType<typeof setTimeout> | null = null;
    private renderTimer: ReturnType<typeof setTimeout> | null = null;
    private lastInteractionTime = 0;
    private updateDropdownCallback: (() => void) | null = null;
    private lastRefreshSignal = "";

    // ─── Lifecycle ────────────────────────────────────────────────────────────

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        console.log("DEBUG - INIT called - Component is starting");
        
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;

        this.applyStyles(container, {
            height: "100%",
            overflowY: "auto",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            color: "#344054",
            boxSizing: "border-box",
        });

        this.root = this.el("div", {
            background: "#ffffff",
            borderRadius: "8px",
            border: "1px solid #e4e7ec",
            overflow: "hidden",
            minHeight: "200px",
        }) as HTMLDivElement;

        container.appendChild(this.root);

        // Force pointer-events on container so Power Apps doesn't suppress clicks
        container.style.pointerEvents = "auto";
        container.style.touchAction = "auto";
        // Track last interaction so renders don't destroy DOM during clicks
        container.addEventListener("pointerdown", () => { this.lastInteractionTime = Date.now(); }, { capture: true, passive: true });
        container.addEventListener("mousedown",   () => { this.lastInteractionTime = Date.now(); }, { capture: true, passive: true });

        if (context.parameters.users?.paging) {
            context.parameters.users.paging.setPageSize(5000);
        }

        // Notify Power Apps of initial values
        this.notifyOutputChanged();
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        console.log("DEBUG - updateView called");
        this.context = context;
        const users = context.parameters.users;

        // Check if refresh signal has changed
        const currentRefreshSignal = context.parameters.refreshSignal?.raw || "";
        if (currentRefreshSignal !== this.lastRefreshSignal && currentRefreshSignal.trim() !== "") {
            this.lastRefreshSignal = currentRefreshSignal;
            console.log("DEBUG - Refresh signal detected:", currentRefreshSignal);
            // Force refresh of the dataset
            if (users.refresh) {
                users.refresh();
            }
            return;
        }

        if (!users.loading && users.paging?.hasNextPage) {
            users.paging.loadNextPage();
            return;
        }

        // Immediately refresh the user dropdown if the form is open (usersJson may have changed)
        this.updateDropdownCallback?.();

        // Debounce: batch rapid updateView calls; delay render if user is actively clicking
        if (this.renderTimer !== null) clearTimeout(this.renderTimer);
        const delay = Math.max(16, 500 - (Date.now() - this.lastInteractionTime));
        this.renderTimer = setTimeout(() => {
            this.renderTimer = null;
            this.render(this.context, this.context.parameters.users);
        }, delay);
    }

    public getOutputs(): IOutputs {
        const outputs = {
            requestedAction: this.outputAction || undefined,
            selectedIds: this.outputSelectedIds,
            formData: this.outputFormData || undefined,
            eventToken: this.outputEventToken,
            saveSignal: this.outputSaveSignal || undefined,
            searchTerm: this.outputSearchTerm || undefined,
            selectedId: this.outputSelectedId,
        };
        
        console.log("DEBUG - getOutputs called - FULL DUMP:", {
            selectedIds: this.outputSelectedIds,
            selectedId: this.outputSelectedId,
            eventToken: this.outputEventToken,
            requestedAction: this.outputAction,
            internalSelectedIdsCount: this.selectedIds.size,
            internalSelectedIdsArray: Array.from(this.selectedIds),
            allOutputs: outputs
        });
        
        return outputs;
    }

    public destroy(): void {
        this.modalOverlay?.remove();
    }

    // ─── Render ───────────────────────────────────────────────────────────────

    private render(context: ComponentFramework.Context<IInputs>, users: ComponentFramework.PropertyTypes.DataSet): void {
        console.log("DEBUG - Render START:", {
            loading: users.loading,
            sortedRecordIdsCount: users.sortedRecordIds?.length || 0,
            hasError: !!users.error
        });
        
        this.root.innerHTML = "";
        try {
            if (users.error) {
                console.error("DEBUG - Dataset error:", users.error);
                this.root.appendChild(this.createMessage("Error loading data."));
                return;
            }
            // ...existing code...
            // Add click handler for row selection to set outputSelectedId
            // (Assume you have a function or code block where you render each row or selectable element)
            // Example:
            // rowElement.addEventListener('click', () => {
            //     this.outputSelectedId = user.id; // user.id = SharePoint ID
            //     this.notifyOutputChanged();
            // });
            // Place this logic where you handle row/item selection in your actual render code.

            const allRecords = this.getRecords(context, users);
            console.log("DEBUG - Render called:", {
                totalRecords: allRecords.length,
                recordsSample: allRecords.slice(0, 3),
                currentSelectedCount: this.selectedIds.size
            });
            const filtered = this.applyFilters(allRecords);
            const sorted = this.applySort(filtered);

            const totalPages = Math.max(1, Math.ceil(sorted.length / this.PAGE_SIZE));
            if (this.currentPage > totalPages) this.currentPage = totalPages;

            const pageStart = (this.currentPage - 1) * this.PAGE_SIZE;
            const paginated = sorted.slice(pageStart, pageStart + this.PAGE_SIZE);

            this.root.appendChild(this.createToolbar(context, allRecords, sorted));

            if (users.loading) {
                this.root.appendChild(this.createMessage("Loading…"));
                return;
            }

            if (sorted.length === 0) {
                this.root.appendChild(this.createMessage(
                    allRecords.length === 0 ? "No users found." : "No results for these filters."
                ));
                return;
            }

            this.root.appendChild(this.createTable(context, paginated, allRecords));
            this.root.appendChild(this.createPagination(context, sorted.length, totalPages));

            // Restore focus on the active filter input after DOM rebuild
            if (this.activeFilterKey) {
                const input = this.root.querySelector<HTMLInputElement>(`input[data-filter-key="${this.activeFilterKey}"]`);
                if (input) {
                    input.focus();
                    const len = input.value.length;
                    input.setSelectionRange(len, len);
                }
                this.activeFilterKey = null;
            }
        } catch (err) {
            this.root.innerHTML = "";
            this.root.appendChild(this.createMessage(`Error: ${err instanceof Error ? err.message : String(err)}`));
        }
    }

    // ─── Data ─────────────────────────────────────────────────────────────────

    private getRecords(context: ComponentFramework.Context<IInputs>, users: ComponentFramework.PropertyTypes.DataSet): UserRecordViewModel[] {
        const userAccountCol = this.resolveCol(users, context.parameters.userAccountColumn?.raw, ["UserAccount"]);
        const typeUserCol    = this.resolveCol(users, context.parameters.typeUserColumn?.raw,    ["TypeUser"]);
        const titleUserCol   = this.resolveCol(users, context.parameters.titleUserColumn?.raw,   ["TitleUser"]);
        const roleCol        = this.resolveCol(users, context.parameters.roleColumn?.raw,        ["Role"]);
        const lieuxCol       = this.resolveCol(users, context.parameters.lieuxColumn?.raw,       ["Lieux"]);
        const createdOnCol   = this.resolveCol(users, context.parameters.createdOnColumn?.raw,   ["Created", "createdon", "Créé"]);

        console.log("DEBUG - Available columns:", users.columns.map(c => c.name));
        console.log("DEBUG - Resolved columns:", { userAccountCol, typeUserCol, titleUserCol, roleCol, lieuxCol, createdOnCol });

        return users.sortedRecordIds
            .map((datasetId) => {
                const record = users.records[datasetId];
                if (!record) return null;

                const rawCreated = createdOnCol ? record.getValue(createdOnCol) : undefined;
                let createdOn: Date | undefined;
                if (rawCreated instanceof Date) {
                    createdOn = rawCreated;
                } else if (typeof rawCreated === "string" && rawCreated.trim()) {
                    const parsed = new Date(rawCreated);
                    if (!isNaN(parsed.getTime())) createdOn = parsed;
                }

                return {
                    id: this.getItemId(record),
                    datasetId,
                    userAccount: this.getVal(record, userAccountCol),
                    typeUser:    this.getVal(record, typeUserCol),
                    titleUser:   this.getVal(record, titleUserCol),
                    role:        this.getVal(record, roleCol),
                    lieux:       this.parseLieux(record, lieuxCol),
                    createdOn,
                } as UserRecordViewModel;
            })
            .filter((r): r is UserRecordViewModel => r !== null);
    }

    private getItemId(record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord): string {
        for (const col of ["ID", "Id", "id"]) {
            try {
                const v = record.getValue(col);
                if (v !== null && v !== undefined) return String(v);
            } catch { /* ignore */ }
        }
        return record.getRecordId();
    }

    private parseLieux(record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, col?: string): string[] {
        if (!col) return [];

        const extractFromArray = (arr: unknown): string[] | null => {
            if (!Array.isArray(arr)) return null;
            const vals = (arr as Record<string, unknown>[])
                .map(item => String(item["Value"] ?? item["value"] ?? item["Title"] ?? "").trim())
                .filter(v => v.length > 0);
            return vals.length > 0 ? vals : null;
        };

        const extractFromString = (str: string): string[] | null => {
            const s = str.trim();
            if (!s.startsWith("[")) return null;
            // Try JSON.parse first
            try {
                const parsed = JSON.parse(s);
                const r = extractFromArray(parsed);
                if (r) return r;
            } catch { /* fall through to regex */ }
            // Regex fallback: extract "Value":"..." pairs (handles malformed JSON)
            const matches = s.match(/"[Vv]alue"\s*:\s*"([^"]+)"/g);
            if (matches) {
                const vals = matches
                    .map(m => { const x = m.match(/"[Vv]alue"\s*:\s*"([^"]+)"/); return x ? x[1].trim() : ""; })
                    .filter(v => v.length > 0);
                if (vals.length > 0) return vals;
            }
            // It looks like JSON but we couldn't parse it — don't show raw garbage
            return [];
        };

        // 1. getValue may return a real JS array or a JSON string
        try {
            const raw = record.getValue(col);
            if (Array.isArray(raw)) {
                const r = extractFromArray(raw);
                if (r) return r;
            } else if (raw !== null && raw !== undefined && typeof raw !== "object") {
                const r = extractFromString(String(raw));
                if (r !== null) return r;
            }
        } catch { /* ignore */ }

        // 2. getFormattedValue
        const formatted = record.getFormattedValue(col);
        if (!formatted || !formatted.trim()) return [];
        const fmtResult = extractFromString(formatted);
        if (fmtResult !== null) return fmtResult;

        // 3. Plain "Val1; Val2"
        return formatted.split(/;\s*/).map(s => s.trim()).filter(s => s.length > 0);
    }

    private getVal(record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, col?: string): string {
        if (!col) return "";
        const fmt = record.getFormattedValue(col);
        if (fmt) return fmt;
        const raw = record.getValue(col);
        if (raw === null || raw === undefined) return "";
        if (typeof raw === "object" && "name" in (raw as object)) return String((raw as { name: unknown }).name ?? "");
        return String(raw);
    }

    private resolveCol(users: ComponentFramework.PropertyTypes.DataSet, configured: string | null | undefined, fallbacks: string[]): string | undefined {
        const byName = new Map<string, string>();
        const byDisplay = new Map<string, string>();
        users.columns.forEach(c => {
            byName.set(c.name.toLowerCase(), c.name);
            byDisplay.set(c.displayName.toLowerCase(), c.name);
        });

        if (configured?.trim()) {
            const key = configured.trim().toLowerCase();
            return byName.get(key) ?? byDisplay.get(key);
        }

        for (const fb of fallbacks) {
            const key = fb.toLowerCase();
            const found = byName.get(key) ?? byDisplay.get(key);
            if (found) return found;
        }

        return undefined;
    }

    private applyFilters(records: UserRecordViewModel[]): UserRecordViewModel[] {
        const f = this.colFilters;
        const norm = (s: string) => s.toLowerCase().trim();
        return records.filter(r =>
            (!f.userAccount || norm(r.userAccount).includes(norm(f.userAccount))) &&
            (!f.typeUser    || norm(r.typeUser).includes(norm(f.typeUser)))        &&
            (!f.titleUser   || norm(r.titleUser).includes(norm(f.titleUser)))      &&
            (!f.role        || norm(r.role).includes(norm(f.role)))                &&
            (!f.lieux       || r.lieux.some(l => norm(l).includes(norm(f.lieux))))
        );
    }

    private applySort(records: UserRecordViewModel[]): UserRecordViewModel[] {
        const key = this.sortKey;
        const dir = this.sortDir === "asc" ? 1 : -1;
        return [...records].sort((a, b) => {
            if (key === "createdOn") {
                const at = a.createdOn?.getTime() ?? 0;
                const bt = b.createdOn?.getTime() ?? 0;
                return (at - bt) * dir;
            }
            if (key === "lieux") {
                return a.lieux.join(", ").localeCompare(b.lieux.join(", ")) * dir;
            }
            return String(a[key] ?? "").localeCompare(String(b[key] ?? "")) * dir;
        });
    }

    // ─── Toolbar ──────────────────────────────────────────────────────────────

    private createToolbar(context: ComponentFramework.Context<IInputs>, allRecords: UserRecordViewModel[], filtered: UserRecordViewModel[]): HTMLDivElement {
        const bar = this.el("div", {
            display: "flex",
            alignItems: "center",
            gap: "8px",
            padding: "12px 16px",
            borderBottom: "1px solid #e4e7ec",
            flexWrap: "wrap",
        }) as HTMLDivElement;

        // Title hidden per request
        // const countLabel = filtered.length < allRecords.length
        //     ? `Users (${filtered.length} / ${allRecords.length})`
        //     : `Users (${allRecords.length})`;
        // bar.appendChild(this.el("span", {
        //     fontWeight: "600",
        //     fontSize: "24px",
        //     color: "#101828",
        //     flex: "1",
        // }, countLabel));

        const hasActiveFilter = Object.values(this.colFilters).some(v => v.trim() !== "");
        if (hasActiveFilter) {
            const clearBtn = this.makeBtn("✕ Clear filters", "#fdecea", "#c62828", () => {
                this.colFilters = { userAccount: "", typeUser: "", titleUser: "", role: "", lieux: "" };
                this.currentPage = 1;
                this.render(context, context.parameters.users);
            });
            bar.appendChild(clearBtn);
        }

        // Add, Edit, Delete buttons hidden per request
        // const addBtn = this.makeBtn("+ Add", "#4f46e5", "#ffffff", () => this.openModal(context, null, allRecords));
        // bar.appendChild(addBtn);

        // const editBtn = this.makeBtn("✎ Edit", "#e3f2fd", "#1565c0", () => {
        //     const rec = allRecords.find(r => this.selectedIds.has(r.id));
        //     if (rec) this.openModal(context, rec, allRecords);
        // });
        // this.setDisabled(editBtn, this.selectedIds.size !== 1);
        // bar.appendChild(editBtn);

        // const delBtn = this.makeBtn("✕ Delete", "#fdecea", "#c62828", () => this.openDeleteConfirm(context));
        // this.setDisabled(delBtn, this.selectedIds.size === 0);
        // bar.appendChild(delBtn);

        return bar;
    }

    private makeBtn(label: string, bg: string, color: string, onClick: () => void): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.type = "button";
        this.applyStyles(btn, {
            background: bg,
            color,
            border: "1px solid transparent",
            borderRadius: "6px",
            padding: "7px 14px",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "13px",
            fontWeight: "600",
            cursor: "pointer",
            transition: "opacity 0.15s",
        });
        btn.textContent = label;
        btn.addEventListener("click", (e) => { e.stopPropagation(); onClick(); });
        return btn;
    }

    private setDisabled(btn: HTMLButtonElement, disabled: boolean): void {
        btn.disabled = disabled;
        btn.style.opacity = disabled ? "0.4" : "1";
        btn.style.cursor = disabled ? "not-allowed" : "pointer";
    }

    // ─── Table ────────────────────────────────────────────────────────────────

    // Column definitions used for sort headers and filter inputs
    private readonly COLS: { key: SortKey; label: string; filterKey: keyof ColFilter | null }[] = [
        { key: "userAccount", label: "User",      filterKey: "userAccount" },
        { key: "typeUser",    label: "Type",      filterKey: "typeUser"    },
        { key: "titleUser",   label: "Title",     filterKey: "titleUser"   },
        { key: "role",        label: "Role(s)",   filterKey: "role"        },
        { key: "lieux",       label: "Store(s)",  filterKey: "lieux"       },
    ];

    private createTable(context: ComponentFramework.Context<IInputs>, records: UserRecordViewModel[], allRecords: UserRecordViewModel[]): HTMLDivElement {
        const wrapper = this.el("div", { overflowX: "auto" }) as HTMLDivElement;

        const table = document.createElement("table");
        this.applyStyles(table, { width: "100%", borderCollapse: "collapse", fontSize: "13px" });

        // ── Thead: sort row + filter row
        const thead = document.createElement("thead");

        // — Sort row
        const sortRow = document.createElement("tr");
        this.applyStyles(sortRow, { background: "#f9fafb", borderBottom: "1px solid #e4e7ec" });

        const allChecked = records.length > 0 && records.every(r => this.selectedIds.has(r.id));
        const cbAll = this.makeCheckbox(allChecked, (checked) => {
            console.log("DEBUG - Select All CLICKED (start):", { checked: checked });
            
            records.forEach(r => checked ? this.selectedIds.add(r.id) : this.selectedIds.delete(r.id));
            const selectedArray = Array.from(this.selectedIds);
            this.outputSelectedIds = JSON.stringify(selectedArray);
            this.outputSelectedId = selectedArray.join(","); // Comma-separated for easy Power Apps use
            this.outputSequence++;
            this.outputEventToken = `selectedIds_${Date.now()}_${this.outputSequence}`;
            const timestamp = Date.now();
            this.outputAction = JSON.stringify({
                action: "selectionChanged",
                selectedIds: selectedArray,
                count: selectedArray.length,
                timestamp: timestamp,
                sequence: this.outputSequence
            });
            
            console.log("DEBUG - Select All toggled (after update):", {
                checked: checked,
                selectedCount: this.selectedIds.size,
                outputSelectedIds: this.outputSelectedIds,
                eventToken: this.outputEventToken,
                requestedAction: this.outputAction
            });
            
            // Double notify: Power Apps may miss a single notify
            this.notifyOutputChanged();
            setTimeout(() => { this.notifyOutputChanged(); }, 30);
        });
        const thCb = this.makeTh(""); thCb.style.width = "40px"; thCb.appendChild(cbAll);
        sortRow.appendChild(thCb);

        this.COLS.forEach(col => {
            const th = this.makeTh("");
            const btn = document.createElement("button");
            btn.type = "button";
            const isActive = this.sortKey === col.key;
            this.applyStyles(btn, {
                background: "none",
                border: "none",
                cursor: "pointer",
                fontFamily: "Inter, Segoe UI, sans-serif",
                fontSize: "11px",
                fontWeight: "700",
                color: isActive ? "#4f46e5" : "#667085",
                textTransform: "uppercase",
                letterSpacing: "0.05em",
                padding: "0",
                display: "inline-flex",
                alignItems: "center",
                gap: "4px",
                whiteSpace: "nowrap",
            });
            const arrow = isActive ? (this.sortDir === "asc" ? " ▲" : " ▼") : "";
            btn.textContent = col.label + arrow;
            btn.title = `Trier par ${col.label}`;
            btn.addEventListener("click", (e) => {
                e.stopPropagation();
                if (this.sortKey === col.key) {
                    this.sortDir = this.sortDir === "asc" ? "desc" : "asc";
                } else {
                    this.sortKey = col.key;
                    this.sortDir = "asc";
                }
                this.render(context, context.parameters.users);
            });
            th.appendChild(btn);
            sortRow.appendChild(th);
        });

        thead.appendChild(sortRow);

        // — Filter row
        const filterRow = document.createElement("tr");
        this.applyStyles(filterRow, { background: "#f5f7ff", borderBottom: "2px solid #e4e7ec" });

        // Empty cell for checkbox column
        const tdCbEmpty = document.createElement("td");
        filterRow.appendChild(tdCbEmpty);

        this.COLS.forEach(col => {
            const td = document.createElement("td");
            this.applyStyles(td, { padding: "4px 8px" });

            if (col.filterKey !== null) {
                const fKey = col.filterKey;
                const input = document.createElement("input");
                this.applyStyles(input, {
                    width: "100%",
                    padding: "4px 7px",
                    border: "1px solid #d0d5dd",
                    borderRadius: "5px",
                    fontSize: "12px",
                    fontFamily: "Inter, Segoe UI, sans-serif",
                    color: "#344054",
                    outline: "none",
                    boxSizing: "border-box",
                    background: this.colFilters[fKey] ? "#eff6ff" : "#ffffff",
                });
                input.placeholder = "Filter…";
                input.value = this.colFilters[fKey];
                input.dataset.filterKey = fKey;

                let debounceTimer: ReturnType<typeof setTimeout>;
                input.addEventListener("input", () => {
                    clearTimeout(debounceTimer);
                    this.colFilters[fKey] = input.value; // update immediately to preserve value on re-render
                    debounceTimer = setTimeout(() => {
                        this.activeFilterKey = fKey;
                        this.currentPage = 1;
                        this.render(context, context.parameters.users);
                    }, 300);
                });
                input.addEventListener("focus", () => { input.style.borderColor = "#4f46e5"; });
                input.addEventListener("blur",  () => { input.style.borderColor = "#d0d5dd"; });
                td.appendChild(input);
            }

            filterRow.appendChild(td);
        });

        thead.appendChild(filterRow);
        table.appendChild(thead);

        // ── Body
        const tbody = document.createElement("tbody");
        records.forEach((record, idx) => tbody.appendChild(this.createRow(context, record, idx)));
        table.appendChild(tbody);

        wrapper.appendChild(table);
        return wrapper;
    }

    private makeTh(label: string): HTMLTableCellElement {
        const th = document.createElement("th");
        this.applyStyles(th, {
            padding: "9px 12px",
            textAlign: "left",
            fontWeight: "600",
            fontSize: "11px",
            color: "#667085",
            textTransform: "uppercase",
            letterSpacing: "0.05em",
            whiteSpace: "nowrap",
        });
        if (label) th.textContent = label;
        return th;
    }

    private createRow(context: ComponentFramework.Context<IInputs>, record: UserRecordViewModel, idx: number): HTMLTableRowElement {
        const tr = document.createElement("tr");
        const bgEven = "#ffffff", bgOdd = "#fafafa", bgHover = "#f0f4ff", bgSelected = "#eff6ff";
        const isSelected = this.selectedIds.has(record.id);
        this.applyStyles(tr, {
            borderBottom: "1px solid #f2f4f7",
            background: isSelected ? bgSelected : idx % 2 === 0 ? bgEven : bgOdd,
        });
        tr.addEventListener("mouseenter", () => { if (!this.selectedIds.has(record.id)) tr.style.background = bgHover; });
        tr.addEventListener("mouseleave", () => { if (!this.selectedIds.has(record.id)) tr.style.background = idx % 2 === 0 ? bgEven : bgOdd; });

        // Sélection d'une ligne : expose l'ID SharePoint sélectionné
        tr.addEventListener("click", () => {
            // Don't override selectedId when it contains multiple selected IDs from checkboxes
            if (this.selectedIds.size === 0) {
                this.outputSelectedId = record.id;
                this.notifyOutputChanged();
            }
        });

        // Checkbox
        const tdCb = this.makeTd();
        const cb = this.makeCheckbox(this.selectedIds.has(record.id), (checked) => {
            console.log("DEBUG - Checkbox CLICKED (start):", { recordId: record.id, checked: checked });
            
            if (checked) { this.selectedIds.add(record.id); } else { this.selectedIds.delete(record.id); }
            const selectedArray = Array.from(this.selectedIds);
            this.outputSelectedIds = JSON.stringify(selectedArray);
            this.outputSelectedId = selectedArray.join(","); // Comma-separated for easy Power Apps use
            this.outputSequence++;
            this.outputEventToken = `selectedIds_${Date.now()}_${this.outputSequence}`;
            const timestamp = Date.now();
            this.outputAction = JSON.stringify({
                action: "selectionChanged",
                selectedIds: selectedArray,
                count: selectedArray.length,
                timestamp: timestamp,
                sequence: this.outputSequence
            });
            
            console.log("DEBUG - Checkbox toggled (after update):", {
                recordId: record.id,
                checked: checked,
                selectedIdsArray: selectedArray,
                outputSelectedIds: this.outputSelectedIds,
                eventToken: this.outputEventToken,
                requestedAction: this.outputAction
            });
            
            // Double notify: Power Apps may miss a single notify
            this.notifyOutputChanged();
            setTimeout(() => { this.notifyOutputChanged(); }, 30);
        });
        tdCb.appendChild(cb);
        tr.appendChild(tdCb);

        // UserAccount
        const tdUser = this.makeTd();
        tdUser.appendChild(this.createPersonCell(record.userAccount));
        tr.appendChild(tdUser);

        // TypeUser
        const tdType = this.makeTd();
        if (record.typeUser) {
            const c = this.getValueBadgeColor(record.typeUser);
            tdType.appendChild(this.makeBadge(record.typeUser, c.bg, c.fg));
        }
        tr.appendChild(tdType);

        // TitleUser
        const tdTitle = this.makeTd();
        if (record.titleUser) {
            const c = this.getValueBadgeColor(record.titleUser);
            tdTitle.appendChild(this.makeBadge(record.titleUser, c.bg, c.fg));
        }
        tr.appendChild(tdTitle);

        // Role
        const tdRole = this.makeTd();
        if (record.role) {
            const c = this.getValueBadgeColor(record.role);
            tdRole.appendChild(this.makeBadge(record.role, c.bg, c.fg));
        }
        tr.appendChild(tdRole);

        // Lieux
        const tdLieux = this.makeTd();
        tdLieux.style.maxWidth = "240px";
        const lieuxWrap = this.el("div", { display: "flex", flexWrap: "wrap", gap: "4px" }) as HTMLDivElement;
        record.lieux.forEach(lieu => lieuxWrap.appendChild(this.createLieuChip(context, record, lieu)));
        tdLieux.appendChild(lieuxWrap);
        tr.appendChild(tdLieux);

        return tr;
    }

    private makeTd(): HTMLTableCellElement {
        const td = document.createElement("td");
        this.applyStyles(td, { padding: "10px 12px", verticalAlign: "middle" });
        return td;
    }

    private makeCheckbox(checked: boolean, onChange: (v: boolean) => void): HTMLInputElement {
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = checked;
        cb.style.cursor = "pointer";
        cb.addEventListener("change", () => onChange(cb.checked));
        return cb;
    }

    private resolveUserList(context: ComponentFramework.Context<IInputs>, allRecords: UserRecordViewModel[]): { display: string; email: string }[] {
        const raw = context.parameters.usersJson?.raw?.trim();
        if (raw && raw.startsWith("[")) {
            try {
                const parsed = JSON.parse(raw) as unknown[];
                const users: { display: string; email: string }[] = [];
                for (const item of parsed) {
                    if (typeof item === "string" && item.trim()) {
                        users.push({ display: item.trim(), email: item.trim() });
                    } else if (item && typeof item === "object") {
                        const obj = item as Record<string, unknown>;
                        const display = String(
                            obj["DisplayName"] ?? obj["displayName"] ??
                            obj["displayname"] ?? obj["mail"] ?? obj["Mail"] ??
                            obj["UserPrincipalName"] ?? obj["userPrincipalName"] ?? ""
                        ).trim();
                        const email = String(
                            obj["Mail"] ?? obj["mail"] ??
                            obj["UserPrincipalName"] ?? obj["userPrincipalName"] ?? display
                        ).trim();
                        if (display) users.push({ display, email });
                    }
                }
                if (users.length > 0) return users.sort((a, b) => a.display.localeCompare(b.display));
            } catch { /* fall through */ }
        }
        // Fallback: unique values from existing records
        return Array.from(new Set(
            allRecords.map(r => r.userAccount).filter(u => u.trim().length > 0)
        )).sort((a, b) => a.localeCompare(b)).map(u => ({ display: u, email: u }));
    }

    private createPersonCell(name: string): HTMLDivElement {
        const cell = this.el("div", { display: "flex", alignItems: "center" }) as HTMLDivElement;
        cell.appendChild(this.el("span", { fontWeight: "500", color: "#344054" }, name || "—"));
        return cell;
    }

    private static readonly BADGE_PALETTE: { bg: string; fg: string }[] = [
        { bg: "#e8f4fd", fg: "#1565c0" },
        { bg: "#fde8ef", fg: "#c2185b" },
        { bg: "#e8f5e9", fg: "#2e7d32" },
        { bg: "#fff3e0", fg: "#e65100" },
        { bg: "#f3e5f5", fg: "#6a1b9a" },
        { bg: "#e0f7fa", fg: "#006064" },
        { bg: "#fce4ec", fg: "#880e4f" },
        { bg: "#e8eaf6", fg: "#283593" },
        { bg: "#f9fbe7", fg: "#558b2f" },
        { bg: "#fff8e1", fg: "#f57f17" },
        { bg: "#fbe9e7", fg: "#bf360c" },
        { bg: "#e0f2f1", fg: "#004d40" },
    ];

    private getValueBadgeColor(value: string): { bg: string; fg: string } {
        const key = value.trim().toLowerCase();
        if (!key) return { bg: "#f2f4f7", fg: "#344054" };
        const overrides: Record<string, { bg: string; fg: string }> = {
            "interne":   { bg: "#e8f5e9", fg: "#2e7d32" },
            "ext":       { bg: "#fff3e0", fg: "#e65100" },
            "vendeur":   { bg: "#fff8e1", fg: "#f57f17" },
            "cx région": { bg: "#f3e5f5", fg: "#6a1b9a" },
        };
        if (overrides[key]) return overrides[key];
        let h = 5381;
        for (let i = 0; i < key.length; i++) {
            h = ((h << 5) + h) + key.charCodeAt(i);
            h = h & h;
        }
        const palette = UserManagerPcf15.BADGE_PALETTE;
        return palette[Math.abs(h) % palette.length];
    }

    private makeBadge(text: string, bg: string, color: string): HTMLSpanElement {
        return this.el("span", {
            background: bg,
            color,
            borderRadius: "12px",
            padding: "2px 8px",
            fontSize: "12px",
            fontWeight: "500",
            whiteSpace: "nowrap",
            display: "inline-block",
        }, text) as HTMLSpanElement;
    }

    private createLieuChip(context: ComponentFramework.Context<IInputs>, record: UserRecordViewModel, lieu: string): HTMLDivElement {
        const chip = this.el("div", {
            display: "inline-flex",
            alignItems: "center",
            gap: "4px",
            background: "#f2f4f7",
            border: "1px solid #d0d5dd",
            borderRadius: "6px",
            padding: "2px 4px 2px 6px",
            fontSize: "12px",
            color: "#344054",
            fontWeight: "500",
            whiteSpace: "nowrap",
        }) as HTMLDivElement;

        const label = this.el("span", {}, lieu);

        // Remove button hidden per user request
        // const closeBtn = document.createElement("button");
        // closeBtn.type = "button";
        // this.applyStyles(closeBtn, {
        //     background: "none",
        //     border: "none",
        //     padding: "0 0 0 2px",
        //     cursor: "pointer",
        //     color: "#98a2b3",
        //     fontSize: "14px",
        //     lineHeight: "1",
        //     display: "inline-flex",
        //     alignItems: "center",
        //     fontFamily: "sans-serif",
        // });
        // closeBtn.textContent = "×";
        // closeBtn.title = `Remove ${lieu}`;
        // closeBtn.addEventListener("mouseover", () => closeBtn.style.color = "#c62828");
        // closeBtn.addEventListener("mouseout", () => closeBtn.style.color = "#98a2b3");
        // closeBtn.addEventListener("mousedown", (e) => e.preventDefault()); // prevent focus loss
        // closeBtn.addEventListener("click", (e) => {
        //     e.stopPropagation();
        //     chip.remove();
        //     const remaining = record.lieux.filter(l => l !== lieu);
        //     this.fireAction({
        //         action: "removeLieu",
        //         id: record.id,
        //         lieuValue: lieu,
        //         remainingLieux: remaining,
        //     });
        // });

        chip.appendChild(label);
        // chip.appendChild(closeBtn);
        return chip;
    }

    // ─── Modals ───────────────────────────────────────────────────────────────

    private openModal(context: ComponentFramework.Context<IInputs>, record: UserRecordViewModel | null, allRecords: UserRecordViewModel[]): void {
        this.closeModal();

        const overlay = this.createOverlay();

        const popup = this.el("div", {
            background: "#ffffff",
            borderRadius: "12px",
            padding: "24px",
            width: "480px",
            maxWidth: "90vw",
            maxHeight: "92vh",
            overflowY: "auto",
            boxShadow: "0 8px 40px rgba(0,0,0,0.22)",
        }) as HTMLDivElement;

        popup.appendChild(this.el("h2", {
            margin: "0 0 20px 0",
            fontSize: "16px",
            fontWeight: "700",
            color: "#101828",
        }, record ? "Edit user" : "Add user"));

        const formValues: Record<string, string> = {
            userAccount:      record?.userAccount ?? "",
            userAccountEmail: "",
            typeUser:         record?.typeUser    ?? "",
            titleUser:        record?.titleUser   ?? "",
            role:             record?.role        ?? "",
        };
        let currentLieux = [...(record?.lieux ?? [])];

        const form = this.el("div", { display: "flex", flexDirection: "column", gap: "14px" }) as HTMLDivElement;

        // ── UserAccount — people picker from usersJson (AAD via Office365Users) or fallback to existing records
        const uniqueUsers = this.resolveUserList(context, allRecords);

        const userAccountGroup = this.el("div", { display: "flex", flexDirection: "column", gap: "5px" });
        userAccountGroup.appendChild(this.el("label", { fontSize: "13px", fontWeight: "600", color: "#344054" }, "User"));

        const userAutocompleteWrap = this.el("div", { position: "relative" }) as HTMLDivElement;

        const personWrapper = this.el("div", {
            display: "flex",
            alignItems: "center",
            gap: "8px",
            padding: "6px 10px",
            border: "1px solid #d0d5dd",
            borderRadius: "6px",
            background: "#f9fafb",
            transition: "border-color 0.15s",
        }) as HTMLDivElement;

        const personInput = document.createElement("input");
        this.applyStyles(personInput, {
            border: "none",
            outline: "none",
            background: "transparent",
            fontSize: "13px",
            color: "#344054",
            fontFamily: "Inter, Segoe UI, sans-serif",
            flex: "1",
            minWidth: "0",
        });
        personInput.value = formValues.userAccount;
        personInput.placeholder = "Rechercher un utilisateur…";

        const userDropdown = this.el("div", {
            position: "absolute",
            top: "100%",
            left: "0",
            right: "0",
            background: "#ffffff",
            border: "1px solid #d0d5dd",
            borderRadius: "6px",
            boxShadow: "0 4px 12px rgba(0,0,0,0.12)",
            maxHeight: "180px",
            overflowY: "auto",
            zIndex: "10000",
            display: "none",
            marginTop: "2px",
        }) as HTMLDivElement;

        const updateUserDropdown = (filter: string) => {
            // Re-read usersJson from the latest context on every keystroke
            const liveUsers = this.resolveUserList(this.context, allRecords);
            userDropdown.innerHTML = "";
            const matches = liveUsers.filter(u =>
                filter.trim() === "" || u.display.toLowerCase().includes(filter.toLowerCase().trim())
            );
            if (matches.length === 0) { userDropdown.style.display = "none"; return; }
            matches.forEach(user => {
                const item = this.el("div", {
                    padding: "8px 12px",
                    cursor: "pointer",
                    fontSize: "13px",
                    color: "#344054",
                    fontFamily: "Inter, Segoe UI, sans-serif",
                    display: "flex",
                    alignItems: "center",
                    gap: "8px",
                });
                const initEl = this.el("div", {
                    width: "22px", height: "22px", borderRadius: "50%",
                    background: "#e0e7ff", color: "#4f46e5",
                    display: "flex", alignItems: "center", justifyContent: "center",
                    fontWeight: "700", fontSize: "10px", flexShrink: "0",
                }, user.display.charAt(0).toUpperCase());
                item.appendChild(initEl);
                item.appendChild(this.el("span", {}, user.display));
                item.addEventListener("mousedown", (e) => {
                    e.preventDefault();
                    formValues.userAccount = user.display;
                    formValues.userAccountEmail = user.email;
                    personInput.value = user.display;
                    userDropdown.style.display = "none";
                });
                item.addEventListener("mouseover", () => { item.style.background = "#f5f5ff"; });
                item.addEventListener("mouseout",  () => { item.style.background = ""; });
                userDropdown.appendChild(item);
            });
            userDropdown.style.display = "block";
        };

        personInput.addEventListener("input", () => {
            formValues.userAccount = personInput.value;
            updateUserDropdown(personInput.value);
            // Debounce searchTerm output to avoid rapid updateView calls during typing
            if (this.searchTermTimer) clearTimeout(this.searchTermTimer);
            this.searchTermTimer = setTimeout(() => {
                this.outputSearchTerm = personInput.value;
                this.notifyOutputChanged();
            }, 400);
        });
        personInput.addEventListener("focus", () => updateUserDropdown(personInput.value));
        personInput.addEventListener("blur",  () => { setTimeout(() => { userDropdown.style.display = "none"; }, 150); });

        // Auto-refresh dropdown when PA returns new usersJson results
        this.updateDropdownCallback = () => updateUserDropdown(personInput.value);

        personWrapper.addEventListener("click", () => personInput.focus());
        personWrapper.addEventListener("focusin",  () => { personWrapper.style.borderColor = "#4f46e5"; personWrapper.style.background = "#ffffff"; });
        personWrapper.addEventListener("focusout", () => { personWrapper.style.borderColor = "#d0d5dd"; personWrapper.style.background = "#f9fafb"; });

        personWrapper.appendChild(personInput);
        userAutocompleteWrap.appendChild(personWrapper);
        userAutocompleteWrap.appendChild(userDropdown);
        userAccountGroup.appendChild(userAutocompleteWrap);
        form.appendChild(userAccountGroup);

        // ── Dropdown fields — choices from input props (pipe-separated), fallback to hardcoded
        const parseChoices = (raw: string | null | undefined, fallback: string[]): string[] => {
            const s = raw?.trim();
            if (s) {
                const parsed = s.split("|").map(v => v.trim()).filter(v => v.length > 0);
                if (parsed.length > 0) return parsed;
            }
            return fallback;
        };

        const typeUserOpts  = parseChoices(context.parameters.typeUserChoices?.raw,  ["Interne", "EXT"]);
        const titleUserOpts = parseChoices(context.parameters.titleUserChoices?.raw, [
            "HDTI (Role = Admin)", "CX HI (Role = Admin)", "CX Région (Role = Admin)",
            "CX Filiale (Role = Admin)", "CX Store (Role = Admin)", "Manager (Role = Admin)",
            "Vendeur (Role = Order Owner)", "Hôte (Role = Order Owner)",
            "Catering Team (Role = Runner Processor)",
        ]);
        const roleOpts      = parseChoices(context.parameters.roleChoices?.raw, [
            "Admin : CX HI", "Admin : CX Région", "Admin : CX Filiale", "Admin : CX Store",
            "Admin : Managers", "Order Owner : Vendeurs", "Order Owner : Hôtes",
            "Order Processor : Catering team",
        ]);

        const selectDefs: { label: string; key: string; options: string[] }[] = [
            { label: "Type",  key: "typeUser",  options: typeUserOpts  },
            { label: "Title", key: "titleUser", options: titleUserOpts },
            { label: "Role",  key: "role",      options: roleOpts      },
        ];

        selectDefs.forEach(({ label, key, options }) => {
            const group = this.el("div", { display: "flex", flexDirection: "column", gap: "5px" });
            group.appendChild(this.el("label", { fontSize: "13px", fontWeight: "600", color: "#344054" }, label));
            const select = document.createElement("select");
            this.applyStyles(select, {
                padding: "8px 12px",
                borderRadius: "6px",
                border: "1px solid #d0d5dd",
                fontSize: "13px",
                color: "#344054",
                fontFamily: "Inter, Segoe UI, sans-serif",
                outline: "none",
                background: "#ffffff",
                cursor: "pointer",
                appearance: "auto",
            });

            // Blank option
            const blank = document.createElement("option");
            blank.value = "";
            blank.textContent = "— Select —";
            select.appendChild(blank);

            // Match current value: exact first, then case-insensitive, then partial
            const currentVal = formValues[key] ?? "";
            let matched = false;
            options.forEach(opt => {
                const o = document.createElement("option");
                o.value = opt;
                o.textContent = opt;
                const isMatch = !matched && (
                    opt === currentVal ||
                    opt.toLowerCase() === currentVal.toLowerCase() ||
                    (currentVal.length > 0 && opt.toLowerCase().includes(currentVal.toLowerCase()))
                );
                if (isMatch) { o.selected = true; matched = true; }
                select.appendChild(o);
            });

            select.addEventListener("change", () => { formValues[key] = select.value; });
            select.addEventListener("focus", () => { select.style.borderColor = "#4f46e5"; });
            select.addEventListener("blur",  () => { select.style.borderColor = "#d0d5dd"; });
            group.appendChild(select);
            form.appendChild(group);
        });

        // Lieux field with search suggestions
        const allLieux = Array.from(new Set(
            allRecords.flatMap(r => r.lieux).filter(l => l.trim().length > 0)
        )).sort((a, b) => a.localeCompare(b));

        const lieuxGroup = this.el("div", { display: "flex", flexDirection: "column", gap: "5px" });
        lieuxGroup.appendChild(this.el("label", { fontSize: "13px", fontWeight: "600", color: "#344054" }, "Store(s)"));

        const lieuxFieldWrap = this.el("div", { position: "relative" }) as HTMLDivElement;

        const chipsBox = this.el("div", {
            display: "flex",
            flexWrap: "wrap",
            gap: "4px",
            padding: "6px",
            border: "1px solid #d0d5dd",
            borderRadius: "6px",
            minHeight: "38px",
            cursor: "text",
        }) as HTMLDivElement;

        const lieuxDropdown = this.el("div", {
            position: "absolute",
            top: "100%",
            left: "0",
            right: "0",
            background: "#ffffff",
            border: "1px solid #d0d5dd",
            borderRadius: "6px",
            boxShadow: "0 4px 12px rgba(0,0,0,0.12)",
            maxHeight: "180px",
            overflowY: "auto",
            zIndex: "10000",
            display: "none",
            marginTop: "2px",
        }) as HTMLDivElement;

        const renderLieuxInForm = () => {
            chipsBox.innerHTML = "";
            currentLieux.forEach(lieu => {
                const chip = this.el("div", {
                    display: "inline-flex",
                    alignItems: "center",
                    gap: "4px",
                    background: "#f2f4f7",
                    border: "1px solid #d0d5dd",
                    borderRadius: "6px",
                    padding: "2px 6px",
                    fontSize: "12px",
                    color: "#344054",
                });
                chip.textContent = lieu;
                // Remove button hidden per user request
                // const x = document.createElement("span");
                // x.textContent = " ×";
                // this.applyStyles(x, { cursor: "pointer", color: "#667085", fontWeight: "700" });
                // x.addEventListener("click", () => {
                //     currentLieux = currentLieux.filter(l => l !== lieu);
                //     renderLieuxInForm();
                // });
                // chip.appendChild(x);
                chipsBox.appendChild(chip);
            });

            const addInput = document.createElement("input");
            this.applyStyles(addInput, {
                border: "none",
                outline: "none",
                fontSize: "13px",
                minWidth: "100px",
                flex: "1",
                fontFamily: "Inter, Segoe UI, sans-serif",
                padding: "0 4px",
            });
            addInput.placeholder = "Rechercher un store…";

            const updateLieuxDropdown = (filter: string) => {
                lieuxDropdown.innerHTML = "";
                const already = new Set(currentLieux.map(l => l.toLowerCase()));
                const matches = allLieux.filter(l =>
                    !already.has(l.toLowerCase()) &&
                    (filter.trim() === "" || l.toLowerCase().includes(filter.toLowerCase().trim()))
                );
                if (matches.length === 0) { lieuxDropdown.style.display = "none"; return; }
                matches.forEach(lieu => {
                    const item = this.el("div", {
                        padding: "8px 12px",
                        cursor: "pointer",
                        fontSize: "13px",
                        color: "#344054",
                        fontFamily: "Inter, Segoe UI, sans-serif",
                    }, lieu);
                    item.addEventListener("mousedown", (e) => {
                        e.preventDefault();
                        currentLieux.push(lieu);
                        lieuxDropdown.style.display = "none";
                        renderLieuxInForm();
                    });
                    item.addEventListener("mouseover", () => { item.style.background = "#f5f5ff"; });
                    item.addEventListener("mouseout",  () => { item.style.background = ""; });
                    lieuxDropdown.appendChild(item);
                });
                lieuxDropdown.style.display = "block";
            };

            addInput.addEventListener("input",  () => updateLieuxDropdown(addInput.value));
            addInput.addEventListener("focus",  () => updateLieuxDropdown(addInput.value));
            addInput.addEventListener("blur",   () => { setTimeout(() => { lieuxDropdown.style.display = "none"; }, 150); });
            addInput.addEventListener("keydown", (e) => {
                if ((e.key === "Enter" || e.key === ",") && addInput.value.trim()) {
                    e.preventDefault();
                    currentLieux.push(addInput.value.trim());
                    lieuxDropdown.style.display = "none";
                    renderLieuxInForm();
                }
            });
            chipsBox.appendChild(addInput);
        };

        renderLieuxInForm();
        chipsBox.addEventListener("click", () => {
            const inp = chipsBox.querySelector("input");
            if (inp) inp.focus();
        });

        lieuxFieldWrap.appendChild(chipsBox);
        lieuxFieldWrap.appendChild(lieuxDropdown);
        lieuxGroup.appendChild(lieuxFieldWrap);
        lieuxGroup.appendChild(this.el("span", { fontSize: "11px", color: "#98a2b3" }, "Tapez pour rechercher ou appuyez Entrée pour ajouter"));
        form.appendChild(lieuxGroup);
        popup.appendChild(form);

        // Buttons
        const btnRow = this.el("div", { display: "flex", justifyContent: "flex-end", gap: "8px", marginTop: "20px" }) as HTMLDivElement;

        btnRow.appendChild(this.makeBtn("Cancel", "#f2f4f7", "#344054", () => this.closeModal()));

        const saveBtn = this.makeBtn(record ? "Save" : "Add", "#4f46e5", "#ffffff", () => {
            const payload: Record<string, unknown> = {
                action: record ? "edit" : "add",
                userAccount:      formValues.userAccount,
                userAccountEmail: formValues.userAccountEmail,
                typeUser:         formValues.typeUser,
                titleUser:        formValues.titleUser,
                role:             formValues.role,
                lieux:            currentLieux,
            };
            if (record) payload.id = record.id;
            this.fireAction(payload);
            this.closeModal();
        });
        btnRow.appendChild(saveBtn);

        popup.appendChild(btnRow);
        overlay.appendChild(popup);
        popup.addEventListener("click", (e) => e.stopPropagation());
        document.body.appendChild(overlay);
        this.modalOverlay = overlay;
    }

    private openDeleteConfirm(context: ComponentFramework.Context<IInputs>): void {
        this.closeModal();

        const count = this.selectedIds.size;
        const overlay = this.createOverlay();

        const popup = this.el("div", {
            background: "#ffffff",
            borderRadius: "12px",
            padding: "32px 28px",
            width: "360px",
            maxWidth: "90vw",
            boxShadow: "0 8px 40px rgba(0,0,0,0.22)",
        }) as HTMLDivElement;

        popup.appendChild(this.el("h2", {
            margin: "0 0 8px 0",
            fontSize: "16px",
            fontWeight: "700",
            color: "#c62828",
        }, "Confirm deletion"));

        popup.appendChild(this.el("p", {
            fontSize: "14px",
            color: "#344054",
            margin: "0 0 20px 0",
        }, `Do you want to delete ${count > 1 ? `these ${count} users` : "this user"}? This action is irreversible.`));

        const btnRow = this.el("div", { display: "flex", justifyContent: "flex-end", gap: "8px" }) as HTMLDivElement;

        btnRow.appendChild(this.makeBtn("Cancel", "#f2f4f7", "#344054", () => this.closeModal()));

        btnRow.appendChild(this.makeBtn("Delete", "#c62828", "#ffffff", () => {
            this.fireAction({ action: "delete", ids: Array.from(this.selectedIds) });
            this.selectedIds.clear();
            this.closeModal();
            this.render(context, context.parameters.users);
        }));

        popup.appendChild(btnRow);
        popup.addEventListener("click", (e) => e.stopPropagation());
        overlay.appendChild(popup);
        document.body.appendChild(overlay);
        this.modalOverlay = overlay;
    }

    private createOverlay(): HTMLDivElement {
        const overlay = this.el("div", {
            position: "fixed",
            inset: "0",
            background: "rgba(16,24,40,0.5)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: "9999",
        }) as HTMLDivElement;
        overlay.addEventListener("click", (e) => {
            if (e.target === overlay) this.closeModal();
        });
        return overlay;
    }

    private closeModal(): void {
        this.modalOverlay?.remove();
        this.modalOverlay = null;
        this.updateDropdownCallback = null;
    }

    // ─── Output ───────────────────────────────────────────────────────────────

    private fireAction(payload: Record<string, unknown>): void {
        this.outputSequence += 1;
        this.outputFormData = JSON.stringify(payload);
        this.outputAction = String(payload.action ?? "");
        this.outputSelectedIds = Array.isArray(payload.ids) ? JSON.stringify(payload.ids) : (payload.id ? String(payload.id) : "");
        this.outputEventToken = `um-${Date.now()}-${this.outputSequence}`;
        if (this.outputAction === "upsert") {
            this.outputSaveSignal = this.outputEventToken;
        }
        // Double notify: Power Apps may miss a single notify during modal close/re-render
        this.notifyOutputChanged();
        setTimeout(() => { this.notifyOutputChanged(); }, 30);
    }

    // ─── Helpers ──────────────────────────────────────────────────────────────

    private createPagination(context: ComponentFramework.Context<IInputs>, totalItems: number, totalPages: number): HTMLDivElement {
        const bar = this.el("div", {
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            padding: "10px 16px",
            borderTop: "1px solid #e4e7ec",
            background: "#f9fafb",
            flexWrap: "wrap",
            gap: "8px",
        }) as HTMLDivElement;

        const pageStart = (this.currentPage - 1) * this.PAGE_SIZE + 1;
        const pageEnd = Math.min(this.currentPage * this.PAGE_SIZE, totalItems);
        bar.appendChild(this.el("span", { fontSize: "13px", color: "#667085" },
            `${pageStart}–${pageEnd} sur ${totalItems}`));

        const nav = this.el("div", { display: "flex", alignItems: "center", gap: "6px" }) as HTMLDivElement;

        const prevBtn = this.makeBtn("‹ Prev", "#f2f4f7", "#344054", () => {
            this.currentPage -= 1;
            this.render(context, context.parameters.users);
        });
        this.setDisabled(prevBtn, this.currentPage === 1);
        nav.appendChild(prevBtn);

        // Page number buttons (show max 5 around current)
        const delta = 2;
        const start = Math.max(1, this.currentPage - delta);
        const end   = Math.min(totalPages, this.currentPage + delta);
        if (start > 1) {
            nav.appendChild(this.makePageBtn(context, 1));
            if (start > 2) nav.appendChild(this.el("span", { color: "#98a2b3", padding: "0 2px" }, "…"));
        }
        for (let p = start; p <= end; p++) { nav.appendChild(this.makePageBtn(context, p)); }
        if (end < totalPages) {
            if (end < totalPages - 1) nav.appendChild(this.el("span", { color: "#98a2b3", padding: "0 2px" }, "…"));
            nav.appendChild(this.makePageBtn(context, totalPages));
        }

        const nextBtn = this.makeBtn("Next ›", "#f2f4f7", "#344054", () => {
            this.currentPage += 1;
            this.render(context, context.parameters.users);
        });
        this.setDisabled(nextBtn, this.currentPage === totalPages);
        nav.appendChild(nextBtn);

        bar.appendChild(nav);
        return bar;
    }

    private makePageBtn(context: ComponentFramework.Context<IInputs>, page: number): HTMLButtonElement {
        const isActive = page === this.currentPage;
        const btn = this.makeBtn(String(page), isActive ? "#4f46e5" : "#f2f4f7", isActive ? "#ffffff" : "#344054", () => {
            this.currentPage = page;
            this.render(context, context.parameters.users);
        });
        this.applyStyles(btn, { minWidth: "32px", padding: "5px 8px", fontWeight: isActive ? "700" : "400" });
        return btn;
    }

    private createMessage(text: string): HTMLDivElement {
        return this.el("div", {
            padding: "32px",
            textAlign: "center",
            color: "#667085",
            fontSize: "14px",
        }, text) as HTMLDivElement;
    }

    private el(tag: string, styles?: Record<string, string>, text?: string): HTMLElement {
        const element = document.createElement(tag);
        if (styles) this.applyStyles(element, styles);
        if (text !== undefined) element.textContent = text;
        return element;
    }

    private applyStyles(element: HTMLElement, styles: Record<string, string>): void {
        Object.assign(element.style, styles);
    }
}
