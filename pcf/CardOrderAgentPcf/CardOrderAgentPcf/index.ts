import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class CardOrderAgentPcfLast23 implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private static readonly DEFAULT_COLUMN_CANDIDATES = {
        createdBy: ["Author", "createdby", "Created By", "Cree par", "Creer par"],
        createdOn: ["Created", "created", "createdon", "Cree", "Creer"],
        itemId: ["ID", "Id", "id"],
        modifiedOn: ["Modified", "modified", "modifiedon", "Modifie", "Modifier"],
        note: ["Notes", "Note", "notes"],
        orderNumber: ["Title", "titre", "name", "ordernumber"],
        products: ["Products", "products", "Items", "items", "Produits", "produits", "OrderItems", "orderItems", "JSONOrderSections", "jsonOrderSections", "JsonOrderSections"],
        quantity: ["Quantit_x00e9_e", "quantity"],
        status: ["StatutCommande", "status", "statuscode"]
    };

    private static readonly CARD_STATUSES = {
        cleaned: {
            background: "#eef2f6",
            foreground: "#475467"
        },
        cancelled: {
            background: "#fef0f0",
            foreground: "#f04438"
        },
        inPrep: {
            background: "#c4e8e3",
            foreground: "#00b69b"
        },
        served: {
            background: "#eaecf0",
            foreground: "#667085"
        },
        toClean: {
            background: "#fff1c2",
            foreground: "#eaaa08"
        },
        toPrepare: {
            background: "#fde7de",
            foreground: "#f3875f"
        },
        unknown: {
            background: "#f2f4f7",
            foreground: "#475467"
        }
    };

    private notifyOutputChanged!: () => void;
    private root!: HTMLDivElement;
    private selectedOrderId = "";
    private requestedAction = "";
    private requestedStatus = "";
    private requestedNotes = "";
    private rawRequestedNotes = "";
    private requestedEventToken = "";
    private outputSequence = 0;
    private lastClickedOrderId = "";
    private requestedColumns = new Set<string>();
    private searchTerm = "";
    private selectedStatus = "all";
    private takeModal?: HTMLDivElement;
    private cancelOrderModal?: HTMLDivElement;

    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.root = document.createElement("div");
        this.applyStyles(this.root, {
            background: "#f5f7fb",
            boxSizing: "border-box",
            padding: "12px",
            width: "100%"
        });

        container.appendChild(this.root);
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const orders = context.parameters.orders;
        if (this.ensureRequestedColumns(context, orders)) {
            return;
        }

        this.render(context, orders);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        if (!this.selectedOrderId && this.lastClickedOrderId) {
            this.selectedOrderId = this.lastClickedOrderId;
        }

        if ((this.requestedAction || this.requestedStatus) && !this.requestedEventToken) {
            this.outputSequence += 1;
            this.requestedEventToken = `fallback-${Date.now()}-${this.outputSequence}`;
        }

        this.requestedNotes = this.buildOutputPayload();

        const outputs: IOutputs = {
            requestedAction: this.requestedAction || undefined,
            requestedStatus: this.requestedStatus || undefined,
            requestedNotes: this.requestedNotes || undefined,
            selectedOrderId: this.selectedOrderId || undefined,
            requestedEventToken: this.requestedEventToken || undefined,
            selectedOrderIdV2: this.selectedOrderId || undefined,
            requestedEventTokenV2: this.requestedEventToken || undefined
        };
        return outputs;
    }

    private buildOutputPayload(): string {
        return JSON.stringify({
            action: this.requestedAction || "",
            id: this.selectedOrderId || this.lastClickedOrderId || "",
            note: this.rawRequestedNotes || "",
            status: this.requestedStatus || "",
            token: this.requestedEventToken || ""
        });
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        this.closeTakeModal();
        this.closeCancelOrderModal();
        if (this.cleanModal && this.cleanModal.parentElement) {
            this.cleanModal.parentElement.removeChild(this.cleanModal);
        }
        this.requestedColumns.clear();
    }

    private applyStyles(element: HTMLElement, styles: Partial<CSSStyleDeclaration>): void {
        Object.assign(element.style, styles);
    }

    private emitOutputEvent(): void {
        this.outputSequence += 1;
        this.requestedEventToken = `${Date.now()}-${this.outputSequence}-${Math.random().toString(36).slice(2, 8)}`;
        this.requestedNotes = this.buildOutputPayload();
        this.notifyOutputChanged();
        // Power Apps can occasionally miss a single notify during modal close/re-render.
        setTimeout(() => {
            this.notifyOutputChanged();
        }, 30);
    }

    private normalizeOutgoingStatus(status: string): string {
        return this.normalize(status) === "to clean" ? "To clear" : status;
    }

    private createElement<K extends keyof HTMLElementTagNameMap>(
        tagName: K,
        styles?: Partial<CSSStyleDeclaration>,
        textContent?: string
    ): HTMLElementTagNameMap[K] {
        const element = document.createElement(tagName);
        if (styles) {
            this.applyStyles(element, styles);
        }
        if (textContent !== undefined) {
            element.textContent = textContent;
        }
        return element;
    }

    private ensureRequestedColumns(context: ComponentFramework.Context<IInputs>, orders: ComponentFramework.PropertyTypes.DataSet): boolean {
        if (!orders.addColumn) {
            return false;
        }

        const candidateColumns = [
            context.parameters.orderNumberColumn.raw,
            context.parameters.createdByColumn.raw,
            context.parameters.statusColumn.raw,
            context.parameters.quantityColumn.raw,
            context.parameters.createdOnColumn.raw,
            ...CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.orderNumber,
            ...CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.createdBy,
            ...CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.status,
            ...CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.quantity,
            ...CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.modifiedOn,
            ...CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.createdOn,
            ...CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.itemId,
            ...CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.note,
            ...CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.products
        ]
            .map((value) => value?.trim() ?? "")
            .filter((value) => value.length > 0);

        let shouldRefresh = false;
        for (const column of candidateColumns) {
            const key = column.toLowerCase();
            const exists = orders.columns.some((item) => {
                const itemNameLower = item.name.trim().toLowerCase();
                const itemDisplayLower = item.displayName.trim().toLowerCase();
                const keyNormalized = key.replace(/_/g, " ");
                return itemNameLower === key || itemDisplayLower === key || itemNameLower.replace(/_/g, " ") === keyNormalized;
            });
            if (!exists && !this.requestedColumns.has(key)) {
                this.requestedColumns.add(key);
                orders.addColumn(column);
                shouldRefresh = true;
            }
        }

        if (shouldRefresh) {
            orders.refresh();
        }

        return shouldRefresh;
    }

    private render(context: ComponentFramework.Context<IInputs>, orders: ComponentFramework.PropertyTypes.DataSet): void {
        this.root.replaceChildren();

        const searchContainer = this.createSearchBar(context, orders);
        this.root.appendChild(searchContainer);

        const controlsSpacer = this.createElement("div", {
            height: "8px",
            flexShrink: "0"
        });
        this.root.appendChild(controlsSpacer);

        const scroller = this.createElement("div", {
            display: "flex",
            flexDirection: "column",
            gap: "14px",
            overflowX: "hidden",
            paddingRight: "4px"
        });

        if (orders.error) {
            scroller.appendChild(this.createStateBlock(
                this.translate(context, "loadErrorTitle"),
                orders.errorMessage || this.translate(context, "loadErrorMessage")
            ));
            this.root.appendChild(scroller);
            return;
        }

        if (orders.loading) {
            scroller.appendChild(this.createStateBlock(this.translate(context, "loadingTitle"), this.translate(context, "loadingMessage")));
            this.root.appendChild(scroller);
            return;
        }

        const records = this.getOrders(context, orders);
        const filteredRecords = this.filterRecords(records);

        if (filteredRecords.length === 0) {
            scroller.appendChild(
                this.createStateBlock(
                    this.translate(context, "emptyTitle"),
                    this.translate(context, "emptyMessage")
                )
            );
            this.root.appendChild(scroller);
            return;
        }

        const tabBar = this.createTabBar(context, filteredRecords, scroller);
        this.root.appendChild(tabBar);

        const spacer = this.createElement("div", {
            flexShrink: "0",
            height: "8px"
        });
        this.root.appendChild(spacer);

        filteredRecords.forEach((record) => {
            scroller.appendChild(this.createOrderCard(context, record));
        });

        if (this.selectedStatus !== "all") {
            scroller.querySelectorAll<HTMLElement>("[data-canonical-status]").forEach((cardEl) => {
                cardEl.style.display = cardEl.dataset.canonicalStatus === this.selectedStatus ? "" : "none";
            });
        }

        this.root.appendChild(scroller);
    }

    private createSearchBar(
        context: ComponentFramework.Context<IInputs>,
        orders: ComponentFramework.PropertyTypes.DataSet
    ): HTMLDivElement {
        const searchContainer = this.createElement("div", {
            alignItems: "center",
            alignSelf: "stretch",
            background: "#ffffff",
            border: "0.613636px solid #000000",
            borderRadius: "4.29545px",
            boxSizing: "border-box",
            display: "flex",
            flexDirection: "row",
            flexGrow: "0",
            flexShrink: "0",
            gap: "0px",
            height: "38px",
            order: "1",
            padding: "0px 8px",
            width: "100%"
        });

        // Add magnifying glass icon
        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("width", "16");
        iconSvg.setAttribute("height", "16");
        iconSvg.setAttribute("viewBox", "0 0 16 16");
        iconSvg.setAttribute("fill", "none");
        iconSvg.setAttribute("stroke", "#666666");
        iconSvg.setAttribute("stroke-width", "1.5");
        iconSvg.setAttribute("stroke-linecap", "round");
        iconSvg.setAttribute("stroke-linejoin", "round");
        iconSvg.style.flexShrink = "0";
        iconSvg.style.flexGrow = "0";

        const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        circle.setAttribute("cx", "6");
        circle.setAttribute("cy", "6");
        circle.setAttribute("r", "5");
        iconSvg.appendChild(circle);

        const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
        line.setAttribute("x1", "10");
        line.setAttribute("y1", "10");
        line.setAttribute("x2", "14");
        line.setAttribute("y2", "14");
        iconSvg.appendChild(line);

        searchContainer.appendChild(iconSvg);

        const input = this.createElement("input", {
            background: "#ffffff",
            border: "none",
            borderRadius: "0px",
            boxSizing: "border-box",
            color: "#101828",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "13px",
            height: "100%",
            outline: "none",
            padding: "0px 8px",
            width: "100%"
        }) as HTMLInputElement;

        input.type = "search";
        input.placeholder = "Search";
        input.value = this.searchTerm;
        input.addEventListener("input", () => {
            const cursorPos = input.selectionStart ?? input.value.length;
            this.searchTerm = input.value;
            this.render(context, orders);
            const newInput = this.root.querySelector<HTMLInputElement>("input[type='search']");
            if (newInput) {
                newInput.focus();
                newInput.setSelectionRange(cursorPos, cursorPos);
            }
        });

        searchContainer.appendChild(input);
        return searchContainer;
    }

    private filterRecords(records: OrderRecordViewModel[]): OrderRecordViewModel[] {
        const keyword = this.normalize(this.searchTerm.trim());
        if (!keyword) {
            return records;
        }

        return records.filter((record) => {
            const fields = [
                record.orderNumber,
                record.createdBy,
                record.status,
                record.note,
                record.quantity,
                ...record.products.map((item) => `${item.quantity} ${item.label}`)
            ];

            return fields.some((value) => this.normalize(value).includes(keyword));
        });
    }

    private createStateBlock(title: string, message: string): HTMLDivElement {
        const block = this.createElement("div", {
            alignItems: "center",
            background: "#ffffff",
            border: "1px solid #eaecf0",
            borderRadius: "11px",
            boxShadow: "0 1px 2px rgba(16, 24, 40, 0.05)",
            display: "flex",
            flexDirection: "column",
            gap: "8px",
            justifyContent: "center",
            minHeight: "160px",
            padding: "24px",
            textAlign: "center"
        });

        block.appendChild(this.createElement("div", {
            color: "#101828",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "16px",
            fontWeight: "700",
            lineHeight: "24px"
        }, title));

        block.appendChild(this.createElement("div", {
            color: "#667085",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "13px",
            lineHeight: "20px",
            maxWidth: "260px"
        }, message));

        return block;
    }

    private createOrderCard(context: ComponentFramework.Context<IInputs>, record: OrderRecordViewModel): HTMLDivElement {
        const card = this.createElement("div", {
            background: "#ffffff",
            border: "1px solid #eaecf0",
            borderRadius: "11px",
            boxShadow: "0 1px 2px rgba(16, 24, 40, 0.05)",
            display: "flex",
            flexDirection: "column",
            overflow: "hidden",
            width: "100%"
        });
        card.dataset.canonicalStatus = this.toCanonicalStatus(record.status);

        const header = this.createElement("div", {
            alignItems: "flex-start",
            display: "flex",
            gap: "12px",
            justifyContent: "space-between",
            padding: "15px 16px 12px 16px"
        });

        const left = this.createElement("div", {
            alignItems: "flex-start",
            display: "flex",
            flex: "1 1 auto",
            gap: "10px",
            minWidth: "0"
        });

        const avatar = this.createAvatar(record, context.parameters.defaultUserPhoto.raw ?? undefined);

        const details = this.createElement("div", {
            display: "flex",
            flexDirection: "column",
            gap: "4px",
            minWidth: "0"
        });

        details.appendChild(this.createElement("div", {
            color: "#667085",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "10px",
            fontStyle: "italic",
            fontWeight: "500",
            lineHeight: "15px"
        }, record.orderNumber));

        details.appendChild(this.createElement("div", {
            color: "#101828",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "600",
            lineHeight: "20px",
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap"
        }, record.createdBy));

        const meta = this.createElement("div", {
            color: "#667085",
            display: "flex",
            flexWrap: "wrap",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "10px",
            fontWeight: "700",
            gap: "4px",
            lineHeight: "16px"
        });
        const hideItemCount = this.shouldHideItemCount(record.status);
        if (!hideItemCount) {
            meta.appendChild(this.createElement("span", {
                fontWeight: "400"
            }, this.formatQuantity(record.quantity, this.getLanguage(context))));
            meta.appendChild(this.createElement("span", undefined, "•"));
        }
        meta.appendChild(this.createElement("span", {
            fontWeight: "400"
        }, record.createdTime));
        details.appendChild(meta);

        left.appendChild(avatar);
        left.appendChild(details);

        const badgeTheme = this.getStatusTheme(record.status);
        const badge = this.createElement("div", {
            alignItems: "center",
            background: badgeTheme.background,
            borderRadius: "2px",
            display: "flex",
            flex: "0 0 auto",
            justifyContent: "center",
            minWidth: "93px",
            padding: "5px 10px"
        });
        badge.appendChild(this.createElement("span", {
            color: badgeTheme.foreground,
            fontFamily: "Nunito Sans, Inter, Segoe UI, sans-serif",
            fontSize: "12px",
            fontWeight: "700",
            lineHeight: "16px",
            textAlign: "center"
        }, this.getLocalizedStatus(record.status, this.getLanguage(context))));

        header.appendChild(left);
        header.appendChild(badge);

        const middle = this.createElement("div", {
            background: "#ffffff",
            display: "flex",
            flexDirection: "column",
            gap: "10px",
            minHeight: "128px",
            padding: "0 16px 10px 16px",
            width: "100%"
        });

        const productsBlock = this.createProductsBlock(record.products);
        if (productsBlock) {
            middle.appendChild(productsBlock);
        }

        const noteBlock = this.createNoteBlock(context, record.note);
        if (noteBlock) {
            middle.appendChild(noteBlock);
        }

        const footer = this.createElement("div", {
            alignItems: "center",
            display: "flex",
            justifyContent: "center",
            minHeight: "60px",
            padding: "10px 16px 16px 16px"
        });

        if (record.actionLabel && record.actionName && record.nextStatus) {
            const actionName = record.actionName;
            const nextStatus = record.nextStatus;
            const isViewAction = actionName === "serve";
            const button = this.createElement("button", {
                alignItems: isViewAction ? "flex-start" : "center",
                background: isViewAction ? "transparent" : "#121926",
                border: isViewAction ? "none" : "1px solid #121926",
                borderRadius: "1px",
                boxShadow: isViewAction ? "none" : "0 1px 2px rgba(16, 24, 40, 0.05)",
                color: isViewAction ? "#344054" : "#ffffff",
                cursor: "pointer",
                display: "flex",
                flexDirection: "row",
                fontFamily: "Inter, Segoe UI, sans-serif",
                fontSize: "14px",
                fontWeight: "500",
                gap: isViewAction ? "0px" : "8px",
                justifyContent: "center",
                height: "36px",
                width: isViewAction ? "62px" : actionName === "clean" ? "190px" : "156px",
                padding: isViewAction ? "0px" : "8px 14px",
                flex: "none",
                order: "0",
                alignSelf: isViewAction ? "stretch" : "auto",
                flexGrow: "0"
            }) as HTMLButtonElement;

            button.type = "button";
            
            // Create button content with icon and text
            const buttonContent = this.createElement("div", {
                display: "flex",
                alignItems: "center",
                gap: isViewAction ? "0px" : "6px"
            });
            
            const buttonLabel = isViewAction ? this.getLocalizedActionLabel("View", this.getLanguage(context)) : this.getLocalizedActionLabel(record.actionLabel, this.getLanguage(context));
            buttonContent.appendChild(this.createElement("span", undefined, buttonLabel));
            if (!isViewAction) {
                const iconSvg = this.createActionIcon(actionName);
                buttonContent.appendChild(iconSvg);
            }
            
            button.appendChild(buttonContent);
            button.addEventListener("click", () => {
                this.lastClickedOrderId = record.id;
                if (actionName === "take" || actionName === "serve") {
                    this.openTakeModal(context, record, actionName, nextStatus);
                    return;
                }

                if (actionName === "clean") {
                    this.openCleanModal(context, record, nextStatus);
                    return;
                }

                this.selectedOrderId = record.id;
                this.requestedAction = actionName;
                this.requestedStatus = this.normalizeOutgoingStatus(nextStatus);
                this.emitOutputEvent();
            });
            footer.appendChild(button);
        }

        card.appendChild(header);
        card.appendChild(middle);
        if (record.actionLabel && record.actionName && record.nextStatus) {
            card.appendChild(footer);
        }

        return card;
    }

    private openTakeModal(
        context: ComponentFramework.Context<IInputs>,
        record: OrderRecordViewModel,
        actionName: string,
        nextStatus: string
    ): void {
        this.closeTakeModal();

        const overlay = this.createElement("div", {
            alignItems: "center",
            background: "rgba(16, 24, 40, 0.45)",
            bottom: "0",
            display: "flex",
            justifyContent: "center",
            left: "0",
            padding: "16px",
            position: "fixed",
            right: "0",
            top: "0",
            zIndex: "9999"
        });

        const modalCard = this.createElement("div", {
            background: "#ffffff",
            border: "1px solid #eaecf0",
            borderRadius: "11px",
            boxShadow: "0 10px 24px rgba(16, 24, 40, 0.2)",
            display: "flex",
            flexDirection: "column",
            gap: "10px",
            maxHeight: "90vh",
            maxWidth: "420px",
            overflowY: "auto",
            position: "relative",
            width: "100%"
        });

        const closeBtn = this.createElement("button", {
            alignItems: "center",
            background: "transparent",
            border: "none",
            color: "#344054",
            cursor: "pointer",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "24px",
            fontWeight: "400",
            height: "28px",
            justifyContent: "center",
            lineHeight: "1",
            padding: "0",
            width: "28px"
        }, "×") as HTMLButtonElement;
        closeBtn.type = "button";
        closeBtn.addEventListener("click", () => {
            this.closeTakeModal();
        });

        const header = this.createElement("div", {
            alignItems: "flex-start",
            display: "flex",
            gap: "12px",
            justifyContent: "space-between",
            padding: "15px 16px 12px 16px"
        });

        const left = this.createElement("div", {
            alignItems: "flex-start",
            display: "flex",
            flex: "1 1 auto",
            gap: "10px",
            minWidth: "0"
        });

        const avatar = this.createAvatar(record, context.parameters.defaultUserPhoto.raw ?? undefined);
        const details = this.createElement("div", {
            display: "flex",
            flexDirection: "column",
            gap: "4px",
            minWidth: "0"
        });

        details.appendChild(this.createElement("div", {
            color: "#667085",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "10px",
            fontStyle: "italic",
            fontWeight: "500",
            lineHeight: "15px"
        }, record.orderNumber));

        details.appendChild(this.createElement("div", {
            color: "#101828",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "600",
            lineHeight: "20px",
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap"
        }, record.createdBy));

        const meta = this.createElement("div", {
            color: "#667085",
            display: "flex",
            flexWrap: "wrap",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "10px",
            fontWeight: "700",
            gap: "4px",
            lineHeight: "16px"
        });
        const hideItemCount = this.shouldHideItemCount(record.status);
        if (!hideItemCount) {
            meta.appendChild(this.createElement("span", {
                fontWeight: "400"
            }, this.formatQuantity(record.quantity, this.getLanguage(context))));
            meta.appendChild(this.createElement("span", undefined, "•"));
        }
        meta.appendChild(this.createElement("span", {
            fontWeight: "400"
        }, record.createdTime));
        details.appendChild(meta);

        left.appendChild(avatar);
        left.appendChild(details);

        const badgeTheme = this.getStatusTheme(record.status);
        const badge = this.createElement("div", {
            alignItems: "center",
            background: badgeTheme.background,
            borderRadius: "2px",
            display: "flex",
            flex: "0 0 auto",
            justifyContent: "center",
            minWidth: "93px",
            padding: "5px 10px"
        });
        badge.appendChild(this.createElement("span", {
            color: badgeTheme.foreground,
            fontFamily: "Nunito Sans, Inter, Segoe UI, sans-serif",
            fontSize: "12px",
            fontWeight: "700",
            lineHeight: "16px",
            textAlign: "center"
        }, this.getLocalizedStatus(record.status, this.getLanguage(context))));

        const headerRight = this.createElement("div", {
            alignItems: "center",
            display: "flex",
            flex: "0 0 auto",
            gap: "6px"
        });
        headerRight.appendChild(badge);
        headerRight.appendChild(closeBtn);

        header.appendChild(left);
        header.appendChild(headerRight);

        const middle = this.createElement("div", {
            borderTop: "1px solid #eaecf0",
            display: "flex",
            flexDirection: "column",
            gap: "12px",
            padding: "14px 16px",
            width: "100%"
        });

        record.takeProducts.forEach((item) => {
            const row = this.createElement("div", {
                alignItems: "center",
                display: "flex",
                gap: "12px"
            });

            row.appendChild(this.createElement("span", {
                color: "#667085",
                fontFamily: "Inter, Segoe UI, sans-serif",
                fontSize: "14px",
                fontWeight: "500",
                lineHeight: "20px"
            }, item.quantity));

            row.appendChild(this.createElement("span", {
                color: "#101828",
                fontFamily: "Inter, Segoe UI, sans-serif",
                fontSize: "14px",
                fontWeight: "700",
                lineHeight: "20px",
                flex: "1 1 auto",
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap"
            }, item.label));

            middle.appendChild(row);
        });

        const noteBlock = this.createNoteBlock(context, record.note);
        if (noteBlock) {
            middle.appendChild(noteBlock);
        }

        const footer = this.createElement("div", {
            alignItems: "center",
            borderTop: "1px solid #eaecf0",
            display: "flex",
            gap: "14px",
            justifyContent: "center",
            padding: "10px 16px 16px 16px"
        });

        const cancelBtn = this.createElement("button", {
            alignItems: "center",
            background: "transparent",
            border: "none",
            color: "#d92d20",
            cursor: "pointer",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "500",
            gap: "8px",
            minHeight: "36px",
            padding: "8px 10px"
        }) as HTMLButtonElement;
        cancelBtn.type = "button";
        const cancelContent = this.createElement("div", {
            alignItems: "center",
            display: "flex",
            gap: "6px"
        });
        cancelContent.appendChild(this.createElement("span", undefined, "Cancel"));
        cancelContent.appendChild(this.createCancelIcon());
        cancelBtn.appendChild(cancelContent);
        cancelBtn.addEventListener("click", () => {
            this.openCancelOrderModal(context, record, actionName, nextStatus);
        });

        const takeBtn = this.createElement("button", {
            alignItems: "center",
            background: "#121926",
            border: "1px solid #121926",
            borderRadius: "1px",
            boxShadow: "0 1px 2px rgba(16, 24, 40, 0.05)",
            color: "#ffffff",
            cursor: "pointer",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "500",
            gap: "8px",
            justifyContent: "center",
            minHeight: "36px",
            minWidth: "92px",
            padding: "8px 10px"
        }) as HTMLButtonElement;
        takeBtn.type = "button";

        const takeContent = this.createElement("div", {
            alignItems: "center",
            display: "flex",
            gap: "6px"
        });
        const confirmLabel = actionName === "serve"
            ? this.getLocalizedActionLabel("Make as served", this.getLanguage(context))
            : this.getLocalizedActionLabel("Take", this.getLanguage(context));
        takeContent.appendChild(this.createElement("span", undefined, confirmLabel));
        takeContent.appendChild(this.createActionIcon("serve"));
        takeBtn.appendChild(takeContent);

        takeBtn.addEventListener("click", () => {
            this.lastClickedOrderId = record.id;
            this.selectedOrderId = record.id;
            this.requestedAction = actionName;
            this.requestedStatus = this.normalizeOutgoingStatus(nextStatus);
            this.rawRequestedNotes = "";
            this.closeTakeModal();
            this.emitOutputEvent();
        });

        if (actionName === "take") {
            footer.appendChild(cancelBtn);
        }
        footer.appendChild(takeBtn);

        modalCard.appendChild(header);
        modalCard.appendChild(middle);
        modalCard.appendChild(footer);
        overlay.appendChild(modalCard);
        document.body.appendChild(overlay);
        this.takeModal = overlay;
    }

    private closeTakeModal(): void {
        if (this.takeModal && this.takeModal.parentElement) {
            this.takeModal.parentElement.removeChild(this.takeModal);
        }
        this.takeModal = undefined;
    }

    private cleanModal?: HTMLDivElement;

    private openCleanModal(
        context: ComponentFramework.Context<IInputs>,
        record: OrderRecordViewModel,
        nextStatus: string
    ): void {
        if (this.cleanModal && this.cleanModal.parentElement) {
            this.cleanModal.parentElement.removeChild(this.cleanModal);
        }
        this.cleanModal = undefined;

        const isFr = this.getLanguage(context) === "fr";

        const overlay = this.createElement("div", {
            alignItems: "center",
            background: "rgba(16, 24, 40, 0.45)",
            bottom: "0",
            display: "flex",
            justifyContent: "center",
            left: "0",
            padding: "16px",
            position: "fixed",
            right: "0",
            top: "0",
            zIndex: "9999"
        });

        const modalCard = this.createElement("div", {
            background: "#ffffff",
            border: "1px solid #eaecf0",
            borderRadius: "11px",
            boxShadow: "0 10px 24px rgba(16, 24, 40, 0.2)",
            display: "flex",
            flexDirection: "column",
            maxHeight: "90vh",
            maxWidth: "420px",
            overflowY: "auto",
            position: "relative",
            width: "100%"
        });

        // Header
        const header = this.createElement("div", {
            alignItems: "flex-start",
            display: "flex",
            gap: "12px",
            justifyContent: "space-between",
            padding: "15px 16px 12px 16px"
        });

        const left = this.createElement("div", {
            alignItems: "flex-start",
            display: "flex",
            flex: "1 1 auto",
            gap: "10px",
            minWidth: "0"
        });

        const avatar = this.createAvatar(record, context.parameters.defaultUserPhoto.raw ?? undefined);
        const details = this.createElement("div", {
            display: "flex",
            flexDirection: "column",
            gap: "4px",
            minWidth: "0"
        });
        details.appendChild(this.createElement("div", {
            color: "#667085",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "10px",
            fontStyle: "italic",
            fontWeight: "500",
            lineHeight: "15px"
        }, record.orderNumber));
        details.appendChild(this.createElement("div", {
            color: "#101828",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "600",
            lineHeight: "20px",
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap"
        }, record.createdBy));
        const meta = this.createElement("div", {
            color: "#667085",
            display: "flex",
            flexWrap: "wrap",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "10px",
            fontWeight: "700",
            gap: "4px",
            lineHeight: "16px"
        });
        const hideItemCount = this.shouldHideItemCount(record.status);
        if (!hideItemCount) {
            meta.appendChild(this.createElement("span", { fontWeight: "400" }, this.formatQuantity(record.quantity, this.getLanguage(context))));
            meta.appendChild(this.createElement("span", undefined, "•"));
        }
        meta.appendChild(this.createElement("span", { fontWeight: "400" }, record.createdTime));
        details.appendChild(meta);
        left.appendChild(avatar);
        left.appendChild(details);

        const badgeTheme = this.getStatusTheme(record.status);
        const badge = this.createElement("div", {
            alignItems: "center",
            background: badgeTheme.background,
            borderRadius: "2px",
            display: "flex",
            flex: "0 0 auto",
            justifyContent: "center",
            minWidth: "93px",
            padding: "5px 10px"
        });
        badge.appendChild(this.createElement("span", {
            color: badgeTheme.foreground,
            fontFamily: "Nunito Sans, Inter, Segoe UI, sans-serif",
            fontSize: "12px",
            fontWeight: "700",
            lineHeight: "16px",
            textAlign: "center"
        }, this.getLocalizedStatus(record.status, this.getLanguage(context))));

        const closeBtn = this.createElement("button", {
            alignItems: "center",
            background: "transparent",
            border: "none",
            color: "#344054",
            cursor: "pointer",
            display: "inline-flex",
            fontSize: "24px",
            fontWeight: "400",
            height: "28px",
            justifyContent: "center",
            lineHeight: "1",
            padding: "0",
            width: "28px"
        }, "×") as HTMLButtonElement;
        closeBtn.type = "button";
        closeBtn.addEventListener("click", () => {
            if (this.cleanModal && this.cleanModal.parentElement) {
                this.cleanModal.parentElement.removeChild(this.cleanModal);
            }
            this.cleanModal = undefined;
        });

        const headerRight = this.createElement("div", {
            alignItems: "center",
            display: "flex",
            flex: "0 0 auto",
            gap: "6px"
        });
        headerRight.appendChild(badge);
        headerRight.appendChild(closeBtn);

        header.appendChild(left);
        header.appendChild(headerRight);

        // Products list
        const middle = this.createElement("div", {
            borderTop: "1px solid #eaecf0",
            display: "flex",
            flexDirection: "column",
            gap: "12px",
            padding: "14px 16px",
            width: "100%"
        });

        record.takeProducts.forEach((item) => {
            const row = this.createElement("div", { alignItems: "center", display: "flex", gap: "12px" });
            row.appendChild(this.createElement("span", {
                color: "#667085",
                fontFamily: "Inter, Segoe UI, sans-serif",
                fontSize: "14px",
                fontWeight: "500",
                lineHeight: "20px"
            }, item.quantity));
            row.appendChild(this.createElement("span", {
                color: "#101828",
                fontFamily: "Inter, Segoe UI, sans-serif",
                fontSize: "14px",
                fontWeight: "700",
                lineHeight: "20px",
                flex: "1 1 auto",
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap"
            }, item.label));
            middle.appendChild(row);
        });

        const noteBlock = this.createNoteBlock(context, record.note);
        if (noteBlock) {
            middle.appendChild(noteBlock);
        }

        // Footer
        const footer = this.createElement("div", {
            alignItems: "center",
            borderTop: "1px solid #eaecf0",
            display: "flex",
            gap: "14px",
            justifyContent: "center",
            padding: "10px 16px 16px 16px"
        });

        const cancelCleanBtn = this.createElement("button", {
            alignItems: "center",
            background: "transparent",
            border: "none",
            color: "#d92d20",
            cursor: "pointer",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "500",
            gap: "8px",
            minHeight: "36px",
            padding: "8px 10px"
        }) as HTMLButtonElement;
        cancelCleanBtn.type = "button";
        const cancelCleanContent = this.createElement("div", { alignItems: "center", display: "flex", gap: "6px" });
        cancelCleanContent.appendChild(this.createElement("span", undefined, isFr ? "Annuler le nettoyage" : "Cancel clear"));
        cancelCleanContent.appendChild(this.createCancelIcon());
        cancelCleanBtn.appendChild(cancelCleanContent);
        cancelCleanBtn.addEventListener("click", () => {
            if (this.cleanModal && this.cleanModal.parentElement) {
                this.cleanModal.parentElement.removeChild(this.cleanModal);
            }
            this.cleanModal = undefined;
            this.openCancelOrderModal(context, record, "clean", nextStatus);
        });

        const markCleanBtn = this.createElement("button", {
            alignItems: "center",
            background: "#121926",
            border: "1px solid #121926",
            borderRadius: "1px",
            boxShadow: "0 1px 2px rgba(16, 24, 40, 0.05)",
            color: "#ffffff",
            cursor: "pointer",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "500",
            gap: "8px",
            justifyContent: "center",
            minHeight: "36px",
            minWidth: "130px",
            padding: "8px 14px"
        }) as HTMLButtonElement;
        markCleanBtn.type = "button";
        const markCleanContent = this.createElement("div", { alignItems: "center", display: "flex", gap: "6px" });
        markCleanContent.appendChild(this.createElement("span", undefined, isFr ? "Marquer comme nettoyé" : "Mark cleared"));
        markCleanContent.appendChild(this.createActionIcon("clean"));
        markCleanBtn.appendChild(markCleanContent);
        markCleanBtn.addEventListener("click", () => {
            this.lastClickedOrderId = record.id;
            this.selectedOrderId = record.id;
            this.requestedAction = "clean";
            this.requestedStatus = this.normalizeOutgoingStatus(nextStatus);
            this.rawRequestedNotes = "";
            if (this.cleanModal && this.cleanModal.parentElement) {
                this.cleanModal.parentElement.removeChild(this.cleanModal);
            }
            this.cleanModal = undefined;
            this.emitOutputEvent();
        });

        footer.appendChild(cancelCleanBtn);
        footer.appendChild(markCleanBtn);

        modalCard.appendChild(header);
        modalCard.appendChild(middle);
        modalCard.appendChild(footer);
        overlay.appendChild(modalCard);
        document.body.appendChild(overlay);
        this.cleanModal = overlay;
    }

    private openCancelOrderModal(
        context: ComponentFramework.Context<IInputs>,
        record: OrderRecordViewModel,
        actionName: string,
        nextStatus: string
    ): void {
        this.closeTakeModal();
        this.closeCancelOrderModal();

        const isFr = this.getLanguage(context) === "fr";
        const logoUrl = context.parameters.logoUrl?.raw ?? "";

        const t = {
            title: isFr
                ? "Êtes-vous certain de vouloir abandonner ?"
                : "Are you certain you wish to abandon?",
            subtitle: isFr
                ? "Cette action est irréversible, la commande apparaîtra comme abandonnée."
                : "This action is irreversable, the order will appear as abandonned.",
            notesLabel: isFr ? "Notes (Optionnel)" : "Notes (Optional)",
            notesPlaceholder: "Write here",
            notesHelper: isFr
                ? "Veuillez indiquer la raison de l'abandon de la commande. Ex. : Aucun des articles n'est disponible."
                : "Please indicate the reason for abandonning the order. Eg. Non of the items are available.",
            backBtn: isFr ? "Retour" : "Back",
            cancelOrderBtn: isFr ? "Annuler la commande" : "Cancel order"
        };

        const overlay = this.createElement("div", {
            alignItems: "center",
            background: "rgba(16, 24, 40, 0.45)",
            bottom: "0",
            display: "flex",
            justifyContent: "center",
            left: "0",
            padding: "16px",
            position: "fixed",
            right: "0",
            top: "0",
            zIndex: "10000"
        });

        const popup = this.createElement("div", {
            background: "#FFFFFF",
            borderRadius: "11px",
            boxSizing: "border-box",
            display: "flex",
            flexDirection: "column",
            maxWidth: "324px",
            overflow: "hidden",
            position: "relative",
            width: "100%"
        });

        // Close button
        const closeBtn = this.createElement("button", {
            alignItems: "center",
            background: "transparent",
            border: "none",
            color: "#667085",
            cursor: "pointer",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "22px",
            fontWeight: "400",
            height: "28px",
            justifyContent: "center",
            lineHeight: "1",
            padding: "0",
            position: "absolute",
            right: "14px",
            top: "12px",
            width: "28px",
            zIndex: "1"
        }, "×") as HTMLButtonElement;
        closeBtn.type = "button";
        closeBtn.addEventListener("click", () => { this.closeCancelOrderModal(); });
        popup.appendChild(closeBtn);

        // Body
        const body = this.createElement("div", {
            display: "flex",
            flexDirection: "column",
            gap: "12px",
            padding: "18px 24px 4px 24px",
            width: "100%"
        });

        // Logo section
        const logoSection = this.createElement("div", {
            alignItems: "center",
            display: "flex",
            flexDirection: "column",
            gap: "2px",
            marginBottom: "4px"
        });
        if (logoUrl) {
            const logoImg = this.createElement("img", {
                display: "block",
                height: "auto",
                maxHeight: "72px",
                maxWidth: "130px",
                objectFit: "contain"
            }) as HTMLImageElement;
            logoImg.src = logoUrl;
            logoImg.alt = "Logo";
            logoSection.appendChild(logoImg);
        } else {
            logoSection.appendChild(this.createElement("div", {
                color: "#101828",
                fontFamily: "Georgia, 'Times New Roman', serif",
                fontSize: "13px",
                fontWeight: "700",
                letterSpacing: "3px",
                lineHeight: "1.2",
                textAlign: "center"
            }, "HERMÈS"));
            logoSection.appendChild(this.createElement("div", {
                color: "#101828",
                fontFamily: "Georgia, 'Times New Roman', serif",
                fontSize: "10px",
                fontWeight: "400",
                letterSpacing: "3px",
                lineHeight: "1.2",
                textAlign: "center"
            }, "PARIS"));
        }
        body.appendChild(logoSection);

        // Title
        body.appendChild(this.createElement("div", {
            color: "#101828",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "18px",
            fontWeight: "700",
            lineHeight: "1.4",
            textAlign: "center"
        }, t.title));

        // Subtitle
        body.appendChild(this.createElement("div", {
            color: "#667085",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "13px",
            lineHeight: "1.5",
            textAlign: "center"
        }, t.subtitle));

        // Notes label
        body.appendChild(this.createElement("div", {
            color: "#344054",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "600",
            lineHeight: "20px",
            marginTop: "2px"
        }, t.notesLabel));

        // Textarea container (with ? icon)
        const textareaWrap = this.createElement("div", {
            position: "relative",
            width: "100%"
        });
        const noteInput = this.createElement("textarea", {
            background: "#f9fafb",
            border: "1px solid #d0d5dd",
            borderRadius: "8px",
            boxSizing: "border-box",
            color: "#667085",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            lineHeight: "20px",
            minHeight: "96px",
            outline: "none",
            padding: "12px 40px 12px 14px",
            resize: "none",
            width: "100%"
        }) as HTMLTextAreaElement;
        noteInput.placeholder = t.notesPlaceholder;
        const helpIcon = this.createElement("div", {
            alignItems: "center",
            border: "1.5px solid #d0d5dd",
            borderRadius: "999px",
            color: "#98a2b3",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "11px",
            fontWeight: "700",
            height: "20px",
            justifyContent: "center",
            lineHeight: "1",
            pointerEvents: "none",
            position: "absolute",
            right: "12px",
            top: "12px",
            width: "20px"
        }, "?");
        textareaWrap.appendChild(noteInput);
        textareaWrap.appendChild(helpIcon);
        body.appendChild(textareaWrap);

        // Helper text
        body.appendChild(this.createElement("div", {
            color: "#667085",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "12px",
            lineHeight: "18px",
            marginBottom: "4px"
        }, t.notesHelper));

        popup.appendChild(body);

        // Footer
        const footer = this.createElement("div", {
            alignItems: "center",
            borderTop: "1px solid #eaecf0",
            display: "flex",
            gap: "12px",
            justifyContent: "space-between",
            padding: "12px 24px 16px 24px"
        });

        const backBtn = this.createElement("button", {
            alignItems: "center",
            background: "#121926",
            border: "1px solid #121926",
            borderRadius: "6px",
            color: "#FFFFFF",
            cursor: "pointer",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "600",
            gap: "8px",
            minHeight: "38px",
            padding: "8px 16px"
        }) as HTMLButtonElement;
        backBtn.type = "button";
        backBtn.appendChild(this.createBackIcon());
        backBtn.appendChild(this.createElement("span", undefined, t.backBtn));
        backBtn.addEventListener("click", () => {
            this.closeCancelOrderModal();
            if (actionName === "clean") {
                this.openCleanModal(context, record, nextStatus);
                return;
            }

            this.openTakeModal(context, record, actionName, nextStatus);
        });

        const cancelOrderBtn = this.createElement("button", {
            alignItems: "center",
            background: "transparent",
            border: "none",
            color: "#d92d20",
            cursor: "pointer",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "14px",
            fontWeight: "600",
            gap: "6px",
            minHeight: "38px",
            padding: "8px 0"
        }) as HTMLButtonElement;
        cancelOrderBtn.type = "button";
        cancelOrderBtn.appendChild(this.createCancelIcon());
        cancelOrderBtn.appendChild(this.createElement("span", undefined, t.cancelOrderBtn));
        cancelOrderBtn.addEventListener("click", () => {
            this.lastClickedOrderId = record.id;
            this.selectedOrderId = record.id;
            this.requestedAction = "cancel";
            this.requestedStatus = "Cancelled";
            this.rawRequestedNotes = noteInput.value.trim();
            this.closeCancelOrderModal();
            this.emitOutputEvent();
        });

        footer.appendChild(backBtn);
        footer.appendChild(cancelOrderBtn);
        popup.appendChild(footer);

        overlay.appendChild(popup);
        document.body.appendChild(overlay);
        this.cancelOrderModal = overlay;
    }

    private closeCancelOrderModal(): void {
        if (this.cancelOrderModal && this.cancelOrderModal.parentElement) {
            this.cancelOrderModal.parentElement.removeChild(this.cancelOrderModal);
        }
        this.cancelOrderModal = undefined;
    }

    private createAvatar(record: OrderRecordViewModel, configuredPhoto?: string): HTMLDivElement {
        const avatar = this.createElement("div", {
            alignItems: "center",
            background: "linear-gradient(135deg, #fda29b 0%, #fec84b 100%)",
            borderRadius: "999px",
            color: "#ffffff",
            display: "flex",
            flex: "0 0 24px",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "11px",
            fontWeight: "700",
            height: "24px",
            justifyContent: "center",
            overflow: "hidden",
            position: "relative",
            width: "24px"
        }, this.getInitials(record.createdBy));

        const primaryPhoto = this.normalizeImageSource(record.createdByImageUrl);
        const defaultPhoto = this.normalizeImageSource(configuredPhoto);

        if (!primaryPhoto && !defaultPhoto) {
            return avatar;
        }

        const image = this.createElement("img", {
            height: "100%",
            left: "0",
            objectFit: "cover",
            position: "absolute",
            top: "0",
            width: "100%"
        }) as HTMLImageElement;

        image.alt = record.createdBy;

        let hasTriedDefault = false;
        if (primaryPhoto) {
            image.src = primaryPhoto;
        } else if (defaultPhoto) {
            image.src = defaultPhoto;
            hasTriedDefault = true;
        }

        image.addEventListener("error", () => {
            if (!hasTriedDefault && defaultPhoto) {
                hasTriedDefault = true;
                image.src = defaultPhoto;
                return;
            }

            image.remove();
        });

        avatar.appendChild(image);
        return avatar;
    }

    private normalizeImageSource(value: string | undefined): string | undefined {
        if (!value) {
            return undefined;
        }

        const trimmed = value.trim();
        if (!trimmed) {
            return undefined;
        }

        if (trimmed.startsWith("\"") && trimmed.endsWith("\"")) {
            try {
                const parsed = JSON.parse(trimmed);
                return typeof parsed === "string" && parsed.trim() ? parsed.trim() : undefined;
            } catch {
                return trimmed;
            }
        }

        return trimmed;
    }

    private getOrders(
        context: ComponentFramework.Context<IInputs>,
        orders: ComponentFramework.PropertyTypes.DataSet
    ): OrderRecordViewModel[] {
        const orderNumberColumn = this.resolveColumnName(
            orders,
            context.parameters.orderNumberColumn.raw,
            CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.orderNumber
        );
        const createdByColumn = this.resolveColumnName(
            orders,
            context.parameters.createdByColumn.raw,
            CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.createdBy
        );
        const statusColumn = this.resolveColumnName(
            orders,
            context.parameters.statusColumn.raw,
            CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.status
        );
        const quantityColumn = this.resolveColumnName(
            orders,
            context.parameters.quantityColumn.raw,
            CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.quantity
        );
        const itemIdColumn = this.resolveColumnName(
            orders,
            null,
            CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.itemId
        );
        const modifiedOnColumn = this.resolveColumnName(
            orders,
            null,
            CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.modifiedOn
        );
        const createdOnColumn = this.resolveColumnName(
            orders,
            context.parameters.createdOnColumn.raw,
            CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.createdOn
        );
        const noteColumn = this.resolveColumnName(
            orders,
            null,
            CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.note
        );
        const productsColumn = this.resolveColumnName(
            orders,
            null,
            CardOrderAgentPcfLast23.DEFAULT_COLUMN_CANDIDATES.products
        );

        return orders.sortedRecordIds
            .map((datasetRecordId) => {
                const record = orders.records[datasetRecordId];
                if (!record) {
                    return undefined;
                }

                return this.toOrderViewModel(
                    record,
                    orderNumberColumn,
                    createdByColumn,
                    statusColumn,
                    quantityColumn,
                    createdOnColumn,
                    itemIdColumn,
                    modifiedOnColumn,
                    noteColumn,
                    productsColumn,
                    context.userSettings.userName,
                    this.getLanguage(context),
                    datasetRecordId
                );
            })
            .filter((record): record is OrderRecordViewModel => Boolean(record))
            .sort((left, right) => {
                const leftTime = (left.modifiedOn ?? left.createdOn)?.getTime() ?? 0;
                const rightTime = (right.modifiedOn ?? right.createdOn)?.getTime() ?? 0;
                return rightTime - leftTime;
            });
    }

    private toOrderViewModel(
        record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
        orderNumberColumn?: string,
        createdByColumn?: string,
        statusColumn?: string,
        quantityColumn?: string,
        createdOnColumn?: string,
        itemIdColumn?: string,
        modifiedOnColumn?: string,
        noteColumn?: string,
        productsColumn?: string,
        currentUserName?: string,
        language?: LanguageCode,
        datasetRecordId?: string
    ): OrderRecordViewModel | undefined {
        const createdOn = createdOnColumn ? this.toDate(record.getValue(createdOnColumn), record.getFormattedValue(createdOnColumn)) : undefined;
        const modifiedOn = modifiedOnColumn ? this.toDate(record.getValue(modifiedOnColumn), record.getFormattedValue(modifiedOnColumn)) : undefined;

        const orderNumber = this.getDisplayValue(record, orderNumberColumn, "Order without number");
        const createdByIdentity = this.getPersonIdentity(record, createdByColumn);
        const createdBy = createdByIdentity?.displayName || this.getDisplayValue(record, createdByColumn, "Unknown user");
        const note = this.getDisplayValue(record, noteColumn, "");
        const status = this.getDisplayValue(record, statusColumn, "Unknown");
        const quantity = this.getDisplayValue(record, quantityColumn, "-");
        const productsRaw = productsColumn ? record.getValue(productsColumn) : undefined;
        const products = productsColumn ? this.getProducts(productsRaw) : [];
        const takeProducts = productsColumn ? this.getTakeProducts(productsRaw, language ?? "en") : [];
        const itemId = this.getSharePointItemId(record, itemIdColumn, datasetRecordId);
        const action = this.getActionForStatus(status, createdBy, currentUserName ?? "", this.getLanguageFromStatus(status));

        return {
            actionLabel: action?.label,
            actionName: action?.name,
            createdBy,
            createdByImageUrl: createdByIdentity?.imageUrl,
            createdOn,
            createdTime: createdOn ? this.formatTime(createdOn) : "-",
            datasetRecordId: datasetRecordId ?? record.getRecordId(),
            id: itemId,
            modifiedOn,
            nextStatus: action?.nextStatus,
            note,
            orderNumber,
            products,
            takeProducts,
            quantity,
            status
        };
    }

    private createProductsBlock(products: ProductLine[]): HTMLDivElement | undefined {
        if (products.length === 0) {
            return undefined;
        }

        const block = this.createElement("div", {
            alignItems: "flex-start",
            boxSizing: "border-box",
            display: "flex",
            flexDirection: "column",
            gap: "10px",
            minHeight: "86px",
            padding: "0",
            width: "100%"
        });

        products.forEach((item) => {
            const row = this.createElement("div", {
                alignItems: "baseline",
                background: "#f5f5f5",
                borderRadius: "999px",
                display: "flex",
                gap: "0",
                maxWidth: "100%",
                padding: "6px 8px",
                width: "fit-content"
            });

            row.appendChild(this.createElement("span", {
                color: "rgba(93, 107, 152, 1)",
                fontFamily: "Inter, Segoe UI, sans-serif",
                fontSize: "12px",
                fontWeight: "700",
                lineHeight: "18px",
                textAlign: "left"
            }, item.quantity));

            row.appendChild(this.createElement("span", {
                color: "rgba(93, 107, 152, 1)",
                fontFamily: "Inter, Segoe UI, sans-serif",
                fontSize: "12px",
                fontWeight: "700",
                lineHeight: "18px",
                marginLeft: "2px"
            }, item.label));

            block.appendChild(row);
        });

        return block;
    }

    private getProducts(value: unknown): ProductLine[] {
        if (!value) {
            return [];
        }

        if (typeof value === "string") {
            const trimmed = value.trim();
            if (!trimmed) {
                return [];
            }

            if ((trimmed.startsWith("[") && trimmed.endsWith("]")) || (trimmed.startsWith("{") && trimmed.endsWith("}"))) {
                try {
                    return this.getProducts(JSON.parse(trimmed));
                } catch {
                    return [];
                }
            }

            return [];
        }

        if (Array.isArray(value)) {
            return value
                .map((item) => this.toProductLine(item))
                .filter((item): item is ProductLine => Boolean(item));
        }

        if (typeof value !== "object") {
            return [];
        }

        const record = value as Record<string, unknown>;
        const nestedCollection = this.getFirstArray(record, ["items", "Items", "products", "Products", "produits", "Produits", "value", "Value"]);
        if (nestedCollection) {
            return this.getProducts(nestedCollection);
        }

        const single = this.toProductLine(record);
        return single ? [single] : [];
    }

    private getTakeProducts(value: unknown, language: LanguageCode): ProductLine[] {
        if (!value) {
            return [];
        }

        if (typeof value === "string") {
            const trimmed = value.trim();
            if (!trimmed) {
                return [];
            }

            try {
                return this.getTakeProducts(JSON.parse(trimmed), language);
            } catch {
                return [];
            }
        }

        if (Array.isArray(value)) {
            return value.flatMap((item) => this.getTakeProducts(item, language));
        }

        if (typeof value !== "object") {
            return [];
        }

        const record = value as Record<string, unknown>;

        // Popup products schema (legacy expected format): { Lignes: [{ NomFR, NomEN, Quantite }] }
        const nestedLines = this.getFirstArray(record, ["Lignes"]);
        if (nestedLines) {
            const result = nestedLines.flatMap((item) => this.getTakeProducts(item, language));
            return result;
        }

        const single = this.toTakeProductLine(record, language);
        if (single) {
            return [single];
        }

        return [];
    }

    private toTakeProductLine(value: unknown, language: LanguageCode): ProductLine | undefined {
        if (!value || typeof value !== "object") {
            return undefined;
        }

        const record = value as Record<string, unknown>;
        const label = language === "fr"
            ? this.getFirstString(record, ["NomFR", "NomEN"])
            : this.getFirstString(record, ["NomEN", "NomFR"]);
        
        // Try both string and number for Quantite
        let quantity: string | undefined = this.getFirstString(record, ["Quantite"]);
        if (!quantity) {
            const numValue = this.getFirstNumber(record, ["Quantite"]);
            quantity = numValue;
        }

        if (!label || !quantity) {
            return undefined;
        }

        return {
            label,
            quantity
        };
    }

    private toProductLine(value: unknown): ProductLine | undefined {
        if (!value || typeof value !== "object") {
            return undefined;
        }

        const record = value as Record<string, unknown>;
        const label = this.getFirstString(record, ["SectionNom", "sectionNom", "libellé", "libelle", "Libelle", "label", "Label", "title", "Title", "name", "Name"]);
        const quantityValue = this.getFirstString(record, ["NbProduits", "nbProduits", "QuantiteeTotale", "quantiteeTotale", "QuantiteTotale", "quantiteTotale", "quantity", "Quantity", "qty", "Qty"])
            ?? this.getFirstNumber(record, ["NbProduits", "nbProduits", "QuantiteeTotale", "quantiteeTotale", "QuantiteTotale", "quantiteTotale", "quantity", "Quantity", "qty", "Qty"]);

        if (!label || !quantityValue) {
            return undefined;
        }

        return {
            label,
            quantity: quantityValue
        };
    }

    private getFirstArray(record: Record<string, unknown>, keys: string[]): unknown[] | undefined {
        for (const key of keys) {
            const value = record[key];
            if (Array.isArray(value)) {
                return value;
            }
        }

        return undefined;
    }

    private getFirstNumber(record: Record<string, unknown>, keys: string[]): string | undefined {
        for (const key of keys) {
            const value = record[key];
            if (typeof value === "number" && Number.isFinite(value)) {
                return String(value);
            }
        }

        return undefined;
    }

    private getSharePointItemId(
        record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
        itemIdColumn?: string,
        datasetRecordId?: string
    ): string {
        const candidates = [itemIdColumn, "ID", "Id", "id"]
            .filter((value): value is string => Boolean(value && value.trim()))
            .map((value) => value.trim());

        for (const candidate of candidates) {
            const formatted = record.getFormattedValue(candidate)?.trim();
            if (formatted && /^\d+$/.test(formatted)) {
                return formatted;
            }

            const raw = record.getValue(candidate);
            if (typeof raw === "number" && Number.isFinite(raw)) {
                return String(Math.trunc(raw));
            }

            if (typeof raw === "string") {
                const trimmed = raw.trim();
                if (/^\d+$/.test(trimmed)) {
                    return trimmed;
                }
            }
        }

        if (typeof datasetRecordId === "string") {
            const trimmedDatasetId = datasetRecordId.trim();
            if (trimmedDatasetId.length > 0) {
                return trimmedDatasetId;
            }
        }

        // Fallback for non-SharePoint sources.
        return record.getRecordId();
    }

    private createNoteBlock(context: ComponentFramework.Context<IInputs>, note: string): HTMLDivElement | undefined {
        const noteValue = note.trim();
        if (!noteValue) {
            return undefined;
        }

        const block = this.createElement("div", {
            alignItems: "flex-start",
            background: "#f5f5f5",
            borderRadius: "6.20339px",
            display: "flex",
            flexDirection: "column",
            gap: "6.2px",
            justifyContent: "center",
            marginTop: "10px",
            minHeight: "86.61px",
            padding: "6.20339px 10px",
            width: "100%"
        });

        const titleRow = this.createElement("div", {
            alignItems: "center",
            color: "#667085",
            display: "flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "13px",
            fontWeight: "700",
            gap: "8px",
            lineHeight: "20px"
        });

        titleRow.appendChild(this.createNoteIcon());
        titleRow.appendChild(this.createElement("span", undefined, this.translate(context, "noteTitle")));

        block.appendChild(titleRow);
        block.appendChild(this.createElement("div", {
            color: "#3f46ff",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "12px",
            fontStyle: "italic",
            fontWeight: "500",
            lineHeight: "18px",
            whiteSpace: "pre-wrap"
        }, noteValue));

        return block;
    }

    private createNoteIcon(): SVGSVGElement {
        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("width", "16");
        iconSvg.setAttribute("height", "16");
        iconSvg.setAttribute("viewBox", "0 0 16 16");
        iconSvg.setAttribute("fill", "none");
        iconSvg.setAttribute("stroke", "#667085");
        iconSvg.setAttribute("stroke-width", "1.7");
        iconSvg.setAttribute("stroke-linecap", "round");
        iconSvg.setAttribute("stroke-linejoin", "round");
        iconSvg.setAttribute("aria-hidden", "true");

        const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
        path.setAttribute("d", "M6 2H3.5A1.5 1.5 0 0 0 2 3.5V6M10 2h2.5A1.5 1.5 0 0 1 14 3.5V6M14 10v2.5A1.5 1.5 0 0 1 12.5 14H10M6 14H3.5A1.5 1.5 0 0 1 2 12.5V10");
        iconSvg.appendChild(path);
        return iconSvg;
    }

    private createCancelIcon(): SVGSVGElement {
        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("width", "18");
        iconSvg.setAttribute("height", "18");
        iconSvg.setAttribute("viewBox", "0 0 18 18");
        iconSvg.setAttribute("fill", "none");
        iconSvg.setAttribute("stroke", "#d92d20");
        iconSvg.setAttribute("stroke-width", "1.7");
        iconSvg.setAttribute("stroke-linecap", "round");
        iconSvg.setAttribute("stroke-linejoin", "round");

        const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        circle.setAttribute("cx", "9");
        circle.setAttribute("cy", "9");
        circle.setAttribute("r", "7");
        iconSvg.appendChild(circle);

        const line1 = document.createElementNS("http://www.w3.org/2000/svg", "line");
        line1.setAttribute("x1", "6.5");
        line1.setAttribute("y1", "6.5");
        line1.setAttribute("x2", "11.5");
        line1.setAttribute("y2", "11.5");
        iconSvg.appendChild(line1);

        const line2 = document.createElementNS("http://www.w3.org/2000/svg", "line");
        line2.setAttribute("x1", "11.5");
        line2.setAttribute("y1", "6.5");
        line2.setAttribute("x2", "6.5");
        line2.setAttribute("y2", "11.5");
        iconSvg.appendChild(line2);

        return iconSvg;
    }

    private createBackIcon(): SVGSVGElement {
        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("width", "18");
        iconSvg.setAttribute("height", "18");
        iconSvg.setAttribute("viewBox", "0 0 18 18");
        iconSvg.setAttribute("fill", "none");
        iconSvg.setAttribute("stroke", "#ffffff");
        iconSvg.setAttribute("stroke-width", "1.7");
        iconSvg.setAttribute("stroke-linecap", "round");
        iconSvg.setAttribute("stroke-linejoin", "round");

        const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        circle.setAttribute("cx", "9");
        circle.setAttribute("cy", "9");
        circle.setAttribute("r", "7");
        iconSvg.appendChild(circle);

        const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
        line.setAttribute("x1", "11");
        line.setAttribute("y1", "9");
        line.setAttribute("x2", "7");
        line.setAttribute("y2", "9");
        iconSvg.appendChild(line);

        const arrow1 = document.createElementNS("http://www.w3.org/2000/svg", "line");
        arrow1.setAttribute("x1", "8.5");
        arrow1.setAttribute("y1", "7.5");
        arrow1.setAttribute("x2", "7");
        arrow1.setAttribute("y2", "9");
        iconSvg.appendChild(arrow1);

        const arrow2 = document.createElementNS("http://www.w3.org/2000/svg", "line");
        arrow2.setAttribute("x1", "8.5");
        arrow2.setAttribute("y1", "10.5");
        arrow2.setAttribute("x2", "7");
        arrow2.setAttribute("y2", "9");
        iconSvg.appendChild(arrow2);

        return iconSvg;
    }

    private createTabBar(
        context: ComponentFramework.Context<IInputs>,
        records: OrderRecordViewModel[],
        scroller: HTMLDivElement
    ): HTMLDivElement {
        const language = this.getLanguage(context);
        const usedStatuses = new Set(records.map((r) => this.toCanonicalStatus(r.status)));

        if (this.selectedStatus !== "all" && !usedStatuses.has(this.selectedStatus as CanonicalStatus)) {
            this.selectedStatus = "all";
        }

        const bar = this.createElement("div", {
            background: "#ffffff",
            border: "0.613636px solid #e5e7eb",
            borderRadius: "4.29545px",
            display: "flex",
            flexWrap: "wrap",
            flexShrink: "0",
            gap: "2px",
            overflowX: "hidden",
            padding: "4px",
            width: "100%"
        });

        const orderedStatuses: Exclude<CanonicalStatus, "cancel" | "unknown">[] = [
            "toPrepare", "inPrep", "toClean", "served", "cleaned", "cancelled"
        ];

        const statusCounts = new Map<CanonicalStatus, number>();
        for (const record of records) {
            const status = this.toCanonicalStatus(record.status);
            statusCounts.set(status, (statusCounts.get(status) ?? 0) + 1);
        }

        const totalCount = records.length;

        const tabs: { key: string; label: string }[] = [
            { key: "all", label: `${this.translate(context, "tabAll")} (${totalCount})` }
        ];

        for (const s of orderedStatuses) {
            if (usedStatuses.has(s)) {
                const count = statusCounts.get(s) ?? 0;
                tabs.push({ key: s, label: `${this.getStatusLabel(s, language)} (${count})` });
            }
        }

        const getTabStyles = (active: boolean): Partial<CSSStyleDeclaration> => ({
            alignItems: "center",
            background: active ? "#121926" : "transparent",
            border: active ? "1px solid #121926" : "1px solid #d0d5dd",
            borderRadius: "999px",
            color: active ? "#ffffff" : "#667085",
            cursor: "pointer",
            display: "inline-flex",
            fontFamily: "Inter, Segoe UI, sans-serif",
            fontSize: "11px",
            fontWeight: active ? "700" : "500",
            justifyContent: "center",
            padding: "3px 10px",
            whiteSpace: "nowrap"
        });

        const tabElements: HTMLButtonElement[] = [];

        for (const { key, label } of tabs) {
            const tab = this.createElement("button", getTabStyles(this.selectedStatus === key), label) as HTMLButtonElement;
            tab.type = "button";
            tab.addEventListener("click", () => {
                this.selectedStatus = key;
                tabElements.forEach((t, i) => {
                    this.applyStyles(t, getTabStyles(tabs[i].key === key));
                });
                scroller.querySelectorAll<HTMLElement>("[data-canonical-status]").forEach((cardEl) => {
                    cardEl.style.display = key === "all" || cardEl.dataset.canonicalStatus === key ? "" : "none";
                });
            });
            tabElements.push(tab);
            bar.appendChild(tab);
        }

        return bar;
    }

    private resolveColumnName(
        orders: ComponentFramework.PropertyTypes.DataSet,
        configuredName: string | null,
        fallbacks: string[]
    ): string | undefined {
        const columnsByName = new Map<string, string>();
        const columnsByDisplay = new Map<string, string>();
        orders.columns.forEach((column) => {
            columnsByName.set(column.name.trim().toLowerCase(), column.name);
            columnsByDisplay.set(column.displayName.trim().toLowerCase(), column.name);
        });

        const configured = configuredName?.trim();
        if (configured) {
            const configLower = configured.toLowerCase();
            
            // Search by internal name first
            const byName = columnsByName.get(configLower);
            if (byName) {
                return byName;
            }

            // Then by display name
            const byDisplay = columnsByDisplay.get(configLower);
            if (byDisplay) {
                return byDisplay;
            }
        }

        // Try fallbacks
        for (const fallback of fallbacks) {
            const fallbackLower = fallback.trim().toLowerCase();
            
            // Search by internal name first
            const byName = columnsByName.get(fallbackLower);
            if (byName) {
                return byName;
            }

            // Then by display name
            const byDisplay = columnsByDisplay.get(fallbackLower);
            if (byDisplay) {
                return byDisplay;
            }

            // Try normalized matching (underscores to spaces)
            const normalized = fallbackLower.replace(/_/g, " ");
            for (const [key, name] of columnsByName.entries()) {
                if (key.replace(/_/g, " ") === normalized) {
                    return name;
                }
            }
            for (const [key, name] of columnsByDisplay.entries()) {
                if (key.replace(/_/g, " ") === normalized) {
                    return name;
                }
            }
        }

        return configured;
    }

    private getDisplayValue(
        record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
        columnName: string | undefined,
        fallback: string
    ): string {
        if (!columnName) {
            return fallback;
        }

        const formatted = record.getFormattedValue(columnName);
        const personFromFormatted = this.tryGetPersonDisplayName(formatted);
        if (personFromFormatted) {
            return personFromFormatted;
        }
        if (formatted) {
            return formatted;
        }

        const raw = record.getValue(columnName);
        const personFromRaw = this.tryGetPersonDisplayName(raw);
        if (personFromRaw) {
            return personFromRaw;
        }
        if (raw instanceof Date) {
            return raw.toLocaleString();
        }

        if (typeof raw === "string" || typeof raw === "number" || typeof raw === "boolean") {
            return String(raw);
        }

        if (raw && typeof raw === "object" && "name" in raw) {
            return String(raw.name ?? fallback);
        }

        return fallback;
    }

    private getPersonIdentity(
        record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
        columnName: string | undefined
    ): PersonIdentity | undefined {
        if (!columnName) {
            return undefined;
        }

        const raw = record.getValue(columnName);
        const fromRaw = this.extractPersonIdentity(raw);
        if (fromRaw) {
            return fromRaw;
        }

        const formatted = record.getFormattedValue(columnName);
        const fromFormatted = this.extractPersonIdentity(formatted);
        if (fromFormatted) {
            return fromFormatted;
        }

        return undefined;
    }

    private extractPersonIdentity(value: unknown): PersonIdentity | undefined {
        if (!value) {
            return undefined;
        }

        if (typeof value === "string") {
            const trimmed = value.trim();
            if (!trimmed) {
                return undefined;
            }

            if ((trimmed.startsWith("{") && trimmed.endsWith("}")) || (trimmed.startsWith("[") && trimmed.endsWith("]"))) {
                try {
                    return this.extractPersonIdentity(JSON.parse(trimmed));
                } catch {
                    return undefined;
                }
            }

            return undefined;
        }

        if (Array.isArray(value)) {
            for (const item of value) {
                const identity = this.extractPersonIdentity(item);
                if (identity) {
                    return identity;
                }
            }

            return undefined;
        }

        if (typeof value !== "object") {
            return undefined;
        }

        const record = value as Record<string, unknown>;
        const displayName = this.getFirstString(record, ["DisplayName", "displayName", "name", "Name", "Title", "title"]);
        let email = this.getFirstString(record, ["Email", "email", "UserPrincipalName", "userPrincipalName"]);
        const imageUrl = this.getFirstString(record, ["Picture", "picture", "PictureUrl", "pictureUrl", "ImageUrl", "imageUrl", "Photo", "photo", "CreatorPhoto", "creatorPhoto"]);

        if (!email) {
            const claims = this.getFirstString(record, ["Claims", "claims"]);
            if (claims) {
                const claimParts = claims.split("|");
                const claimEmail = claimParts[claimParts.length - 1]?.trim();
                if (claimEmail && claimEmail.includes("@")) {
                    email = claimEmail;
                }
            }
        }

        if ("Value" in record) {
            const nested = this.extractPersonIdentity(record.Value);
            if (nested) {
                return {
                    displayName: displayName || nested.displayName,
                    email: email || nested.email,
                    imageUrl: imageUrl || nested.imageUrl
                };
            }
        }

        if (!displayName && !email && !imageUrl) {
            return undefined;
        }

        return {
            displayName,
            email,
            imageUrl
        };
    }

    private getFirstString(record: Record<string, unknown>, keys: string[]): string | undefined {
        for (const key of keys) {
            const value = record[key];
            if (typeof value === "string" && value.trim().length > 0) {
                return value.trim();
            }
        }

        return undefined;
    }

    private tryGetPersonDisplayName(value: unknown): string | undefined {
        if (!value) {
            return undefined;
        }

        if (typeof value === "string") {
            const trimmed = value.trim();
            if (!trimmed) {
                return undefined;
            }

            if ((trimmed.startsWith("{") && trimmed.endsWith("}")) || (trimmed.startsWith("[") && trimmed.endsWith("]"))) {
                try {
                    return this.extractPersonDisplayName(JSON.parse(trimmed));
                } catch {
                    return undefined;
                }
            }

            return undefined;
        }

        if (typeof value === "object") {
            return this.extractPersonDisplayName(value);
        }

        return undefined;
    }

    private extractPersonDisplayName(value: unknown): string | undefined {
        if (!value) {
            return undefined;
        }

        if (Array.isArray(value)) {
            for (const item of value) {
                const candidate = this.extractPersonDisplayName(item);
                if (candidate) {
                    return candidate;
                }
            }
            return undefined;
        }

        if (typeof value !== "object") {
            return undefined;
        }

        const record = value as Record<string, unknown>;
        const candidates = ["DisplayName", "displayName", "name", "Name", "Title", "title", "Email", "email"];
        for (const key of candidates) {
            const data = record[key];
            if (typeof data === "string" && data.trim().length > 0) {
                return data.trim();
            }
        }

        if ("Value" in record) {
            return this.extractPersonDisplayName(record.Value);
        }

        return undefined;
    }

    private toDate(value: string | number | boolean | Date | number[] | ComponentFramework.EntityReference | ComponentFramework.EntityReference[] | ComponentFramework.LookupValue | ComponentFramework.LookupValue[] | undefined, formattedValue?: string): Date | undefined {
        if (value instanceof Date) {
            return value;
        }

        if (typeof value === "string" || typeof value === "number") {
            const date = new Date(value);
            if (!Number.isNaN(date.getTime())) {
                return date;
            }
        }

        if (formattedValue) {
            const formattedDate = new Date(formattedValue);
            if (!Number.isNaN(formattedDate.getTime())) {
                return formattedDate;
            }
        }

        return undefined;
    }

    private formatTime(value: Date): string {
        return value.toLocaleTimeString([], {
            hour: "numeric",
            minute: "2-digit",
            hour12: true
        });
    }

    private formatQuantity(quantity: string, language: LanguageCode): string {
        const trimmedQuantity = quantity.trim();
        const singular = trimmedQuantity === "1";

        if (language === "fr") {
            return `${trimmedQuantity} ${singular ? "Item" : "Items"}`;
        }

        return `${trimmedQuantity} ${singular ? "Item" : "Items"}`;
    }

    private createActionIcon(actionName: string): SVGSVGElement {
        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("width", "16");
        iconSvg.setAttribute("height", "16");
        iconSvg.setAttribute("viewBox", "0 0 16 16");
        iconSvg.setAttribute("fill", "none");
        iconSvg.setAttribute("stroke", "currentColor");
        iconSvg.setAttribute("stroke-width", "1.67");
        iconSvg.setAttribute("stroke-linecap", "round");
        iconSvg.setAttribute("stroke-linejoin", "round");
        iconSvg.setAttribute("aria-hidden", "true");

        if (actionName === "take") {
            const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
            circle.setAttribute("cx", "8");
            circle.setAttribute("cy", "5");
            circle.setAttribute("r", "2.5");

            const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
            path.setAttribute("d", "M3.5 13c0-2.2 2-4 4.5-4s4.5 1.8 4.5 4");

            iconSvg.appendChild(circle);
            iconSvg.appendChild(path);
            return iconSvg;
        }

        if (actionName === "serve") {
            iconSvg.setAttribute("width", "19");
            iconSvg.setAttribute("height", "19");
            iconSvg.setAttribute("viewBox", "0 0 19 19");

            const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
            path.setAttribute("d", "M17.5016 8.4064V9.17306C17.5006 10.9701 16.9187 12.7186 15.8427 14.1579C14.7668 15.5972 13.2544 16.6501 11.5311 17.1596C9.80782 17.6692 7.96601 17.608 6.28035 16.9852C4.5947 16.3625 3.15551 15.2115 2.17743 13.704C1.19935 12.1964 0.734787 10.4131 0.853025 8.61999C0.971263 6.82687 1.66597 5.12 2.83353 3.75396C4.00109 2.38791 5.57895 1.43588 7.33179 1.03985C9.08462 0.643822 10.9185 0.825012 12.56 1.5564M17.5016 2.5064L9.1683 10.8481L6.6683 8.34806");

            iconSvg.appendChild(path);
            return iconSvg;
        }

        if (actionName === "clean") {
            iconSvg.setAttribute("width", "19");
            iconSvg.setAttribute("height", "19");
            iconSvg.setAttribute("viewBox", "0 0 19 19");

            const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
            path.setAttribute("d", "M17.5016 8.4064V9.17306C17.5006 10.9701 16.9187 12.7186 15.8427 14.1579C14.7668 15.5972 13.2544 16.6501 11.5311 17.1596C9.80782 17.6692 7.96601 17.608 6.28035 16.9852C4.5947 16.3625 3.15551 15.2115 2.17743 13.704C1.19935 12.1964 0.734787 10.4131 0.853025 8.61999C0.971263 6.82687 1.66597 5.12 2.83353 3.75396C4.00109 2.38791 5.57895 1.43588 7.33179 1.03985C9.08462 0.643822 10.9185 0.825012 12.56 1.5564M17.5016 2.5064L9.1683 10.8481L6.6683 8.34806");

            iconSvg.appendChild(path);
            return iconSvg;
        }

        if (actionName === "view") {
            const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
            path.setAttribute("d", "M1 8s2.5-5 7-5 7 5 7 5-2.5 5-7 5-7-5-7-5z");

            const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
            circle.setAttribute("cx", "8");
            circle.setAttribute("cy", "8");
            circle.setAttribute("r", "2");

            iconSvg.appendChild(path);
            iconSvg.appendChild(circle);
            return iconSvg;
        }

        const polyline = document.createElementNS("http://www.w3.org/2000/svg", "polyline");
        polyline.setAttribute("points", "2 8 6 12 14 4");
        iconSvg.appendChild(polyline);
        return iconSvg;
    }

    private getInitials(fullName: string): string {
        const parts = fullName
            .split(" ")
            .map((value) => value.trim())
            .filter((value) => value.length > 0);

        if (parts.length === 0) {
            return "?";
        }

        const first = parts[0][0] ?? "";
        const second = parts.length > 1 ? parts[1][0] ?? "" : "";
        return `${first}${second}`.toUpperCase();
    }

    private getActionForStatus(status: string, createdBy: string, currentUserName: string, outputLanguage: LanguageCode): OrderActionMapping | undefined {
        const normalizedStatus = this.toCanonicalStatus(status);
        const isOwner = this.normalize(createdBy) === this.normalize(currentUserName);

        if (normalizedStatus === "toPrepare") {
            return {
                label: "Take",
                name: "take",
                nextStatus: "In prep"
            };
        }

        if (normalizedStatus === "inPrep") {
            return {
                label: "Make as served",
                name: "serve",
                nextStatus: "Served"
            };
        }

        if (normalizedStatus === "toClean") {
            return {
                label: "Mark cleared",
                name: "clean",
                nextStatus: "Cleared"
            };
        }

        if (normalizedStatus === "cancel" && isOwner) {
            return {
                label: "Cancel",
                name: "cancel",
                nextStatus: "Cancelled"
            };
        }

        return undefined;
    }

    private getStatusTheme(status: string): StatusTheme {
        const normalized = this.toCanonicalStatus(status);

        if (normalized === "toPrepare") {
            return CardOrderAgentPcfLast23.CARD_STATUSES.toPrepare;
        }

        if (normalized === "inPrep") {
            return CardOrderAgentPcfLast23.CARD_STATUSES.inPrep;
        }

        if (normalized === "served") {
            return CardOrderAgentPcfLast23.CARD_STATUSES.served;
        }

        if (normalized === "toClean") {
            return CardOrderAgentPcfLast23.CARD_STATUSES.toClean;
        }

        if (normalized === "cleaned") {
            return CardOrderAgentPcfLast23.CARD_STATUSES.cleaned;
        }

        if (normalized === "cancelled") {
            return CardOrderAgentPcfLast23.CARD_STATUSES.cancelled;
        }

        return CardOrderAgentPcfLast23.CARD_STATUSES.unknown;
    }

    private shouldHideItemCount(status: string): boolean {
        const normalized = this.toCanonicalStatus(status);
        return normalized === "toClean" || normalized === "cleaned";
    }

    private normalize(value: string): string {
        return value
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase()
            .replace(/[_-]+/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    private getLanguage(context: ComponentFramework.Context<IInputs>): LanguageCode {
        const rawValue = (context.parameters.language.raw ?? "EN").trim();
        if (rawValue === "FR") {
            return "fr";
        }

        if (rawValue === "EN") {
            return "en";
        }

        const normalized = this.normalize(rawValue);
        return normalized === "fr" ? "fr" : "en";
    }

    private translate(context: ComponentFramework.Context<IInputs>, key: TranslationKey): string {
        const language = this.getLanguage(context);
        return TRANSLATIONS[language][key];
    }

    private getLocalizedActionLabel(label: string, language: LanguageCode): string {
        if (language === "fr") {
            if (label === "Take") {
                return "Prendre";
            }

            if (label === "View") {
                return "Voir";
            }

            if (label === "Make as served") {
                return "Marquer comme servi";
            }

            if (label === "Mark cleared") {
                return "Marquer comme nettoyé";
            }

            if (label === "Cancel") {
                return "Annuler";
            }
        }

        return label;
    }

    private getLocalizedStatus(status: string, language: LanguageCode): string {
        const canonical = this.toCanonicalStatus(status);
        if (!canonical || canonical === "cancel" || canonical === "unknown") {
            return status;
        }

        return this.getStatusLabel(canonical, language);
    }

    private getStatusLabel(status: Exclude<CanonicalStatus, "cancel" | "unknown">, language: LanguageCode): string {
        const labels = {
            cancelled: { en: "Cancelled", fr: "Annulee" },
            cleaned: { en: "Cleared", fr: "Nettoyee" },
            inPrep: { en: "In prep", fr: "En prep" },
            served: { en: "Served", fr: "Served" },
            toClean: { en: "To clear", fr: "A nettoyer" },
            toPrepare: { en: "To prepare", fr: "A preparer" }
        } as const;

        return labels[status][language];
    }

    private getLanguageFromStatus(status: string): LanguageCode {
        const normalized = this.normalize(status);
        if (normalized.includes("prepare") || normalized.includes("clean") || normalized.includes("served") || normalized.includes("cancel")) {
            return "en";
        }

        if (normalized.includes("annul") || normalized.includes("nettoy") || normalized.includes("prepar") || normalized.includes("servi")) {
            return "fr";
        }

        return "en";
    }

    private toCanonicalStatus(status: string): CanonicalStatus {
        const normalized = this.normalize(status);

        if (normalized === "to prepare" || normalized === "a preparer") {
            return "toPrepare";
        }

        if (normalized === "in prep" || normalized === "en prep" || normalized === "en preparation") {
            return "inPrep";
        }

        if (normalized === "served" || normalized === "servi") {
            return "served";
        }

        if (normalized === "to clean" || normalized === "to clear" || normalized === "a nettoyer") {
            return "toClean";
        }

        if (normalized === "cleaned" || normalized === "cleared" || normalized === "nettoye") {
            return "cleaned";
        }

        if (normalized === "cancelled" || normalized === "canceled" || normalized === "annule") {
            return "cancelled";
        }

        if (normalized === "cancel" || normalized === "annuler") {
            return "cancel";
        }

        return "unknown";
    }
}

interface OrderActionMapping {
    label: string;
    name: string;
    nextStatus: string;
}

interface OrderRecordViewModel {
    actionLabel?: string;
    actionName?: string;
    createdBy: string;
    createdByImageUrl?: string;
    createdOn?: Date;
    createdTime: string;
    datasetRecordId: string;
    id: string;
    modifiedOn?: Date;
    nextStatus?: string;
    note: string;
    orderNumber: string;
    products: ProductLine[];
    takeProducts: ProductLine[];
    quantity: string;
    status: string;
}

interface ProductLine {
    label: string;
    quantity: string;
}

interface StatusTheme {
    background: string;
    foreground: string;
}

interface PersonIdentity {
    displayName?: string;
    email?: string;
    imageUrl?: string;
}

type CanonicalStatus = "cancel" | "cancelled" | "cleaned" | "inPrep" | "served" | "toClean" | "toPrepare" | "unknown";
type LanguageCode = "en" | "fr";
type TranslationKey = "emptyMessage" | "emptyTitle" | "loadErrorMessage" | "loadErrorTitle" | "loadingMessage" | "loadingTitle" | "noteTitle" | "quantityLabel" | "tabAll";

const TRANSLATIONS: Record<LanguageCode, Record<TranslationKey, string>> = {
    en: {
        emptyMessage: "No order is available in the selected data source.",
        emptyTitle: "No orders",
        loadErrorMessage: "The data source could not be loaded.",
        loadErrorTitle: "Loading error",
        loadingMessage: "Orders are being loaded.",
        loadingTitle: "Loading",
        noteTitle: "Note",
        quantityLabel: "Qty",
        tabAll: "All"
    },
    fr: {
        emptyMessage: "Aucune commande disponible dans la source de donnees selectionnee.",
        emptyTitle: "Aucune commande",
        loadErrorMessage: "La source de donnees n'a pas pu etre chargee.",
        loadErrorTitle: "Erreur de chargement",
        loadingMessage: "Les commandes sont en cours de recuperation.",
        loadingTitle: "Chargement",
        noteTitle: "Note",
        quantityLabel: "Qte",
        tabAll: "Tout"
    }
};



