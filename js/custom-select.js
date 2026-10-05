/* ============================================================
   VGRAT FMS — CUSTOM DROPDOWNS
   ============================================================
   Native <select> menus are drawn by the browser / operating
   system, so their highlight colour (usually blue) ignores the
   site's theme and colour palette, and the current option is
   always shown highlighted when the menu opens.

   enhanceSelect() keeps the real <select> (hidden) as the
   source of truth and draws a themed button + list on top:

   - no option is highlighted until you hover it or use the
     arrow keys; the current choice is marked with a ✓
   - hover / keyboard highlight uses the palette colour
   - setting select.value from code, or replacing its options,
     updates the button automatically
   - change events still fire on the <select>, so existing
     filter code keeps working unchanged

   On touch phones the native picker is kept (it is the better
   experience there).
   ============================================================ */

const PHONE_QUERY =
    "(max-width: 699px) and (pointer: coarse), (max-height: 500px) and (max-width: 1000px) and (pointer: coarse)";

const valueSetter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value");
const indexSetter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "selectedIndex");

let openInstance = null;
let uid = 0;


function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
}


function enhanceSelect(select) {
    if (!select || select.dataset.enhanced === "true") return;

    if (window.matchMedia?.(PHONE_QUERY).matches) return;

    select.dataset.enhanced = "true";

    const id = `vsel-${++uid}`;

    const wrap = document.createElement("div");
    wrap.className = "vselect";

    const button = document.createElement("button");
    button.type = "button";
    button.className = "vselect-button";
    button.setAttribute("aria-haspopup", "listbox");
    button.setAttribute("aria-expanded", "false");
    button.setAttribute("aria-controls", `${id}-list`);

    const label = select.closest("label")?.querySelector("span")?.textContent.trim();

    if (label) button.setAttribute("aria-label", label);

    const list = document.createElement("ul");
    list.className = "vselect-list";
    list.id = `${id}-list`;
    list.setAttribute("role", "listbox");
    list.hidden = true;

    select.parentNode.insertBefore(wrap, select);
    wrap.append(select, button, list);

    select.classList.add("vselect-native");
    select.tabIndex = -1;
    select.setAttribute("aria-hidden", "true");

    const instance = { select, wrap, button, list, active: -1 };


    /* ---------- render ---------- */

    function syncButton() {
        const option = select.options[select.selectedIndex];

        button.innerHTML = `<span class="vselect-value">${escapeHtml(option?.textContent ?? "")}</span><span class="vselect-caret" aria-hidden="true">▾</span>`;

        // keep the "filter in use" styling the page adds to the select
        wrap.classList.toggle("is-filtered", select.classList.contains("is-filtered"));
    }

    function renderList() {
        list.innerHTML = [...select.options]
            .map((option, index) => `
                <li
                    class="vselect-option${option.selected ? " is-selected" : ""}"
                    role="option"
                    aria-selected="${option.selected}"
                    data-index="${index}"
                >
                    <span>${escapeHtml(option.textContent)}</span>
                    <span class="vselect-check" aria-hidden="true">✓</span>
                </li>
            `)
            .join("");

        instance.active = -1;   // nothing highlighted until hover / keys
    }

    function setActive(index) {
        const items = list.querySelectorAll(".vselect-option");

        if (!items.length) return;

        instance.active = Math.max(0, Math.min(items.length - 1, index));

        items.forEach((item, i) => item.classList.toggle("is-active", i === instance.active));
        items[instance.active].scrollIntoView({ block: "nearest" });
    }

    function open() {
        if (openInstance && openInstance !== instance) openInstance.close();

        renderList();

        list.hidden = false;
        wrap.classList.add("is-open");
        button.setAttribute("aria-expanded", "true");
        openInstance = instance;

        // Show the current choice without highlighting it
        list.querySelector(".is-selected")?.scrollIntoView({ block: "nearest" });
    }

    function close() {
        list.hidden = true;
        wrap.classList.remove("is-open");
        button.setAttribute("aria-expanded", "false");

        if (openInstance === instance) openInstance = null;
    }

    function choose(index) {
        if (index < 0 || index >= select.options.length) return;

        const changed = select.selectedIndex !== index;

        indexSetter.set.call(select, index);
        syncButton();
        close();
        button.focus();

        if (changed) {
            select.dispatchEvent(new Event("input", { bubbles: true }));
            select.dispatchEvent(new Event("change", { bubbles: true }));
        }
    }

    instance.close = close;


    /* ---------- events ---------- */

    button.addEventListener("click", () => (list.hidden ? open() : close()));

    button.addEventListener("keydown", event => {
        const isOpen = !list.hidden;

        if (event.key === "ArrowDown" || event.key === "ArrowUp") {
            event.preventDefault();

            if (!isOpen) {
                open();
                setActive(select.selectedIndex);
                return;
            }

            const step = event.key === "ArrowDown" ? 1 : -1;

            setActive(instance.active < 0 ? select.selectedIndex + (step > 0 ? 0 : 0) : instance.active + step);
        } else if (event.key === "Enter" || event.key === " ") {
            if (isOpen && instance.active >= 0) {
                event.preventDefault();
                choose(instance.active);
            } else if (!isOpen) {
                event.preventDefault();
                open();
            }
        } else if (event.key === "Escape" || event.key === "Tab") {
            close();
        } else if (event.key === "Home" && isOpen) {
            event.preventDefault();
            setActive(0);
        } else if (event.key === "End" && isOpen) {
            event.preventDefault();
            setActive(select.options.length - 1);
        }
    });

    // mousedown keeps focus on the button
    list.addEventListener("mousedown", event => event.preventDefault());

    list.addEventListener("click", event => {
        // Stop a wrapping <label> from also toggling its checkbox
        event.preventDefault();

        const item = event.target.closest(".vselect-option");

        if (item) choose(Number(item.dataset.index));
    });

    // hovering clears any keyboard highlight; CSS handles the hover colour
    list.addEventListener("mousemove", () => {
        if (instance.active < 0) return;

        list.querySelectorAll(".is-active").forEach(item => item.classList.remove("is-active"));
        instance.active = -1;
    });


    /* ---------- stay in sync with the real <select> ---------- */

    select.addEventListener("change", syncButton);

    // value / selectedIndex set from code (e.g. "Clear filters")
    Object.defineProperty(select, "value", {
        configurable: true,
        get() {
            return valueSetter.get.call(this);
        },
        set(value) {
            valueSetter.set.call(this, value);
            syncButton();
        }
    });

    Object.defineProperty(select, "selectedIndex", {
        configurable: true,
        get() {
            return indexSetter.get.call(this);
        },
        set(value) {
            indexSetter.set.call(this, value);
            syncButton();
        }
    });

    // options rebuilt (filters filled once data loads), or the
    // page toggling the "is-filtered" class
    new MutationObserver(() => {
        syncButton();

        if (!list.hidden) renderList();
    }).observe(select, { childList: true, subtree: true, attributes: true, attributeFilter: ["class"] });

    syncButton();
}


/* Close the open menu when clicking elsewhere */
document.addEventListener("mousedown", event => {
    if (openInstance && !openInstance.wrap.contains(event.target)) {
        openInstance.close();
    }
});

window.addEventListener("resize", () => openInstance?.close());


function enhanceSelects(selector, root = document) {
    root.querySelectorAll(selector).forEach(enhanceSelect);
}


export {
    enhanceSelect,
    enhanceSelects
};
