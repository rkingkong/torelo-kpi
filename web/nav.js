/**
 * Torelo KPI — Shared Navigation Sidebar
 * 
 * Usage: Add this to any page:
 *   <script src="nav.js" data-active="gantt"></script>
 * 
 * data-active values: dashboard, gantt, cashflow, attendance, stock, po, inventory, search
 * 
 * This script:
 *   1. Injects the sidebar HTML + CSS
 *   2. Shifts the page content right by 260px
 *   3. Adds a mobile hamburger toggle
 *   4. Highlights the active nav item
 */

(function() {
    const SIDEBAR_W = 260;

    // Determine active page from script tag attribute
    const scriptTag = document.currentScript;
    const activePage = scriptTag?.getAttribute('data-active') || '';

    // ─── CSS ───
    const style = document.createElement('style');
    style.textContent = `
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,500;9..40,700&display=swap');

        .tkpi-sidebar {
            position: fixed; top: 0; left: 0;
            width: ${SIDEBAR_W}px; height: 100vh;
            background: #181a20; border-right: 1px solid #2a2d37;
            display: flex; flex-direction: column;
            z-index: 99990; font-family: 'DM Sans', system-ui, sans-serif;
            transition: transform .3s cubic-bezier(.4,0,.2,1);
        }
        .tkpi-sidebar * { box-sizing: border-box; margin: 0; padding: 0; }
        .tkpi-sidebar.collapsed { transform: translateX(-${SIDEBAR_W}px); }

        .tkpi-sb-brand {
            height: 64px; display: flex; align-items: center; gap: 12px;
            padding: 0 20px; border-bottom: 1px solid #2a2d37; flex-shrink: 0;
        }
        .tkpi-sb-brand .logo {
            width: 34px; height: 34px; background: linear-gradient(135deg, #6c5ce7, #a78bfa);
            border-radius: 8px; display: grid; place-items: center;
            font-weight: 700; font-size: 15px; color: #fff; flex-shrink: 0;
        }
        .tkpi-sb-brand-text { font-size: 16px; font-weight: 700; letter-spacing: -.3px;
            background: linear-gradient(135deg, #e4e5e9, #a78bfa);
            -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; }
        .tkpi-sb-brand small { display: block; font-size: 11px; font-weight: 500;
            background: none; -webkit-text-fill-color: #5c5f6b; margin-top: 1px; }

        .tkpi-sb-nav { flex: 1; overflow-y: auto; padding: 12px 0; }
        .tkpi-sb-section { margin-bottom: 8px; }
        .tkpi-sb-label { padding: 8px 20px 4px; font-size: 10px; font-weight: 600;
            text-transform: uppercase; letter-spacing: .8px; color: #5c5f6b; }

        .tkpi-sb-item {
            display: flex; align-items: center; gap: 10px;
            padding: 9px 20px; margin: 1px 8px; border-radius: 8px;
            color: #8b8f9a; text-decoration: none; font-size: 13.5px; font-weight: 500;
            transition: all .15s; cursor: pointer; border: none; background: none;
            width: calc(100% - 16px); text-align: left; font-family: inherit;
        }
        .tkpi-sb-item:hover { background: #1f2128; color: #e4e5e9; }
        .tkpi-sb-item.active { background: rgba(108,92,231,.12); color: #a78bfa; }
        .tkpi-sb-item .icon { font-size: 16px; width: 22px; text-align: center; flex-shrink: 0; }

        .tkpi-sb-badge {
            margin-left: auto; padding: 2px 8px; border-radius: 10px;
            font-size: 10px; font-weight: 600;
        }
        .badge-green { background: rgba(0,214,143,.12); color: #00d68f; }
        .badge-purple { background: rgba(108,92,231,.12); color: #a78bfa; }
        .badge-amber { background: rgba(240,180,41,.12); color: #f0b429; }
        .badge-blue { background: rgba(72,191,227,.12); color: #48bfe3; }
        .badge-red { background: rgba(255,107,107,.12); color: #ff6b6b; }

        .tkpi-sb-footer {
            padding: 16px 20px; border-top: 1px solid #2a2d37;
            font-size: 12px; color: #5c5f6b;
        }
        .tkpi-sb-footer .dot {
            display: inline-block; width: 6px; height: 6px; border-radius: 50%;
            background: #00d68f; margin-right: 6px; vertical-align: middle;
        }
        .tkpi-sb-footer .dot.offline { background: #ff6b6b; }

        /* Hamburger */
        .tkpi-hamburger {
            display: none; position: fixed; top: 14px; left: 14px; z-index: 99991;
            width: 40px; height: 40px; border-radius: 10px;
            background: #181a20; border: 1px solid #2a2d37;
            place-items: center; cursor: pointer; color: #e4e5e9; font-size: 18px;
        }

        /* Overlay for mobile */
        .tkpi-overlay {
            display: none; position: fixed; inset: 0;
            background: rgba(0,0,0,.5); z-index: 99989;
        }
        .tkpi-overlay.show { display: block; }

        /* Push page content over */
        .tkpi-page-shifted {
            margin-left: ${SIDEBAR_W}px !important;
            transition: margin-left .3s cubic-bezier(.4,0,.2,1);
        }

        @media (max-width: 840px) {
            .tkpi-sidebar { transform: translateX(-${SIDEBAR_W}px); }
            .tkpi-sidebar.open { transform: translateX(0); }
            .tkpi-hamburger { display: grid; }
            .tkpi-page-shifted { margin-left: 0 !important; }
        }
    `;
    document.head.appendChild(style);

    // ─── Navigation items ───
    const navItems = [
        { section: 'Principal', items: [
            { id: 'dashboard', icon: '📊', label: 'Dashboard', href: '/' },
        ]},
        { section: 'Herramientas', items: [
            { id: 'gantt', icon: '📋', label: 'Plan de Proyectos', href: 'gantt-viewer.html', badge: 'Gantt', badgeClass: 'badge-purple' },
            { id: 'cashflow', icon: '💰', label: 'Flujo de Pagos', href: 'cashflow-viewer.html' },
            { id: 'attendance', icon: '👷', label: 'Asistencia', href: 'attendance-viewer.html', badge: 'Live', badgeClass: 'badge-green' },
            { id: 'stock', icon: '📦', label: 'Movimientos de Stock', href: 'stock-explorer.html' },
            { id: 'po', icon: '🛒', label: 'Órdenes de Compra', href: 'po-viewer.html', badge: 'Nuevo', badgeClass: 'badge-green' },
            { id: 'inventory', icon: '🏗️', label: 'Inventario', href: 'inventory-viewer.html', badge: 'Nuevo', badgeClass: 'badge-green' },
            { id: 'search', icon: '🔍', label: 'Buscar Documentos', href: 'search.html' },
        ]},
    ];

    // ─── Build sidebar HTML ───
    let navHTML = '';
    navItems.forEach(section => {
        navHTML += `<div class="tkpi-sb-section"><div class="tkpi-sb-label">${section.section}</div>`;
        section.items.forEach(item => {
            const isActive = activePage === item.id ? ' active' : '';
            const badge = item.badge ? `<span class="tkpi-sb-badge ${item.badgeClass || ''}">${item.badge}</span>` : '';
            navHTML += `<a href="${item.href}" class="tkpi-sb-item${isActive}"><span class="icon">${item.icon}</span>${item.label}${badge}</a>`;
        });
        navHTML += '</div>';
    });

    // ─── Create sidebar element ───
    const sidebar = document.createElement('aside');
    sidebar.className = 'tkpi-sidebar';
    sidebar.id = 'tkpiSidebar';
    sidebar.innerHTML = `
        <div class="tkpi-sb-brand">
            <div class="logo">T</div>
            <div>
                <div class="tkpi-sb-brand-text">Torelo KPI</div>
                <small>Centro de Inteligencia</small>
            </div>
        </div>
        <nav class="tkpi-sb-nav">${navHTML}</nav>
        <div class="tkpi-sb-footer">
            <span class="dot" id="tkpiServerDot"></span>
            <span id="tkpiServerStatus">Verificando...</span>
        </div>
    `;

    // ─── Create hamburger ───
    const hamburger = document.createElement('button');
    hamburger.className = 'tkpi-hamburger';
    hamburger.id = 'tkpiHamburger';
    hamburger.innerHTML = '☰';
    hamburger.setAttribute('aria-label', 'Abrir menú');

    // ─── Create overlay ───
    const overlay = document.createElement('div');
    overlay.className = 'tkpi-overlay';
    overlay.id = 'tkpiOverlay';

    // ─── Insert into DOM ───
    document.body.prepend(overlay);
    document.body.prepend(sidebar);
    document.body.prepend(hamburger);

    // ─── Shift page content ───
    // Find the first main content element (.main, main, #app, or first child)
    const mainEl = document.querySelector('.main') || document.querySelector('main') || document.querySelector('#app');
    if (mainEl) {
        mainEl.classList.add('tkpi-page-shifted');
    }

    // ─── Toggle handlers ───
    function toggleSidebar() {
        const sb = document.getElementById('tkpiSidebar');
        const ov = document.getElementById('tkpiOverlay');
        sb.classList.toggle('open');
        ov.classList.toggle('show');
    }

    hamburger.addEventListener('click', toggleSidebar);
    overlay.addEventListener('click', toggleSidebar);

    // ─── Server status check ───
    async function checkServer() {
        const dot = document.getElementById('tkpiServerDot');
        const status = document.getElementById('tkpiServerStatus');
        try {
            const resp = await fetch('/latest/health.json?' + Date.now(), { signal: AbortSignal.timeout(5000) });
            if (resp.ok) {
                const data = await resp.json();
                dot.className = 'dot';
                const ts = data.last_run ? new Date(data.last_run) : null;
                status.textContent = ts
                    ? `Conectado — ${ts.toLocaleTimeString('es-GT', { hour: '2-digit', minute: '2-digit' })}`
                    : 'Conectado';
            } else {
                throw new Error('not ok');
            }
        } catch {
            // Fallback: try legacy status.json
            try {
                const resp2 = await fetch('/latest/status.json?' + Date.now(), { signal: AbortSignal.timeout(5000) });
                if (resp2.ok) {
                    const data = await resp2.json();
                    dot.className = 'dot';
                    status.textContent = `Conectado — ${data.date_folder || 'hoy'}`;
                } else {
                    throw new Error('not ok');
                }
            } catch {
                dot.className = 'dot offline';
                status.textContent = 'Sin conexión';
            }
        }
    }
    checkServer();
    setInterval(checkServer, 60000); // Re-check every minute

})();