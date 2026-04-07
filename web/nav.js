/**
 * Torelo KPI — Shared Navigation Sidebar
 * 
 * Usage: Add this to any page:
 *   <script src="nav.js" data-active="gantt"></script>
 * 
 * data-active values: dashboard, gantt, cashflow, attendance, stock, search
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
        .tkpi-sb-brand small { display: block; font-size: 11px; color: #5c5f6b; font-weight: 400;
            -webkit-text-fill-color: #5c5f6b; }

        .tkpi-sb-nav { flex: 1; overflow-y: auto; padding: 16px 12px; }
        .tkpi-sb-section { margin-bottom: 24px; }
        .tkpi-sb-label {
            font-size: 10px; font-weight: 700; letter-spacing: 1.2px; text-transform: uppercase;
            color: #5c5f6b; padding: 0 8px; margin-bottom: 8px;
        }
        .tkpi-sb-item {
            display: flex; align-items: center; gap: 10px;
            padding: 10px 12px; border-radius: 8px;
            color: #8b8f9a; text-decoration: none;
            font-size: 13.5px; font-weight: 500;
            transition: all .15s; cursor: pointer;
        }
        .tkpi-sb-item:hover { background: #1f2128; color: #e4e5e9; }
        .tkpi-sb-item.active { background: rgba(108,92,231,.15); color: #7c6ff0; }
        .tkpi-sb-item .icon { font-size: 17px; width: 24px; text-align: center; flex-shrink: 0; }
        .tkpi-sb-item .badge {
            margin-left: auto; font-size: 10px; font-weight: 700;
            padding: 2px 7px; border-radius: 6px;
        }
        .tkpi-sb-item .badge-purple { background: rgba(108,92,231,.15); color: #7c6ff0; }
        .tkpi-sb-item .badge-green { background: rgba(0,214,143,.12); color: #00d68f; }

        /* Hamburger for mobile */
        .tkpi-hamburger {
            display: none; position: fixed; top: 12px; left: 12px; z-index: 99995;
            width: 40px; height: 40px; border-radius: 8px;
            background: #181a20; border: 1px solid #2a2d37;
            color: #8b8f9a; font-size: 20px; cursor: pointer;
            place-items: center;
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
            { id: 'stock', icon: '📦', label: 'Movimientos de Stock', href: 'stock-explorer.html', badge: 'Nuevo', badgeClass: 'badge-green' },
            { id: 'search', icon: '🔍', label: 'Buscar Documentos', href: 'search.html' },
        ]},
    ];

    // ─── Build sidebar HTML ───
    let navHTML = '';
    navItems.forEach(section => {
        navHTML += `<div class="tkpi-sb-section"><div class="tkpi-sb-label">${section.section}</div>`;
        section.items.forEach(item => {
            const isActive = activePage === item.id ? ' active' : '';
            const badge = item.badge ? `<span class="badge ${item.badgeClass || ''}">${item.badge}</span>` : '';
            navHTML += `<a href="${item.href}" class="tkpi-sb-item${isActive}">
                <span class="icon">${item.icon}</span>${item.label}${badge}</a>`;
        });
        navHTML += '</div>';
    });

    const sidebar = document.createElement('div');
    sidebar.className = 'tkpi-sidebar';
    sidebar.id = 'tkpiSidebar';
    sidebar.innerHTML = `
        <div class="tkpi-sb-brand">
            <div class="logo">T</div>
            <div><span class="tkpi-sb-brand-text">Torelo KPI</span><small>Centro de Inteligencia</small></div>
        </div>
        <nav class="tkpi-sb-nav">${navHTML}</nav>
    `;

    // Hamburger button
    const hamburger = document.createElement('button');
    hamburger.className = 'tkpi-hamburger';
    hamburger.id = 'tkpiHamburger';
    hamburger.textContent = '☰';
    hamburger.onclick = function() {
        sidebar.classList.toggle('open');
        overlay.classList.toggle('show');
    };

    // Mobile overlay
    const overlay = document.createElement('div');
    overlay.className = 'tkpi-overlay';
    overlay.id = 'tkpiOverlay';
    overlay.onclick = function() {
        sidebar.classList.remove('open');
        overlay.classList.remove('show');
    };

    // ─── Inject into page ───
    document.body.prepend(overlay);
    document.body.prepend(sidebar);
    document.body.prepend(hamburger);

    // Shift existing page content
    // Find the first major container element to shift
    const shiftTargets = document.querySelectorAll('body > div, body > main, body > section');
    shiftTargets.forEach(el => {
        if (el !== sidebar && el !== overlay && el !== hamburger && !el.classList.contains('tkpi-sidebar')) {
            el.classList.add('tkpi-page-shifted');
        }
    });

    // Also shift any fixed-position elements that aren't our sidebar
    // (loading screens, headers, etc.)
    requestAnimationFrame(() => {
        document.querySelectorAll('.loading-container, .error-container').forEach(el => {
            el.style.marginLeft = SIDEBAR_W + 'px';
        });
    });

})();