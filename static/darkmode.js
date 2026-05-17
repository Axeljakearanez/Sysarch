/* darkmode.js — shared across all pages */
(function () {
    const STORAGE_KEY = 'ccs_dark_mode';

    // Apply IMMEDIATELY to <html> element to prevent flash
    // This runs before <body> exists
    if (localStorage.getItem(STORAGE_KEY) === '1') {
        document.documentElement.classList.add('dark');
    }

    function applyTheme(dark) {
        if (dark) {
            document.documentElement.classList.add('dark');
            if (document.body) document.body.classList.add('dark');
        } else {
            document.documentElement.classList.remove('dark');
            if (document.body) document.body.classList.remove('dark');
        }

        document.querySelectorAll('.dm-toggle').forEach(btn => {
            btn.textContent = dark ? '☀️' : '🌙';
            btn.title = dark ? 'Switch to Light Mode' : 'Switch to Dark Mode';
        });
    }

    function toggle() {
        const isDark = document.body.classList.contains('dark');
        localStorage.setItem(STORAGE_KEY, isDark ? '0' : '1');
        applyTheme(!isDark);
    }

    document.addEventListener('DOMContentLoaded', function () {
        const isDark = localStorage.getItem(STORAGE_KEY) === '1';
        applyTheme(isDark);

        document.querySelectorAll('.dm-toggle').forEach(btn => {
            btn.addEventListener('click', toggle);
        });
    });
})();