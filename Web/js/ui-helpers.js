// ============================================
// ui-helpers.js - UI Component Helpers
// ============================================
// Helper functions for loading states, skeletons, and UI enhancements

const UIHelpers = {
    /**
     * Show loading spinner on a button
     * @param {string} buttonId - The button element ID
     */
    setButtonLoading(buttonId, loading = true) {
        const btn = document.getElementById(buttonId);
        if (!btn) return;

        if (loading) {
            btn.classList.add('btn-loading');
            btn.disabled = true;
        } else {
            btn.classList.remove('btn-loading');
            btn.disabled = false;
        }
    },

    /**
     * Show skeleton loader in an element
     * @param {string} elementId - The element ID to show skeleton in
     * @param {string} type - Type of skeleton (text, card, metric, table)
     */
    showSkeleton(elementId, type = 'card') {
        const el = document.getElementById(elementId);
        if (!el) return;

        const skeletonMap = {
            text: `<div class="skeleton skeleton-text"></div>`.repeat(3),
            card: `<div class="skeleton skeleton-card"></div>`,
            metric: `<div class="skeleton skeleton-metric"></div>`,
            table: `<div class="skeleton skeleton-table-row"></div>`.repeat(5)
        };

        el.innerHTML = skeletonMap[type] || skeletonMap.card;
    },

    /**
     * Show skeleton for metric cards
     */
    showMetricSkeletons() {
        const metrics = ['metric-sites', 'metric-users', 'metric-groups', 'metric-external',
                        'metric-roles', 'metric-inheritance', 'metric-sharing'];

        metrics.forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                el.innerHTML = '<div class="skeleton" style="height: 32px; width: 60px; margin: 0 auto;"></div>';
            }
        });
    },

    /**
     * Show loading overlay
     * @param {string} message - Optional loading message
     */
    showLoadingOverlay(message = 'Loading...') {
        let overlay = document.getElementById('loading-overlay');

        if (!overlay) {
            overlay = document.createElement('div');
            overlay.id = 'loading-overlay';
            overlay.className = 'loading-overlay';
            overlay.innerHTML = `
                <div class="loading-content">
                    <div class="spinner spinner-primary spinner-lg"></div>
                    <p>${message}</p>
                </div>
            `;
            document.body.appendChild(overlay);
        } else {
            overlay.classList.remove('hidden');
            const messageEl = overlay.querySelector('p');
            if (messageEl) messageEl.textContent = message;
        }
    },

    /**
     * Hide loading overlay
     */
    hideLoadingOverlay() {
        const overlay = document.getElementById('loading-overlay');
        if (overlay) {
            overlay.classList.add('hidden');
        }
    },

    /**
     * Show progress bar
     * @param {string} containerId - Container element ID
     * @param {number} percent - Progress percentage (0-100)
     */
    updateProgress(containerId, percent) {
        const container = document.getElementById(containerId);
        if (!container) return;

        let progressBar = container.querySelector('.progress-bar');
        if (!progressBar) {
            container.innerHTML = `
                <div class="progress-container">
                    <div class="progress-bar" style="width: 0%"></div>
                </div>
            `;
            progressBar = container.querySelector('.progress-bar');
        }

        progressBar.style.width = `${Math.min(100, Math.max(0, percent))}%`;
    },

    /**
     * Show indeterminate progress (when exact progress unknown)
     * @param {string} containerId - Container element ID
     */
    showIndeterminateProgress(containerId) {
        const container = document.getElementById(containerId);
        if (!container) return;

        container.innerHTML = '<div class="progress-indeterminate"></div>';
    },

    /**
     * Add validation state to input
     * @param {string} inputId - Input element ID
     * @param {boolean} isValid - Whether input is valid
     * @param {string} message - Validation message
     */
    setInputValidation(inputId, isValid, message = '') {
        const input = document.getElementById(inputId);
        if (!input) return;

        // Remove existing validation classes
        input.classList.remove('input-valid', 'input-error');

        // Remove existing feedback
        const existingFeedback = input.parentElement.querySelector('.form-feedback');
        if (existingFeedback) existingFeedback.remove();

        if (isValid === null) return; // No validation state

        // Add new validation class
        input.classList.add(isValid ? 'input-valid' : 'input-error');

        // Add feedback message if provided
        if (message) {
            const feedback = document.createElement('div');
            feedback.className = `form-feedback ${isValid ? 'feedback-success' : 'feedback-error'}`;
            feedback.innerHTML = `<span>${isValid ? '✓' : '✕'}</span><span>${message}</span>`;
            input.parentElement.appendChild(feedback);
        }
    },

    /**
     * Show empty state in container
     * @param {string} containerId - Container element ID
     * @param {string} title - Empty state title
     * @param {string} message - Empty state message
     * @param {string} icon - Optional icon (emoji or text)
     */
    showEmptyState(containerId, title = 'No Data Available', message = 'Try running an analysis first.', icon = '📊') {
        const container = document.getElementById(containerId);
        if (!container) return;

        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">${icon}</div>
                <h3>${title}</h3>
                <p>${message}</p>
            </div>
        `;
    },

    /**
     * Create a badge element
     * @param {string} text - Badge text
     * @param {string} variant - Badge variant (primary, success, warning, error, neutral)
     * @param {boolean} withDot - Show dot indicator
     */
    createBadge(text, variant = 'primary', withDot = false) {
        const badge = document.createElement('span');
        badge.className = `badge badge-${variant}${withDot ? ' badge-dot' : ''}`;
        badge.textContent = text;
        return badge;
    },

    /**
     * Create a chip/tag element
     * @param {string} text - Chip text
     * @param {Function} onRemove - Callback when remove is clicked
     */
    createChip(text, onRemove = null) {
        const chip = document.createElement('span');
        chip.className = 'chip';
        chip.innerHTML = `<span>${text}</span>`;

        if (onRemove) {
            const removeBtn = document.createElement('span');
            removeBtn.className = 'chip-remove';
            removeBtn.innerHTML = '×';
            removeBtn.onclick = (e) => {
                e.stopPropagation();
                onRemove();
                chip.remove();
            };
            chip.appendChild(removeBtn);
        }

        return chip;
    },

    /**
     * Animate number counting up
     * @param {string} elementId - Element ID to animate
     * @param {number} target - Target number
     * @param {number} duration - Animation duration in ms
     */
    animateCounter(elementId, target, duration = 1000) {
        const el = document.getElementById(elementId);
        if (!el) return;

        const start = parseInt(el.textContent) || 0;
        const increment = (target - start) / (duration / 16);
        let current = start;

        const timer = setInterval(() => {
            current += increment;
            if ((increment > 0 && current >= target) || (increment < 0 && current <= target)) {
                el.textContent = target;
                clearInterval(timer);
            } else {
                el.textContent = Math.round(current);
            }
        }, 16);
    },

    /**
     * Add ripple effect to element on click
     * @param {HTMLElement} element - Element to add ripple to
     * @param {MouseEvent} event - Click event
     */
    createRipple(element, event) {
        const ripple = document.createElement('span');
        const rect = element.getBoundingClientRect();
        const size = Math.max(rect.width, rect.height);
        const x = event.clientX - rect.left - size / 2;
        const y = event.clientY - rect.top - size / 2;

        ripple.style.cssText = `
            position: absolute;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.6);
            width: ${size}px;
            height: ${size}px;
            left: ${x}px;
            top: ${y}px;
            pointer-events: none;
            animation: ripple-animation 0.6s ease-out;
        `;

        element.style.position = 'relative';
        element.style.overflow = 'hidden';
        element.appendChild(ripple);

        setTimeout(() => ripple.remove(), 600);
    },

    /**
     * Smooth scroll to element
     * @param {string} elementId - Element ID to scroll to
     */
    smoothScrollTo(elementId) {
        const element = document.getElementById(elementId);
        if (!element) return;

        element.scrollIntoView({
            behavior: 'smooth',
            block: 'start'
        });
    },

    /**
     * Copy text to clipboard and show feedback
     * @param {string} text - Text to copy
     * @param {string} successMessage - Success toast message
     */
    async copyToClipboard(text, successMessage = 'Copied to clipboard') {
        try {
            await navigator.clipboard.writeText(text);
            if (typeof toast === 'function') {
                toast(successMessage, 'success');
            }
            return true;
        } catch (err) {
            console.error('Failed to copy:', err);
            if (typeof toast === 'function') {
                toast('Failed to copy to clipboard', 'error');
            }
            return false;
        }
    },

    /**
     * Format file size
     * @param {number} bytes - Size in bytes
     */
    formatFileSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
    },

    /**
     * Format date relative to now
     * @param {Date|string} date - Date to format
     */
    formatRelativeDate(date) {
        const now = new Date();
        const then = new Date(date);
        const diffMs = now - then;
        const diffMins = Math.floor(diffMs / 60000);
        const diffHours = Math.floor(diffMs / 3600000);
        const diffDays = Math.floor(diffMs / 86400000);

        if (diffMins < 1) return 'just now';
        if (diffMins < 60) return `${diffMins}m ago`;
        if (diffHours < 24) return `${diffHours}h ago`;
        if (diffDays < 7) return `${diffDays}d ago`;
        return then.toLocaleDateString();
    },

    // ============================================================
    // PHASE 3: ADVANCED INTERACTIONS
    // ============================================================

    /**
     * Show enhanced toast notification with icon and title
     * @param {string} title - Toast title
     * @param {string} message - Toast message
     * @param {string} type - Toast type (success, error, warning, info)
     * @param {number} duration - Duration in ms (default 4000)
     */
    showEnhancedToast(title, message, type = 'info', duration = 4000) {
        const container = document.getElementById('toast-container');
        if (!container) return;

        const icons = {
            success: '✓',
            error: '✕',
            warning: '⚠',
            info: 'ℹ'
        };

        const toast = document.createElement('div');
        toast.className = `toast-enhanced ${type}`;
        toast.innerHTML = `
            <div class="toast-icon">${icons[type] || icons.info}</div>
            <div class="toast-content">
                <div class="toast-title">${title}</div>
                ${message ? `<div class="toast-message">${message}</div>` : ''}
            </div>
            <button class="toast-close">×</button>
        `;

        const closeBtn = toast.querySelector('.toast-close');
        closeBtn.onclick = () => toast.remove();

        container.appendChild(toast);

        if (duration > 0) {
            setTimeout(() => toast.remove(), duration);
        }

        return toast;
    },

    /**
     * Make table sortable
     * @param {string} tableId - Table element ID
     */
    makeSortable(tableId) {
        const table = document.getElementById(tableId);
        if (!table) return;

        const headers = table.querySelectorAll('th.sortable');

        headers.forEach((header, index) => {
            header.addEventListener('click', () => {
                const tbody = table.querySelector('tbody');
                const rows = Array.from(tbody.querySelectorAll('tr'));
                const isAscending = header.classList.contains('sort-asc');

                // Remove sort classes from all headers
                headers.forEach(h => h.classList.remove('sort-asc', 'sort-desc'));

                // Add appropriate sort class
                header.classList.add(isAscending ? 'sort-desc' : 'sort-asc');

                // Sort rows
                rows.sort((a, b) => {
                    const aVal = a.cells[index].textContent.trim();
                    const bVal = b.cells[index].textContent.trim();

                    // Try numeric sort first
                    const aNum = parseFloat(aVal.replace(/[^0-9.-]/g, ''));
                    const bNum = parseFloat(bVal.replace(/[^0-9.-]/g, ''));

                    if (!isNaN(aNum) && !isNaN(bNum)) {
                        return isAscending ? bNum - aNum : aNum - bNum;
                    }

                    // Fallback to string sort
                    return isAscending
                        ? bVal.localeCompare(aVal)
                        : aVal.localeCompare(bVal);
                });

                // Reattach rows
                rows.forEach(row => tbody.appendChild(row));
            });
        });
    },

    /**
     * Add scroll shadow indicators to container
     * @param {string} containerId - Container element ID
     */
    addScrollShadows(containerId) {
        const container = document.getElementById(containerId);
        if (!container) return;

        container.classList.add('scroll-shadow-container');

        const updateShadows = () => {
            const isScrolledFromTop = container.scrollTop > 10;
            const isScrolledToBottom =
                container.scrollHeight - container.scrollTop - container.clientHeight < 10;

            container.classList.toggle('show-top-shadow', isScrolledFromTop);
            container.classList.toggle('show-bottom-shadow', !isScrolledToBottom);
        };

        container.addEventListener('scroll', updateShadows);
        updateShadows();
    },

    /**
     * Add stagger animation to list items
     * @param {string} containerId - Container element ID
     */
    staggerItems(containerId) {
        const container = document.getElementById(containerId);
        if (!container) return;

        const items = container.children;
        Array.from(items).forEach((item, index) => {
            item.classList.add('stagger-item');
            item.style.animationDelay = `${index * 0.05}s`;
        });
    },

    /**
     * Highlight search term in text
     * @param {string} text - Text to search in
     * @param {string} searchTerm - Term to highlight
     * @returns {string} HTML with highlighted terms
     */
    highlightSearchTerm(text, searchTerm) {
        if (!searchTerm) return text;

        const regex = new RegExp(`(${searchTerm})`, 'gi');
        return text.replace(regex, '<span class="search-highlight">$1</span>');
    },

    /**
     * Filter table rows by search term
     * @param {string} tableId - Table element ID
     * @param {string} searchTerm - Search term
     * @param {number[]} columnIndexes - Columns to search (default: all)
     */
    filterTable(tableId, searchTerm, columnIndexes = null) {
        const table = document.getElementById(tableId);
        if (!table) return;

        const tbody = table.querySelector('tbody');
        const rows = tbody.querySelectorAll('tr');
        const term = searchTerm.toLowerCase();

        let visibleCount = 0;

        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            const cellsToSearch = columnIndexes
                ? columnIndexes.map(i => cells[i])
                : Array.from(cells);

            const matches = cellsToSearch.some(cell =>
                cell && cell.textContent.toLowerCase().includes(term)
            );

            if (matches) {
                row.style.display = '';
                visibleCount++;
            } else {
                row.style.display = 'none';
            }
        });

        return visibleCount;
    },

    /**
     * Create breadcrumb navigation
     * @param {Array} items - Array of {label, url} objects
     * @returns {HTMLElement} Breadcrumb element
     */
    createBreadcrumbs(items) {
        const breadcrumbs = document.createElement('nav');
        breadcrumbs.className = 'breadcrumbs';

        items.forEach((item, index) => {
            const breadcrumbItem = document.createElement('div');
            breadcrumbItem.className = 'breadcrumb-item';

            const link = document.createElement('a');
            link.href = item.url || '#';
            link.className = 'breadcrumb-link';
            link.textContent = item.label;

            if (!item.url) {
                link.onclick = (e) => e.preventDefault();
            }

            breadcrumbItem.appendChild(link);

            if (index < items.length - 1) {
                const separator = document.createElement('span');
                separator.className = 'breadcrumb-separator';
                separator.textContent = '/';
                breadcrumbItem.appendChild(separator);
            }

            breadcrumbs.appendChild(breadcrumbItem);
        });

        return breadcrumbs;
    },

    /**
     * Add shake animation to element (for errors)
     * @param {string} elementId - Element ID
     */
    shake(elementId) {
        const el = document.getElementById(elementId);
        if (!el) return;

        el.classList.add('animate-shake');
        setTimeout(() => el.classList.remove('animate-shake'), 500);
    },

    /**
     * Debounce function
     * @param {Function} func - Function to debounce
     * @param {number} wait - Wait time in ms
     * @returns {Function} Debounced function
     */
    debounce(func, wait = 300) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    },

    /**
     * Throttle function
     * @param {Function} func - Function to throttle
     * @param {number} limit - Limit time in ms
     * @returns {Function} Throttled function
     */
    throttle(func, limit = 300) {
        let inThrottle;
        return function(...args) {
            if (!inThrottle) {
                func.apply(this, args);
                inThrottle = true;
                setTimeout(() => inThrottle = false, limit);
            }
        };
    },

    /**
     * Create floating action button
     * @param {string} icon - Button icon (text or emoji)
     * @param {Function} onClick - Click handler
     * @returns {HTMLElement} FAB element
     */
    createFAB(icon, onClick) {
        const fab = document.createElement('button');
        fab.className = 'fab';
        fab.innerHTML = icon;
        fab.onclick = onClick;
        document.body.appendChild(fab);
        return fab;
    },

    /**
     * Show context menu
     * @param {number} x - X position
     * @param {number} y - Y position
     * @param {Array} items - Menu items [{label, onClick, divider}]
     */
    showContextMenu(x, y, items) {
        // Remove existing context menu
        const existing = document.querySelector('.context-menu');
        if (existing) existing.remove();

        const menu = document.createElement('div');
        menu.className = 'context-menu';
        menu.style.left = `${x}px`;
        menu.style.top = `${y}px`;

        items.forEach(item => {
            if (item.divider) {
                const divider = document.createElement('div');
                divider.className = 'context-menu-divider';
                menu.appendChild(divider);
            } else {
                const menuItem = document.createElement('div');
                menuItem.className = 'context-menu-item';
                menuItem.textContent = item.label;
                menuItem.onclick = () => {
                    item.onClick();
                    menu.remove();
                };
                menu.appendChild(menuItem);
            }
        });

        document.body.appendChild(menu);

        // Close on click outside
        setTimeout(() => {
            document.addEventListener('click', function closeMenu() {
                menu.remove();
                document.removeEventListener('click', closeMenu);
            });
        }, 10);

        return menu;
    },

    /**
     * Add notification badge to element
     * @param {string} elementId - Element ID
     * @param {number} count - Notification count
     */
    addNotificationBadge(elementId, count) {
        const el = document.getElementById(elementId);
        if (!el) return;

        // Remove existing badge
        const existing = el.querySelector('.notification-badge');
        if (existing) existing.remove();

        if (count > 0) {
            const badge = document.createElement('span');
            badge.className = 'notification-badge';
            badge.textContent = count > 99 ? '99+' : count;
            el.style.position = 'relative';
            el.appendChild(badge);
        }
    },

    /**
     * Get element position relative to viewport
     * @param {HTMLElement} element - Element to get position of
     * @returns {Object} {top, left, bottom, right}
     */
    getElementPosition(element) {
        const rect = element.getBoundingClientRect();
        return {
            top: rect.top + window.scrollY,
            left: rect.left + window.scrollX,
            bottom: rect.bottom + window.scrollY,
            right: rect.right + window.scrollX,
            width: rect.width,
            height: rect.height
        };
    }
};

// Add ripple animation CSS if not exists
if (!document.getElementById('ripple-animation-style')) {
    const style = document.createElement('style');
    style.id = 'ripple-animation-style';
    style.textContent = `
        @keyframes ripple-animation {
            from {
                transform: scale(0);
                opacity: 1;
            }
            to {
                transform: scale(1);
                opacity: 0;
            }
        }
    `;
    document.head.appendChild(style);
}

// Auto-add ripple effect to all buttons
document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('.btn').forEach(btn => {
        btn.addEventListener('click', function(e) {
            if (!this.classList.contains('btn-disabled') && !this.classList.contains('btn-loading')) {
                UIHelpers.createRipple(this, e);
            }
        });
    });
});
