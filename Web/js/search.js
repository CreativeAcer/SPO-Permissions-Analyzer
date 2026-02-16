// ============================================
// search.js - Global search functionality
// ============================================

let globalSearchData = {
    sites: [],
    users: [],
    groups: [],
    permissions: [],
    inheritance: [],
    loaded: false
};

let selectedResultIndex = -1;
let currentSearchResults = [];

// Initialize global search
function initGlobalSearch() {
    const searchInput = document.getElementById('global-search-input');
    const searchResults = document.getElementById('global-search-results');

    if (!searchInput || !searchResults) return;

    // Keyboard shortcut (Ctrl+K or Cmd+K)
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
            e.preventDefault();
            searchInput.focus();
        }
    });

    // Search input handler with debounce
    let debounceTimer;
    searchInput.addEventListener('input', (e) => {
        clearTimeout(debounceTimer);
        const query = e.target.value.trim();

        if (query.length < 2) {
            searchResults.classList.add('hidden');
            return;
        }

        if (!appState.dataLoaded) {
            searchResults.innerHTML = '<div class="search-no-results">Please connect to SharePoint or start Demo Mode first</div>';
            searchResults.classList.remove('hidden');
            return;
        }

        debounceTimer = setTimeout(() => {
            performGlobalSearch(query);
        }, 300);
    });

    // Keyboard navigation in search results
    searchInput.addEventListener('keydown', (e) => {
        if (!searchResults.classList.contains('hidden')) {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                selectedResultIndex = Math.min(selectedResultIndex + 1, currentSearchResults.length - 1);
                highlightSelectedResult();
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                selectedResultIndex = Math.max(selectedResultIndex - 1, -1);
                highlightSelectedResult();
            } else if (e.key === 'Enter' && selectedResultIndex >= 0) {
                e.preventDefault();
                const result = currentSearchResults[selectedResultIndex];
                if (result) {
                    navigateToSearchResult(result.type, result.item);
                }
            } else if (e.key === 'Escape') {
                searchResults.classList.add('hidden');
                searchInput.blur();
            }
        }
    });

    // Focus handler - lazy load data
    searchInput.addEventListener('focus', async () => {
        if (!globalSearchData.loaded && appState.dataLoaded) {
            await loadGlobalSearchData();
        }
    });

    // Click outside to close
    document.addEventListener('click', (e) => {
        if (!searchInput.contains(e.target) && !searchResults.contains(e.target)) {
            searchResults.classList.add('hidden');
        }
    });
}

// Load all data for searching (lazy loading)
async function loadGlobalSearchData() {
    try {
        const [sites, users, groups, permissions, inheritance] = await Promise.all([
            API.getData('sites'),
            API.getData('users'),
            API.getData('groups'),
            API.getData('roleassignments'),
            API.getData('inheritance')
        ]);

        globalSearchData = {
            sites: sites.data || [],
            users: users.data || [],
            groups: groups.data || [],
            permissions: permissions.data || [],
            inheritance: inheritance.data || [],
            loaded: true
        };
    } catch (e) {
        console.error('Failed to load search data:', e);
    }
}

// Perform search across all data types
function performGlobalSearch(query) {
    const q = query.toLowerCase();
    const results = {
        sites: [],
        users: [],
        groups: [],
        permissions: [],
        inheritance: []
    };

    // Search sites
    results.sites = globalSearchData.sites.filter(s =>
        (s.Title || '').toLowerCase().includes(q) ||
        (s.Url || '').toLowerCase().includes(q) ||
        (s.Owner || '').toLowerCase().includes(q)
    ).slice(0, 5);

    // Search users
    results.users = globalSearchData.users.filter(u =>
        (u.Name || '').toLowerCase().includes(q) ||
        (u.Email || '').toLowerCase().includes(q)
    ).slice(0, 5);

    // Search groups
    results.groups = globalSearchData.groups.filter(g =>
        (g.Name || '').toLowerCase().includes(q) ||
        (g.Description || '').toLowerCase().includes(q)
    ).slice(0, 5);

    // Search permissions
    results.permissions = globalSearchData.permissions.filter(p =>
        (p.Principal || '').toLowerCase().includes(q) ||
        (p.Role || '').toLowerCase().includes(q)
    ).slice(0, 5);

    // Search inheritance
    results.inheritance = globalSearchData.inheritance.filter(i =>
        (i.Title || '').toLowerCase().includes(q) ||
        (i.SiteTitle || '').toLowerCase().includes(q)
    ).slice(0, 5);

    renderSearchResults(results, query);
}

// Render search results dropdown
function renderSearchResults(results, query) {
    const resultsContainer = document.getElementById('global-search-results');
    const totalResults = results.sites.length + results.users.length + results.groups.length +
                        results.permissions.length + results.inheritance.length;

    if (totalResults === 0) {
        resultsContainer.innerHTML = '<div class="search-no-results">No results found</div>';
        resultsContainer.classList.remove('hidden');
        currentSearchResults = [];
        selectedResultIndex = -1;
        return;
    }

    currentSearchResults = [];
    let html = '';

    // Helper to add result group
    const addGroup = (title, items, type, icon) => {
        if (items.length > 0) {
            html += `<div class="search-result-group">
                <div class="search-result-group-title">${icon} ${title} (${items.length})</div>`;

            items.forEach((item, index) => {
                const resultIndex = currentSearchResults.length;
                currentSearchResults.push({ type, item });

                let primaryText = '';
                let secondaryText = '';

                if (type === 'sites') {
                    primaryText = esc(item.Title);
                    secondaryText = esc(item.Url);
                } else if (type === 'users') {
                    primaryText = esc(item.Name);
                    secondaryText = esc(item.Email);
                } else if (type === 'groups') {
                    primaryText = esc(item.Name);
                    secondaryText = `${item.MemberCount || 0} members`;
                } else if (type === 'permissions') {
                    primaryText = esc(item.Principal);
                    secondaryText = `${item.Role} on ${item.Scope}`;
                } else if (type === 'inheritance') {
                    primaryText = esc(item.Title);
                    secondaryText = esc(item.SiteTitle);
                }

                html += `<div class="search-result-item" data-result-index="${resultIndex}" onclick="window.navigateToSearchResult('${type}', ${resultIndex})">
                    <div class="search-result-primary">${primaryText}</div>
                    <div class="search-result-secondary">${secondaryText}</div>
                </div>`;
            });

            html += '</div>';
        }
    };

    addGroup('Sites', results.sites, 'sites', '🌐');
    addGroup('Users', results.users, 'users', '👤');
    addGroup('Groups', results.groups, 'groups', '👥');
    addGroup('Permissions', results.permissions, 'permissions', '🔐');
    addGroup('Inheritance', results.inheritance, 'inheritance', '🔗');

    resultsContainer.innerHTML = html;
    resultsContainer.classList.remove('hidden');
    selectedResultIndex = -1;
}

// Highlight selected result
function highlightSelectedResult() {
    const items = document.querySelectorAll('.search-result-item');
    items.forEach((item, index) => {
        if (index === selectedResultIndex) {
            item.classList.add('selected');
            item.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        } else {
            item.classList.remove('selected');
        }
    });
}

// Navigate to selected search result
window.navigateToSearchResult = function(type, indexOrItem) {
    let item;

    if (typeof indexOrItem === 'number') {
        item = currentSearchResults[indexOrItem].item;
    } else {
        item = indexOrItem;
    }

    // Hide search results
    document.getElementById('global-search-results').classList.add('hidden');
    document.getElementById('global-search-input').value = '';

    // Switch to analytics tab
    const analyticsTab = document.querySelector('.tab-btn[data-tab="analytics"]');
    if (analyticsTab) {
        analyticsTab.click();
    }

    // Open appropriate deep dive with item pre-selected
    setTimeout(() => {
        if (type === 'sites') {
            openSiteDetailDeepDive(item.Title);
        } else if (type === 'users') {
            openDeepDive('users');
            setTimeout(() => {
                const searchInput = document.getElementById('dd-search');
                if (searchInput) {
                    searchInput.value = item.Name || item.Email;
                    searchInput.dispatchEvent(new Event('input'));
                }
            }, 100);
        } else if (type === 'groups') {
            openDeepDive('groups');
            setTimeout(() => {
                const searchInput = document.getElementById('dd-search');
                if (searchInput) {
                    searchInput.value = item.Name;
                    searchInput.dispatchEvent(new Event('input'));
                }
            }, 100);
        } else if (type === 'permissions') {
            openFilteredPermissionsDeepDive(item.Role);
        } else if (type === 'inheritance') {
            openDeepDive('inheritance');
            setTimeout(() => {
                const searchInput = document.getElementById('dd-search');
                if (searchInput) {
                    searchInput.value = item.Title;
                    searchInput.dispatchEvent(new Event('input'));
                }
            }, 100);
        }
    }, 200);
};
