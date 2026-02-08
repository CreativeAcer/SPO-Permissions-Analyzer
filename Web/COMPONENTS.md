# Modern UI Components Guide

This guide documents all the modern UI components and helpers available in the web interface after Phase 2 modernization.

---

## 🎨 Phase 2 Components

### Loading States

#### 1. **Button Loading Spinner**
Shows an animated spinner inside the button while an action is in progress.

**Usage:**
```javascript
// Show loading
UIHelpers.setButtonLoading('btn-connect', true);

// Hide loading
UIHelpers.setButtonLoading('btn-connect', false);
```

**CSS Classes:**
```html
<!-- Button automatically gets loading state -->
<button id="my-button" class="btn btn-primary">Click Me</button>
```

---

#### 2. **Skeleton Loaders**
Animated placeholder content while data is loading.

**Usage:**
```javascript
// Show skeleton in an element
UIHelpers.showSkeleton('content-div', 'card'); // Types: text, card, metric, table

// Show skeleton for all metric cards
UIHelpers.showMetricSkeletons();
```

**HTML:**
```html
<!-- Manual skeleton -->
<div class="skeleton skeleton-card"></div>
<div class="skeleton skeleton-metric"></div>
<div class="skeleton skeleton-text"></div>
<div class="skeleton skeleton-circle"></div>
```

---

#### 3. **Progress Bars**

**Determinate Progress:**
```javascript
// Update progress (0-100)
UIHelpers.updateProgress('progress-container', 45);
```

```html
<div id="progress-container"></div>
```

**Indeterminate Progress:**
```javascript
// Show when exact progress is unknown
UIHelpers.showIndeterminateProgress('progress-container');
```

```html
<div class="progress-indeterminate"></div>
```

---

#### 4. **Loading Overlay**
Full-screen loading overlay with blur effect.

**Usage:**
```javascript
// Show overlay
UIHelpers.showLoadingOverlay('Processing data...');

// Hide overlay
UIHelpers.hideLoadingOverlay();
```

---

#### 5. **Spinner Components**

```html
<!-- Primary spinner -->
<div class="spinner spinner-primary"></div>

<!-- Large spinner -->
<div class="spinner spinner-lg"></div>

<!-- Small spinner -->
<div class="spinner spinner-sm"></div>
```

---

### Form Enhancements

#### 1. **Input Validation States**

**Usage:**
```javascript
// Set validation state
UIHelpers.setInputValidation('email-input', true, 'Email is valid');
UIHelpers.setInputValidation('email-input', false, 'Invalid email format');

// Clear validation
UIHelpers.setInputValidation('email-input', null);
```

**HTML:**
```html
<div class="form-group">
    <label for="email">Email</label>
    <input type="email" id="email" class="input-valid">
    <div class="form-feedback feedback-success">
        <span>✓</span><span>Email is valid</span>
    </div>
</div>
```

**CSS Classes:**
- `.input-valid` - Green border, success background
- `.input-error` - Red border, error background

---

#### 2. **Floating Labels**
Modern material-design style floating labels.

```html
<div class="form-group-floating">
    <input type="text" id="tenant" placeholder=" " required>
    <label for="tenant">Tenant URL</label>
</div>
```

---

#### 3. **Input with Icons**

```html
<div class="form-group form-group-with-icon">
    <input type="text" id="search" placeholder="Search...">
    <span class="form-icon">🔍</span>
</div>
```

---

### Badges & Tags

#### 1. **Badges**

**Usage:**
```javascript
// Create badge programmatically
const badge = UIHelpers.createBadge('Active', 'success', true);
document.getElementById('container').appendChild(badge);
```

**HTML:**
```html
<!-- Primary badge -->
<span class="badge badge-primary">New</span>

<!-- With dot indicator -->
<span class="badge badge-success badge-dot">Online</span>

<!-- Other variants -->
<span class="badge badge-warning">Warning</span>
<span class="badge badge-error">Error</span>
<span class="badge badge-neutral">Neutral</span>
```

---

#### 2. **Chips/Tags**

**Usage:**
```javascript
// Create removable chip
const chip = UIHelpers.createChip('JavaScript', () => {
    console.log('Chip removed');
});
document.getElementById('container').appendChild(chip);
```

**HTML:**
```html
<span class="chip">
    <span>React</span>
    <span class="chip-remove">×</span>
</span>
```

---

### Empty States

**Usage:**
```javascript
UIHelpers.showEmptyState(
    'table-body',
    'No Sites Found',
    'Try connecting to SharePoint first.',
    '📊'
);
```

**HTML:**
```html
<div class="empty-state">
    <div class="empty-state-icon">📊</div>
    <h3>No Data Available</h3>
    <p>Try running an analysis first.</p>
</div>
```

---

### Cards

#### 1. **Basic Card**
```html
<div class="card">
    <h3>Card Title</h3>
    <p>Card content...</p>
</div>
```

#### 2. **Card with Header & Footer**
```html
<div class="card">
    <div class="card-header">
        <h3>Title</h3>
        <button class="btn btn-secondary">Action</button>
    </div>

    <p>Content...</p>

    <div class="card-footer">
        <span class="text-muted">Last updated: 2 hours ago</span>
        <button class="btn btn-primary">Save</button>
    </div>
</div>
```

#### 3. **Elevated Card**
```html
<div class="card card-elevated">
    <p>This card has more shadow</p>
</div>
```

---

### Stats Card

```html
<div class="stat-card">
    <div class="stat-label">Total Users</div>
    <div class="stat-value">1,234</div>
    <div class="stat-change positive">
        <span>↑</span> 12% from last week
    </div>
</div>
```

---

### Accordion

```html
<div class="accordion-item">
    <div class="accordion-header">
        <span>Click to expand</span>
        <span class="accordion-icon">▼</span>
    </div>
    <div class="accordion-content">
        <div class="accordion-body">
            Hidden content here...
        </div>
    </div>
</div>
```

**JavaScript:**
```javascript
document.querySelectorAll('.accordion-header').forEach(header => {
    header.addEventListener('click', () => {
        const item = header.closest('.accordion-item');
        item.classList.toggle('active');
    });
});
```

---

### Dividers

```html
<!-- Horizontal divider -->
<div class="divider"></div>

<!-- Vertical divider -->
<div class="divider-vertical"></div>

<!-- Divider with text -->
<div class="divider-text">OR</div>
```

---

### Tooltips

```html
<div class="tooltip-wrapper">
    <button class="btn btn-primary">Hover me</button>
    <div class="tooltip">Helpful information</div>
</div>
```

---

## 🛠️ Utility Functions

### Counter Animation
```javascript
// Animate number from current value to target
UIHelpers.animateCounter('metric-sites', 150, 1000); // element, target, duration
```

### Clipboard Copy
```javascript
// Copy text with feedback toast
await UIHelpers.copyToClipboard('https://example.com', 'Link copied!');
```

### Smooth Scroll
```javascript
// Scroll to element
UIHelpers.smoothScrollTo('section-id');
```

### Ripple Effect
```javascript
// Add ripple effect (auto-added to all buttons)
button.addEventListener('click', (e) => {
    UIHelpers.createRipple(button, e);
});
```

### Format Helpers
```javascript
// Format file size
UIHelpers.formatFileSize(1536000); // "1.46 MB"

// Format relative date
UIHelpers.formatRelativeDate(new Date(Date.now() - 3600000)); // "1h ago"
```

---

## 🎯 Best Practices

### Loading States
1. **Always show loading feedback** for async operations
2. **Use skeleton loaders** for initial page loads
3. **Use spinners** for quick actions (< 3 seconds)
4. **Use progress bars** for longer operations with known progress

### Form Validation
1. **Validate on blur** (when user leaves input)
2. **Show success state** only when valid
3. **Show error state** with helpful message
4. **Clear validation** when user starts typing

### Empty States
1. **Always provide context** (why is it empty?)
2. **Suggest next action** (what should user do?)
3. **Use friendly icons/illustrations**

### Accessibility
1. All components have **focus-visible** styles
2. Use **semantic HTML** (button, label, etc.)
3. Provide **aria-labels** where needed
4. Ensure **keyboard navigation** works

---

## 📱 Responsive Behavior

All components are mobile-responsive:

- **Metric cards**: 4 cols → 2 cols → 1 col
- **Charts**: 2 cols → 1 col
- **Tables**: Horizontal scroll on mobile
- **Modals**: 90vw width on mobile
- **Buttons**: Full width on mobile in button-row

---

## 🎨 CSS Variables Reference

Use these variables for consistent styling:

```css
/* Colors */
--primary-500: #0078D4;
--success-500: #10B981;
--warning-500: #F59E0B;
--error-500: #EF4444;
--neutral-500: #71717A;

/* Spacing */
--spacing-xs: 0.25rem;   /* 4px */
--spacing-sm: 0.5rem;    /* 8px */
--spacing-md: 0.75rem;   /* 12px */
--spacing-lg: 1rem;      /* 16px */
--spacing-xl: 1.5rem;    /* 24px */
--spacing-2xl: 2rem;     /* 32px */

/* Typography */
--font-size-xs: 0.6875rem;  /* 11px */
--font-size-sm: 0.8125rem;  /* 13px */
--font-size-base: 0.875rem; /* 14px */
--font-size-lg: 1rem;       /* 16px */
--font-size-xl: 1.125rem;   /* 18px */

/* Shadows */
--shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
--shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
--shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1);

/* Border Radius */
--radius-sm: 0.375rem;  /* 6px */
--radius-md: 0.5rem;    /* 8px */
--radius-lg: 0.75rem;   /* 12px */
--radius-xl: 1rem;      /* 16px */
--radius-full: 9999px;

/* Transitions */
--transition-fast: 150ms cubic-bezier(0.4, 0, 0.2, 1);
--transition-base: 200ms cubic-bezier(0.4, 0, 0.2, 1);
--transition-slow: 300ms cubic-bezier(0.4, 0, 0.2, 1);
```

---

## 🚀 Examples

### Complete Form with Validation
```html
<div class="card">
    <h3>Connect to SharePoint</h3>

    <div class="form-group">
        <label for="tenant-url">Tenant URL</label>
        <input type="url" id="tenant-url" placeholder="https://yourtenant.sharepoint.com">
        <div class="form-feedback feedback-error hidden">
            <span>✕</span><span>Please enter a valid URL</span>
        </div>
    </div>

    <div class="button-row">
        <button id="btn-connect" class="btn btn-primary">
            Connect
        </button>
        <button class="btn btn-secondary">Cancel</button>
    </div>
</div>

<script>
// Validation
document.getElementById('tenant-url').addEventListener('blur', (e) => {
    const isValid = e.target.value.startsWith('https://');
    UIHelpers.setInputValidation('tenant-url', isValid,
        isValid ? 'Valid URL' : 'URL must start with https://');
});

// Connect with loading
document.getElementById('btn-connect').addEventListener('click', async () => {
    UIHelpers.setButtonLoading('btn-connect', true);
    try {
        await API.connect();
        toast('Connected successfully!', 'success');
    } finally {
        UIHelpers.setButtonLoading('btn-connect', false);
    }
});
</script>
```

---

## 🎓 Migration Guide

### Old → New

**Button Loading:**
```javascript
// OLD
btn.textContent = 'Loading...';
btn.disabled = true;

// NEW
UIHelpers.setButtonLoading('btn-id', true);
```

**Empty Tables:**
```javascript
// OLD
tbody.innerHTML = '<tr><td colspan="5">No data</td></tr>';

// NEW
UIHelpers.showEmptyState('tbody', 'No Sites Found', 'Run analysis first', '📊');
```

**Counter Updates:**
```javascript
// OLD
element.textContent = newValue;

// NEW
UIHelpers.animateCounter('element-id', newValue, 800);
```

---

## 📊 Performance Tips

1. **Debounce input validation** for search/filter inputs
2. **Use skeleton loaders** instead of spinners for better perceived performance
3. **Animate counters** for metric updates to draw attention
4. **Lazy load modals** - only render content when opened
5. **Virtualize long tables** (consider for 1000+ rows)

---

Made with ❤️ for modern web experiences
