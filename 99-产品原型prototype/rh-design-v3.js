/* RH Design Token v3.0 interaction layer. Business handlers remain page-owned. */
(() => {
  const tabSelectors = '.tab, .trace-tab, .seg button, .segmented button';
  const focusableSelector = [
    'input:not([disabled]):not([readonly]):not([type="hidden"]):not([type="checkbox"]):not([type="radio"])',
    'select:not([disabled])',
    'textarea:not([disabled]):not([readonly])',
    'button:not([disabled])',
    'a[href]',
    '[tabindex]:not([tabindex="-1"])'
  ].join(',');

  function syncSelectedState(root = document) {
    root.querySelectorAll(tabSelectors).forEach((item) => {
      item.setAttribute('aria-selected', String(item.classList.contains('active')));
    });
  }

  function focusFirstField(layer) {
    if (!layer || layer.getAttribute('aria-hidden') === 'true') return;
    const firstRequired = layer.querySelector('input[required]:not([disabled]):not([readonly]), select[required]:not([disabled]), textarea[required]:not([disabled]):not([readonly])');
    const firstEditable = layer.querySelector('input:not([disabled]):not([readonly]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]):not([readonly])');
    const target = firstRequired || firstEditable;
    if (target && !layer.contains(document.activeElement)) requestAnimationFrame(() => target.focus({ preventScroll: true }));
  }

  document.addEventListener('click', (event) => {
    const tab = event.target.closest(tabSelectors);
    if (tab) requestAnimationFrame(() => syncSelectedState(tab.parentElement || document));
  });

  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Enter' || event.isComposing || event.ctrlKey || event.altKey || event.metaKey) return;
    const current = event.target;
    const layer = current.closest('.modal.show, .drawer.show, .confirm.show');
    if (!layer) return;
    if (current.matches('button, a, input[type="checkbox"], input[type="radio"]')) return;
    if (current.matches('textarea') && event.shiftKey) return;
    const focusables = [...layer.querySelectorAll(focusableSelector)].filter((el) => el.offsetParent !== null);
    const index = focusables.indexOf(current);
    if (index < 0 || index === focusables.length - 1) return;
    event.preventDefault();
    focusables[index + 1].focus();
  });

  const observer = new MutationObserver((records) => {
    for (const record of records) {
      if (record.type !== 'attributes' || record.attributeName !== 'class') continue;
      const layer = record.target;
      if (layer.matches('.modal.show, .drawer.show, .confirm.show')) focusFirstField(layer);
    }
  });

  document.querySelectorAll('.modal, .drawer, .confirm').forEach((layer) => {
    observer.observe(layer, { attributes: true, attributeFilter: ['class'] });
  });
  syncSelectedState();
})();
