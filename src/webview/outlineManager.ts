export class OutlineManager {
  private sidebarEl: HTMLElement | null = null;
  private btnToggleOutline: HTMLElement | null = null;
  private container: HTMLElement;
  private isOpen: boolean = false;

  constructor(container: HTMLElement) {
    this.container = container;
    this.sidebarEl = document.getElementById('outline-sidebar');
    this.btnToggleOutline = document.getElementById('btn-toggle-outline');

    this.btnToggleOutline?.addEventListener('click', () => {
      this.toggle();
    });
  }

  public updateHeadings(): number {
    if (!this.sidebarEl) return 0;
    const listContainer = this.sidebarEl.querySelector('.outline-list');
    if (!listContainer) return 0;

    listContainer.innerHTML = '';
    const headings = Array.from(this.container.querySelectorAll('h1, h2, h3, h4, h5, h6'));

    if (headings.length === 0) {
      this.btnToggleOutline?.classList.add('hidden');
      this.close();
      return 0;
    }

    this.btnToggleOutline?.classList.remove('hidden');

    headings.forEach((h, index) => {
      const level = parseInt(h.tagName.substring(1), 10);
      const text = h.textContent?.trim() || `Başlık ${index + 1}`;

      const item = document.createElement('div');
      item.className = `outline-item level-${level}`;
      item.textContent = text;
      item.title = text;

      item.addEventListener('click', () => {
        h.scrollIntoView({ behavior: 'smooth', block: 'start' });
        h.classList.add('heading-target-highlight');
        setTimeout(() => {
          h.classList.remove('heading-target-highlight');
        }, 1500);
      });

      listContainer.appendChild(item);
    });

    return headings.length;
  }

  public toggle() {
    if (this.isOpen) {
      this.close();
    } else {
      this.open();
    }
  }

  public open() {
    this.isOpen = true;
    this.sidebarEl?.classList.remove('hidden');
    this.btnToggleOutline?.classList.add('active');
  }

  public close() {
    this.isOpen = false;
    this.sidebarEl?.classList.add('hidden');
    this.btnToggleOutline?.classList.remove('active');
  }
}
