export type ThemeMode = 'auto' | 'dark' | 'light';
export type VsCodeThemeKind = 'dark' | 'light' | 'high-contrast';

export class ThemeManager {
  private mode: ThemeMode = 'auto';
  private vsCodeTheme: VsCodeThemeKind = 'dark';
  private btnThemeToggle: HTMLButtonElement | null = null;
  private themeIcon: HTMLElement | null = null;
  private themeText: HTMLElement | null = null;

  constructor() {
    this.btnThemeToggle = document.getElementById('btn-theme-toggle') as HTMLButtonElement | null;
    this.themeIcon = document.getElementById('theme-icon');
    this.themeText = document.getElementById('theme-text');

    if (this.btnThemeToggle) {
      this.btnThemeToggle.addEventListener('click', () => this.cycleTheme());
    }

    this.applyTheme();
  }

  public setVsCodeTheme(kind: VsCodeThemeKind) {
    this.vsCodeTheme = kind;
    if (this.mode === 'auto') {
      this.applyTheme();
    }
  }

  public cycleTheme() {
    if (this.mode === 'auto') {
      this.mode = 'dark';
    } else if (this.mode === 'dark') {
      this.mode = 'light';
    } else {
      this.mode = 'auto';
    }
    this.applyTheme();
  }

  public getEffectiveTheme(): 'dark' | 'light' | 'high-contrast' {
    if (this.mode === 'dark') return 'dark';
    if (this.mode === 'light') return 'light';
    return this.vsCodeTheme;
  }

  private applyTheme() {
    const effective = this.getEffectiveTheme();

    document.body.classList.remove('theme-dark', 'theme-light', 'theme-high-contrast');
    document.body.classList.add(`theme-${effective}`);

    if (this.themeIcon && this.themeText) {
      if (this.mode === 'auto') {
        this.themeIcon.textContent = effective === 'dark' ? '🌙' : '☀️';
        this.themeText.textContent = `Oto (${effective === 'dark' ? 'Koyu' : 'Açık'})`;
      } else if (this.mode === 'dark') {
        this.themeIcon.textContent = '🌙';
        this.themeText.textContent = 'Koyu';
      } else {
        this.themeIcon.textContent = '☀️';
        this.themeText.textContent = 'Açık';
      }
    }
  }
}

