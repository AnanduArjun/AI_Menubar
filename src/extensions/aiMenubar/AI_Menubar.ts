import styles from "./AI_Menubar.module.scss";

// ------------------------------------------------------------------
// LOGO
// ------------------------------------------------------------------
const LOGO_URL: string = "https://aciesintranet.sharepoint.com/sites/IntranetDemo/Acies%20%20Images%20Icons/Acies%20Logo.png";
// ------------------------------------------------------------------

export interface INavItem {
  label: string;
  url: string;
  openInNewTab?: boolean;
  children?: INavItem[];
}

// ------------------------------------------------------------------
// FALLBACK MENU
// Used only when the AI_Navigation list is missing, empty or can't be
// read. Learning and Communities have dummy dropdown items for now.
// ------------------------------------------------------------------
export const DEFAULT_NAV_ITEMS: INavItem[] = [
  { label: "Home", url: "/sites/IntranetDemo/SitePages/Home.aspx" },
  {
    label: "Learning",
    url: "#",
    children: [
      { label: "Training Calendar", url: "#" },
      { label: "E-Learning Courses", url: "#" },
      { label: "Certifications", url: "#" }
    ]
  },
  { label: "Help", url: "/sites/IntranetDemo/SitePages/Help.aspx" },
  {
    label: "Communities",
    url: "#",
    children: [
      { label: "Tech Community", url: "#" },
      { label: "Sports & Wellness", url: "#" },
      { label: "Social Club", url: "#" }
    ]
  }
];
// ------------------------------------------------------------------

const MOBILE_BREAKPOINT: number = 768;

const CHEVRON_SVG: string = `
  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M6 9l6 6 6-6" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>`;

function escapeHtml(value: string): string {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function isCurrentUrl(url: string): boolean {
  if (!url || url === "#") {
    return false;
  }
  try {
    const target: string = new URL(url, window.location.origin).pathname.toLowerCase();
    return target === window.location.pathname.toLowerCase();
  } catch {
    return false;
  }
}

function isItemCurrent(item: INavItem): boolean {
  return isCurrentUrl(item.url) || (item.children || []).some(isItemCurrent);
}

export class AI_Menubar {
  private rootEl: HTMLElement;
  private userEmail: string;
  private navItems: INavItem[];
  private isMenuOpen: boolean = false;

  private hamburgerBtn: HTMLButtonElement | undefined;
  private mobilePanel: HTMLDivElement | undefined;

  constructor(container: HTMLElement | string, userEmail: string = "", navItems: INavItem[] = DEFAULT_NAV_ITEMS) {
    const el: HTMLElement | null =
      typeof container === "string" ? document.getElementById(container) : container;

    if (!el) {
      throw new Error(`AI_Menubar: container "${String(container)}" not found`);
    }
    this.rootEl = el;
    this.userEmail = userEmail;
    this.navItems = navItems && navItems.length ? navItems : DEFAULT_NAV_ITEMS;
  }

  public render(): void {
    this.removeListeners();
    this.rootEl.innerHTML = this.getMarkup();
    this.cacheDom();
    this.bindEvents();
  }

  public dispose(): void {
    this.removeListeners();
    this.rootEl.innerHTML = "";
  }

  // Re-checks which page we're on and moves the highlight/underline.
  // Called after SharePoint navigates to another page without a full reload.
  public refreshActive(): void {
    const desktopNav: HTMLElement | null = this.rootEl.querySelector(`.${styles.navRow}`);
    if (desktopNav) {
      desktopNav.innerHTML = this.getDesktopNavMarkup();
    }
    const mobileNav: HTMLElement | null = this.rootEl.querySelector(`.${styles.mobileNav}`);
    if (mobileNav) {
      mobileNav.innerHTML = this.getMobileNavMarkup();
    }
    if (this.isMenuOpen) {
      this.closeMenu();
    }
  }

  // ---------------------------------------------------------------- markup

  private getMarkup(): string {
    return `
      <header class="${styles.aiMenubar}">
        <div class="${styles.logoWrap}">
          <img class="${styles.logo}" src="${LOGO_URL}" alt="Logo" />
        </div>

        <div class="${styles.searchWrap}">
          ${this.getSearchMarkup()}
        </div>

        <div class="${styles.profileWrap}">
          ${this.getProfileMarkup()}
        </div>

        <button class="${styles.hamburger}" type="button" aria-label="Open menu" aria-expanded="false">
          <span></span>
          <span></span>
          <span></span>
        </button>
      </header>

      <nav class="${styles.navRow}" aria-label="Main">
        ${this.getDesktopNavMarkup()}
      </nav>

      <div class="${styles.mobilePanel}" aria-hidden="true">
        <div class="${styles.mobileSearch}">
          ${this.getSearchMarkup()}
        </div>
        <div class="${styles.mobileNav}">
          ${this.getMobileNavMarkup()}
        </div>
        <div class="${styles.mobileProfile}">
          ${this.getProfileMarkup()}
          <span class="${styles.mobileProfileLabel}">Profile</span>
        </div>
      </div>
    `;
  }

  private getSearchMarkup(): string {
    return `
      <div class="${styles.search}">
        <svg class="${styles.searchIcon}" width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <circle cx="11" cy="11" r="7" stroke="currentColor" stroke-width="2"/>
          <line x1="21" y1="21" x2="16.65" y2="16.65" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
        </svg>
        <input class="${styles.searchInput}" type="text" placeholder="Search" />
      </div>
    `;
  }

  private getProfileMarkup(): string {
    const photoUrl: string =
      `/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(this.userEmail)}`;
    return `
      <button class="${styles.profile}" type="button" aria-label="Profile" title="${escapeHtml(this.userEmail)}">
        <img src="${photoUrl}" alt="" />
      </button>
    `;
  }

  private getLinkAttrs(item: INavItem): string {
    const target: string = item.openInNewTab ? ` target="_blank" rel="noopener noreferrer"` : "";
    return `href="${escapeHtml(item.url || "#")}"${target}`;
  }

  // Desktop: top-level links, plus hover/click dropdowns for items with children
  private getDesktopNavMarkup(): string {
    return this.navItems
      .map((item: INavItem): string => {
        const current: boolean = isItemCurrent(item);
        const linkClasses: string = current ? `${styles.navLink} ${styles.navLinkActive}` : styles.navLink;
        const children: INavItem[] = item.children || [];

        if (!children.length) {
          return `<a class="${linkClasses}" ${this.getLinkAttrs(item)}>${escapeHtml(item.label)}</a>`;
        }

        const childLinks: string = children
          .map((child: INavItem): string => {
            const childClasses: string = isCurrentUrl(child.url)
              ? `${styles.dropdownLink} ${styles.dropdownLinkActive}`
              : styles.dropdownLink;
            return `<a class="${childClasses}" ${this.getLinkAttrs(child)}>${escapeHtml(child.label)}</a>`;
          })
          .join("");

        return `
          <div class="${styles.navItem}">
            <button class="${linkClasses} ${styles.navToggle}" type="button"
                    data-nav-toggle aria-haspopup="true" aria-expanded="false">
              ${escapeHtml(item.label)}
              <span class="${styles.chevron}">${CHEVRON_SVG}</span>
            </button>
            <div class="${styles.dropdown}" role="menu">
              ${childLinks}
            </div>
          </div>
        `;
      })
      .join("");
  }

  // Mobile: items with children expand/collapse inside the hamburger panel
  private getMobileNavMarkup(): string {
    return this.navItems
      .map((item: INavItem): string => {
        const current: boolean = isItemCurrent(item);
        const linkClasses: string = current
          ? `${styles.mobileNavLink} ${styles.mobileNavLinkActive}`
          : styles.mobileNavLink;
        const children: INavItem[] = item.children || [];

        if (!children.length) {
          return `<a class="${linkClasses}" ${this.getLinkAttrs(item)}>${escapeHtml(item.label)}</a>`;
        }

        const childLinks: string = children
          .map((child: INavItem): string =>
            `<a class="${styles.mobileSubLink}" ${this.getLinkAttrs(child)}>${escapeHtml(child.label)}</a>`
          )
          .join("");

        return `
          <div class="${styles.mobileGroup}">
            <button class="${linkClasses} ${styles.mobileGroupToggle}" type="button"
                    data-mobile-toggle aria-expanded="false">
              ${escapeHtml(item.label)}
              <span class="${styles.chevron}">${CHEVRON_SVG}</span>
            </button>
            <div class="${styles.mobileSubmenu}">
              ${childLinks}
            </div>
          </div>
        `;
      })
      .join("");
  }

  // ---------------------------------------------------------------- events

  private cacheDom(): void {
    this.hamburgerBtn = this.rootEl.querySelector(`.${styles.hamburger}`) as HTMLButtonElement;
    this.mobilePanel = this.rootEl.querySelector(`.${styles.mobilePanel}`) as HTMLDivElement;
  }

  private bindEvents(): void {
    this.rootEl.addEventListener("click", this.onRootClick);
    window.addEventListener("resize", this.onResize);
    document.addEventListener("click", this.onDocumentClick);
    document.addEventListener("keydown", this.onKeyDown);
  }

  private removeListeners(): void {
    this.rootEl.removeEventListener("click", this.onRootClick);
    window.removeEventListener("resize", this.onResize);
    document.removeEventListener("click", this.onDocumentClick);
    document.removeEventListener("keydown", this.onKeyDown);
  }

  // One listener for all clicks inside the menubar
  private onRootClick = (e: MouseEvent): void => {
    const target: HTMLElement = e.target as HTMLElement;

    if (this.hamburgerBtn && this.hamburgerBtn.contains(target)) {
      this.toggleMenu();
      return;
    }

    const navToggle: HTMLElement | null = target.closest("[data-nav-toggle]");
    if (navToggle) {
      const wrapper: HTMLElement = navToggle.parentElement as HTMLElement;
      const willOpen: boolean = !wrapper.classList.contains(styles.navItemOpen);
      this.closeDropdowns();
      if (willOpen) {
        wrapper.classList.add(styles.navItemOpen);
        navToggle.setAttribute("aria-expanded", "true");
      }
      return;
    }

    const mobileToggle: HTMLElement | null = target.closest("[data-mobile-toggle]");
    if (mobileToggle) {
      const group: HTMLElement = mobileToggle.parentElement as HTMLElement;
      const open: boolean = group.classList.toggle(styles.mobileGroupOpen);
      mobileToggle.setAttribute("aria-expanded", String(open));
    }
  };

  private onResize = (): void => {
    if (window.innerWidth > MOBILE_BREAKPOINT && this.isMenuOpen) {
      this.closeMenu();
    }
  };

  private onDocumentClick = (e: MouseEvent): void => {
    if (!this.rootEl.contains(e.target as Node)) {
      this.closeDropdowns();
      if (this.isMenuOpen) {
        this.closeMenu();
      }
    }
  };

  private onKeyDown = (e: KeyboardEvent): void => {
    if (e.key === "Escape") {
      this.closeDropdowns();
      if (this.isMenuOpen) {
        this.closeMenu();
      }
    }
  };

  private closeDropdowns(): void {
    const open: NodeListOf<HTMLElement> = this.rootEl.querySelectorAll(`.${styles.navItemOpen}`);
    open.forEach((el: HTMLElement) => {
      el.classList.remove(styles.navItemOpen);
      const btn: HTMLElement | null = el.querySelector("[data-nav-toggle]");
      if (btn) {
        btn.setAttribute("aria-expanded", "false");
      }
    });
  }

  private toggleMenu(): void {
    if (this.isMenuOpen) {
      this.closeMenu();
    } else {
      this.openMenu();
    }
  }

  private openMenu(): void {
    if (!this.mobilePanel || !this.hamburgerBtn) {
      return;
    }
    this.isMenuOpen = true;
    this.mobilePanel.classList.add(styles.isOpen);
    this.mobilePanel.setAttribute("aria-hidden", "false");
    this.hamburgerBtn.classList.add(styles.isActive);
    this.hamburgerBtn.setAttribute("aria-expanded", "true");
  }

  private closeMenu(): void {
    if (!this.mobilePanel || !this.hamburgerBtn) {
      return;
    }
    this.isMenuOpen = false;
    this.mobilePanel.classList.remove(styles.isOpen);
    this.mobilePanel.setAttribute("aria-hidden", "true");
    this.hamburgerBtn.classList.remove(styles.isActive);
    this.hamburgerBtn.setAttribute("aria-expanded", "false");
  }
}