import styles from "./AI_Footer.module.scss";
 
// ------------------------------------------------------------------
// EDIT FOOTER SETTINGS HERE
// ------------------------------------------------------------------
const FOOTER_LOGO_URL: string =
  "https://aciesintranet.sharepoint.com/sites/IntranetDemo/Acies%20%20Images%20Icons/Acies%20Logo.png";
 
const WEBSITE_URL: string = "https://aciesinnovations.com/";
const LINKEDIN_URL: string = "https://www.linkedin.com/company/myelos/";
 
// Paste the LinkedIn icon image from your design here (upload it to
// SharePoint first). Leave empty to show the word "LinkedIn" instead.
const LINKEDIN_ICON_URL: string =
  "https://aciesintranet.sharepoint.com/sites/IntranetDemo/Acies%20%20Images%20Icons/Linkdin.png";
// ------------------------------------------------------------------
 
export class AI_Footer {
  private rootEl: HTMLElement;
 
  constructor(container: HTMLElement) {
    this.rootEl = container;
  }
 
  public render(): void {
    this.rootEl.innerHTML = `
      <div class="${styles.footerOuter}">
        <footer class="${styles.aiFooter}">
          <a class="${styles.logoLink}" href="${WEBSITE_URL}" target="_blank" rel="noopener noreferrer">
            <img class="${styles.logo}" src="${FOOTER_LOGO_URL}" alt="Acies Innovations" />
          </a>
 
          <div class="${styles.links}">
            <a class="${styles.iconLink}" href="${WEBSITE_URL}" target="_blank" rel="noopener noreferrer"
               aria-label="Acies Innovations website" title="Website">
              ${this.getGlobeIcon()}
            </a>
            <a class="${styles.iconLink}" href="${LINKEDIN_URL}" target="_blank" rel="noopener noreferrer"
               aria-label="Acies Innovations on LinkedIn" title="LinkedIn">
              ${this.getLinkedInMarkup()}
            </a>
          </div>
        </footer>
      </div>
    `;
  }
 
  public dispose(): void {
    this.rootEl.innerHTML = "";
  }
 
  // Generic globe icon (website)
  private getGlobeIcon(): string {
    return `
      <svg class="${styles.icon}" viewBox="0 0 24 24" width="26" height="26" fill="none"
           stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <circle cx="12" cy="12" r="10"/>
        <line x1="2" y1="12" x2="22" y2="12"/>
        <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>
      </svg>
    `;
  }
 
  private getLinkedInMarkup(): string {
    if (LINKEDIN_ICON_URL) {
      return `<img class="${styles.iconImg}" src="${LINKEDIN_ICON_URL}" alt="" />`;
    }
    return `<span class="${styles.iconText}">LinkedIn</span>`;
  }
}
 