import { override } from "@microsoft/decorators";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from "@microsoft/sp-application-base";

import { AI_Menubar, INavItem, DEFAULT_NAV_ITEMS } from "./AI_Menubar";
import { AI_Footer } from "./AI_Footer";
import { getNavItems } from "./NavigationService";
import './assets/style.css';

export interface IAiMenubarApplicationCustomizerProperties {
  testMessage?: string;
}

// The scrolling area of a modern SharePoint page. The footer is added at
// the end of it, so it only shows once the user scrolls to the bottom.
const SCROLL_REGION_SELECTOR: string = '[data-automation-id="contentScrollRegion"]';

export default class AiMenubarApplicationCustomizer
  extends BaseApplicationCustomizer<IAiMenubarApplicationCustomizerProperties> {

  private topPlaceholder: PlaceholderContent | undefined;
  private bottomPlaceholder: PlaceholderContent | undefined;
  private menubar: AI_Menubar | undefined;
  private footer: AI_Footer | undefined;
  private footerHost: HTMLDivElement | undefined;
  private footerRetryTimer: number | undefined;
  private navItems: INavItem[] = DEFAULT_NAV_ITEMS;

  @override
  public async onInit(): Promise<void> {
    // Load menu items from the AI_Navigation list; fall back to the dummy menu.
    try {
      const items: INavItem[] = await getNavItems(
        this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl
      );
      this.navItems = items.length ? items : DEFAULT_NAV_ITEMS;
    } catch (error) {
      console.warn("AI_Menubar: using fallback menu.", error);
      this.navItems = DEFAULT_NAV_ITEMS;
    }

    this.context.placeholderProvider.changedEvent.add(this, this.renderPlaceholders);
    this.context.application.navigatedEvent.add(this, this.onNavigated);
    this.renderPlaceholders();
    this.renderFooter();
  }

  // ---------------------------------------------------------------- navbar

  private renderPlaceholders(): void {
    if (!this.topPlaceholder) {
      this.topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this.onTopDispose }
      );
      if (this.topPlaceholder) {
        this.menubar = new AI_Menubar(
          this.topPlaceholder.domElement,
          this.context.pageContext.user.email,
          this.navItems
        );
        this.menubar.render();
      }
    }
  }

  // SharePoint switches pages without a full reload: update the highlighted
  // menu item and make sure the footer is still at the end of the new page.
  private onNavigated(): void {
    window.setTimeout(() => {
      if (this.menubar) {
        this.menubar.refreshActive();
      }
      this.renderFooter();
    }, 0);
  }

  private onTopDispose = (): void => {
    if (this.menubar) {
      this.menubar.dispose();
      this.menubar = undefined;
    }
    this.topPlaceholder = undefined;
  };

  // ---------------------------------------------------------------- footer

  // Adds the footer at the end of the page's scrolling content.
  // The scroll area can appear a moment after the extension starts,
  // so this retries for a few seconds before falling back.
  private renderFooter(attempt: number = 0): void {
    if (this.footerRetryTimer) {
      window.clearTimeout(this.footerRetryTimer);
      this.footerRetryTimer = undefined;
    }

    const region: HTMLElement | null = document.querySelector(SCROLL_REGION_SELECTOR);

    if (!region) {
      if (attempt < 20) {
        this.footerRetryTimer = window.setTimeout(() => this.renderFooter(attempt + 1), 250);
      } else {
        this.renderFooterInBottomPlaceholder();
      }
      return;
    }

    // The page lives in an inner wrapper inside the scroll area (the same
    // wrapper SharePoint puts its own site footer in). Add ours at its end.
    let container: HTMLElement = region;
    for (let i: number = 0; i < region.children.length; i++) {
      const child: Element = region.children[i];
      if (!child.hasAttribute("data-ai-footer")) {
        container = child as HTMLElement;
        break;
      }
    }

    // Already in place and still at the end of the page: nothing to do.
    if (this.footerHost && container.lastElementChild === this.footerHost) {
      return;
    }

    this.removeFooter();

    this.footerHost = document.createElement("div");
    this.footerHost.setAttribute("data-ai-footer", "true");
    container.appendChild(this.footerHost);

    this.footer = new AI_Footer(this.footerHost);
    this.footer.render();
  }

  // Fallback for pages without the usual scroll area (pinned to the bottom).
  private renderFooterInBottomPlaceholder(): void {
    if (this.bottomPlaceholder) {
      return;
    }
    this.bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Bottom,
      { onDispose: this.onBottomDispose }
    );
    if (this.bottomPlaceholder) {
      this.footer = new AI_Footer(this.bottomPlaceholder.domElement);
      this.footer.render();
    }
  }

  private removeFooter(): void {
    if (this.footer) {
      this.footer.dispose();
      this.footer = undefined;
    }
    if (this.footerHost && this.footerHost.parentElement) {
      this.footerHost.parentElement.removeChild(this.footerHost);
    }
    this.footerHost = undefined;
  }

  private onBottomDispose = (): void => {
    if (this.footer) {
      this.footer.dispose();
      this.footer = undefined;
    }
    this.bottomPlaceholder = undefined;
  };

  @override
  public onDispose(): void {
    if (this.footerRetryTimer) {
      window.clearTimeout(this.footerRetryTimer);
    }
    this.onTopDispose();
    this.removeFooter();
    this.onBottomDispose();
    super.onDispose();
  }
}