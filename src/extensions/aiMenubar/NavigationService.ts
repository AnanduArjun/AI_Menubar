import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { INavItem } from "./AI_Menubar";

// ------------------------------------------------------------------
// Name of the SharePoint list that holds the menu items.
// ------------------------------------------------------------------
export const NAV_LIST_TITLE: string = "AI_Navigation";

// Empty columns come back from SharePoint as null; the code below
// treats null and undefined the same way, so optional fields are enough.
interface INavListRow {
  Id: number;
  Title: string;
  NavUrl?: string;
  ParentId?: number;
  SortOrder?: number;
  OpenInNewTab?: boolean;
}

/**
 * Reads the navigation list and turns it into a menu tree.
 * Rows with no Parent become top-level items; rows with a Parent
 * become dropdown items under that parent.
 */
export async function getNavItems(spHttpClient: SPHttpClient, siteUrl: string): Promise<INavItem[]> {
  const endpoint: string =
    `${siteUrl}/_api/web/lists/getbytitle('${NAV_LIST_TITLE}')/items` +
    `?$select=Id,Title,NavUrl,ParentId,SortOrder,OpenInNewTab` +
    `&$orderby=SortOrder asc,Title asc&$top=500`;

  const response: SPHttpClientResponse = await spHttpClient.get(
    endpoint,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } }
  );

  if (!response.ok) {
    throw new Error(`Navigation list request failed (${response.status})`);
  }

  const json: { value: INavListRow[] } = await response.json();
  const rows: INavListRow[] = json.value || [];

  const toItem = (row: INavListRow): INavItem => ({
    label: row.Title,
    url: row.NavUrl || "#",
    openInNewTab: !!row.OpenInNewTab,
    children: []
  });

  const topLevel: INavItem[] = [];
  const byId: { [id: number]: INavItem } = {};

  // First pass: create every item
  rows.forEach((row: INavListRow) => {
    byId[row.Id] = toItem(row);
  });

  // Second pass: attach children to parents (rows are already sorted)
  rows.forEach((row: INavListRow) => {
    const item: INavItem = byId[row.Id];
    if (row.ParentId && byId[row.ParentId]) {
      (byId[row.ParentId].children as INavItem[]).push(item);
    } else {
      topLevel.push(item);
    }
  });

  return topLevel;
}