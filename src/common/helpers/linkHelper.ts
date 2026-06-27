/**
 * Build deep link URL trỏ thẳng tới một request cụ thể.
 * pageUrl: full URL của SharePoint page chứa web part
 *   ví dụ: https://tenant.sharepoint.com/sites/xxx/SitePages/VehicleBooking.aspx
 */
export function buildApprovalLink(pageUrl: string, requestId: number): string {
  const url = new URL(pageUrl);
  url.searchParams.set('requestId', String(requestId));
  return url.toString();
}
