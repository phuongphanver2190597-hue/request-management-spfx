/**
 * Build deep link URL trỏ thẳng tới một request cụ thể.
 * Tự lấy page URL hiện tại, chỉ cần truyền requestId.
 */
export function buildApprovalLink(requestId: number): string {
  const base = window.location.origin + window.location.pathname;
  const url = new URL(base);
  url.searchParams.set('requestId', String(requestId));
  return url.toString();
}
