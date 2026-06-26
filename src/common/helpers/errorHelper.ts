export function extractErrorMessage(err: unknown): string {
  if (!err) return 'Đã xảy ra lỗi không xác định';
  if (typeof err === 'string') return err;
  if (err instanceof Error) return err.message;
  const e = err as { message?: string; error?: { message?: string } };
  if (e.error?.message) return e.error.message;
  if (e.message) return e.message;
  return JSON.stringify(err);
}

export async function parseSpError(res: { json: () => Promise<unknown> }, status: number): Promise<string> {
  try {
    const body = await res.json() as { error?: { message?: string } };
    return body?.error?.message || `HTTP ${status}`;
  } catch {
    return `HTTP ${status}`;
  }
}
