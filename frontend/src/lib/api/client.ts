export class ApiError extends Error {
  status: number;

  constructor(message: string, status = 0) {
    super(message);
    this.name = 'ApiError';
    this.status = status;
  }
}

const DEFAULT_API_BASE = 'http://127.0.0.1:8000';

function trimTrailingSlash(value: string) {
  return value.replace(/\/+$/, '');
}

export function getApiBaseUrl() {
  const configuredBase =
    typeof import.meta !== 'undefined' && import.meta.env?.VITE_API_BASE
      ? String(import.meta.env.VITE_API_BASE)
      : '';

  return trimTrailingSlash(configuredBase || DEFAULT_API_BASE);
}

async function requestJson<T>(path: string, init: RequestInit = {}, timeoutMs = 8000): Promise<T> {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(`${getApiBaseUrl()}${path}`, {
      ...init,
      signal: controller.signal,
      headers: {
        'Content-Type': 'application/json',
        ...(init.headers ?? {})
      }
    });

    const contentType = response.headers.get('content-type') || '';
    const payload = contentType.includes('application/json')
      ? ((await response.json()) as T & { message?: string })
      : ({ message: await response.text() } as T & { message?: string });

    if (!response.ok) {
      throw new ApiError(payload.message || 'Istek basarisiz oldu.', response.status);
    }

    return payload;
  } catch (error) {
    if (error instanceof ApiError) {
      throw error;
    }

    if (error instanceof DOMException && error.name === 'AbortError') {
      throw new ApiError('Istek zaman asimina ugradi.');
    }

    throw new ApiError('API servisine ulasilamadi.');
  } finally {
    window.clearTimeout(timeout);
  }
}

export function getJson<T>(path: string, timeoutMs?: number) {
  return requestJson<T>(path, { method: 'GET' }, timeoutMs);
}

export function postJson<T>(path: string, body: unknown, timeoutMs?: number) {
  return requestJson<T>(
    path,
    {
      method: 'POST',
      body: JSON.stringify(body)
    },
    timeoutMs
  );
}
