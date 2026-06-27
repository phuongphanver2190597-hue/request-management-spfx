export interface IEnvironmentConfig {
  env: 'local' | 'dev' | 'prod';
  apiBaseUrl: string;
  enableDebugLog: boolean;
}

const CONFIGS: Record<string, IEnvironmentConfig> = {
  local: {
    env: 'local',
    apiBaseUrl: 'http://localhost:7071/api',
    enableDebugLog: true,
  },
  dev: {
    env: 'dev',
    apiBaseUrl: 'https://api-dev.greenfeed.com.vn/api',
    enableDebugLog: true,
  },
  prod: {
    env: 'prod',
    apiBaseUrl: 'https://api.greenfeed.com.vn/api',
    enableDebugLog: false,
  },
};

function detectEnv(siteUrl: string): 'local' | 'dev' | 'prod' {
  if (siteUrl.includes('localhost')) return 'local';
  if (/\/(dev|test|uat|staging)/i.test(siteUrl)) return 'dev';
  return 'prod';
}

export function getConfig(siteUrl: string, overrideApiUrl?: string): IEnvironmentConfig {
  const base = CONFIGS[detectEnv(siteUrl)];
  if (!overrideApiUrl) return base;
  return { ...base, apiBaseUrl: overrideApiUrl };
}
