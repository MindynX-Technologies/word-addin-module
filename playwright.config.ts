import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  timeout: 35_000,
  workers: process.env.WORD_CONVERTER_TESTS ? 2 : undefined,
  use: { ...devices['Desktop Chrome'], baseURL: 'http://localhost:4174', channel: process.env.PLAYWRIGHT_CHANNEL },
  projects: [
    { name: 'chromium', use: { browserName: 'chromium' } },
    ...(process.env.PLAYWRIGHT_ALL_BROWSERS ? [
      { name: 'firefox', use: { ...devices['Desktop Firefox'], browserName: 'firefox' as const, channel: undefined } },
      { name: 'webkit', use: { ...devices['Desktop Safari'], browserName: 'webkit' as const, channel: undefined } },
    ] : []),
  ],
  webServer: {
    command: 'npm run dev -- --host localhost --port 4174 --strictPort',
    url: 'http://localhost:4174',
    reuseExistingServer: !process.env.CI,
  },
});
