import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: true,
    environment: 'node',
    include: ['src/**/*.test.ts'],
    // Upstream's tests exercise write tools; read-only.test.ts covers the default.
    env: { MSTEAMS_ALLOW_WRITES: '1' },
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html'],
      include: ['src/**/*.ts'],
      exclude: [
        'src/**/*.test.ts',
        'src/research/**',
        'src/index.ts',
      ],
    },
  },
});
