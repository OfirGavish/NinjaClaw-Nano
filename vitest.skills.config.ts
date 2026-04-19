import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['.copilot/skills/**/tests/*.test.ts'],
  },
});
