import type { Config } from 'tailwindcss';

export default {
  darkMode: 'class',
  content: ['./frontend/index.html', './frontend/src/**/*.{svelte,ts}'],
  theme: {
    extend: {
      colors: {
        enterprise: {
          primary: 'var(--enterprise-primary)',
          accent: 'var(--enterprise-accent)',
          bg: {
            DEFAULT: 'var(--enterprise-bg)',
            soft: 'var(--enterprise-bg-soft)'
          },
          surface: {
            DEFAULT: 'var(--enterprise-surface)',
            alt: 'var(--enterprise-surface-alt)'
          },
          border: {
            DEFAULT: 'var(--enterprise-border)',
            strong: 'var(--enterprise-border-strong)'
          },
          success: 'var(--enterprise-success)',
          warning: 'var(--enterprise-warning)',
          danger: 'var(--enterprise-danger)',
          info: 'var(--enterprise-info)',
          text: {
            primary: 'var(--enterprise-text-primary)',
            secondary: 'var(--enterprise-text-secondary)',
            muted: 'var(--enterprise-text-muted)'
          }
        }
      },
      fontFamily: {
        sans: ['Inter', '"Segoe UI"', 'Roboto', 'ui-sans-serif', 'system-ui', 'sans-serif'],
        mono: ['"IBM Plex Mono"', '"Cascadia Code"', '"SFMono-Regular"', 'Consolas', 'monospace']
      },
      fontSize: {
        'data-xs': ['0.6875rem', { lineHeight: '0.95rem' }],
        'data-sm': ['0.75rem', { lineHeight: '1.05rem' }]
      },
      boxShadow: {
        enterprise: '0 8px 24px rgba(15, 23, 42, 0.10)',
        window: '0 18px 42px rgba(15, 23, 42, 0.16)',
        'enterprise-focus': '0 22px 48px rgba(15, 23, 42, 0.20)'
      },
      borderRadius: {
        enterprise: '0.25rem'
      }
    }
  },
  plugins: []
} satisfies Config;
