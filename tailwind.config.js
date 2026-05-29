/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        // AM/NS Brand v1.1 — Smart Red ramp derived from #E52726 (canonical).
        red: {
          50: '#FDEBEA', 100: '#FBD0CE', 200: '#F7A6A1', 300: '#F17C75',
          400: '#EA524A', 500: '#E52726', 600: '#C71F1E', 700: '#A11816',
          800: '#7B1110', 900: '#560A0A',
        },
        graphite: {
          50: '#F6F7F9', 100: '#EBEDF1', 200: '#D6DAE2', 300: '#B4BBC8',
          400: '#8892A3', 500: '#5F6B80', 600: '#434E63', 700: '#2E3647',
          800: '#1C222E', 900: '#0E1219',
        },
        // Status semantics (kept orthogonal to brand palette)
        ember: '#D9761E',
        leaf: '#3F9460',
        signal: '#1B5EA6',
        // Brand secondaries — accents, never dominant
        'accent-yellow': '#FFA700',
        'accent-green':  '#C0F353',
        'accent-blue':   '#A8E0FF',
      },
      fontFamily: {
        display: ['"Albert Sans"', 'system-ui', 'sans-serif'],
        sans:    ['"Albert Sans"', 'system-ui', 'sans-serif'],
        mono:    ['"JetBrains Mono"', 'ui-monospace', 'monospace'],
      },
      borderRadius: {
        'brand': '2px',
      },
      transitionTimingFunction: {
        'brand': 'cubic-bezier(0.2, 0.7, 0.2, 1)',
      },
      transitionDuration: {
        'brand-fast': '120ms',
        'brand-base': '200ms',
        'brand-slow': '360ms',
      },
    },
  },
  plugins: [],
}
