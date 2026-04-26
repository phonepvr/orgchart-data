/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        // AM/NS Brand v1.0
        red: {
          50: '#FDF2F2', 100: '#FBE0E0', 200: '#F5B5B5', 300: '#EC8585',
          400: '#DF5454', 500: '#D12D2D', 600: '#B81F1F', 700: '#971717',
          800: '#741212', 900: '#541010',
        },
        graphite: {
          50: '#F6F7F9', 100: '#EBEDF1', 200: '#D6DAE2', 300: '#B4BBC8',
          400: '#8892A3', 500: '#5F6B80', 600: '#434E63', 700: '#2E3647',
          800: '#1C222E', 900: '#0E1219',
        },
        ember: '#D9761E',
        leaf: '#3F9460',
        signal: '#1B5EA6',
      },
      fontFamily: {
        display: ['Fraunces', 'Georgia', 'serif'],
        sans: ['"Inter Tight"', 'Inter', 'system-ui', 'sans-serif'],
        mono: ['"JetBrains Mono"', 'ui-monospace', 'monospace'],
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
