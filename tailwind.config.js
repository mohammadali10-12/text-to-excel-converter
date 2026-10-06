/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./public/**/*.html', './public/assets/js/**/*.js'],
  theme: {
    extend: {
      maxWidth: { app: '1320px' },
    },
  },
  plugins: [],
};
