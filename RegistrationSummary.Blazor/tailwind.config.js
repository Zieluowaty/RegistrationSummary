module.exports = {
    content: [
        './Pages/**/*.razor',
        './Shared/**/*.razor',
        './**/*.cshtml',
        './**/*.html'
    ],
    theme: {
        extend: {
            colors: {
                brand: {
                    brown: '#522930',     // RGB 82, 41, 48
                    red: '#d14033',       // RGB 209, 64, 51
                    gold: '#c79466',      // RGB 199, 148, 102
                    white: '#FAFAFA',     // RGB 250, 250, 250
                }
            }
        },
    },
    plugins: [],
};