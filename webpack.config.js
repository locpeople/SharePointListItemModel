const path = require('path');

module.exports = {
    entry: './SPListItemModel.ts',
    module: {
        loaders: [
            {
                test: /\.js$/,
                exclude: /node_modules/,
                loader: 'babel-loader',
                query: {presets: ["es2015"]}
            },
            {
                test: /\.ts?$/,
                use: 'ts-loader',
                exclude: /node_modules/
            }

        ]
    },
    resolve: {
        extensions: ['.ts', '.js']
    },
    output: {
        filename: 'bundle.js',
        path: path.resolve(__dirname, 'dist')
    },
    externals: {
        "moment": "moment",
        "reflect-metadata": "reflect-metadata",
        "sp-pnp-js": "sp-pnp-js"

    }
};
