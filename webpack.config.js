const path = require('path');
const MinifyPlugin = require('babel-minify-webpack-plugin');

module.exports = {
    entry: ['babel-polyfill', './SPListItemModel.ts'],
    module: {
        loaders: [
            {
                test: /\.ts?$/, loaders: ["babel-loader", "ts-loader"],
                exclude: /node_modules/
            }
        ]
    },
    resolve: {
        extensions: ['.ts', '.js']
    },
    output: {
        filename: 'bundle.js',
        path:
            path.resolve(__dirname, 'dist')
    }
    ,
    externals: {
        "moment":
            "moment",
        "reflect-metadata":
            "reflect-metadata",
        "sp-pnp-js":
            "sp-pnp-js"

    },
    plugins: [
        new MinifyPlugin()
    ]
}
;
