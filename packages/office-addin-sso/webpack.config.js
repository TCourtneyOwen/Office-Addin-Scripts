const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin} = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const webpack = require("webpack");
const path = require('path');

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      server: "./src/server.ts",
      auth: "./src/auth.ts",
      errors: "./src/errors.ts",
      graphHelper: "./src/msgraph-helper.ts",
      serverStorage: "./src/server-storage.ts"
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
    },
    resolve: {
      extensions: [
                ".ts",
                ".tsx",
                ".html",
                ".js"
            ]
        },
    node: {
      fs: 'empty',
      net: 'empty',
      child_process: 'empty'
    },
    module: {
      rules: [
                {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: "babel-loader"
                },
                {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "ts-loader"
                },
                {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
                },
                {
          test: /\.(png|jpg|jpeg|gif)$/,
          use: "file-loader"
                }
            ]
        },
    plugins: [
      new CleanWebpackPlugin(),
      new CopyWebpackPlugin([
                {
          to: "index.html",
          from: "./public/index.html"
                }
            ]),
        ]
    };

  return config;
};
