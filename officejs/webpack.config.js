const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");

// Where the built add-in is served from. Dev uses the local https dev-server;
// override urlProd with your hosting origin (GitHub Pages, Azure Static Web
// Apps, etc.) before running a production build + publishing the manifest.
const urlDev = "https://localhost:3000/";
const urlProd = "https://localhost:3000/";

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      functions: "./src/functions/functions.ts",
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, "dist"),
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        { test: /\.ts$/, use: "ts-loader", exclude: /node_modules/ },
        { test: /\.html$/, use: "html-loader" },
        { test: /\.css$/, use: ["style-loader", "css-loader"] },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext][query]" },
        },
      ],
    },
    plugins: [
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: path.resolve(__dirname, "src/functions/functions.ts"),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["functions"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets", noErrorOnMissing: true },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              if (dev) return content;
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: "https",
      port: 3000,
      static: [path.join(__dirname, "dist")],
    },
  };
  return config;
};
