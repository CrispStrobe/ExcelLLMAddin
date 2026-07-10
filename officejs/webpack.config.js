const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");
const devCerts = require("office-addin-dev-certs");

// Where the built add-in is served from. Dev uses the local https dev-server;
// override urlProd with your hosting origin (GitHub Pages, Azure Static Web
// Apps, etc.) before running a production build + publishing the manifest.
// Production hosting origin. GitHub Pages project site for CrispStrobe/ExcelLLMAddin.
// Change these if you host elsewhere (Azure Static Web Apps, Cloudflare Pages, a
// custom domain, ...).
const prodOrigin = "https://crispstrobe.github.io";
const urlProd = `${prodOrigin}/ExcelLLMAddin/`;

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  // Serve the SAME cert that `office-addin-dev-certs install` trusted in the
  // system keychain. Without this, webpack-dev-server self-signs a different,
  // untrusted cert and Excel's WKWebView blocks the add-in ("not signed by a
  // valid security certificate").
  const httpsOptions = dev ? await devCerts.getHttpsServerOptions() : undefined;
  const config = {
    devtool: "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      functions: "./src/functions/functions.ts",
      // Dev-only browser harness: runs the task pane with Office mocked.
      ...(dev ? { harness: "./src/harness/harness.ts" } : {}),
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, "dist"),
      // Relative so injected asset URLs resolve correctly under a Pages subpath
      // (https://…/ExcelLLMAddin/taskpane.html -> ./taskpane.js).
      publicPath: "",
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "ts-loader",
            // transpileOnly: don't type-check during bundling. ts-loader's
            // watch-mode incremental type checker produced phantom "Cannot find
            // name" errors misattributed to the wrong files. Types are enforced
            // separately by `npm run typecheck`, jest (ts-jest), and CI.
            options: { transpileOnly: true },
          },
        },
        { test: /\.html$/, use: "html-loader" },
        { test: /\.css$/, use: [MiniCssExtractPlugin.loader, "css-loader"] },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext][query]" },
        },
      ],
    },
    plugins: [
      new MiniCssExtractPlugin(),
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: path.resolve(__dirname, "src/functions/functions.ts"),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        // Shared runtime: the task-pane page also hosts the custom functions, so
        // both bundles load here and the functions register in the same runtime.
        chunks: ["functions", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["functions"],
      }),
      ...(dev
        ? [
            new HtmlWebpackPlugin({
              filename: "harness.html",
              template: "./src/harness/harness.html",
              chunks: ["harness"],
            }),
          ]
        : []),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets", noErrorOnMissing: true },
          { from: "src/site", to: "." },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              if (dev) return content;
              // Resource URLs (with trailing slash) -> prod subpath; then the
              // bare AppDomain origin -> prod origin.
              return content
                .toString()
                .replace(/https:\/\/localhost:3000\//g, urlProd)
                .replace(/https:\/\/localhost:3000/g, prodOrigin);
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: httpsOptions
        ? { type: "https", options: { key: httpsOptions.key, cert: httpsOptions.cert, ca: httpsOptions.ca } }
        : "https",
      port: 3000,
      // Do NOT serve dist/ statically: webpack-dev-server already serves every
      // emitted asset (taskpane.html, functions.js, functions.json, icons) from
      // memory. Pointing static at the build output collides with those in-memory
      // assets ("Multiple assets emit ... functions.json") and corrupts the build.
      static: false,
      // Don't show the fullscreen overlay for cross-origin "Script error." runtime
      // events (they come from office.js / the host, not our code).
      client: { overlay: { errors: true, warnings: false, runtimeErrors: false } },
    },
  };
  return config;
};
