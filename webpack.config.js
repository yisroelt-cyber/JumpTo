/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const path = require("path");
const urlDev = "https://localhost:3000/";
const urlProd = "https://yisroelt-cyber.github.io/JumpTo/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: {
        // Iteration 36: return to the full React taskpane as the primary UI.
        import: ["./src/taskpane/index.jsx", "./src/taskpane/taskpane.html"],
      },
      commands: "./src/commands/commands.js",
      dialog: {
        import: ["./src/dialog/dialog.jsx", "./src/dialog/dialog.html"],
      },
    },
    output: {
      
clean: true,
// Cache-bust in production to avoid Office/WebView and GitHub Pages serving stale JS (common cause of #321 persisting)
filename: dev ? "[name].js" : "[name].[contenthash].js",
chunkFilename: dev ? "[name].js" : "[name].[contenthash].js",
publicPath: "", // keep relative for GitHub Pages subpath deployments
    },

optimization: {
  runtimeChunk: "single",
  splitChunks: {
    chunks: "all",
    cacheGroups: {
      reactVendor: {
        test: /[\\/]node_modules[\\/](react|react-dom)[\\/]/,
        name: "react-vendor",
        chunks: "all",
        enforce: true,
        priority: 40,
      },
    },
  },
},
    resolve: {
  extensions: [".js", ".jsx", ".html"],
  // Ensure a single physical React + ReactDOM instance is used (prevents invalid hook call #321).
  alias: {
    react: path.resolve(__dirname, "node_modules/react"),
    "react-dom": path.resolve(__dirname, "node_modules/react-dom"),
  },
  // If anything is symlinked (e.g., linked packages), don't create duplicate module instances.
  symlinks: false,
},
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          use: {
            loader: "babel-loader",
          },
          exclude: /node_modules/,
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        scriptLoading: "defer",
        // Iteration 19: Inject taskpane bundles so the dev "Open JumpTo" button works.
        inject: true,
        chunks: ["polyfill", "taskpane", "commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
          { from: "src/index.html", to: "index.html" },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        scriptLoading: "defer",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/dialog/dialog.html",
        scriptLoading: "defer",
        chunks: ["polyfill", "dialog"],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};