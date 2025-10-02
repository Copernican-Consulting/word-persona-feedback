const path = require("path");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CERT_DIR = path.join(process.env.USERPROFILE || process.env.HOME || "", ".office-addin-dev-certs");
const CERT_KEY = path.join(CERT_DIR, "localhost.key");
const CERT_CRT = path.join(CERT_DIR, "localhost.crt");
function readIfExists(p) { try { return fs.readFileSync(p); } catch { return undefined; } }
const key = readIfExists(CERT_KEY);
const cert = readIfExists(CERT_CRT);
module.exports = {
  mode: "development",
  devtool: "inline-source-map",
  entry: { taskpane: "./src/taskpane.ts" },
  output: { filename: "[name].bundle.js", path: path.resolve(__dirname, "dist"), publicPath: "/", clean: true },
  resolve: { extensions: [".ts", ".tsx", ".js"] },
  module: {
    rules: [
      { test: /\.tsx?$/, use: "ts-loader", exclude: /node_modules/ },
      { test: /\.css$/i, use: ["style-loader", "css-loader"] },
      { test: /\.(png|jpe?g|gif|svg)$/i, type: "asset/resource", generator: { filename: "assets/[name][ext]" } },
      { test: /\.html$/i, loader: "html-loader", options: { sources: false } }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({ template: "./src/taskpane.html", filename: "taskpane.html", chunks: ["taskpane"], cache: false }),
    new CopyWebpackPlugin({ patterns: [ { from: "src/assets", to: "assets", noErrorOnMissing: true } ] })
  ],
  devServer: {
    host: "0.0.0.0",
    port: 3000,
    allowedHosts: "all",
    server: key && cert ? { type: "https", options: { key, cert } } : { type: "https" },
    static: { directory: path.resolve(__dirname, "dist"), publicPath: "/", watch: true },
    hot: false,
    liveReload: true,
    headers: { "Access-Control-Allow-Origin": "*" },
    client: { overlay: true, logging: "info" }
  }
};