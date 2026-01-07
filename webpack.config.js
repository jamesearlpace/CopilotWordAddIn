const path = require("path");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

// Get Office Add-in dev certificates
const certPath = path.join(process.env.USERPROFILE || process.env.HOME, ".office-addin-dev-certs");
const httpsOptions = fs.existsSync(path.join(certPath, "localhost.crt")) ? {
  key: fs.readFileSync(path.join(certPath, "localhost.key")),
  cert: fs.readFileSync(path.join(certPath, "localhost.crt")),
  ca: fs.readFileSync(path.join(certPath, "ca.crt")),
} : true;

module.exports = {
  entry: {
    taskpane: "./src/taskpane/taskpane.ts",
    commands: "./src/commands/commands.ts",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/taskpane/taskpane.html",
      filename: "taskpane.html",
      chunks: ["taskpane"],
    }),
    new HtmlWebpackPlugin({
      template: "./src/commands/commands.html",
      filename: "commands.html",
      chunks: ["commands"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets", to: "assets" },
      ],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    port: 3000,
    server: {
      type: "https",
      options: httpsOptions,
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    hot: true,
  },
};
