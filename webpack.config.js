const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

const urlDev = "https://localhost:3000/";
const urlProd = "https://globalmoo.github.io/gmoo-excel-plugin/";

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  let devServerConfig = {};
  if (dev) {
    const httpsOptions = await devCerts.getHttpsServerOptions();
    devServerConfig = {
      devServer: {
        headers: {
          "Access-Control-Allow-Origin": "*",
        },
        server: {
          type: "https",
          options: httpsOptions,
        },
        port: 3000,
        hot: true,
      },
    };
  }

  return {
    ...devServerConfig,
    entry: {
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
          type: "asset/resource",
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        inject: true,
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["commands"],
        inject: true,
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext]",
            noErrorOnMissing: true,
          },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              const url = dev ? urlDev : urlProd;
              return content.toString().replace(/https:\/\/localhost:3000\//g, url);
            },
          },
        ],
      }),
    ],
    devtool: dev ? "source-map" : false,
  };
};
