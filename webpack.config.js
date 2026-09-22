/* eslint-disable no-undef */

const fs = require("fs");
const path = require("path");
const webpack = require("webpack");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3001/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

/* global require, module, process */

function parseEnvFile(envPath) {
  const env = {};
  if (!fs.existsSync(envPath)) {
    return env;
  }

  fs.readFileSync(envPath, "utf8")
    .split(/\r?\n/)
    .forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) {
        return;
      }

      const separator = trimmed.indexOf("=");
      if (separator === -1) {
        return;
      }

      const key = trimmed.slice(0, separator).trim();
      let value = trimmed.slice(separator + 1).trim();
      if (
        (value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      }
      env[key] = value;
    });

  return env;
}

/**
 * npm run build (production) → .env
 * npm start / build:dev (development) → .env.local (falls back to .env)
 */
function loadEnvFile(isDev) {
  const prodPath = path.resolve(__dirname, ".env");
  const localPath = path.resolve(__dirname, ".env.local");

  if (isDev) {
    if (fs.existsSync(localPath)) {
      console.log("[MethodTech] Loaded env from .env.local (local/dev)");
      return parseEnvFile(localPath);
    }
    console.warn("[MethodTech] .env.local not found — falling back to .env for local/dev");
    return parseEnvFile(prodPath);
  }

  console.log("[MethodTech] Loaded env from .env (production build)");
  return parseEnvFile(prodPath);
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const fileEnv = loadEnvFile(dev);
  // Real Django host (used as webpack proxy target)
  const djangoProxyTarget = (process.env.DJANGO_BASE_URL || fileEnv.DJANGO_BASE_URL || "").replace(/\/$/, "");
  // In the Excel WebView (HTTPS), call same-origin so mixed-content HTTP is not blocked.
  // webpack-dev-server proxies /api and /v1 to djangoProxyTarget.
  const djangoBaseUrlForBrowser = dev ? "" : djangoProxyTarget;
  const microsoftAppId = process.env.MICROSOFT_APP_ID || fileEnv.MICROSOFT_APP_ID || "";
  const microsoftTenantId = process.env.MICROSOFT_TENANT_ID || fileEnv.MICROSOFT_TENANT_ID || "common";
  const microsoftApiScope =
    process.env.MICROSOFT_API_SCOPE ||
    fileEnv.MICROSOFT_API_SCOPE ||
    "api://612da946-27be-427c-b997-2279839c4e56/access_as_user";

  const replaceManifestPlaceholders = (content) => {
    let text = content.toString();
    if (!dev) {
      text = text.replace(urlDev, urlProd);
    }
    return text;
  };

  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
      functions: "./src/functions/functions.js",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
      alias: {
        // webpack-dev-server requests "process/browser" (no extension); ESM requires .js
        "process/browser": require.resolve("process/browser.js"),
      },
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        "process.env.DJANGO_BASE_URL": JSON.stringify(djangoBaseUrlForBrowser),
        "process.env.MICROSOFT_APP_ID": JSON.stringify(microsoftAppId),
        "process.env.MICROSOFT_TENANT_ID": JSON.stringify(microsoftTenantId),
        "process.env.MICROSOFT_API_SCOPE": JSON.stringify(microsoftApiScope),
      }),
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.js",
      }),
      // CustomFunctionsMetadataPlugin writes an empty file when there are no JSDoc
      // functions. Excel rejects a 0-byte functions.json ("failed to download a
      // required resource"). Merge cached catalog metadata and always emit valid JSON.
      {
        apply(compiler) {
          compiler.hooks.afterEmit.tap("MethodTechEnsureFunctionsJson", () => {
            const catalogMetaPath = path.resolve(__dirname, ".methodtech", "catalog-functions.json");
            const distFunctionsPath = path.resolve(__dirname, "dist", "functions.json");
            let base = { allowCustomDataForDataTypeAny: true, functions: [] };
            try {
              if (fs.existsSync(distFunctionsPath)) {
                const raw = fs.readFileSync(distFunctionsPath, "utf8").trim();
                if (raw) base = JSON.parse(raw);
              }
            } catch (error) {
              console.warn("[MethodTech] Could not read emitted functions.json:", error.message);
            }
            let catalogFns = [];
            try {
              if (fs.existsSync(catalogMetaPath)) {
                catalogFns = JSON.parse(fs.readFileSync(catalogMetaPath, "utf8"));
              }
            } catch (error) {
              console.warn("[MethodTech] Could not read catalog-functions.json:", error.message);
            }
            const byId = new Map();
            (base.functions || []).forEach((fn) => {
              if (fn && fn.id) byId.set(String(fn.id).toUpperCase(), fn);
            });
            (catalogFns || []).forEach((fn) => {
              if (fn && fn.id) byId.set(String(fn.id).toUpperCase(), fn);
            });
            const merged = {
              allowCustomDataForDataTypeAny: true,
              functions: Array.from(byId.values()),
            };
            fs.mkdirSync(path.dirname(distFunctionsPath), { recursive: true });
            fs.writeFileSync(distFunctionsPath, JSON.stringify(merged, null, 2));
            console.log(
              `[MethodTech] Ensured functions.json (${merged.functions.length} function(s), ${catalogFns.length} from catalog)`
            );
          });
        },
      },
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "functions", "commands"],
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
              return replaceManifestPlaceholders(content);
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      proxy: djangoProxyTarget
        ? [
            {
              context: ["/api", "/v1"],
              target: djangoProxyTarget,
              changeOrigin: true,
              secure: false,
              logLevel: "debug",
            },
          ]
        : undefined,
      setupMiddlewares: (middlewares, devServer) => {
        if (!devServer) {
          throw new Error("webpack-dev-server is not defined");
        }

        const catalogMetaPath = path.resolve(__dirname, ".methodtech", "catalog-functions.json");
        const distFunctionsPath = path.resolve(__dirname, "dist", "functions.json");
        let catalogReadyWaiters = [];

        const readCatalogMeta = () => {
          try {
            if (fs.existsSync(catalogMetaPath)) {
              return JSON.parse(fs.readFileSync(catalogMetaPath, "utf8"));
            }
          } catch (error) {
            console.warn("[MethodTech] Could not read catalog metadata cache:", error.message);
          }
          return [];
        };

        const readBaseFunctions = () => {
          try {
            if (fs.existsSync(distFunctionsPath)) {
              return JSON.parse(fs.readFileSync(distFunctionsPath, "utf8"));
            }
          } catch (error) {
            console.warn("[MethodTech] Could not read base functions.json:", error.message);
          }
          return { allowCustomDataForDataTypeAny: true, functions: [] };
        };

        const mergeFunctionsJson = (baseObj, catalogFns) => {
          const byId = new Map();
          (baseObj.functions || []).forEach((fn) => {
            if (fn && fn.id) {
              byId.set(String(fn.id).toUpperCase(), fn);
            }
          });
          (catalogFns || []).forEach((fn) => {
            if (fn && fn.id) {
              byId.set(String(fn.id).toUpperCase(), fn);
            }
          });
          return {
            allowCustomDataForDataTypeAny: true,
            functions: Array.from(byId.values()),
          };
        };

        const writeMergedFunctionsJson = (catalogFns) => {
          const merged = mergeFunctionsJson(readBaseFunctions(), catalogFns);
          try {
            fs.mkdirSync(path.dirname(distFunctionsPath), { recursive: true });
            fs.writeFileSync(distFunctionsPath, JSON.stringify(merged, null, 2));
          } catch (error) {
            console.warn("[MethodTech] Could not write merged functions.json:", error.message);
          }
          return merged;
        };

        const notifyCatalogReady = () => {
          const waiters = catalogReadyWaiters;
          catalogReadyWaiters = [];
          waiters.forEach((resolve) => resolve());
        };

        // Excel asks for functions.json at startup. Hold that response until the
        // catalog API succeeds and publishes metadata — then suggestions are complete.
        const waitForCatalogMetadata = (timeoutMs = 90000) => {
          if (readCatalogMeta().length > 0) {
            return Promise.resolve(true);
          }
          return new Promise((resolve) => {
            const timer = setTimeout(() => {
              catalogReadyWaiters = catalogReadyWaiters.filter((item) => item !== onReady);
              console.warn("[MethodTech] Timed out waiting for catalog metadata");
              resolve(false);
            }, timeoutMs);
            const onReady = () => {
              clearTimeout(timer);
              resolve(true);
            };
            catalogReadyWaiters.push(onReady);
          });
        };

        const sendMergedFunctionsJson = (res) => {
          const catalogFns = readCatalogMeta();
          const merged = mergeFunctionsJson(readBaseFunctions(), catalogFns);
          res.setHeader("Content-Type", "application/json; charset=utf-8");
          res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate");
          res.setHeader("Pragma", "no-cache");
          res.statusCode = 200;
          res.end(JSON.stringify(merged));
          console.log(
            `[MethodTech] Served functions.json with ${merged.functions.length} function(s) (${catalogFns.length} from catalog)`
          );
        };

        // Publish catalog → Excel metadata (unblocks waiting functions.json)
        devServer.app.post("/__methodtech/catalog-metadata", (req, res) => {
          let raw = "";
          req.on("data", (chunk) => {
            raw += chunk;
          });
          req.on("end", () => {
            try {
              const body = raw ? JSON.parse(raw) : {};
              const functions = Array.isArray(body.functions) ? body.functions : [];
              fs.mkdirSync(path.dirname(catalogMetaPath), { recursive: true });
              fs.writeFileSync(catalogMetaPath, JSON.stringify(functions, null, 2));
              writeMergedFunctionsJson(functions);
              console.log(`[MethodTech] Catalog API metadata saved: ${functions.length} function(s) — suggestions ready`);
              notifyCatalogReady();
              res.setHeader("Content-Type", "application/json");
              res.end(JSON.stringify({ ok: true, count: functions.length }));
            } catch (error) {
              res.statusCode = 400;
              res.setHeader("Content-Type", "application/json");
              res.end(JSON.stringify({ ok: false, error: error.message }));
            }
          });
        });

        // Prefer Express route so we reliably intercept Excel's metadata request
        devServer.app.get("/functions.json", (req, res) => {
          console.log("[MethodTech] Excel requested functions.json — waiting for catalog if needed…");
          waitForCatalogMetadata(90000)
            .then(() => sendMergedFunctionsJson(res))
            .catch((error) => {
              console.error("[MethodTech] functions.json wait failed:", error);
              sendMergedFunctionsJson(res);
            });
        });

        // Also keep middleware as a fallback
        middlewares.unshift({
          name: "methodtech-merged-functions-json",
          middleware: (req, res, next) => {
            const url = req.url || "";
            if (req.method === "GET" && (url === "/functions.json" || url.startsWith("/functions.json?"))) {
              // Handled by app.get above when registered; if we get here, still wait.
              waitForCatalogMetadata(90000)
                .then(() => sendMergedFunctionsJson(res))
                .catch(next);
              return;
            }
            next();
          },
        });

        return middlewares;
      },
    },
  };

  if (dev && djangoProxyTarget) {
    console.log(`[MethodTech] Proxying /api and /v1 -> ${djangoProxyTarget}`);
  }

  // If catalog was saved from a previous Excel session, merge it now so the
  // NEXT Excel launch already has full =METHODTECH. IntelliSense on first load.
  const catalogMetaPath = path.resolve(__dirname, ".methodtech", "catalog-functions.json");
  const distFunctionsPath = path.resolve(__dirname, "dist", "functions.json");
  try {
    if (fs.existsSync(catalogMetaPath)) {
      const catalogFns = JSON.parse(fs.readFileSync(catalogMetaPath, "utf8"));
      let base = { allowCustomDataForDataTypeAny: true, functions: [] };
      if (fs.existsSync(distFunctionsPath)) {
        base = JSON.parse(fs.readFileSync(distFunctionsPath, "utf8"));
      }
      const byId = new Map();
      (base.functions || []).forEach((fn) => {
        if (fn && fn.id) byId.set(String(fn.id).toUpperCase(), fn);
      });
      (catalogFns || []).forEach((fn) => {
        if (fn && fn.id) byId.set(String(fn.id).toUpperCase(), fn);
      });
      const merged = {
        allowCustomDataForDataTypeAny: true,
        functions: Array.from(byId.values()),
      };
      fs.mkdirSync(path.dirname(distFunctionsPath), { recursive: true });
      fs.writeFileSync(distFunctionsPath, JSON.stringify(merged, null, 2));
      console.log(
        `[MethodTech] Preloaded ${catalogFns.length} catalog function(s) into functions.json for Excel suggestions`
      );
    }
  } catch (error) {
    console.warn("[MethodTech] Could not preload catalog metadata:", error.message);
  }

  return config;
};
