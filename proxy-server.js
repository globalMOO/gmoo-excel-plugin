// Lightweight CORS proxy for development.
// Forwards requests from https://localhost:3001/api/* to https://app.globalmoo.com/api/*
// and adds CORS headers so the Office add-in WebView can reach the API.

const https = require("https");
const fs = require("fs");
const path = require("path");

const PROXY_PORT = 3001;
const API_TARGET = "app.globalmoo.com";

// Use the same dev certs that the webpack dev server uses
const certDir = path.join(
  process.env.HOME || process.env.USERPROFILE,
  ".office-addin-dev-certs"
);

const serverOptions = {
  key: fs.readFileSync(path.join(certDir, "localhost.key")),
  cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
};

const server = https.createServer(serverOptions, (req, res) => {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
      "Access-Control-Allow-Headers": "Authorization, Content-Type, Accept",
      "Access-Control-Max-Age": "86400",
    });
    res.end();
    return;
  }

  // Only proxy /api/* requests
  if (!req.url.startsWith("/api")) {
    res.writeHead(404, { "Content-Type": "text/plain" });
    res.end("Not found. This proxy only handles /api/* requests.");
    return;
  }

  console.log(`[Proxy] ${req.method} ${req.url} -> https://${API_TARGET}${req.url}`);

  // Collect request body
  const bodyChunks = [];
  req.on("data", (chunk) => bodyChunks.push(chunk));
  req.on("end", () => {
    const body = Buffer.concat(bodyChunks);

    const proxyOptions = {
      hostname: API_TARGET,
      port: 443,
      path: req.url,
      method: req.method,
      headers: {
        ...req.headers,
        host: API_TARGET,
      },
    };

    // Remove headers that shouldn't be forwarded
    delete proxyOptions.headers["origin"];
    delete proxyOptions.headers["referer"];

    const proxyReq = https.request(proxyOptions, (proxyRes) => {
      console.log(`[Proxy] ${req.method} ${req.url} <- ${proxyRes.statusCode}`);

      // Add CORS headers to the response
      const responseHeaders = { ...proxyRes.headers };
      responseHeaders["access-control-allow-origin"] = "*";
      responseHeaders["access-control-allow-methods"] = "GET, POST, PUT, DELETE, OPTIONS";
      responseHeaders["access-control-allow-headers"] = "Authorization, Content-Type, Accept";

      res.writeHead(proxyRes.statusCode, responseHeaders);
      proxyRes.pipe(res);
    });

    proxyReq.on("error", (err) => {
      console.error(`[Proxy] Error:`, err.message);
      res.writeHead(502, {
        "Content-Type": "text/plain",
        "Access-Control-Allow-Origin": "*",
      });
      res.end(`Proxy error: ${err.message}`);
    });

    if (body.length > 0) {
      proxyReq.write(body);
    }
    proxyReq.end();
  });
});

server.listen(PROXY_PORT, () => {
  console.log(`[Proxy] CORS proxy running at https://localhost:${PROXY_PORT}`);
  console.log(`[Proxy] Forwarding /api/* -> https://${API_TARGET}/api/*`);
});
