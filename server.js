const https = require("https");
const http = require("http");

const CLIENT_ID = process.env.CLIENT_ID;
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_ID = process.env.SITE_ID;
const LIST_ID = "247ff344-20d5-4433-bdfa-fde5fd8b8b23";

function getToken(cb) {
  const body = `grant_type=client_credentials&client_id=${CLIENT_ID}&client_secret=${encodeURIComponent(CLIENT_SECRET)}&scope=https://graph.microsoft.com/.default`;
  const req = https.request({ hostname: "login.microsoftonline.com", path: `/${TENANT_ID}/oauth2/v2.0/token`, method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" } }, res => {
    let d = ""; res.on("data", c => d += c); res.on("end", () => cb(JSON.parse(d).access_token));
  });
  req.write(body); req.end();
}

function postItem(token, fields, cb) {
  const body = JSON.stringify({ fields });
  const req = https.request({ hostname: "graph.microsoft.com", path: `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`, method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json", "Content-Length": Buffer.byteLength(body) } }, res => {
    let d = ""; res.on("data", c => d += c); res.on("end", () => cb(res.statusCode, d));
  });
  req.write(body); req.end();
}

http.createServer((req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") { res.writeHead(204); res.end(); return; }
  if (req.method === "POST" && req.url === "/submit") {
    let body = ""; req.on("data", c => body += c);
    req.on("end", () => {
      const items = JSON.parse(body);
      getToken(token => {
        let ok = 0, fail = 0, done = 0;
        items.forEach(fields => {
          postItem(token, fields, (status) => {
            if (status === 201) ok++; else fail++;
            done++;
            if (done === items.length) {
              res.writeHead(200, { "Content-Type": "application/json" });
              res.end(JSON.stringify({ ok, fail }));
            }
          });
        });
      });
    });
  } else { res.writeHead(404); res.end(); }
}).listen(process.env.PORT || 3000);
