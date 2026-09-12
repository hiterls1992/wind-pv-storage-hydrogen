/**
 * 本地静态服务（开发/测试用）
 * ============================================================================
 * 用途：直接用 file:// 打开 index.html 时，浏览器会禁止创建 Web Worker
 *       （Script at 'file:///...' cannot be accessed from origin 'null'），
 *       优化会自动回退到主线程分片调度。若希望使用 Web Worker 后台线程，
 *       用本脚本起一个本地 HTTP 服务即可。
 *
 * 运行：node tests/dev-server.js  [端口，默认 8099]
 * 访问：http://127.0.0.1:8099/index.html
 *
 * 说明：仅监听 127.0.0.1，纯静态文件服务，无任何上传/写入能力。
 */

'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const ROOT = path.resolve(__dirname, '..');
const PORT = parseInt(process.argv[2], 10) || 8099;

const MIME = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.svg': 'image/svg+xml',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.txt': 'text/plain; charset=utf-8',
};

const server = http.createServer((req, res) => {
    let rel = decodeURIComponent(url.parse(req.url).pathname);
    if (rel === '/' || rel === '') rel = '/index.html';

    // 防目录穿越
    const target = path.normalize(path.join(ROOT, rel));
    if (target.indexOf(ROOT) !== 0) {
        res.writeHead(403); res.end('403 Forbidden'); return;
    }
    if (!fs.existsSync(target) || fs.statSync(target).isDirectory()) {
        res.writeHead(404); res.end('404 Not Found'); return;
    }

    res.writeHead(200, {
        'Content-Type': MIME[path.extname(target).toLowerCase()] || 'application/octet-stream',
        'Cache-Control': 'no-store',
    });
    fs.createReadStream(target).pipe(res);
});

server.listen(PORT, '127.0.0.1', () => {
    process.stdout.write('多能互补 WEB-V2.1 本地服务已启动：http://127.0.0.1:' + PORT + '/index.html\n');
});
