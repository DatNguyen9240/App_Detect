@echo off
chcp 65001 >nul
echo ========================================================
echo   KHỞI ĐỘNG SERVER TOOL KIỂM TRA LỖI BÁO CÔNG MES/ERP
echo ========================================================
start "" http://localhost:5173
node -e "const http=require('http'),fs=require('fs'),path=require('path');const mime={'.html':'text/html','.css':'text/css','.js':'application/javascript','.json':'application/json'};http.createServer((req,res)=>{let p=req.url.split('?')[0];if(p==='/')p='/index.html';let f=path.join(__dirname,p);if(fs.existsSync(f)){res.writeHead(200,{'Content-Type':mime[path.extname(f)]||'text/plain'});fs.createReadStream(f).pipe(res);}else{res.writeHead(404);res.end('Not found');}}).listen(5173,()=>console.log('Server dang chay tai: http://localhost:5173'));"
pause
