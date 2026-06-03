# 无锡车路云服务修复 APP

当浏览器访问 `http://127.0.0.1:4173/` 出现“拒绝连接”时，双击：

```text
tools/WuxiRoadsideServiceRepair.app
```

它会自动执行：

1. 检查 `http://127.0.0.1:4173/api/health`
2. 如果服务未运行，重新安装并启动后台服务
3. 如果后台服务仍未起来，用兜底方式直接拉起 Node 服务
4. 服务恢复后自动打开应用主页

日志位置：

```text
~/Library/Application Support/wuxi-roadside-prototype/logs/repair-app.log
```
