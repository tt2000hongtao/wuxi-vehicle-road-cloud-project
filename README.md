# 无锡车路云项目管理平台

一期原型与设计文档仓库，覆盖点位管理、设备管理、合同管理、导入治理、坐标异常、一点一档和资料管理。

## 核心文档

- [完整 PRD](./docs/无锡车路云项目管理平台-完整PRD.md)
- [项目审计与优化报告](./docs/项目审计与优化报告-2026-06-11.md)
- [文档索引](./docs/文档索引.md)

## 本地预览

直接打开 `prototype/index.html`，或在 `prototype` 目录启动静态服务后访问页面。

## 本地服务与数据持久化

运维管理页的路侧设备运行状态历史数据支持保存到本地硬盘文件。启动本地 Node.js 服务后访问原型：

```bash
node prototype/server.js
```

访问地址：

```text
http://127.0.0.1:4173
```

服务会自动托管 `prototype/` 静态文件，并将路侧设备运行状态保存到：

```text
prototype/storage/roadside-status-state.json
```

如果仍直接打开 `prototype/index.html`，页面会继续使用浏览器 localStorage 作为兜底缓存，但不会自动写入本地硬盘 JSON 文件。

### 安装为本机常驻服务

为避免每次手动启动服务，在 macOS 上执行一次：

```bash
bash prototype/install-local-service.sh
```

安装后服务会在当前用户登录时自动启动，并保持运行。页面请始终访问：

```text
http://127.0.0.1:4173
```

如需移除自启动服务：

```bash
bash prototype/uninstall-local-service.sh
```
