# MAP XML 转 SVG 工具

该工具用于将 RSU 中的 C-V2X MAP XML 转换为物理路口 SVG 可视化，便于在无锡车路云项目管理平台中核查车道、停止线、人行横道、待行区、PhaseID 和车道下游关系。

## 使用方式

```bash
python3 tools/map-xml-to-svg/scripts/map_xml_to_physical_svg.py \
  "/path/to/map.xml" \
  -o "tools/map-xml-to-svg/examples/output.svg" \
  --size 1200 \
  --radius-m 82 \
  --rotation-deg 0
```

示例输出：

```text
tools/map-xml-to-svg/examples/MAP_24158_惠育路-政和大道-physical.svg
```

## 渲染原则

- `Lane.points` 是车道中心线，只用于构造车道面和放置箭头，不直接绘制。
- 车道由 `Lane.points + laneWidth` 生成 polygon。
- 停止线宽度等于对应车道宽度。
- 人行横道图层放在车道线后面，并使用较淡透明度。
- `guidedLaneWidth / guidedLaneLength` 用于生成直行待行区或导向区。
- `connectsTo` 的 `connectingLane/lane` 只在当前 Link 内追踪，避免重复 laneID 跨进口串联。

## 后续接入建议

当前工具先作为离线转换脚本并入项目。后续可在 `prototype/server.js` 中增加接口，例如：

```text
POST /api/map-xml-to-svg
```

由前端上传 MAP XML，服务端调用该脚本生成 SVG，再在点位详情页中展示 SVG 图层。
