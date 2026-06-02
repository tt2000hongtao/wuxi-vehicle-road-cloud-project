#!/usr/bin/env python3
"""Render C-V2X MAP XML to a physical-world SVG.

SVG is the preferred vector visualization artifact for web integration:
- MAP lane centerlines are not rendered directly.
- Lane polygons are constructed from centerline + laneWidth.
- Road arrows are placed on the lane centerline at fixed distance from stopLine.
- Physical elements only: asphalt, lane boundaries, stop lines, crosswalk zebra,
  waiting/guide lanes, road arrows, scale/north aids.
"""
from __future__ import annotations

import argparse
import html
import math
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

BG = "#262a2b"
ASPHALT = "#444846"
ASPHALT_DARK = "#363938"
LANE_WHITE = "#f0f2f0"
GUIDE_GREEN = "#00ff2d"
CROSSWALK_WHITE = "#f7f7f7"

M_PER_LAT_UNIT = 111_320 / 1e7


def text(el: Optional[ET.Element], path: str, default: str = "") -> str:
    if el is None:
        return default
    v = el.findtext(path)
    return v.strip() if v and v.strip() else default


def intval(v, default=0) -> int:
    try:
        return int(str(v).strip())
    except Exception:
        return default


@dataclass
class StopLine:
    point_m: tuple[float, float]


@dataclass
class Lane:
    lane_id: str
    width_m: float
    type_tag: str
    type_bits: str
    veh_ext: str
    maneuver: str
    phase_id: str
    connection_remote: str
    connection_lane: str
    points_m: list[tuple[float, float]]
    stop_lines: list[StopLine] = field(default_factory=list)
    guided_width_m: float = 0.0
    guided_length_m: float = 0.0

    @property
    def is_crosswalk(self) -> bool:
        return self.type_tag == "crosswalk"

    @property
    def is_guide_or_waiting(self) -> bool:
        return bool(self.veh_ext) or (intval(self.lane_id) >= 240 and not self.is_crosswalk)


@dataclass
class Link:
    name: str
    upstream: str
    width_m: float
    points_m: list[tuple[float, float]]
    lanes: list[Lane]


@dataclass
class MapData:
    name: str
    node_id: str
    region: str
    ref_lat: int
    ref_lon: int
    links: list[Link]


def converters(ref_lat: int, ref_lon: int):
    lat_deg = ref_lat / 1e7
    m_per_lon_unit = 111_320 * math.cos(math.radians(lat_deg)) / 1e7

    def offset_to_m(lon_off: int, lat_off: int) -> tuple[float, float]:
        return lon_off * m_per_lon_unit, lat_off * M_PER_LAT_UNIT

    def abs_to_m(lon: int, lat: int) -> tuple[float, float]:
        return (lon - ref_lon) * m_per_lon_unit, (lat - ref_lat) * M_PER_LAT_UNIT

    return offset_to_m, abs_to_m


def road_points(parent: Optional[ET.Element], conv) -> list[tuple[float, float]]:
    if parent is None:
        return []
    pts = []
    for rp in parent.findall("./RoadPoint"):
        lon = rp.findtext("./posOffset/offsetLL/position-LL4/lon")
        lat = rp.findtext("./posOffset/offsetLL/position-LL4/lat")
        if lon is not None and lat is not None:
            pts.append(conv(intval(lon), intval(lat)))
    return pts


def parse_map(xml_path: Path) -> MapData:
    root = ET.parse(xml_path).getroot()
    node = root.find(".//Node")
    if node is None:
        raise ValueError("MAP XML does not contain Node")
    ref_lat = intval(text(node, "refPos/lat"))
    ref_lon = intval(text(node, "refPos/long"))
    off_to_m, abs_to_m = converters(ref_lat, ref_lon)
    links: list[Link] = []
    for link_el in node.findall("./inLinks/Link"):
        lanes: list[Lane] = []
        for lane_el in link_el.findall("./lanes/Lane"):
            lane_type = lane_el.find("./laneAttributes/laneType")
            type_tag, type_bits = "unknown", ""
            if lane_type is not None and len(lane_type):
                child = list(lane_type)[0]
                type_tag = child.tag
                type_bits = (child.text or "").strip()
            phase_id = ""
            connection_remote = ""
            connection_lane = ""
            for ce in lane_el.findall("./connectsTo/Connection"):
                if not connection_remote:
                    connection_remote = text(ce, "remoteIntersection/id")
                    connection_lane = text(ce, "connectingLane/lane")
                ph = text(ce, "phaseId")
                if ph and not phase_id:
                    phase_id = ph
            stops = []
            for sl in lane_el.findall("./stopLines/StopLine"):
                cp = sl.find("./centerPoint")
                if cp is not None and cp.findtext("lat") and cp.findtext("long"):
                    stops.append(StopLine(abs_to_m(intval(cp.findtext("long")), intval(cp.findtext("lat")))))
            lanes.append(Lane(
                lane_id=text(lane_el, "laneID"),
                width_m=max(1.8, intval(text(lane_el, "laneWidth")) / 100.0),
                type_tag=type_tag,
                type_bits=type_bits,
                veh_ext=text(lane_el, "laneTypeAttrVehExt"),
                maneuver=text(lane_el, "maneuvers"),
                phase_id=phase_id,
                connection_remote=connection_remote,
                connection_lane=connection_lane,
                points_m=road_points(lane_el.find("./points"), off_to_m),
                stop_lines=stops,
                guided_width_m=intval(text(lane_el, "guidedLaneWidth")) / 100.0,
                guided_length_m=intval(text(lane_el, "guidedLaneLength")) / 100.0,
            ))
        links.append(Link(
            name=text(link_el, "name"),
            upstream=text(link_el, "upstreamNodeId/id"),
            width_m=max(5.0, intval(text(link_el, "linkWidth")) / 100.0),
            points_m=road_points(link_el.find("./points"), off_to_m),
            lanes=lanes,
        ))
    return MapData(text(node, "name"), text(node, "id/id"), text(node, "id/region"), ref_lat, ref_lon, links)


class Transform:
    def __init__(self, pts: Iterable[tuple[float, float]], size: int, margin: int, rotation_deg: float, radius_m: Optional[float]):
        self.size = size
        self.margin = margin
        self.rot = math.radians(rotation_deg)
        if radius_m and radius_m > 0:
            self.center = (0.0, 0.0)
            self.scale = (size - 2 * margin) / (2 * radius_m)
        else:
            rpts = [self._rot(p) for p in pts]
            xs, ys = [p[0] for p in rpts], [p[1] for p in rpts]
            minx, maxx = min(xs), max(xs)
            miny, maxy = min(ys), max(ys)
            self.center = ((minx + maxx) / 2, (miny + maxy) / 2)
            self.scale = min((size - 2 * margin) / max(1, maxx - minx), (size - 2 * margin) / max(1, maxy - miny))

    def _rot(self, p):
        x, y = p
        return x * math.cos(self.rot) - y * math.sin(self.rot), x * math.sin(self.rot) + y * math.cos(self.rot)

    def __call__(self, p):
        x, y = self._rot(p)
        cx, cy = self.center
        return self.size / 2 + (x - cx) * self.scale, self.size / 2 - (y - cy) * self.scale

    def px(self, meters: float) -> float:
        return max(1.0, meters * self.scale)


def unit(v):
    x, y = v
    L = math.hypot(x, y)
    return (0.0, 0.0) if L <= 1e-9 else (x / L, y / L)


def offset_polyline(points: list[tuple[float, float]], offset: float) -> list[tuple[float, float]]:
    if len(points) < 2:
        return points[:]
    out = []
    n = len(points)
    for i, p in enumerate(points):
        if i == 0:
            tan = unit((points[1][0]-p[0], points[1][1]-p[1]))
        elif i == n - 1:
            tan = unit((p[0]-points[i-1][0], p[1]-points[i-1][1]))
        else:
            t1 = unit((p[0]-points[i-1][0], p[1]-points[i-1][1]))
            t2 = unit((points[i+1][0]-p[0], points[i+1][1]-p[1]))
            tan = unit((t1[0]+t2[0], t1[1]+t2[1])) or t2
        nx, ny = -tan[1], tan[0]
        out.append((p[0] + nx * offset, p[1] + ny * offset))
    return out


def lane_polygon_m(lane: Lane) -> list[tuple[float, float]]:
    if len(lane.points_m) < 2:
        return []
    half = max(lane.width_m / 2.0, 0.75)
    left = offset_polyline(lane.points_m, half)
    right = offset_polyline(lane.points_m, -half)
    return left + list(reversed(right))


def guided_lane_polygon_m(lane: Lane) -> list[tuple[float, float]]:
    """Construct visible waiting/guide lane area from guidedLaneWidth/Length.

    The MAP lane centerline ends near the stop line. guidedLaneWidth and
    guidedLaneLength describe a downstream guide/waiting area inside the
    intersection. This is why the north-south straight waiting areas were not
    visible when only laneTypeAttrVehExt was used.
    """
    if len(lane.points_m) < 2 or lane.guided_width_m <= 0 or lane.guided_length_m <= 0:
        return []
    hdg = heading(lane.points_m)
    start = lane.stop_lines[0].point_m if lane.stop_lines else lane.points_m[-1]
    end = (start[0] + math.cos(hdg) * lane.guided_length_m, start[1] + math.sin(hdg) * lane.guided_length_m)
    center = [start, end]
    half = max(lane.guided_width_m / 2.0, 0.75)
    left = offset_polyline(center, half)
    right = offset_polyline(center, -half)
    return left + list(reversed(right))


def heading(points):
    if len(points) < 2:
        return 0.0
    a, b = points[-2], points[-1]
    return math.atan2(b[1]-a[1], b[0]-a[0])


def screen_angle(points, t: Transform):
    if len(points) < 2:
        return 0.0
    a, b = t(points[-2]), t(points[-1])
    return math.atan2(b[1]-a[1], b[0]-a[0])


def point_along_polyline(points: list[tuple[float, float]], fraction: float) -> tuple[float, float]:
    """Return a point along a polyline at 0..1 fraction of arc length."""
    if not points:
        return (0.0, 0.0)
    if len(points) == 1:
        return points[0]
    segs=[]; total=0.0
    for a,b in zip(points, points[1:]):
        L=math.hypot(b[0]-a[0], b[1]-a[1]); segs.append((a,b,L)); total+=L
    target=max(0.0,min(1.0,fraction))*total
    acc=0.0
    for a,b,L in segs:
        if acc+L>=target and L>1e-9:
            u=(target-acc)/L
            return (a[0]+(b[0]-a[0])*u, a[1]+(b[1]-a[1])*u)
        acc+=L
    return points[-1]


def point_before_polyline_end(points: list[tuple[float, float]], distance_m: float) -> tuple[float, float]:
    """Return point a fixed distance upstream from the polyline end."""
    if not points:
        return (0.0, 0.0)
    if len(points) == 1:
        return points[0]
    remaining = max(0.0, distance_m)
    for i in range(len(points)-1, 0, -1):
        a = points[i-1]
        b = points[i]
        L = math.hypot(b[0]-a[0], b[1]-a[1])
        if L >= remaining and L > 1e-9:
            # move from b back toward a by remaining
            u = remaining / L
            return (b[0] + (a[0]-b[0]) * u, b[1] + (a[1]-b[1]) * u)
        remaining -= L
    return points[0]


def phase_label_position(lane: Lane, t: Transform) -> tuple[tuple[float, float], float]:
    """Place PhaseID inside the lane, near the stop-line side like the sample PNG."""
    if len(lane.points_m) < 2:
        return ((0.0,0.0),0.0)
    if lane.stop_lines:
        hdg=heading(lane.points_m)
        sx,sy=lane.stop_lines[0].point_m
        p=(sx - math.cos(hdg)*2.2, sy - math.sin(hdg)*2.2)
    else:
        p=point_along_polyline(lane.points_m, 0.82)
    return t(p), screen_angle(lane.points_m,t)


def maneuver_kind(bits: str) -> str:
    b = (bits or "")[:3]
    if b in ("110", "111"):
        return "left_straight"
    if b == "101":
        return "right_straight"
    if b.startswith("010"):
        return "left"
    if b.startswith("001"):
        return "right"
    return "straight"


def is_early_right_turn_lane(lane: Lane) -> bool:
    """Heuristic for slip/advance right-turn lanes in this MAP XML family.

    These lanes are physically separated curved right-turn lanes. Their arrows
    should be sampled inside the slip lane itself instead of being positioned
    from a stop line at the intersection throat.
    """
    return maneuver_kind(lane.maneuver) == "right" and not lane.is_crosswalk and not lane.is_guide_or_waiting


def arrow_center(lane: Lane, t: Transform, distance_m=7.0):
    if len(lane.points_m) < 2:
        return None
    # Advance/slip right-turn lanes: place arrow inside the curved right-turn
    # lane itself, not by projecting back from the stop line.
    if is_early_right_turn_lane(lane):
        # Use a point near the visible/intersection end of the slip lane.
        # A fixed upstream distance avoids putting arrows far outside the view
        # on long advance-right lane centerlines.
        return t(point_before_polyline_end(lane.points_m, 9.0))
    hdg = heading(lane.points_m)
    if lane.stop_lines:
        sx, sy = lane.stop_lines[0].point_m
        return t((sx - math.cos(hdg) * distance_m, sy - math.sin(hdg) * distance_m))
    pts = [t(p) for p in lane.points_m]
    p1, p2 = pts[max(0, len(pts)-3)], pts[-1]
    return ((p1[0]*0.55+p2[0]*0.45), (p1[1]*0.55+p2[1]*0.45))


def downstream_node_for_lane(lane: Lane, lane_by_id: dict[str, Lane], current_node_id: str, depth: int = 0) -> str:
    """Resolve downstream NodeID in the lane travel direction.

    `connectingLane/lane` is only unique within the current Link context in this
    XML family. Do not resolve it through a global laneID map, because lane IDs
    like 1, 2, 3, 240, and 241 repeat under every inLink.

    If a lane first connects to an internal lane within current NodeID (common
    for waiting/guide lanes 240/241), follow that connecting lane within the
    same Link until the remoteIntersection is outside current node.
    """
    if depth > 4:
        return lane.connection_remote or current_node_id
    remote = lane.connection_remote or current_node_id
    if remote != current_node_id:
        return remote
    if lane.connection_lane and lane.connection_lane in lane_by_id:
        nxt = lane_by_id[lane.connection_lane]
        if nxt is not lane:
            return downstream_node_for_lane(nxt, lane_by_id, current_node_id, depth + 1)
    return remote


def arrow_rear_label_position(lane: Lane, t: Transform, distance_behind_arrow_m: float = 12.0) -> tuple[tuple[float, float], float]:
    """Place label at fixed distance behind arrow center, aligned to lane."""
    c = arrow_center(lane, t, 7.0)
    if c is None:
        p = point_along_polyline(lane.points_m, 0.34)
        c = t(p)
    ang = screen_angle(lane.points_m, t)
    px_back = t.px(distance_behind_arrow_m)
    pos = (c[0] - math.cos(ang) * px_back, c[1] - math.sin(ang) * px_back)
    deg = math.degrees(ang)
    if deg > 90 or deg < -90:
        deg += 180
    return pos, deg



def arrow_scale_for_lane(lane: Lane, t: Transform, kind: str, base: float) -> float:
    """Choose arrow size. Turn arrows keep original readable size."""
    if kind in ("left", "right", "left_straight", "right_straight"):
        return max(0.55, min(base, 1.05))
    lane_w_px = t.px(lane.width_m)
    cap = (lane_w_px * 0.36) / 8.5
    return max(0.42, min(base, cap))


def shifted_arrow_center_for_turn(lane: Lane, t: Transform, center: tuple[float, float], angle: float, kind: str, scale: float) -> tuple[float, float]:
    """Shift entire left/right arrow laterally into lane without changing shape.

    Local arrow +Y is vehicle-right, -Y is vehicle-left.
    - left arrow extends to -Y, so move whole arrow to +Y / right side.
    - right arrow extends to +Y, so move whole arrow to -Y / left side.
    Shape, size, stroke width, and arrow head remain unchanged.
    """
    if kind not in ("left", "right", "left_straight", "right_straight"):
        return center
    L = 34 * scale
    turn_extent = L * 0.48
    half_lane = t.px(lane.width_m) / 2.0
    margin = max(1.5, t.px(0.22))
    needed = max(0.0, turn_extent + margin - half_lane)
    # Always move turn arrows a bit away from their turning-side boundary;
    # if the arrow envelope would overflow, move further.
    preferred = t.px(lane.width_m) * 0.18
    shift = min(max(preferred, needed), max(0.0, half_lane - margin))
    # screen vector of local +Y/right side.
    rx, ry = -math.sin(angle), math.cos(angle)
    if kind in ("left", "left_straight"):
        return (center[0] + rx * shift, center[1] + ry * shift)
    return (center[0] - rx * shift, center[1] - ry * shift)

def fmt(n: float) -> str:
    return f"{n:.2f}".rstrip("0").rstrip(".")


def pts_attr(points: list[tuple[float, float]]) -> str:
    return " ".join(f"{fmt(x)},{fmt(y)}" for x, y in points)


def path_attr(points: list[tuple[float, float]]) -> str:
    if not points:
        return ""
    return "M " + " L ".join(f"{fmt(x)} {fmt(y)}" for x, y in points)


def svg_arrow(center, angle, kind, scale, color=LANE_WHITE) -> str:
    """SVG road arrow with fixed shape/size; no clipping or lateral compression."""
    x, y = center
    ca, sa = math.cos(angle), math.sin(angle)
    def tr(local):
        lx, ly = local
        return (x + lx * ca - ly * sa, y + lx * sa + ly * ca)
    L = 34 * scale
    shaft = max(3.0, 4.2 * scale)
    head_w = 8.5 * scale
    def line(points):
        return f'<polyline points="{pts_attr(points)}" fill="none" stroke="{color}" stroke-width="{fmt(shaft)}" stroke-linecap="round" stroke-linejoin="round" opacity="0.96"/>'
    def poly(points):
        return f'<polygon points="{pts_attr(points)}" fill="{color}" opacity="0.96"/>'
    if kind == "straight":
        return line([tr((-L*.48,0)), tr((L*.18,0))]) + poly([tr((L*.46,0)), tr((L*.16,-head_w)), tr((L*.16,head_w))])
    if kind == "left":
        pts = [tr((-L*.46,0)), tr((-L*.08,0)), tr((L*.05,-L*.10)), tr((L*.05,-L*.28))]
        return line(pts) + poly([tr((L*.05,-L*.48)), tr((-head_w,-L*.25)), tr((head_w,-L*.25))])
    if kind == "right":
        pts = [tr((-L*.46,0)), tr((-L*.08,0)), tr((L*.05,L*.10)), tr((L*.05,L*.28))]
        return line(pts) + poly([tr((L*.05,L*.48)), tr((-head_w,L*.25)), tr((head_w,L*.25))])
    if kind == "left_straight":
        return svg_arrow(center, angle, "straight", scale, color) + svg_arrow(center, angle, "left", scale*.86, color)
    if kind == "right_straight":
        return svg_arrow(center, angle, "straight", scale, color) + svg_arrow(center, angle, "right", scale*.86, color)
    return svg_arrow(center, angle, "straight", scale, color)


def render_svg(m: MapData, out: Path, size=1200, radius_m=82, rotation_deg=0.0, title=True):
    allpts = [(0.0, 0.0)]
    for link in m.links:
        allpts.extend(link.points_m)
        for lane in link.lanes:
            allpts.extend(lane.points_m)
            allpts.extend(sl.point_m for sl in lane.stop_lines)
    t = Transform(allpts, size, 55, rotation_deg, radius_m or None)
    parts = []
    parts.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 {size} {size}" role="img">')
    parts.append(f'<title>MAP {html.escape(m.region)}-{html.escape(m.node_id)} {html.escape(m.name)}</title>')
    parts.append(f'<rect width="100%" height="100%" fill="{BG}"/>')
    parts.append('<g id="asphalt">')
    for link in m.links:
        pts = [t(p) for p in link.points_m]
        if len(pts) >= 2:
            w = t.px(link.width_m)
            parts.append(f'<path d="{path_attr(pts)}" fill="none" stroke="{ASPHALT_DARK}" stroke-width="{fmt(w+t.px(.8))}" stroke-linecap="round" stroke-linejoin="round"/>')
            parts.append(f'<path d="{path_attr(pts)}" fill="none" stroke="{ASPHALT}" stroke-width="{fmt(w)}" stroke-linecap="round" stroke-linejoin="round"/>')
    parts.append('</g>')

    parts.append('<g id="crosswalks">')
    for link in m.links:
        for lane in link.lanes:
            if not lane.is_crosswalk or len(lane.points_m) < 2:
                continue
            poly = [t(p) for p in lane_polygon_m(lane)]
            if len(poly) >= 3:
                parts.append(f'<polygon data-lane="{html.escape(lane.lane_id)}" points="{pts_attr(poly)}" fill="#232624" opacity="0.28"/>')
            half = max(lane.width_m/2, .75)
            for side in (half, -half):
                bpts = [t(p) for p in offset_polyline(lane.points_m, side)]
                parts.append(f'<path d="{path_attr(bpts)}" fill="none" stroke="{GUIDE_GREEN}" stroke-width="{fmt(t.px(.16))}" opacity="0.38"/>')
            pts = [t(p) for p in lane.points_m]
            width_px = max(t.px(lane.width_m), 12)
            for i in range(len(pts)-1):
                x1,y1=pts[i]; x2,y2=pts[i+1]
                L=math.hypot(x2-x1,y2-y1)
                if L<2: continue
                ux,uy=(x2-x1)/L,(y2-y1)/L; nx,ny=-uy,ux
                step=max(8,t.px(1.2)); stripe_w=max(3,t.px(.45)); d=0; k=0
                while d<=L:
                    if k%2==0:
                        cx,cy=x1+ux*d,y1+uy*d
                        a=(cx-nx*width_px*.46, cy-ny*width_px*.46); b=(cx+nx*width_px*.46, cy+ny*width_px*.46)
                        parts.append(f'<line x1="{fmt(a[0])}" y1="{fmt(a[1])}" x2="{fmt(b[0])}" y2="{fmt(b[1])}" stroke="{CROSSWALK_WHITE}" stroke-width="{fmt(stripe_w)}" opacity="0.42"/>')
                    k+=1; d+=step
    parts.append('</g>')

    parts.append('<g id="lane-polygons">')
    for link in m.links:
        for lane in link.lanes:
            if lane.is_crosswalk:
                continue
            poly = [t(p) for p in lane_polygon_m(lane)]
            if len(poly) >= 3:
                fill = '#2d5232' if lane.is_guide_or_waiting else '#484c4a'
                opacity = '0.55' if lane.is_guide_or_waiting else '0.78'
                parts.append(f'<polygon data-link="{html.escape(link.upstream)}" data-lane="{html.escape(lane.lane_id)}" points="{pts_attr(poly)}" fill="{fill}" opacity="{opacity}"/>')
                half = max(lane.width_m/2, .75)
                edge = GUIDE_GREEN if lane.is_guide_or_waiting else LANE_WHITE
                ew = t.px(.22 if lane.is_guide_or_waiting else .12)
                for side in (half, -half):
                    bpts = [t(p) for p in offset_polyline(lane.points_m, side)]
                    parts.append(f'<path d="{path_attr(bpts)}" fill="none" stroke="{edge}" stroke-width="{fmt(ew)}" stroke-linecap="round" stroke-linejoin="round" opacity="0.92"/>')
    parts.append('</g>')

    parts.append('<g id="guided-waiting-areas">')
    for link in m.links:
        for lane in link.lanes:
            if lane.is_crosswalk:
                continue
            gpoly_m = guided_lane_polygon_m(lane)
            if len(gpoly_m) >= 3:
                kind = maneuver_kind(lane.maneuver)
                gpoly = [t(p) for p in gpoly_m]
                parts.append(f'<polygon data-link="{html.escape(link.upstream)}" data-lane="{html.escape(lane.lane_id)}" data-kind="guided-{kind}" points="{pts_attr(gpoly)}" fill="#174f22" opacity="0.50" stroke="{GUIDE_GREEN}" stroke-width="{fmt(t.px(.22))}"/>')
                # inner center arrow only as road marking; the centerline itself is not drawn.
                c0, c1 = lane.stop_lines[0].point_m if lane.stop_lines else lane.points_m[-1], None
                hdg = heading(lane.points_m)
                mid = (c0[0] + math.cos(hdg) * lane.guided_length_m * 0.52, c0[1] + math.sin(hdg) * lane.guided_length_m * 0.52)
                gbase = max(.42, min(.80, t.px(1.0)/8.0))
                gscale = arrow_scale_for_lane(lane, t, kind, gbase)
                parts.append(f'<g data-layer="guided-arrow" data-lane="{html.escape(lane.lane_id)}">{svg_arrow(shifted_arrow_center_for_turn(lane, t, t(mid), screen_angle(lane.points_m,t), kind, gscale), screen_angle(lane.points_m,t), kind, gscale, GUIDE_GREEN)}</g>')
    parts.append('</g>')

    parts.append('<g id="stop-lines">')
    for link in m.links:
        for lane in link.lanes:
            if lane.is_crosswalk or len(lane.points_m) < 2:
                continue
            hdg=heading(lane.points_m); nx,ny=-math.sin(hdg),math.cos(hdg)
            half=max(lane.width_m/2,.75)
            for sl in lane.stop_lines:
                a=(sl.point_m[0]-nx*half, sl.point_m[1]-ny*half); b=(sl.point_m[0]+nx*half, sl.point_m[1]+ny*half)
                ascr, bscr = t(a), t(b)
                parts.append(f'<line data-lane="{html.escape(lane.lane_id)}" x1="{fmt(ascr[0])}" y1="{fmt(ascr[1])}" x2="{fmt(bscr[0])}" y2="{fmt(bscr[1])}" stroke="{GUIDE_GREEN}" stroke-width="{fmt(t.px(.55))}" stroke-linecap="butt"/>')
    parts.append('</g>')

    parts.append('<g id="lane-arrow-rear-labels">')
    for link in m.links:
        lane_by_id = {lane.lane_id: lane for lane in link.lanes}
        for lane in link.lanes:
            if lane.is_crosswalk or len(lane.points_m) < 2:
                continue
            # Label only lanes with physical arrow semantics.
            kind = maneuver_kind(lane.maneuver)
            if not lane.maneuver and not lane.is_guide_or_waiting:
                continue
            downstream = downstream_node_for_lane(lane, lane_by_id, m.node_id)
            pos, deg = arrow_rear_label_position(lane, t, 12.0)
            label = f'{link.name} → {downstream}'
            parts.append(f'<text data-layer="lane-arrow-rear-label" data-lane="{html.escape(lane.lane_id)}" data-road="{html.escape(link.name)}" data-downstream-node="{html.escape(downstream)}" x="{fmt(pos[0])}" y="{fmt(pos[1])}" transform="rotate({fmt(deg)} {fmt(pos[0])} {fmt(pos[1])})" text-anchor="middle" dominant-baseline="central" fill="{LANE_WHITE}" font-family="Arial, PingFang SC, sans-serif" font-size="{fmt(max(10,t.px(1.15)))}" font-weight="700" opacity="0.82">{html.escape(label)}</text>')
    parts.append('</g>')

    parts.append('<g id="phase-labels">')
    for link in m.links:
        for lane in link.lanes:
            if lane.is_crosswalk or not lane.phase_id or len(lane.points_m) < 2:
                continue
            pos, ang = phase_label_position(lane, t)
            # Keep label readable but aligned with lane direction, like the MAP sample.
            deg = math.degrees(ang)
            if deg > 90 or deg < -90:
                deg += 180
            parts.append(f'<text data-lane="{html.escape(lane.lane_id)}" data-phase="{html.escape(lane.phase_id)}" x="{fmt(pos[0])}" y="{fmt(pos[1])}" transform="rotate({fmt(deg)} {fmt(pos[0])} {fmt(pos[1])})" text-anchor="middle" dominant-baseline="central" fill="{LANE_WHITE}" font-family="Arial, PingFang SC, sans-serif" font-size="{fmt(max(10,t.px(1.15)))}" font-weight="700" opacity="0.96">{html.escape(lane.phase_id)}</text>')
    parts.append('</g>')

    parts.append('<g id="arrows">')
    for link in m.links:
        for lane in link.lanes:
            if lane.is_crosswalk or lane.is_guide_or_waiting:
                continue
            c = arrow_center(lane, t, 7.0)
            if c:
                kind = maneuver_kind(lane.maneuver)
                base_scale=max(.42,min(.95,t.px(1.0)/7.5))
                scale=arrow_scale_for_lane(lane, t, kind, base_scale)
                parts.append(f'<g data-lane="{html.escape(lane.lane_id)}" data-maneuver="{html.escape(lane.maneuver)}">{svg_arrow(shifted_arrow_center_for_turn(lane, t, c, screen_angle(lane.points_m,t), kind, scale), screen_angle(lane.points_m,t), kind, scale, LANE_WHITE)}</g>')
    parts.append('</g>')

    if title:
        label = f'{m.name}  {m.region}-{m.node_id}'
        parts.append('<g id="map-label" opacity="0.92">')
        parts.append(f'<rect x="12" y="12" width="300" height="32" fill="#000" opacity="0.55"/>')
        parts.append(f'<text x="20" y="34" fill="#eeeeee" font-family="Arial, PingFang SC, sans-serif" font-size="16" font-weight="700">{html.escape(label)}</text>')
        parts.append('</g>')
    bar=t.px(10); x0=24; y0=size-28
    parts.append('<g id="scale-north">')
    parts.append(f'<line x1="{x0}" y1="{y0}" x2="{fmt(x0+bar)}" y2="{y0}" stroke="{LANE_WHITE}" stroke-width="3"/>')
    parts.append(f'<text x="{fmt(x0+bar/2-20)}" y="{y0-14}" fill="{LANE_WHITE}" font-size="15" font-family="Arial">10 m</text>')
    nx,ny=size-35,size-55
    parts.append(f'<polygon points="{nx},{ny-36} {nx-10},{ny} {nx+10},{ny}" fill="{LANE_WHITE}" opacity="0.85"/><text x="{nx-7}" y="{ny-58}" fill="{LANE_WHITE}" font-size="14" font-weight="700">N</text>')
    parts.append('</g>')
    parts.append('</svg>')
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(parts), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="Convert MAP XML to physical-world SVG")
    ap.add_argument("xml", type=Path)
    ap.add_argument("-o", "--output", type=Path)
    ap.add_argument("--size", type=int, default=1200)
    ap.add_argument("--radius-m", type=float, default=82)
    ap.add_argument("--rotation-deg", type=float, default=0.0)
    ap.add_argument("--no-title", action="store_true")
    args = ap.parse_args()
    m = parse_map(args.xml)
    out = args.output or args.xml.with_suffix(".svg")
    render_svg(m, out, args.size, args.radius_m, args.rotation_deg, not args.no_title)
    print(out)

if __name__ == "__main__":
    main()
