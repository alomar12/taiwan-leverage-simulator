from __future__ import annotations

# V3.6.60：治理現況底圖允許藍色「未指定現況」線先行儲存，後續再補登治理狀態。
# V3.6.59：修正 st_folium 重繪後整個 Streamlit 頁面往上捲動；改用地圖錨點恢復頁面位置。
# V3.6.58：總覽顯示設定／縮放線寬與各 GIS 圖台視角保持；暫時隱藏 PNG 輸出。
# V3.6.57：重複工程與共用圖資新增多組、多工程批次整理；保留 V3.6.56 修正。
# V3.6.51：重複工程盤點＋SPJ共用圖資＋跨計畫同一實體工程去重

import base64
import copy
import hashlib
import inspect
import io
import json
import re
import time
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import quote

from branca.element import MacroElement, Template
import folium
import requests
import streamlit as st
import streamlit.components.v1 as components
from folium.plugins import Draw, MarkerCluster
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Protection, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from pyproj import Transformer
from streamlit_folium import st_folium

try:
    _ST_FOLIUM_HAS_ON_CHANGE = "on_change" in inspect.signature(st_folium).parameters
except Exception:
    _ST_FOLIUM_HAS_ON_CHANGE = False


# V3.6.33：治理現況底圖新增資料庫工程快速定位＋沿用代表點/工程圖資重構
# ============================================================
# 基本設定
# ============================================================
TARGET_SHEETS = ["治理工程", "應急工程", "前瞻治理工程", "前瞻應急工程"]
PROJECT_ID_HEADER = "系統工程ID"
PROJECT_ID_RE = re.compile(r"^PRJ-(\d{7})(?:-(\d{2}))*$")
ROOT_PROJECT_ID_RE = re.compile(r"^PRJ-(\d{7})$")
GEO_ID_RE = re.compile(r"^GEO-(\d{7})$")
HIS_ID_RE = re.compile(r"^HIS-(\d{7})$")
SPJ_ID_RE = re.compile(r"^SPJ-(\d{7})$")

DEFAULT_GEO_PATH = "system_data/engineering_geo.geojson"
DEFAULT_HISTORY_PATH = "system_data/project_history.json"
DEFAULT_REGISTRY_PATH = "system_data/id_registry.json"
DEFAULT_EXCEL_PATH = "data/current.xlsx"

TAIWAN_CENTER = [23.70, 120.95]
TAIWAN_ZOOM = 7

STATUS_STYLE = {
    "未發包": ("#808080", "gray"),
    "招標中": ("#f1c40f", "orange"),
    "訂約中": ("#f39c12", "orange"),
    "待開工": ("#f39c12", "orange"),
    "施工中": ("#2980b9", "blue"),
    "停工中": ("#c0392b", "red"),
    "落後": ("#e74c3c", "red"),
    "已完工": ("#27ae60", "green"),
    "已解約": ("#2c3e50", "black"),
    "已取消": ("#2c3e50", "black"),
}
DEFAULT_STYLE = ("#7f8c8d", "gray")

TRANSFORMER_121 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
TRANSFORMER_119 = Transformer.from_crs("EPSG:3825", "EPSG:4326", always_xy=True)
TRANSFORMER_TO_121 = Transformer.from_crs("EPSG:4326", "EPSG:3826", always_xy=True)
TRANSFORMER_TO_119 = Transformer.from_crs("EPSG:4326", "EPSG:3825", always_xy=True)

# V3.6.44：TWD67 僅用於座標品質診斷／候選預覽；不會自動覆寫正式資料。
TRANSFORMER_67_121 = Transformer.from_crs("EPSG:3828", "EPSG:4326", always_xy=True)
TRANSFORMER_67_119 = Transformer.from_crs("EPSG:3827", "EPSG:4326", always_xy=True)
TRANSFORMER_67_GEO = Transformer.from_crs("EPSG:3821", "EPSG:4326", always_xy=True)


# ============================================================
# 共用工具
# ============================================================
def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _parent_busy_overlay_script(message: str, note: str, overlay_id: str = "wra-gis-global-busy-mask") -> str:
    """產生會把遮罩直接掛到 Streamlit parent document.body 的 JS。

    掛在 body 而不是一般 Streamlit 元件內，可跨過 st.rerun 的元件重建，
    避免 GitHub 已寫完、但 Streamlit 還在重建地圖的幾秒鐘出現空窗。
    """
    msg_js = json.dumps(str(message or "處理中…"), ensure_ascii=False)
    note_js = json.dumps(str(note or "請稍候"), ensure_ascii=False)
    oid_js = json.dumps(str(overlay_id), ensure_ascii=False)
    return f"""
    <script>
    (function() {{
      try {{
        var p = window.parent;
        var d = p.document;
        var oid = {oid_js};
        var old = d.getElementById(oid);
        if (old) old.remove();
        var mask = d.createElement('div');
        mask.id = oid;
        mask.setAttribute('role','status');
        mask.setAttribute('aria-live','assertive');
        mask.style.position='fixed';
        mask.style.inset='0';
        mask.style.zIndex='2147483646';
        mask.style.background='rgba(15,23,42,.48)';
        mask.style.backdropFilter='blur(1.5px)';
        mask.style.webkitBackdropFilter='blur(1.5px)';
        mask.style.display='flex';
        mask.style.alignItems='center';
        mask.style.justifyContent='center';
        mask.style.pointerEvents='all';
        mask.style.cursor='wait';

        var card=d.createElement('div');
        card.style.minWidth='min(440px,86vw)';
        card.style.maxWidth='86vw';
        card.style.padding='27px 31px 25px 31px';
        card.style.borderRadius='16px';
        card.style.background='#ffffff';
        card.style.boxShadow='0 18px 55px rgba(0,0,0,.30)';
        card.style.textAlign='center';
        card.style.color='#111827';
        card.style.webkitTextFillColor='#111827';
        card.style.fontFamily='sans-serif';

        var spinner=d.createElement('div');
        spinner.style.width='42px'; spinner.style.height='42px';
        spinner.style.margin='0 auto 14px auto';
        spinner.style.border='5px solid #e5e7eb';
        spinner.style.borderTopColor='#2563eb';
        spinner.style.borderRadius='50%';
        spinner.style.animation='wraGisBusySpin .8s linear infinite';

        if (!d.getElementById('wra-gis-busy-style')) {{
          var style=d.createElement('style');
          style.id='wra-gis-busy-style';
          style.textContent='@keyframes wraGisBusySpin{{0%{{transform:rotate(0deg)}}100%{{transform:rotate(360deg)}}}}';
          d.head.appendChild(style);
        }}

        var title=d.createElement('div');
        title.textContent={msg_js};
        title.style.fontSize='22px'; title.style.fontWeight='800';
        title.style.lineHeight='1.35'; title.style.marginBottom='8px';

        var desc=d.createElement('div');
        desc.textContent={note_js};
        desc.style.fontSize='15px'; desc.style.lineHeight='1.55';
        desc.style.whiteSpace='pre-line'; desc.style.color='#4b5563';
        desc.style.webkitTextFillColor='#4b5563';

        card.appendChild(spinner); card.appendChild(title); card.appendChild(desc);
        mask.appendChild(card); d.body.appendChild(mask);
      }} catch(e) {{}}
    }})();
    </script>
    """


def _inject_parent_busy_overlay(message: str, note: str) -> None:
    components.html(
        _parent_busy_overlay_script(message, note),
        height=0,
        width=0,
    )


def _saving_overlay(message: str = "資料儲存中…"):
    """V3.6.41：全畫面儲存遮罩跨 st.rerun 保留到新畫面穩定。

    舊版在 GitHub 寫入完成後就立刻拿掉遮罩，但 Streamlit / Folium 還會再花
    幾秒重建元件。新版把遮罩直接掛在 parent body，並以 session flag 延續到
    下一輪頁面全部輸出完畢後，再等 DOM 安靜才自動移除。
    """
    msg = str(message or "資料儲存中…")
    st.session_state["_gis_v641_busy_message"] = msg
    st.session_state["_gis_v641_busy_release_pending"] = True
    _inject_parent_busy_overlay(
        msg,
        "請勿切換頁面、重新整理或關閉視窗\n系統正在寫入並重新整理畫面，完成後會自動恢復",
    )
    # 讓前端先收到遮罩 delta，再進行 GitHub / Excel 寫入。
    time.sleep(0.08)
    return None


def _close_saving_overlay(_ph=None) -> None:
    """不立即移除遮罩；由 render_engineering_gis_page 結尾統一釋放。

    這可消除『遮罩先消失、右上角 Streamlit 還跑好幾秒』的空窗。
    """
    st.session_state["_gis_v641_busy_release_pending"] = True


def _resume_busy_overlay_if_needed() -> None:
    """若上一輪在 st.rerun 前完成寫入，下一輪一開始立刻補回中央遮罩。"""
    if not st.session_state.get("_gis_v641_busy_release_pending"):
        return
    msg = str(st.session_state.get("_gis_v641_busy_message") or "正在重新整理圖資畫面…")
    _inject_parent_busy_overlay(
        msg,
        "資料已進入更新流程，正在完成 Streamlit / 地圖重新載入\n請稍候，完成後遮罩會自動消失",
    )


def _release_busy_overlay_when_stable() -> None:
    """在本輪 Streamlit 已輸出完畢後，等畫面 DOM 安靜再移除遮罩。"""
    if not st.session_state.get("_gis_v641_busy_release_pending"):
        return
    components.html(
        r"""
        <script>
        (function(){
          try {
            var p=window.parent, d=p.document;
            var oid='wra-gis-global-busy-mask';
            var root=d.querySelector('[data-testid="stAppViewContainer"]') || d.querySelector('.stApp') || d.body;
            var done=false, quietTimer=null, hardTimer=null, observer=null;
            function finish(){
              if(done) return; done=true;
              try{ if(observer) observer.disconnect(); }catch(e){}
              try{ if(quietTimer) clearTimeout(quietTimer); }catch(e){}
              try{ if(hardTimer) clearTimeout(hardTimer); }catch(e){}
              try{ var el=d.getElementById(oid); if(el) el.remove(); }catch(e){}
            }
            function arm(){
              try{ if(quietTimer) clearTimeout(quietTimer); }catch(e){}
              quietTimer=setTimeout(finish, 900);
            }
            try{
              observer=new MutationObserver(function(){ arm(); });
              observer.observe(root,{childList:true,subtree:true,attributes:false});
            }catch(e){}
            arm();
            hardTimer=setTimeout(finish, 12000);
          }catch(e){}
        })();
        </script>
        """,
        height=0,
        width=0,
    )
    st.session_state.pop("_gis_v641_busy_release_pending", None)
    st.session_state.pop("_gis_v641_busy_message", None)


def _install_gis_scroll_memory(restore: bool = True) -> None:
    """V3.6.59：整個 GIS 頁面記住捲動位置與鄰近地圖錨點。

    舊版只還原絕對 scrollY；Streamlit rerun 時頁面高度與元件位置會短暫改變，
    瀏覽器自動捲動又可能覆寫正確值。新版在按下會觸發 rerun 的控制項時，另記
    鄰近地圖 iframe 在可視視窗中的 top，重建後以同一錨點反算 scrollY。
    """
    restore_js = "true" if restore else "false"
    components.html(
        f"""
        <script>
        (function(){{
          try {{
            var p=window.parent;
            var key='wra_gis_page_scroll_v3659';
            var intentKey='wra_gis_page_scroll_intent_v3659';
            var store=p.sessionStorage;
            function pageY(){{ return Number(p.scrollY||p.pageYOffset||0); }}
            function readJson(k){{
              try{{ var raw=store.getItem(k); return raw?JSON.parse(raw):null; }}catch(e){{ return null; }}
            }}
            function recentIntent(){{
              var obj=readJson(intentKey);
              if(obj && isFinite(obj.y) && isFinite(obj.ts) && Date.now()-Number(obj.ts)<8000) return obj;
              try{{ if(obj) store.removeItem(intentKey); }}catch(e){{}}
              return null;
            }}
            function largeFrames(){{
              try{{
                return Array.prototype.slice.call(p.document.querySelectorAll('iframe')).filter(function(fr){{
                  var r=fr.getBoundingClientRect();
                  return Number(r.height)>=280 && Number(r.width)>=280;
                }});
              }}catch(e){{ return []; }}
            }}
            function nearestFrame(clientY){{
              var fs=largeFrames(), best=null, bestDist=Infinity;
              fs.forEach(function(fr,idx){{
                try{{
                  var r=fr.getBoundingClientRect();
                  var d=clientY<r.top ? r.top-clientY : (clientY>r.bottom ? clientY-r.bottom : 0);
                  if(d<bestDist){{ bestDist=d; best={{index:idx,top:Number(r.top)}}; }}
                }}catch(e){{}}
              }});
              return best;
            }}
            function savePassive(){{
              try{{
                if(recentIntent()) return;
                store.setItem(key, JSON.stringify({{y:pageY(),ts:Date.now()}}));
              }}catch(e){{}}
            }}
            function isRerunControl(target){{
              try{{
                return !!(target && target.closest && target.closest(
                  'button,[role="button"],[role="radio"],[role="checkbox"],input[type="submit"],input[type="button"],input[type="radio"],input[type="checkbox"],input[type="range"],[data-baseweb="select"]'
                ));
              }}catch(e){{ return false; }}
            }}
            function captureIntent(ev){{
              if(!isRerunControl(ev && ev.target)) return;
              try{{
                var obj={{y:pageY(),ts:Date.now(),source:'page-control'}};
                var anchor=nearestFrame(Number(ev && ev.clientY));
                if(anchor){{ obj.frame_index=anchor.index; obj.frame_top=anchor.top; }}
                store.setItem(intentKey,JSON.stringify(obj));
                store.setItem(key,JSON.stringify(obj));
              }}catch(e){{}}
            }}
            function restoreState(obj){{
              if(!obj || !isFinite(obj.y)) return;
              var targetY=Math.max(0,Number(obj.y));
              try{{
                if(isFinite(obj.frame_index) && isFinite(obj.frame_top)){{
                  var fs=largeFrames(), fr=fs[Number(obj.frame_index)];
                  if(fr){{
                    var currentTop=Number(fr.getBoundingClientRect().top);
                    if(isFinite(currentTop)) targetY=Math.max(0,pageY()+currentTop-Number(obj.frame_top));
                  }}
                }}
              }}catch(e){{}}
              try{{ p.scrollTo({{top:targetY,left:0,behavior:'auto'}}); }}
              catch(e){{ try{{ p.scrollTo(0,targetY); }}catch(_e){{}} }}
            }}
            if(!p.__wraGisScrollMemoryInstalledV3659){{
              p.__wraGisScrollMemoryInstalledV3659=true;
              p.addEventListener('scroll',savePassive,{{passive:true}});
              p.addEventListener('pointerdown',captureIntent,true);
              p.addEventListener('keydown',function(ev){{
                if(ev && (ev.key==='Enter' || ev.key===' ')) captureIntent(ev);
              }},true);
            }}
            if({restore_js}){{
              var intent=recentIntent();
              var obj=intent || readJson(key);
              if(obj && isFinite(obj.y) && (!obj.ts || Date.now()-Number(obj.ts)<30000)){{
                [40,140,320,620,1000,1500,2100].forEach(function(ms){{
                  setTimeout(function(){{ restoreState(obj); }},ms);
                }});
                if(intent) setTimeout(function(){{
                  try{{
                    var latest=readJson(intentKey);
                    if(latest && Number(latest.ts)===Number(intent.ts)) store.removeItem(intentKey);
                    store.setItem(key,JSON.stringify({{y:pageY(),ts:Date.now()}}));
                  }}catch(e){{}}
                }},2350);
              }}
            }} else {{ savePassive(); }}
          }} catch(e) {{}}
        }})();
        </script>
        """,
        height=0,
        width=0,
    )


def _scroll_parent_to_current_component(offset_px: int = 88) -> None:
    """把 Streamlit 主頁自動捲到這個 components.html 所在位置。"""
    try:
        offset = max(0, int(offset_px))
    except Exception:
        offset = 88
    components.html(
        f"""
        <script>
        (function() {{
          function go() {{
            try {{
              var p = window.parent;
              var frame = window.frameElement;
              if (!p || !frame) return;
              var rect = frame.getBoundingClientRect();
              var top = Number(p.scrollY || p.pageYOffset || 0) + Number(rect.top || 0) - {offset};
              p.scrollTo({{top: Math.max(0, top), behavior: 'smooth'}});
              try {{
                p.sessionStorage.removeItem('wra_gis_page_scroll_intent_v3659');
                p.sessionStorage.setItem('wra_gis_page_scroll_v3659', JSON.stringify({{y:Math.max(0,top),ts:Date.now()}}));
              }} catch(e) {{}}
            }} catch(e) {{}}
          }}
          setTimeout(go, 80);
          setTimeout(go, 300);
          setTimeout(go, 650);
        }})();
        </script>
        """,
        height=1,
        width=1,
    )

def norm_text(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\s+", "", s)
    s = s.replace("（", "(").replace("）", ")").replace("％", "%")
    return s.strip()


def display_text(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.lower() == "nan":
        return ""
    return s


def parse_number(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return float(v)
        except Exception:
            return None
    s = str(v).strip().replace(",", "")
    if not s or s in {"-", "--"}:
        return None
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None


def status_color(status: Any) -> str:
    s = display_text(status)
    if s in STATUS_STYLE:
        return STATUS_STYLE[s][0]
    # 對可能含補充文字的狀態做保守判斷
    for key, (color, _) in STATUS_STYLE.items():
        if key and key in s:
            return color
    return DEFAULT_STYLE[0]


def status_icon_color(status: Any) -> str:
    s = display_text(status)
    if s in STATUS_STYLE:
        return STATUS_STYLE[s][1]
    for key, (_, icon_color) in STATUS_STYLE.items():
        if key and key in s:
            return icon_color
    return DEFAULT_STYLE[1]


def is_valid_wgs84(lon: Optional[float], lat: Optional[float]) -> bool:
    if lon is None or lat is None:
        return False
    # 涵蓋臺灣本島、澎湖、金門、馬祖附近
    return 117.8 <= lon <= 122.8 and 20.5 <= lat <= 26.8


def normalize_wgs84_pair(a: Any, b: Any) -> Optional[Tuple[float, float, bool]]:
    x = parse_number(a)
    y = parse_number(b)
    if x is None or y is None:
        return None
    if is_valid_wgs84(x, y):
        return x, y, False
    if is_valid_wgs84(y, x):
        return y, x, True
    return None


def normalize_twd97_pair(a: Any, b: Any, county: str = "") -> Optional[Tuple[float, float, bool, str]]:
    x = parse_number(a)
    y = parse_number(b)
    if x is None or y is None:
        return None

    swapped = False
    if 80_000 <= x <= 420_000 and 2_300_000 <= y <= 2_900_000:
        tx, ty = x, y
    elif 80_000 <= y <= 420_000 and 2_300_000 <= x <= 2_900_000:
        tx, ty = y, x
        swapped = True
    else:
        return None

    county_n = norm_text(county)
    prefer_119 = any(k in county_n for k in ["澎湖", "金門", "連江"])
    candidates = []
    for epsg, tr in (["3825", TRANSFORMER_119], ["3826", TRANSFORMER_121]) if prefer_119 else (["3826", TRANSFORMER_121], ["3825", TRANSFORMER_119]):
        try:
            lon, lat = tr.transform(tx, ty)
            if is_valid_wgs84(lon, lat):
                candidates.append((lon, lat, epsg))
        except Exception:
            pass
    if not candidates:
        return None
    lon, lat, epsg = candidates[0]
    return lon, lat, swapped, epsg


def normalize_header_map(ws, header_row: int) -> Dict[str, int]:
    result: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        n = norm_text(v)
        if n and n not in result:
            result[n] = c
    return result


def find_header_row(ws, target: str = "工程名稱", max_scan: int = 10) -> Optional[int]:
    target_n = norm_text(target)
    for r in range(1, min(ws.max_row, max_scan) + 1):
        for c in range(1, ws.max_column + 1):
            if norm_text(ws.cell(r, c).value) == target_n:
                return r
    return None


def find_col(header_map: Dict[str, int], candidates: Sequence[str]) -> Optional[int]:
    # 先完全相符，再做包含判斷
    norms = [norm_text(x) for x in candidates]
    for n in norms:
        if n in header_map:
            return header_map[n]
    for k, c in header_map.items():
        for n in norms:
            # 單一字母（例如 E / N）只允許完全相符，避免誤配到其他英文標題
            if len(n) >= 3 and (n in k or k in n):
                return c
    return None


def ensure_id_column(ws, header_row: int) -> int:
    hmap = normalize_header_map(ws, header_row)
    existing = find_col(hmap, [PROJECT_ID_HEADER])
    if existing:
        return existing

    c = ws.max_column + 1
    ws.cell(header_row, c).value = PROJECT_ID_HEADER
    # 複製左側標題的基本格式，避免新增欄看起來突兀
    if c > 1:
        src = ws.cell(header_row, c - 1)
        dst = ws.cell(header_row, c)
        if src.has_style:
            dst._style = copy.copy(src._style)
        if src.font:
            dst.font = copy.copy(src.font)
        if src.fill:
            dst.fill = copy.copy(src.fill)
        if src.border:
            dst.border = copy.copy(src.border)
        if src.alignment:
            dst.alignment = copy.copy(src.alignment)
    ws.column_dimensions[get_column_letter(c)].width = 18
    return c


def project_type_from_sheet(sheet_name: str) -> str:
    if "應急" in sheet_name:
        return "應急工程"
    return "治理工程"


def make_row_key(sheet_name: str, row: int) -> str:
    return f"{sheet_name}::{row}"


def split_row_key(key: str) -> Tuple[str, int]:
    sheet, row = key.rsplit("::", 1)
    return sheet, int(row)


def root_number(project_id: str) -> Optional[int]:
    m = re.match(r"^PRJ-(\d{7})", display_text(project_id))
    return int(m.group(1)) if m else None


def infer_parent_and_child_no(project_id: str) -> Optional[Tuple[str, int]]:
    s = display_text(project_id)
    m = re.match(r"^(PRJ-\d{7}(?:-\d{2})*)-(\d{2})$", s)
    if not m:
        return None
    return m.group(1), int(m.group(2))


def empty_geojson() -> Dict[str, Any]:
    return {"type": "FeatureCollection", "schema_version": 1, "features": []}


def empty_history() -> Dict[str, Any]:
    return {"schema_version": 1, "events": []}


def empty_registry() -> Dict[str, Any]:
    return {
        "schema_version": 2,
        "counters": {"PRJ": 0, "GEO": 0, "HIS": 0, "SPJ": 0},
        "child_counters": {},
        # 永久保存曾經出現過的工程識別資料。即使日後從 current.xlsx 移除，
        # 分標／併標管理仍可選回原母工程／來源工程。
        "project_catalog": {},
        # V3.6.46：人工座標檢核的最新結論；只存審核紀錄與候選值，
        # 不會直接改 current.xlsx 或正式 GIS。
        "coordinate_reviews": {},
        # V3.6.48：每次人工檢核都追加歷程；coordinate_reviews 仍保存每件最新結論。
        "coordinate_review_history": [],
        # V3.6.50：縣市回填座標的中央最新審核結果與歷程。
        "coordinate_return_reviews": {},
        "coordinate_return_review_history": [],
        # V3.6.51：實體空間工程群組。PRJ 保留管控資料身份，SPJ 表示共用空間圖資身份。
        "spatial_project_groups": {},
        # 人工確認「同名但不是同一實體工程」後，避免候選反覆出現。
        "duplicate_project_reviews": {},
    }


def json_load_bytes(data: Optional[bytes], default: Dict[str, Any]) -> Dict[str, Any]:
    if not data:
        return copy.deepcopy(default)
    try:
        return json.loads(data.decode("utf-8-sig"))
    except Exception:
        return copy.deepcopy(default)


def json_dump_bytes(obj: Any) -> bytes:
    return (json.dumps(obj, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def geometry_type(feature: Dict[str, Any]) -> str:
    g = feature.get("geometry") or {}
    return display_text(g.get("type"))


def normalize_keywords(v: Any) -> str:
    """統一圖資關鍵字格式；允許逗號、頓號、分號或換行分隔。"""
    raw = display_text(v)
    if not raw:
        return ""
    parts = re.split(r"[\n\r,，、;；]+", raw)
    out: List[str] = []
    seen = set()
    for part in parts:
        item = display_text(part)
        key = norm_text(item)
        if not item or not key or key in seen:
            continue
        seen.add(key)
        out.append(item)
    return "、".join(out)


def feature_keywords_text(features: Sequence[Dict[str, Any]]) -> str:
    """把既有多筆圖資關鍵字合併成編輯欄位預設值。"""
    values: List[str] = []
    seen = set()
    for f in features or []:
        p = (f or {}).get("properties") or {}
        for item in re.split(r"[\n\r,，、;；]+", display_text(p.get("keywords"))):
            item = display_text(item)
            key = norm_text(item)
            if item and key and key not in seen:
                seen.add(key)
                values.append(item)
    return "、".join(values)


def feature_keyword_match(feature: Dict[str, Any], keyword_norm: str) -> bool:
    """河道總覽用：比對單一 GeoJSON feature 的名稱／關鍵字等屬性。"""
    if not keyword_norm:
        return True
    p = (feature or {}).get("properties") or {}
    hay = "|".join(norm_text(p.get(k, "")) for k in (
        "keywords", "geo_name", "project_name", "water_name", "reach_name", "basis", "notes"
    ))
    return keyword_norm in hay


# ============================================================
# GitHub：同一 commit 原子更新多個檔案
# ============================================================
def _clear_gis_read_caches() -> None:
    """正式寫入 GitHub 後，清除唯讀快取與本次 Session 快照。"""
    for _name in (
        "_cached_gis_page_snapshot_v351",
        "_cached_query_geo_snapshot_v351",
        "_cached_scan_rows_v351",
    ):
        _fn = globals().get(_name)
        if _fn is not None and hasattr(_fn, "clear"):
            try:
                _fn.clear()
            except Exception:
                pass
    # V3.6.6：一般正式異動後才清 Session 快照。
    # 代表點批次同步會自行把最新 Excel/GEO 寫回 Session，不需要整頁重載。
    try:
        st.session_state.pop("_gis_v366_snapshot", None)
        st.session_state.pop("_gis_v366_snapshot_key", None)
    except Exception:
        pass


class GitHubError(RuntimeError):
    pass


@dataclass
class GitHubSettings:
    owner: str
    repo: str
    branch: str
    token: str
    excel_path: str = DEFAULT_EXCEL_PATH
    geo_path: str = DEFAULT_GEO_PATH
    history_path: str = DEFAULT_HISTORY_PATH
    registry_path: str = DEFAULT_REGISTRY_PATH

    @classmethod
    def from_streamlit(cls) -> "GitHubSettings":
        try:
            cfg = st.secrets["github"]
            return cls(
                owner=str(cfg["owner"]),
                repo=str(cfg["repo"]),
                branch=str(cfg.get("branch", "main")),
                token=str(cfg["token"]),
                excel_path=str(cfg.get("excel_path", DEFAULT_EXCEL_PATH)),
                geo_path=str(cfg.get("geo_path", DEFAULT_GEO_PATH)),
                history_path=str(cfg.get("history_path", DEFAULT_HISTORY_PATH)),
                registry_path=str(cfg.get("registry_path", DEFAULT_REGISTRY_PATH)),
            )
        except Exception as exc:
            raise GitHubError(
                "找不到 Streamlit Secrets 的 [github] 設定。請依安裝說明設定 owner、repo、branch、token。"
            ) from exc


class GitHubRepoStore:
    API = "https://api.github.com"

    def __init__(self, settings: GitHubSettings, timeout: int = 8):
        self.s = settings
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/vnd.github+json",
                "Authorization": f"Bearer {self.s.token}",
                "X-GitHub-Api-Version": "2022-11-28",
                "User-Agent": "WRA-Engineering-GIS",
            }
        )

    def _url(self, path: str) -> str:
        return f"{self.API}/repos/{quote(self.s.owner)}/{quote(self.s.repo)}/{path.lstrip('/')}"

    def _request(self, method: str, path: str, **kwargs):
        r = self.session.request(method, self._url(path), timeout=self.timeout, **kwargs)
        return r

    def get_head_commit(self) -> Tuple[str, str]:
        r = self._request("GET", f"git/ref/heads/{quote(self.s.branch, safe='')}")
        if r.status_code != 200:
            raise GitHubError(f"讀取 GitHub 分支失敗 ({r.status_code})：{r.text[:300]}")
        commit_sha = r.json()["object"]["sha"]

        r2 = self._request("GET", f"git/commits/{commit_sha}")
        if r2.status_code != 200:
            raise GitHubError(f"讀取 GitHub commit 失敗 ({r2.status_code})：{r2.text[:300]}")
        tree_sha = r2.json()["tree"]["sha"]
        return commit_sha, tree_sha

    def read_file(self, path: str, ref: Optional[str] = None, allow_missing: bool = False) -> Optional[bytes]:
        ref = ref or self.s.branch
        # 使用 raw media type，可直接取得 xlsx / GeoJSON 原始 bytes，並支援 1~100MB 檔案。
        r = self._request(
            "GET",
            f"contents/{quote(path, safe='/')}",
            params={"ref": ref},
            headers={"Accept": "application/vnd.github.raw+json"},
        )
        if r.status_code == 404 and allow_missing:
            return None
        if r.status_code != 200:
            raise GitHubError(f"讀取 {path} 失敗 ({r.status_code})：{r.text[:300]}")
        return r.content

    def atomic_update(
        self,
        paths: Sequence[str],
        mutator: Callable[[Dict[str, Optional[bytes]]], Tuple[Dict[str, bytes], Any]],
        message: str,
        max_retries: int = 5,
        clear_read_caches: bool = True,
    ) -> Any:
        """
        以一個 Git commit 同時更新多個檔案。
        mutator 接收「同一個 HEAD commit」下的檔案內容，回傳 changed_files 與 result。
        若 branch 在最後更新 ref 前被別人更新，重新從最新版執行整個 mutator。
        """
        last_err = None
        for attempt in range(max_retries):
            base_commit, base_tree = self.get_head_commit()
            # V3.6.6：同一 commit 下的檔案平行讀取，避免大型 current.xlsx 與 JSON 逐檔等待。
            current: Dict[str, Optional[bytes]] = {}
            path_list = list(paths)
            if path_list:
                def _read_path(_p):
                    return _p, self.read_file(_p, ref=base_commit, allow_missing=True)
                with ThreadPoolExecutor(max_workers=min(6, len(path_list))) as pool:
                    for _p, _content in pool.map(_read_path, path_list):
                        current[_p] = _content

            changed, result = mutator(current)
            if not changed:
                return result

            tree_entries = []
            for path, content in changed.items():
                rb = self._request(
                    "POST",
                    "git/blobs",
                    json={"content": base64.b64encode(content).decode("ascii"), "encoding": "base64"},
                )
                if rb.status_code != 201:
                    raise GitHubError(f"建立 Git blob 失敗 ({rb.status_code})：{rb.text[:300]}")
                tree_entries.append(
                    {"path": path, "mode": "100644", "type": "blob", "sha": rb.json()["sha"]}
                )

            rt = self._request(
                "POST",
                "git/trees",
                json={"base_tree": base_tree, "tree": tree_entries},
            )
            if rt.status_code != 201:
                raise GitHubError(f"建立 Git tree 失敗 ({rt.status_code})：{rt.text[:300]}")

            rc = self._request(
                "POST",
                "git/commits",
                json={
                    "message": message,
                    "tree": rt.json()["sha"],
                    "parents": [base_commit],
                },
            )
            if rc.status_code != 201:
                raise GitHubError(f"建立 Git commit 失敗 ({rc.status_code})：{rc.text[:300]}")

            rr = self._request(
                "PATCH",
                f"git/refs/heads/{quote(self.s.branch, safe='')}",
                json={"sha": rc.json()["sha"], "force": False},
            )
            if rr.status_code == 200:
                # 一般正式圖資異動仍清除快取；代表點校正可保留目前頁面快取，
                # 避免儲存第1點後，選第2點又重新下載整本 Excel 與全部 GIS 檔案。
                if clear_read_caches:
                    _clear_gis_read_caches()
                return result

            # 409/422：通常代表 branch 已被別人更新；重新讀最新版本再做一次
            if rr.status_code in (409, 422):
                last_err = GitHubError(f"GitHub 版本衝突，正在重試：{rr.text[:200]}")
                time.sleep(0.5 + attempt * 0.3)
                continue
            raise GitHubError(f"更新 GitHub branch 失敗 ({rr.status_code})：{rr.text[:300]}")

        raise last_err or GitHubError("GitHub 多人更新衝突次數過多，請重新整理頁面後再試。")


# ============================================================
# Excel / 工程資料解析
# ============================================================
@dataclass
class ProjectRow:
    row_key: str
    sheet_name: str
    row: int
    project_id: str
    project_name: str
    status: str
    county: str
    unit: str
    project_type: str
    address: str
    lon: Optional[float]
    lat: Optional[float]
    coord_source: str
    coord_swapped: bool
    coord_status: str
    water_system: str = ""
    project_content: str = ""


def _load_wb(excel_bytes: bytes, excel_path: str):
    keep_vba = excel_path.lower().endswith(".xlsm")
    return load_workbook(io.BytesIO(excel_bytes), data_only=False, keep_vba=keep_vba)


def _save_wb_bytes(wb) -> bytes:
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


_XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XLSX_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XLSX_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", _XLSX_MAIN_NS)
ET.register_namespace("r", _XLSX_DOC_REL_NS)


def _xlsx_sheet_xml_map(excel_bytes: bytes) -> Dict[str, str]:
    """取得工作表名稱 -> XLSX ZIP 內 worksheet XML 路徑。"""
    with zipfile.ZipFile(io.BytesIO(excel_bytes), "r") as zin:
        workbook = ET.fromstring(zin.read("xl/workbook.xml"))
        rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))

    rel_map = {}
    for rel in rels.findall(f"{{{_XLSX_REL_NS}}}Relationship"):
        rid = rel.attrib.get("Id", "")
        target = rel.attrib.get("Target", "")
        if rid and target:
            if target.startswith("/"):
                target = target.lstrip("/")
            elif not target.startswith("xl/"):
                target = "xl/" + target
            rel_map[rid] = target

    result = {}
    sheets = workbook.find(f"{{{_XLSX_MAIN_NS}}}sheets")
    if sheets is None:
        return result
    for sheet in sheets.findall(f"{{{_XLSX_MAIN_NS}}}sheet"):
        name = sheet.attrib.get("name", "")
        rid = sheet.attrib.get(f"{{{_XLSX_DOC_REL_NS}}}id", "")
        target = rel_map.get(rid)
        if name and target:
            result[name] = target
    return result


def _xlsx_cell_col_number(cell_ref: str) -> int:
    m = re.match(r"([A-Z]+)", cell_ref or "")
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n


def _patch_xlsx_numeric_cells(excel_bytes: bytes, patches: Sequence[Dict[str, Any]]) -> bytes:
    """只修改指定 XLSX 儲存格的 XML，不用 OpenPyXL 重新儲存整本活頁簿。

    目的：保留其它公式儲存格原本的 cached values，避免 GIS 校正座標後
    工程查詢的經費公式欄全部變成空白/0。
    patches: {sheet, row, col, value}
    """
    if not patches:
        return excel_bytes

    sheet_map = _xlsx_sheet_xml_map(excel_bytes)
    by_path: Dict[str, List[Dict[str, Any]]] = {}
    for patch in patches:
        sheet_name = display_text(patch.get("sheet"))
        path = sheet_map.get(sheet_name)
        if not path:
            raise ValueError(f"XLSX 找不到工作表 XML：{sheet_name}")
        by_path.setdefault(path, []).append(dict(patch))

    with zipfile.ZipFile(io.BytesIO(excel_bytes), "r") as zin:
        file_map = {info.filename: (info, zin.read(info.filename)) for info in zin.infolist()}

    for path, items in by_path.items():
        if path not in file_map:
            raise ValueError(f"XLSX 找不到工作表檔案：{path}")
        info, raw = file_map[path]
        root = ET.fromstring(raw)
        sheet_data = root.find(f"{{{_XLSX_MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise ValueError(f"工作表缺少 sheetData：{path}")

        row_map = {int(r.attrib.get("r", "0") or 0): r for r in sheet_data.findall(f"{{{_XLSX_MAIN_NS}}}row")}
        for patch in items:
            row_num = int(patch["row"])
            col_num = int(patch["col"])
            value = float(patch["value"])
            cell_ref = f"{get_column_letter(col_num)}{row_num}"

            row_el = row_map.get(row_num)
            if row_el is None:
                row_el = ET.Element(f"{{{_XLSX_MAIN_NS}}}row", {"r": str(row_num)})
                inserted = False
                for idx, existing in enumerate(list(sheet_data)):
                    er = int(existing.attrib.get("r", "0") or 0)
                    if er > row_num:
                        sheet_data.insert(idx, row_el)
                        inserted = True
                        break
                if not inserted:
                    sheet_data.append(row_el)
                row_map[row_num] = row_el

            cell_el = None
            cells = row_el.findall(f"{{{_XLSX_MAIN_NS}}}c")
            for c in cells:
                if c.attrib.get("r") == cell_ref:
                    cell_el = c
                    break
            if cell_el is None:
                cell_el = ET.Element(f"{{{_XLSX_MAIN_NS}}}c", {"r": cell_ref})
                inserted = False
                for idx, existing in enumerate(cells):
                    if _xlsx_cell_col_number(existing.attrib.get("r", "")) > col_num:
                        row_el.insert(idx, cell_el)
                        inserted = True
                        break
                if not inserted:
                    row_el.append(cell_el)

            # 座標欄正式改成數字值；若原格意外含公式/字串，移除舊內容。
            cell_el.attrib.pop("t", None)
            for tag in ("f", "is"):
                old = cell_el.find(f"{{{_XLSX_MAIN_NS}}}{tag}")
                if old is not None:
                    cell_el.remove(old)
            v_el = cell_el.find(f"{{{_XLSX_MAIN_NS}}}v")
            if v_el is None:
                v_el = ET.SubElement(cell_el, f"{{{_XLSX_MAIN_NS}}}v")
            v_el.text = format(value, ".15g")

        new_raw = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        file_map[path] = (info, new_raw)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zout:
        for filename, (info, raw) in file_map.items():
            zout.writestr(info, raw)
    return out.getvalue()


def extract_coordinate(ws, row: int, hmap: Dict[str, int], county: str) -> Tuple[Optional[float], Optional[float], str, bool, str]:
    gx = find_col(
        hmap,
        [
            "google X座標-E",
            "google X座標",
            "googleX座標-E",
            "googleX座標",
            "E",
        ],
    )
    gy = find_col(
        hmap,
        [
            "google Y座標-N",
            "google Y座標",
            "googleY座標-N",
            "googleY座標",
            "N",
        ],
    )
    if gx and gy:
        p = normalize_wgs84_pair(ws.cell(row, gx).value, ws.cell(row, gy).value)
        if p:
            lon, lat, swapped = p
            return lon, lat, "WGS84/Google", swapped, "有效"

    tx = find_col(
        hmap,
        [
            "TWD_97\n經度-X座標",
            "TWD_97經度-X座標",
            "TWD_97\n經度",
            "TWD97經度-X座標",
            "TWD97經度",
            "X座標",
        ],
    )
    ty = find_col(
        hmap,
        [
            "TWD_97\n緯度-Y座標",
            "TWD_97緯度-Y座標",
            "TWD_97\n緯度",
            "TWD97緯度-Y座標",
            "TWD97緯度",
            "Y座標",
        ],
    )
    if tx and ty:
        p = normalize_twd97_pair(ws.cell(row, tx).value, ws.cell(row, ty).value, county)
        if p:
            lon, lat, swapped, epsg = p
            return lon, lat, f"TWD97/EPSG:{epsg}", swapped, "有效"

    return None, None, "", False, "無有效座標"


def _coordinate_write_columns(hmap: Dict[str, int]) -> Dict[str, Optional[int]]:
    """找出 current.xlsx 既有的 WGS84/Google 與 TWD97 座標欄位。

    只依欄名尋找，不依固定 Excel 欄號；使用者在表內插入欄位不會影響。
    不自動新增座標欄，避免破壞既有表格結構。
    """
    return {
        "google_x": find_col(
            hmap,
            ["google X座標-E", "google X座標", "googleX座標-E", "googleX座標", "E"],
        ),
        "google_y": find_col(
            hmap,
            ["google Y座標-N", "google Y座標", "googleY座標-N", "googleY座標", "N"],
        ),
        "twd_x": find_col(
            hmap,
            [
                "TWD_97\n經度-X座標", "TWD_97經度-X座標", "TWD_97\n經度",
                "TWD97經度-X座標", "TWD97經度", "X座標",
            ],
        ),
        "twd_y": find_col(
            hmap,
            [
                "TWD_97\n緯度-Y座標", "TWD_97緯度-Y座標", "TWD_97\n緯度",
                "TWD97緯度-Y座標", "TWD97緯度", "Y座標",
            ],
        ),
    }


def _write_project_coordinate_to_workbook(
    wb,
    project: "ProjectRow",
    lon: float,
    lat: float,
) -> Dict[str, Any]:
    """把 GIS 校正後的代表點同步回 current.xlsx 的既有座標欄位。"""
    if not is_valid_wgs84(lon, lat):
        raise ValueError("新座標不在臺灣及離島的合理 WGS84 範圍內。")
    if project.sheet_name not in wb.sheetnames:
        raise ValueError(f"找不到工程工作表：{project.sheet_name}")

    ws = wb[project.sheet_name]
    header_row = find_header_row(ws)
    if not header_row:
        raise ValueError(f"{project.sheet_name} 找不到工程名稱標題列。")
    hmap = normalize_header_map(ws, header_row)
    cols = _coordinate_write_columns(hmap)

    written = []
    gx, gy = cols.get("google_x"), cols.get("google_y")
    if gx and gy:
        ws.cell(project.row, gx).value = round(float(lon), 8)
        ws.cell(project.row, gy).value = round(float(lat), 8)
        written.extend([display_text(ws.cell(header_row, gx).value), display_text(ws.cell(header_row, gy).value)])

    tx, ty = cols.get("twd_x"), cols.get("twd_y")
    if tx and ty:
        county_n = norm_text(project.county)
        use_119 = (
            "3825" in display_text(project.coord_source)
            or any(k in county_n for k in ["澎湖", "金門", "連江"])
        )
        transformer = TRANSFORMER_TO_119 if use_119 else TRANSFORMER_TO_121
        twd_x, twd_y = transformer.transform(float(lon), float(lat))
        ws.cell(project.row, tx).value = round(float(twd_x), 3)
        ws.cell(project.row, ty).value = round(float(twd_y), 3)
        written.extend([display_text(ws.cell(header_row, tx).value), display_text(ws.cell(header_row, ty).value)])

    if not written:
        raise ValueError(
            f"{project.sheet_name} 找不到可寫入的 Google/WGS84 或 TWD97 座標欄位，"
            "本次未修改 current.xlsx。"
        )

    return {
        "sheet_name": project.sheet_name,
        "row": project.row,
        "written_headers": [x for x in written if x],
    }


def _project_coordinate_xml_patches(
    wb,
    project: "ProjectRow",
    lon: float,
    lat: float,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """產生座標 XML patch，不直接修改/儲存 workbook。"""
    if not is_valid_wgs84(lon, lat):
        raise ValueError("新座標不在臺灣及離島的合理 WGS84 範圍內。")
    if project.sheet_name not in wb.sheetnames:
        raise ValueError(f"找不到工程工作表：{project.sheet_name}")

    ws = wb[project.sheet_name]
    header_row = find_header_row(ws)
    if not header_row:
        raise ValueError(f"{project.sheet_name} 找不到工程名稱標題列。")
    hmap = normalize_header_map(ws, header_row)
    cols = _coordinate_write_columns(hmap)
    patches: List[Dict[str, Any]] = []
    written: List[str] = []

    gx, gy = cols.get("google_x"), cols.get("google_y")
    if gx and gy:
        patches.extend([
            {"sheet": project.sheet_name, "row": project.row, "col": gx, "value": round(float(lon), 8)},
            {"sheet": project.sheet_name, "row": project.row, "col": gy, "value": round(float(lat), 8)},
        ])
        written.extend([display_text(ws.cell(header_row, gx).value), display_text(ws.cell(header_row, gy).value)])

    tx, ty = cols.get("twd_x"), cols.get("twd_y")
    if tx and ty:
        county_n = norm_text(project.county)
        use_119 = (
            "3825" in display_text(project.coord_source)
            or any(k in county_n for k in ["澎湖", "金門", "連江"])
        )
        transformer = TRANSFORMER_TO_119 if use_119 else TRANSFORMER_TO_121
        twd_x, twd_y = transformer.transform(float(lon), float(lat))
        patches.extend([
            {"sheet": project.sheet_name, "row": project.row, "col": tx, "value": round(float(twd_x), 3)},
            {"sheet": project.sheet_name, "row": project.row, "col": ty, "value": round(float(twd_y), 3)},
        ])
        written.extend([display_text(ws.cell(header_row, tx).value), display_text(ws.cell(header_row, ty).value)])

    if not patches:
        raise ValueError(
            f"{project.sheet_name} 找不到可寫入的 Google/WGS84 或 TWD97 座標欄位，"
            "本次未修改 current.xlsx。"
        )
    return patches, {
        "sheet_name": project.sheet_name,
        "row": project.row,
        "written_headers": [x for x in written if x],
    }


def _haversine_m(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    import math
    r = 6371008.8
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(min(1.0, math.sqrt(a)))


TAIWAN_COUNTY_NAMES = [
    "基隆市", "臺北市", "新北市", "桃園市", "新竹市", "新竹縣", "苗栗縣",
    "臺中市", "彰化縣", "南投縣", "雲林縣", "嘉義市", "嘉義縣", "臺南市",
    "高雄市", "屏東縣", "宜蘭縣", "花蓮縣", "臺東縣", "澎湖縣", "金門縣", "連江縣",
]


def normalize_county_name(value: Any) -> str:
    """將『○○縣政府／○○市政府』等文字統一成縣市名稱。"""
    s = display_text(value).replace("台", "臺")
    if not s:
        return ""
    for name in TAIWAN_COUNTY_NAMES:
        if name in s:
            return name
    # 欄位本身若已是簡短縣市名稱，保留原值
    if s.endswith(("縣", "市")) and len(s) <= 4:
        return s
    return ""


def infer_county_from_values(*values: Any) -> str:
    for value in values:
        county = normalize_county_name(value)
        if county:
            return county
    return ""


def scan_workbook(excel_bytes: bytes, excel_path: str, ensure_id_cols: bool = False) -> Tuple[Any, List[ProjectRow], Dict[str, int]]:
    wb = _load_wb(excel_bytes, excel_path)
    rows: List[ProjectRow] = []
    sheet_id_cols: Dict[str, int] = {}

    for sheet_name in TARGET_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_row = find_header_row(ws)
        if not header_row:
            continue
        hmap = normalize_header_map(ws, header_row)
        name_col = find_col(hmap, ["工程名稱"])
        if not name_col:
            continue
        id_col = find_col(hmap, [PROJECT_ID_HEADER])
        if ensure_id_cols and not id_col:
            id_col = ensure_id_column(ws, header_row)
            hmap = normalize_header_map(ws, header_row)
        if id_col:
            sheet_id_cols[sheet_name] = id_col

        status_col = find_col(hmap, ["執行情形", "執行狀態判斷(程式用)"])
        county_col = find_col(hmap, ["縣市", "縣市別", "縣市政府", "縣市名稱", "補助縣市"])
        unit_col = find_col(hmap, ["執行單位"])
        address_col = find_col(hmap, ["工程地點", "實施鄉鎮", "實施\n鄉鎮", "鄉鎮市"])
        water_col = find_col(hmap, ["水系修正", "水系名稱", "所屬水系", "河川／排水名稱", "河川/排水名稱", "排水系統"])
        content_col = find_col(hmap, ["工程內容", "工程概要", "工程內容概述", "主要工程內容"])

        for r in range(header_row + 1, ws.max_row + 1):
            name = display_text(ws.cell(r, name_col).value)
            if not name:
                continue
            if name in {"合計", "總計"}:
                continue
            pid = display_text(ws.cell(r, id_col).value) if id_col else ""
            status = display_text(ws.cell(r, status_col).value) if status_col else ""
            raw_county = display_text(ws.cell(r, county_col).value) if county_col else ""
            unit = display_text(ws.cell(r, unit_col).value) if unit_col else ""
            address = display_text(ws.cell(r, address_col).value) if address_col else ""
            water_system = display_text(ws.cell(r, water_col).value) if water_col else ""
            project_content = display_text(ws.cell(r, content_col).value) if content_col else ""
            # V3.5.2a：除縣市欄本身外，亦可從執行單位／工程地點推回縣市，
            # 避免不同工作表使用『縣市政府』或沒有獨立縣市欄時清單變空。
            county = infer_county_from_values(raw_county, unit, address) or raw_county
            lon, lat, src, swapped, cstatus = extract_coordinate(ws, r, hmap, county)
            rows.append(
                ProjectRow(
                    row_key=make_row_key(sheet_name, r),
                    sheet_name=sheet_name,
                    row=r,
                    project_id=pid,
                    project_name=name,
                    status=status,
                    county=county,
                    unit=unit,
                    project_type=project_type_from_sheet(sheet_name),
                    address=address,
                    lon=lon,
                    lat=lat,
                    coord_source=src,
                    coord_swapped=swapped,
                    coord_status=cstatus,
                    water_system=water_system,
                    project_content=project_content,
                )
            )
    return wb, rows, sheet_id_cols


def validate_duplicate_ids(rows: Sequence[ProjectRow]) -> Dict[str, List[ProjectRow]]:
    found: Dict[str, List[ProjectRow]] = {}
    for p in rows:
        if p.project_id:
            found.setdefault(p.project_id, []).append(p)
    return {k: v for k, v in found.items() if len(v) > 1}


def _catalog_upsert_project(
    registry: Dict[str, Any],
    project_id: Any,
    project_name: Any = "",
    *,
    sheet_name: Any = "",
    county: Any = "",
    source: str = "",
) -> bool:
    """將曾經出現過的工程永久保存在 registry.project_catalog。

    這個目錄只保存識別／查找用資料，不代表工程仍存在 current.xlsx。
    回傳是否真的有新增或更新。
    """
    pid = display_text(project_id)
    if not pid or not PROJECT_ID_RE.match(pid):
        return False
    catalog = registry.setdefault("project_catalog", {})
    old = catalog.get(pid) if isinstance(catalog.get(pid), dict) else {}
    entry = dict(old or {})
    changed = pid not in catalog

    name = display_text(project_name)
    sheet = display_text(sheet_name)
    cty = display_text(county)
    src = display_text(source)
    for key, value in (("project_name", name), ("sheet_name", sheet), ("county", cty)):
        if value and display_text(entry.get(key)) != value:
            entry[key] = value
            changed = True
    if src and not display_text(entry.get("source")):
        entry["source"] = src
        changed = True
    if not entry.get("first_recorded_at"):
        entry["first_recorded_at"] = now_iso()
        changed = True
    entry["project_id"] = pid
    catalog[pid] = entry
    return changed


def project_catalog_map(registry: Dict[str, Any]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    catalog = registry.get("project_catalog") or {}
    if isinstance(catalog, dict):
        for pid, raw in catalog.items():
            pid2 = display_text(pid)
            if not pid2:
                continue
            ent = raw if isinstance(raw, dict) else {}
            out[pid2] = display_text(ent.get("project_name"))
    return out


def _history_project_refs(event: Dict[str, Any]) -> List[Tuple[str, str, str]]:
    """統一讀取舊式單一 from/to 與 V3.6.39 多對多沿革。

    回傳 (方向, project_id, project_name)，方向為 from / to。
    """
    out: List[Tuple[str, str, str]] = []

    # 舊格式：一筆 event 一個來源、一個去向
    for direction, ik, nk in (
        ("from", "from_project_id", "from_project_name"),
        ("to", "to_project_id", "to_project_name"),
    ):
        pid = display_text(event.get(ik))
        if pid:
            out.append((direction, pid, display_text(event.get(nk))))

    # 新格式：一筆 event 保存整組來源與整組新工程
    for direction, ik, nk in (
        ("from", "from_project_ids", "from_project_names"),
        ("to", "to_project_ids", "to_project_names"),
    ):
        ids = event.get(ik) or []
        names = event.get(nk) or []
        if not isinstance(ids, (list, tuple)):
            ids = [ids]
        if not isinstance(names, (list, tuple)):
            names = [names]
        for i, raw_pid in enumerate(ids):
            pid = display_text(raw_pid)
            if not pid:
                continue
            name = display_text(names[i]) if i < len(names) else ""
            out.append((direction, pid, name))

    # 同一 event 可能為相容性同時帶有 scalar/list；去重。
    dedup: List[Tuple[str, str, str]] = []
    seen = set()
    for direction, pid, name in out:
        key = (direction, pid)
        if key in seen:
            continue
        seen.add(key)
        dedup.append((direction, pid, name))
    return dedup


def history_consumed_source_ids(history: Optional[Dict[str, Any]]) -> set:
    """已經完成過分標／併標／多對多重組的來源工程，不再允許重複選取。"""
    used = set()
    if not history:
        return used
    for ev in history.get("events", []):
        et = display_text(ev.get("event_type"))
        if et not in {"分標", "併標", "分併標重組"}:
            continue
        for direction, pid, _name in _history_project_refs(ev):
            if direction == "from" and pid:
                used.add(pid)
    return used


def eligible_historical_source_projects(
    rows: Sequence[ProjectRow],
    geo: Dict[str, Any],
    history: Optional[Dict[str, Any]],
    registry: Optional[Dict[str, Any]],
) -> Dict[str, str]:
    """只回傳「有歷史 ID、目前 Excel 已不存在、且尚未被分併標使用」的來源工程。"""
    current_ids = {p.project_id for p in rows if p.project_id}
    used_ids = history_consumed_source_ids(history)
    known = all_known_projects(rows, geo, history, registry)
    return {
        pid: name
        for pid, name in known.items()
        if PROJECT_ID_RE.match(display_text(pid))
        and pid not in current_ids
        and pid not in used_ids
    }


def sync_registry_with_existing_ids(registry: Dict[str, Any], rows: Sequence[ProjectRow], geo: Dict[str, Any], history: Dict[str, Any]) -> None:
    registry.setdefault("counters", {}).setdefault("PRJ", 0)
    registry["counters"].setdefault("GEO", 0)
    registry["counters"].setdefault("HIS", 0)
    registry["counters"].setdefault("SPJ", 0)
    registry.setdefault("child_counters", {})
    registry.setdefault("project_catalog", {})
    registry.setdefault("spatial_project_groups", {})
    registry.setdefault("duplicate_project_reviews", {})

    # 先收歷史與 GeoJSON，再以 current.xlsx 的最新名稱覆蓋，
    # 因此舊母工程即使已離開 current.xlsx 仍會永久留在目錄中。
    for ev in history.get("events", []):
        for _direction, _pid, _name in _history_project_refs(ev):
            _catalog_upsert_project(
                registry, _pid, _name, source="history"
            )
    for f in geo.get("features", []):
        prop = f.get("properties") or {}
        _catalog_upsert_project(
            registry, prop.get("project_id"), prop.get("project_name"),
            county=prop.get("county"), source="geojson"
        )
    for p in rows:
        if p.project_id:
            _catalog_upsert_project(
                registry, p.project_id, p.project_name,
                sheet_name=p.sheet_name, county=p.county, source="current.xlsx"
            )

    for p in rows:
        if p.project_id:
            rn = root_number(p.project_id)
            if rn is not None:
                registry["counters"]["PRJ"] = max(registry["counters"]["PRJ"], rn)
            pc = infer_parent_and_child_no(p.project_id)
            if pc:
                parent, child = pc
                registry["child_counters"][parent] = max(int(registry["child_counters"].get(parent, 0)), child)

    for f in geo.get("features", []):
        gid = display_text((f.get("properties") or {}).get("geo_id"))
        m = GEO_ID_RE.match(gid)
        if m:
            registry["counters"]["GEO"] = max(registry["counters"]["GEO"], int(m.group(1)))

    for ev in history.get("events", []):
        hid = display_text(ev.get("history_id"))
        m = HIS_ID_RE.match(hid)
        if m:
            registry["counters"]["HIS"] = max(registry["counters"]["HIS"], int(m.group(1)))

    for spj_id in (registry.get("spatial_project_groups") or {}).keys():
        m = SPJ_ID_RE.match(display_text(spj_id))
        if m:
            registry["counters"]["SPJ"] = max(
                int(registry["counters"].get("SPJ", 0)),
                int(m.group(1)),
            )


def next_root_id(registry: Dict[str, Any]) -> str:
    registry["counters"]["PRJ"] = int(registry["counters"].get("PRJ", 0)) + 1
    return f"PRJ-{registry['counters']['PRJ']:07d}"


def next_geo_id(registry: Dict[str, Any]) -> str:
    registry["counters"]["GEO"] = int(registry["counters"].get("GEO", 0)) + 1
    return f"GEO-{registry['counters']['GEO']:07d}"


def next_history_id(registry: Dict[str, Any]) -> str:
    registry["counters"]["HIS"] = int(registry["counters"].get("HIS", 0)) + 1
    return f"HIS-{registry['counters']['HIS']:07d}"


def next_spatial_project_id(registry: Dict[str, Any]) -> str:
    registry.setdefault("counters", {}).setdefault("SPJ", 0)
    registry["counters"]["SPJ"] = int(registry["counters"].get("SPJ", 0)) + 1
    return f"SPJ-{registry['counters']['SPJ']:07d}"


def next_child_id(registry: Dict[str, Any], parent_id: str) -> str:
    cc = registry.setdefault("child_counters", {})
    n = int(cc.get(parent_id, 0)) + 1
    if n > 99:
        raise ValueError(f"{parent_id} 的直接子編號已超過 99。")
    cc[parent_id] = n
    return f"{parent_id}-{n:02d}"



def _spatial_group_signature_v3651(project_ids: Sequence[str]) -> str:
    ids = sorted({display_text(x) for x in project_ids if display_text(x)})
    return "|".join(ids)


def _active_spatial_group_for_project_v3651(
    registry: Dict[str, Any],
    project_id: str,
) -> Tuple[str, Optional[Dict[str, Any]]]:
    pid = display_text(project_id)
    groups = (registry or {}).get("spatial_project_groups") or {}
    if not isinstance(groups, dict) or not pid:
        return "", None
    for spj_id, group in groups.items():
        if not isinstance(group, dict) or not group.get("active", True):
            continue
        members = [display_text(x) for x in group.get("linked_project_ids") or []]
        if pid in members:
            return display_text(spj_id), group
    return "", None


def _spatial_group_is_shared_v3651(group: Optional[Dict[str, Any]]) -> bool:
    if not isinstance(group, dict):
        return False
    return display_text(group.get("share_mode")) == "shared"


def _spatial_primary_from_registry_v3651(
    registry: Dict[str, Any],
    project_id: str,
) -> str:
    spj_id, group = _active_spatial_group_for_project_v3651(registry, project_id)
    if spj_id and _spatial_group_is_shared_v3651(group):
        return display_text(group.get("primary_project_id")) or display_text(project_id)
    return display_text(project_id)


def _spatial_primary_project_id_from_geo_v3651(
    geo: Dict[str, Any],
    project_id: str,
) -> str:
    pid = display_text(project_id)
    if not pid:
        return ""
    for f in geo.get("features", []):
        props = f.get("properties") or {}
        if display_text(props.get("project_id")) != pid:
            continue
        if display_text(props.get("spatial_share_mode")) != "shared":
            continue
        primary = display_text(props.get("spatial_primary_project_id"))
        if primary:
            return primary
    return pid


def _spatial_is_secondary_v3651(
    geo: Dict[str, Any],
    project_id: str,
) -> bool:
    pid = display_text(project_id)
    primary = _spatial_primary_project_id_from_geo_v3651(geo, pid)
    return bool(pid and primary and pid != primary)


def _spatial_dedupe_rows_v3651(
    rows: Sequence["ProjectRow"],
    geo: Dict[str, Any],
) -> List["ProjectRow"]:
    """共用圖資群組只保留一筆可視工程；優先使用 primary PRJ 的 ProjectRow。"""
    by_pid = {display_text(p.project_id): p for p in rows if display_text(p.project_id)}
    result: List["ProjectRow"] = []
    seen = set()
    for p in rows:
        pid = display_text(p.project_id)
        if not pid:
            result.append(p)
            continue
        canonical = _spatial_primary_project_id_from_geo_v3651(geo, pid) or pid
        if canonical in seen:
            continue
        seen.add(canonical)
        result.append(by_pid.get(canonical) or p)
    return result


def _raw_project_point_v3651(
    p: "ProjectRow",
    geo: Dict[str, Any],
) -> Optional[Tuple[float, float]]:
    """不套 SPJ 共用解析，專供重複工程盤點查看每筆原始/現況位置。"""
    f = original_feature_for_project(geo, p.project_id) if p.project_id else None
    geom = (f or {}).get("geometry") or {}
    coords = geom.get("coordinates") or []
    if geom.get("type") == "Point" and len(coords) >= 2:
        try:
            lon, lat = float(coords[0]), float(coords[1])
            if is_valid_wgs84(lon, lat):
                return lat, lon
        except Exception:
            pass
    if p.lon is not None and p.lat is not None and is_valid_wgs84(p.lon, p.lat):
        return float(p.lat), float(p.lon)
    return None


def _duplicate_candidate_key_v3651(name: str, county: str) -> str:
    return f"{norm_text(name)}::{normalize_county_name(county) or display_text(county)}"


def _duplicate_project_candidates_v3651(
    rows: Sequence["ProjectRow"],
    geo: Dict[str, Any],
    registry: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """只提出候選，不自動合併。

    目前第一階段以『同縣市＋同名』為主要規則，特別適合治理/前瞻治理重複列。
    """
    buckets: Dict[str, List["ProjectRow"]] = {}
    for p in rows:
        if not p.project_id or not display_text(p.project_name):
            continue
        key = _duplicate_candidate_key_v3651(p.project_name, p.county)
        buckets.setdefault(key, []).append(p)

    distinct_reviews = (registry or {}).get("duplicate_project_reviews") or {}
    out = []

    for _key, members0 in buckets.items():
        # 同 PRJ 不重複算
        unique = []
        seen_pid = set()
        for p in members0:
            if p.project_id not in seen_pid:
                unique.append(p)
                seen_pid.add(p.project_id)
        if len(unique) < 2:
            continue

        member_ids = sorted(p.project_id for p in unique)
        signature = _spatial_group_signature_v3651(member_ids)

        review = distinct_reviews.get(signature, {}) if isinstance(distinct_reviews, dict) else {}
        if isinstance(review, dict) and review.get("decision") == "confirmed_distinct":
            continue

        # 若全部成員已在同一個 active SPJ 中，視為已整理完成
        active_spjs = []
        for p in unique:
            spj_id, _g = _active_spatial_group_for_project_v3651(registry, p.project_id)
            if spj_id:
                active_spjs.append(spj_id)
        if active_spjs and len(active_spjs) == len(unique) and len(set(active_spjs)) == 1:
            continue

        points = []
        for p in unique:
            pt = _raw_project_point_v3651(p, geo)
            if pt:
                points.append((p.project_id, pt[0], pt[1]))

        max_dist = None
        if len(points) >= 2:
            md = 0.0
            for i in range(len(points)):
                for j in range(i + 1, len(points)):
                    d = _haversine_m(
                        points[i][1], points[i][2],
                        points[j][1], points[j][2],
                    )
                    md = max(md, float(d))
            max_dist = md

        sheets = sorted({display_text(p.sheet_name) for p in unique if display_text(p.sheet_name)})
        cross_plan = len(sheets) >= 2
        if max_dist is None:
            suggestion = "同名，但部分案件沒有有效座標，建議人工確認。"
            risk = "需人工確認"
        elif max_dist <= 30:
            suggestion = (
                "高度疑似跨計畫同一實體工程，位置幾乎相同。"
                if cross_plan else "高度疑似同一實體工程，位置幾乎相同。"
            )
            risk = "高度疑似"
        elif max_dist <= 150:
            suggestion = (
                "同名且跨計畫、位置接近，建議確認是否共用圖資。"
                if cross_plan else "同名且位置接近，建議人工確認。"
            )
            risk = "疑似"
        elif max_dist <= 500:
            suggestion = "同名但位置有明顯差距，可能是同一工程座標誤差，也可能是不同工區。"
            risk = "需人工確認"
        else:
            suggestion = "同名但位置差異很大，請確認是否一筆座標錯誤，或其實是不同工程。"
            risk = "高差異"

        out.append({
            "candidate_id": hashlib.sha1(signature.encode("utf-8")).hexdigest()[:12],
            "signature": signature,
            "project_name": unique[0].project_name,
            "county": unique[0].county,
            "members": unique,
            "member_ids": member_ids,
            "member_count": len(unique),
            "sheets": sheets,
            "cross_plan": cross_plan,
            "max_distance_m": max_dist,
            "suggestion": suggestion,
            "risk": risk,
            "existing_spj_ids": sorted(set(active_spjs)),
        })

    risk_order = {"高度疑似": 0, "疑似": 1, "需人工確認": 2, "高差異": 3}
    out.sort(key=lambda x: (
        risk_order.get(x.get("risk"), 9),
        display_text(x.get("county")),
        display_text(x.get("project_name")),
    ))
    return out



def original_feature_for_project(geo: Dict[str, Any], project_id: str) -> Optional[Dict[str, Any]]:
    for f in geo.get("features", []):
        p = f.get("properties") or {}
        if p.get("project_id") == project_id and p.get("role") == "original_point":
            return f
    return None


def sync_original_features(geo: Dict[str, Any], registry: Dict[str, Any], rows: Sequence[ProjectRow]) -> int:
    created = 0
    current_ids = {p.project_id for p in rows if p.project_id}

    # 目前 Excel 中存在的工程都視為目前有效
    for f in geo.get("features", []):
        prop = f.get("properties") or {}
        pid = prop.get("project_id")
        if pid in current_ids:
            prop["project_active"] = True

    for p in rows:
        if not p.project_id:
            continue
        f = original_feature_for_project(geo, p.project_id)
        geometry = None
        if p.lon is not None and p.lat is not None:
            geometry = {"type": "Point", "coordinates": [p.lon, p.lat]}

        props = {
            "project_id": p.project_id,
            "project_name": p.project_name,
            "sheet_name": p.sheet_name,
            "source_row": p.row,
            "project_type": p.project_type,
            "county": p.county,
            "unit": p.unit,
            "address": p.address,
            "status_snapshot": p.status,
            "coord_source": p.coord_source,
            "coord_swapped": bool(p.coord_swapped),
            "coord_status": p.coord_status,
            "role": "original_point",
            "source": "excel_auto",
            "feature_active": True,
            "project_active": True,
            "updated_at": now_iso(),
        }
        if f is None:
            props["geo_id"] = next_geo_id(registry)
            props["created_at"] = now_iso()
            geo.setdefault("features", []).append({"type": "Feature", "properties": props, "geometry": geometry})
            created += 1
        else:
            old = f.setdefault("properties", {})
            geo_id = old.get("geo_id")
            created_at = old.get("created_at")
            # V3.6.8：GIS 校正後，工程代表點以 GIS 為座標主資料。
            # 後續「同步工程資料」只更新名稱／狀態等屬性，不允許 Excel 舊座標覆蓋 GIS。
            gis_authoritative = bool(
                old.get("coordinate_authority") == "gis"
                or old.get("coordinate_edit_source") == "gis_map"
                or old.get("source") == "gis_representative"
            )
            preserve = {
                k: old.get(k) for k in (
                    "coordinate_authority", "coordinate_edit_source", "coordinate_updated_at",
                    "coordinate_updated_by", "excel_sync_status", "excel_synced_at",
                    "excel_synced_by", "source"
                )
            }
            old.update(props)
            if geo_id:
                old["geo_id"] = geo_id
            if created_at:
                old["created_at"] = created_at
            if gis_authoritative:
                for k, v in preserve.items():
                    if v not in (None, ""):
                        old[k] = v
                old["source"] = "gis_representative"
                if display_text(old.get("spatial_share_mode")) == "shared":
                    old["coord_source"] = "共用GIS代表點"
                    old["coord_status"] = (
                        "共用圖資代表點；Excel待管理者同步"
                        if old.get("excel_sync_status") != "synced"
                        else "共用圖資；GIS／Excel已同步"
                    )
                else:
                    old["coord_source"] = "GIS代表點"
                    old["coord_status"] = "GIS為最新座標；Excel待管理者同步" if old.get("excel_sync_status") != "synced" else "GIS／Excel已同步"
                # geometry 保留 GIS 最新位置
            else:
                f["geometry"] = geometry
    return created


def set_project_inactive(geo: Dict[str, Any], project_id: str) -> None:
    for f in geo.get("features", []):
        p = f.get("properties") or {}
        if p.get("project_id") == project_id:
            p["project_active"] = False
            p["project_inactive_at"] = now_iso()


def set_row_id(wb, sheet_name: str, row: int, expected_name: str, new_id: str) -> None:
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"找不到工作表：{sheet_name}")
    ws = wb[sheet_name]
    header_row = find_header_row(ws)
    if not header_row:
        raise ValueError(f"{sheet_name} 找不到標題列。")
    hmap = normalize_header_map(ws, header_row)
    name_col = find_col(hmap, ["工程名稱"])
    id_col = find_col(hmap, [PROJECT_ID_HEADER]) or ensure_id_column(ws, header_row)
    current_name = display_text(ws.cell(row, name_col).value)
    current_id = display_text(ws.cell(row, id_col).value)
    if current_name != expected_name:
        raise ValueError(f"{sheet_name} 第 {row} 列工程名稱已變更，請重新整理後再操作。")
    if current_id:
        raise ValueError(f"{sheet_name} 第 {row} 列已經有工程ID：{current_id}")
    ws.cell(row, id_col).value = new_id


def workbook_to_project_map(rows: Sequence[ProjectRow]) -> Dict[str, ProjectRow]:
    return {p.project_id: p for p in rows if p.project_id}


def all_known_projects(
    rows: Sequence[ProjectRow],
    geo: Dict[str, Any],
    history: Optional[Dict[str, Any]] = None,
    registry: Optional[Dict[str, Any]] = None,
) -> Dict[str, str]:
    """回傳 current + 永久歷史目錄 + history + GeoJSON 曾出現過的所有工程。

    current.xlsx 的名稱優先級最高；即使舊工程已從 current.xlsx 消失，
    仍可從 project_catalog / history / inactive GeoJSON 找回。
    """
    result: Dict[str, str] = {}
    if registry:
        result.update(project_catalog_map(registry))
    if history:
        for ev in history.get("events", []):
            for _direction, pid, name in _history_project_refs(ev):
                if pid and (name or pid not in result):
                    result[pid] = name or result.get(pid, "")
    for f in geo.get("features", []):
        prop = f.get("properties") or {}
        pid = display_text(prop.get("project_id"))
        name = display_text(prop.get("project_name"))
        if pid and (name or pid not in result):
            result[pid] = name or result.get(pid, "")
    # current.xlsx 最後覆蓋，確保目前工程名稱永遠最新。
    for p in rows:
        if p.project_id:
            result[p.project_id] = p.project_name
    return result


# ============================================================
# GitHub 上的系統資料操作
# ============================================================
class EngineeringGISService:
    def __init__(self, store: GitHubRepoStore):
        self.store = store
        self.cfg = store.s

    @property
    def paths(self) -> List[str]:
        return [self.cfg.excel_path, self.cfg.geo_path, self.cfg.history_path, self.cfg.registry_path]

    def bootstrap(self) -> Dict[str, Any]:
        paths = [self.cfg.geo_path, self.cfg.history_path, self.cfg.registry_path]

        def mutate(cur):
            changed = {}
            if cur[self.cfg.geo_path] is None:
                changed[self.cfg.geo_path] = json_dump_bytes(empty_geojson())
            if cur[self.cfg.history_path] is None:
                changed[self.cfg.history_path] = json_dump_bytes(empty_history())
            if cur[self.cfg.registry_path] is None:
                changed[self.cfg.registry_path] = json_dump_bytes(empty_registry())
            return changed, {"created": list(changed.keys())}

        return self.store.atomic_update(paths, mutate, "初始化工程空間圖資系統檔案")

    def load_snapshot(self) -> Tuple[bytes, Dict[str, Any], Dict[str, Any], Dict[str, Any]]:
        commit, _ = self.store.get_head_commit()
        excel = self.store.read_file(self.cfg.excel_path, ref=commit)
        if excel is None:
            raise GitHubError(f"找不到 {self.cfg.excel_path}")
        geo = json_load_bytes(self.store.read_file(self.cfg.geo_path, ref=commit, allow_missing=True), empty_geojson())
        hist = json_load_bytes(self.store.read_file(self.cfg.history_path, ref=commit, allow_missing=True), empty_history())
        reg = json_load_bytes(self.store.read_file(self.cfg.registry_path, ref=commit, allow_missing=True), empty_registry())
        return excel, geo, hist, reg

    def initialize_existing_projects(self) -> Dict[str, Any]:
        def mutate(cur):
            excel_bytes = cur[self.cfg.excel_path]
            if not excel_bytes:
                raise ValueError(f"找不到 {self.cfg.excel_path}")
            geo = json_load_bytes(cur[self.cfg.geo_path], empty_geojson())
            hist = json_load_bytes(cur[self.cfg.history_path], empty_history())
            reg = json_load_bytes(cur[self.cfg.registry_path], empty_registry())

            wb, rows, _ = scan_workbook(excel_bytes, self.cfg.excel_path, ensure_id_cols=True)
            sync_registry_with_existing_ids(reg, rows, geo, hist)

            # 安全防呆：系統曾經發過 PRJ/GEO/HIS 編號，但目前 Excel 完全沒有工程ID，
            # 通常代表使用者覆蓋 current.xlsx 時遺失「系統工程ID」欄。
            # 此時禁止重新初始化，避免既有工程被重新發號而與歷史圖資失聯。
            registry_used = any(int(reg.get("counters", {}).get(k, 0) or 0) > 0 for k in ("PRJ", "GEO", "HIS"))
            geo_used = bool(geo.get("features"))
            hist_used = bool(hist.get("events"))
            workbook_has_any_id = any(bool(p.project_id) for p in rows)
            if (registry_used or geo_used or hist_used) and not workbook_has_any_id:
                raise ValueError(
                    "系統已有既有工程ID／圖資歷史，但目前 current.xlsx 找不到任何『系統工程ID』。"
                    "可能是覆蓋 Excel 時遺失了 ID 欄。為避免重新編號造成圖資錯接，系統已停止初始化。"
                )

            # 重新掃描，因 ensure_id_cols 可能剛新增欄位
            excel_tmp = _save_wb_bytes(wb)
            wb, rows, _ = scan_workbook(excel_tmp, self.cfg.excel_path, ensure_id_cols=True)

            assigned = 0
            for p in rows:
                if not p.project_id:
                    new_id = next_root_id(reg)
                    set_row_id(wb, p.sheet_name, p.row, p.project_name, new_id)
                    assigned += 1

            new_excel = _save_wb_bytes(wb)
            _, rows2, _ = scan_workbook(new_excel, self.cfg.excel_path, ensure_id_cols=False)
            duplicates = validate_duplicate_ids(rows2)
            if duplicates:
                raise ValueError("初始化後偵測到重複工程ID，已取消儲存。")
            new_geo_count = sync_original_features(geo, reg, rows2)

            changed = {
                self.cfg.excel_path: new_excel,
                self.cfg.geo_path: json_dump_bytes(geo),
                self.cfg.history_path: json_dump_bytes(hist),
                self.cfg.registry_path: json_dump_bytes(reg),
            }
            return changed, {
                "assigned": assigned,
                "geo_created": new_geo_count,
                "projects": len(rows2),
                "valid_coords": sum(1 for p in rows2 if p.lon is not None and p.lat is not None),
                "invalid_coords": sum(1 for p in rows2 if p.lon is None or p.lat is None),
            }

        return self.store.atomic_update(self.paths, mutate, "初始化工程ID與工程原始點位")

    def sync_current_projects(self) -> Dict[str, Any]:
        def mutate(cur):
            excel_bytes = cur[self.cfg.excel_path]
            if not excel_bytes:
                raise ValueError(f"找不到 {self.cfg.excel_path}")
            geo = json_load_bytes(cur[self.cfg.geo_path], empty_geojson())
            hist = json_load_bytes(cur[self.cfg.history_path], empty_history())
            reg = json_load_bytes(cur[self.cfg.registry_path], empty_registry())
            _, rows, _ = scan_workbook(excel_bytes, self.cfg.excel_path, ensure_id_cols=False)
            dups = validate_duplicate_ids(rows)
            if dups:
                msg = "、".join(list(dups.keys())[:10])
                raise ValueError(f"工程資料庫有重複工程ID：{msg}")
            sync_registry_with_existing_ids(reg, rows, geo, hist)
            created = sync_original_features(geo, reg, rows)
            changed = {
                self.cfg.geo_path: json_dump_bytes(geo),
                self.cfg.registry_path: json_dump_bytes(reg),
            }
            return changed, {"geo_created": created, "projects": len(rows)}

        return self.store.atomic_update(self.paths, mutate, "同步工程資料與空間圖資")

    def save_coordinate_reviews(
        self,
        reviews: Sequence[Dict[str, Any]],
        editor: str,
    ) -> Dict[str, Any]:
        """V3.6.48：保存人工座標檢核最新結論＋歷程，只更新 id_registry.json。

        - coordinate_reviews：每件工程最新人工結論。
        - coordinate_review_history：每次儲存都追加，不覆蓋舊判斷。
        - 不改 current.xlsx、不改 engineering_geo.geojson。
        """
        items = [dict(x) for x in reviews if isinstance(x, dict)]
        if not items:
            raise ValueError("沒有可儲存的人工座標檢核結果。")

        def mutate(cur):
            reg = json_load_bytes(cur.get(self.cfg.registry_path), empty_registry())
            bucket = reg.setdefault("coordinate_reviews", {})
            history_bucket = reg.setdefault("coordinate_review_history", [])
            if not isinstance(history_bucket, list):
                history_bucket = []
                reg["coordinate_review_history"] = history_bucket

            saved = []
            history_saved = []
            reviewed_at = now_iso()
            reviewer = display_text(editor)

            for item in items:
                pid = display_text(item.get("project_id"))
                row_key = display_text(item.get("row_key"))
                key = pid or row_key
                if not key:
                    continue
                decision = display_text(item.get("manual_decision"))
                if not decision:
                    continue

                base = {
                    "project_id": pid,
                    "row_key": row_key,
                    "project_name": display_text(item.get("project_name")),
                    "county": display_text(item.get("county")),
                    "system_suggestion": display_text(item.get("system_suggestion")),
                    "system_action": display_text(item.get("system_action")),
                    "manual_decision": decision,
                    "raw_pair_source": display_text(item.get("raw_pair_source")),
                    "tested_crs": display_text(item.get("tested_crs")),
                    "candidate_lon": parse_number(item.get("candidate_lon")),
                    "candidate_lat": parse_number(item.get("candidate_lat")),
                    "review_note": display_text(item.get("review_note")),
                    "reviewed_by": reviewer,
                    "reviewed_at": reviewed_at,
                }

                # 每件最新結論
                bucket[key] = dict(base)
                saved.append(dict(base))

                # 每次審核歷程都追加
                rid_seed = (
                    f"{key}|{reviewed_at}|{reviewer}|{decision}|"
                    f"{len(history_bucket)}"
                )
                hist_entry = dict(base)
                hist_entry["review_id"] = "CRV-" + hashlib.sha1(
                    rid_seed.encode("utf-8")
                ).hexdigest()[:12]
                history_bucket.append(hist_entry)
                history_saved.append(hist_entry)

            if not saved:
                raise ValueError("人工檢核結果缺少工程ID／列識別或人工結論。")

            return {
                self.cfg.registry_path: json_dump_bytes(reg),
            }, {
                "count": len(saved),
                "reviews": saved,
                "history_entries": history_saved,
            }

        # atomic_update 每次都以 GitHub 最新 HEAD 重新讀 registry，
        # 因此不同使用者分批儲存不會把別人的工程審核結果整包覆蓋。
        return self.store.atomic_update(
            [self.cfg.registry_path],
            mutate,
            f"人工座標檢核 {len(items)} 件",
            clear_read_caches=False,
        )

    def save_county_coordinate_return_reviews(
        self,
        items: Sequence[Dict[str, Any]],
        editor: str,
    ) -> Dict[str, Any]:
        """V3.6.50：儲存縣市回填座標的中央審核結果。

        決策：
        - 核准縣市新座標：只更新 engineering_geo.geojson 代表點，
          標記 Excel 待系統管理者同步。
        - 退回縣市再修正 / 待確認 / 採用中央既有GIS點位：
          只保存審核紀錄，不修改座標。

        本方法不讀、不寫 current.xlsx。
        """
        payload = [dict(x) for x in items if isinstance(x, dict)]
        if not payload:
            raise ValueError("沒有可儲存的縣市座標回填審核結果。")

        reviewer = display_text(editor) or "未填審核者"

        def mutate(cur):
            geo = json_load_bytes(cur.get(self.cfg.geo_path), empty_geojson())
            reg = json_load_bytes(cur.get(self.cfg.registry_path), empty_registry())

            latest = reg.setdefault("coordinate_return_reviews", {})
            history = reg.setdefault("coordinate_return_review_history", [])
            if not isinstance(latest, dict):
                latest = {}
                reg["coordinate_return_reviews"] = latest
            if not isinstance(history, list):
                history = []
                reg["coordinate_return_review_history"] = history

            reviewed_at = now_iso()
            saved = []
            approved = []
            rejected = []
            pending = []
            adopted_central = []

            for item in payload:
                pid = display_text(item.get("project_id"))
                decision = display_text(item.get("review_decision"))
                submission_hash = display_text(item.get("submission_hash"))
                if not pid:
                    raise ValueError("回填審核項目缺少系統工程ID。")
                if decision not in {
                    "核准縣府新座標",
                    "退回縣府再修正",
                    "待確認",
                    "採用中央既有GIS點位",
                }:
                    raise ValueError(f"{pid} 的中央審核結果無效：{decision}")

                lon = parse_number(item.get("candidate_lon"))
                lat = parse_number(item.get("candidate_lat"))

                if decision == "核准縣府新座標":
                    if lon is None or lat is None or not is_valid_wgs84(lon, lat):
                        raise ValueError(f"{pid} 的縣府新座標無效，不能核准。")

                    original = original_feature_for_project(geo, pid)
                    if original is None:
                        original = {
                            "type": "Feature",
                            "geometry": {
                                "type": "Point",
                                "coordinates": [float(lon), float(lat)],
                            },
                            "properties": {
                                "geo_id": next_geo_id(reg),
                                "project_id": pid,
                                "role": "original_point",
                                "feature_active": True,
                                "project_active": True,
                                "created_at": now_iso(),
                            },
                        }
                        geo.setdefault("features", []).append(original)
                    else:
                        original["geometry"] = {
                            "type": "Point",
                            "coordinates": [float(lon), float(lat)],
                        }

                    props = original.setdefault("properties", {})
                    props.update({
                        "project_id": pid,
                        "project_name": display_text(item.get("project_name"))
                            or display_text(props.get("project_name")),
                        "county": display_text(item.get("county"))
                            or display_text(props.get("county")),
                        "status_snapshot": display_text(item.get("status"))
                            or display_text(props.get("status_snapshot")),
                        "role": "original_point",
                        "source": "gis_representative",
                        "feature_active": True,
                        "project_active": True,
                        "coordinate_authority": "gis",
                        "coordinate_edit_source": "county_return_review",
                        "coordinate_origin": "縣市回填核准",
                        "coordinate_updated_at": reviewed_at,
                        "coordinate_updated_by": reviewer,
                        "excel_sync_status": "pending",
                        "coord_source": "縣市回填核准",
                        "coord_status": "縣市回填已核准；Excel待系統管理者同步",
                        "county_return_crs": display_text(item.get("returned_crs")),
                        "county_return_raw_x": parse_number(item.get("raw_x")),
                        "county_return_raw_y": parse_number(item.get("raw_y")),
                        "county_return_note": display_text(item.get("county_note")),
                        "county_return_submission_hash": submission_hash,
                        "county_return_reviewed_at": reviewed_at,
                        "county_return_reviewed_by": reviewer,
                        "updated_at": reviewed_at,
                    })
                    approved.append(pid)
                elif decision == "退回縣府再修正":
                    rejected.append(pid)
                elif decision == "待確認":
                    pending.append(pid)
                else:
                    adopted_central.append(pid)

                entry = {
                    "project_id": pid,
                    "row_key": display_text(item.get("row_key")),
                    "project_name": display_text(item.get("project_name")),
                    "county": display_text(item.get("county")),
                    "source_file": display_text(item.get("source_file")),
                    "submission_hash": submission_hash,
                    "returned_crs": display_text(item.get("returned_crs")),
                    "raw_x": parse_number(item.get("raw_x")),
                    "raw_y": parse_number(item.get("raw_y")),
                    "candidate_lon": lon,
                    "candidate_lat": lat,
                    "county_note": display_text(item.get("county_note")),
                    "validation_status": display_text(item.get("validation_status")),
                    "validation_message": display_text(item.get("validation_message")),
                    "distance_from_gis_m": parse_number(item.get("distance_from_gis_m")),
                    "review_decision": decision,
                    "review_note": display_text(item.get("review_note")),
                    "reviewed_by": reviewer,
                    "reviewed_at": reviewed_at,
                    "gis_applied": decision == "核准縣府新座標",
                    "excel_sync_status": (
                        "pending"
                        if decision == "核准縣府新座標"
                        else ""
                    ),
                }
                latest[pid] = dict(entry)

                seed = (
                    f"{pid}|{submission_hash}|{reviewed_at}|{reviewer}|"
                    f"{decision}|{len(history)}"
                )
                hist = dict(entry)
                hist["review_id"] = "CRR-" + hashlib.sha1(
                    seed.encode("utf-8")
                ).hexdigest()[:12]
                history.append(hist)
                saved.append(entry)

            changed = {
                self.cfg.registry_path: json_dump_bytes(reg),
            }
            if approved:
                changed[self.cfg.geo_path] = json_dump_bytes(geo)

            return changed, {
                "count": len(saved),
                "approved_count": len(approved),
                "rejected_count": len(rejected),
                "pending_count": len(pending),
                "adopted_central_count": len(adopted_central),
                "approved_project_ids": approved,
                "reviews": saved,
                "_new_geo": geo,
                "_new_registry": reg,
            }

        # 只讀 / 寫 GeoJSON + registry，不碰 current.xlsx。
        return self.store.atomic_update(
            [self.cfg.geo_path, self.cfg.registry_path],
            mutate,
            f"縣市座標回填中央審核：{len(payload)}件",
            clear_read_caches=False,
        )

    def save_spatial_project_group_v3651(
        self,
        members: Sequence[Dict[str, Any]],
        primary_project_id: str,
        relation_type: str,
        share_geometry: bool,
        note: str,
        editor: str,
    ) -> Dict[str, Any]:
        """建立或延伸 SPJ 實體空間工程群組。

        PRJ-ID 不合併、不刪除；Excel 的計畫/經費/進度仍各自保留。
        share_geometry=True 時：
        - primary PRJ 為正式圖資來源；
        - secondary 人工線/面保留但設 spatial_suppressed，不刪除；
        - 所有 linked PRJ 的 original_point 統一成 primary 代表點；
        - secondary Excel 等待系統管理者既有同步功能更新。
        """
        member_payload = [dict(x) for x in members if isinstance(x, dict)]
        ids = []
        metadata = {}
        for item in member_payload:
            pid = display_text(item.get("project_id"))
            if pid and pid not in ids:
                ids.append(pid)
                metadata[pid] = item
        primary = display_text(primary_project_id)
        if len(ids) < 2:
            raise ValueError("共用圖資群組至少需要 2 個不同的系統工程ID。")
        if primary not in ids:
            raise ValueError("圖資主工程必須是群組成員之一。")

        def mutate(cur):
            geo = json_load_bytes(cur.get(self.cfg.geo_path), empty_geojson())
            reg = json_load_bytes(cur.get(self.cfg.registry_path), empty_registry())
            reg.setdefault("spatial_project_groups", {})
            reg.setdefault("duplicate_project_reviews", {})
            reg.setdefault("counters", {}).setdefault("SPJ", 0)

            existing_ids = set()
            for pid in ids:
                spj0, _g0 = _active_spatial_group_for_project_v3651(reg, pid)
                if spj0:
                    existing_ids.add(spj0)
            if len(existing_ids) > 1:
                raise ValueError(
                    "選取工程目前分屬不同的既有 SPJ 群組，請先解除或整理既有群組。"
                )

            if existing_ids:
                spj_id = next(iter(existing_ids))
                group = reg["spatial_project_groups"][spj_id]
                old_shared = _spatial_group_is_shared_v3651(group)
                if bool(old_shared) != bool(share_geometry):
                    raise ValueError(
                        "既有 SPJ 的共用模式與本次選擇不同；請先解除群組後重新建立。"
                    )
                existing_primary = display_text(group.get("primary_project_id"))
                if existing_primary and primary != existing_primary:
                    raise ValueError(
                        f"既有 {spj_id} 的圖資主工程為 {existing_primary}，延伸群組時不可更換主工程。"
                    )
                combined = []
                for pid in list(group.get("linked_project_ids") or []) + ids:
                    pid = display_text(pid)
                    if pid and pid not in combined:
                        combined.append(pid)
                ids2 = combined
                primary2 = existing_primary or primary
                group.setdefault("extension_history", []).append({
                    "added_project_ids": [x for x in ids if x not in (group.get("linked_project_ids") or [])],
                    "extended_by": display_text(editor),
                    "extended_at": now_iso(),
                })
            else:
                spj_id = next_spatial_project_id(reg)
                ids2 = list(ids)
                primary2 = primary
                group = {
                    "spatial_project_id": spj_id,
                    "active": True,
                    "created_at": now_iso(),
                    "created_by": display_text(editor),
                    "member_original_points_before_group": {},
                    "extension_history": [],
                }

            # 將既有成員 metadata 併回 project_catalog / 群組快照
            member_info = dict(group.get("member_info") or {})
            for pid in ids2:
                item = metadata.get(pid) or {}
                cat = (reg.get("project_catalog") or {}).get(pid) or {}
                member_info[pid] = {
                    "project_id": pid,
                    "project_name": display_text(item.get("project_name")) or display_text(cat.get("project_name")),
                    "sheet_name": display_text(item.get("sheet_name")) or display_text(cat.get("sheet_name")),
                    "county": display_text(item.get("county")) or display_text(cat.get("county")),
                }

            primary_original = original_feature_for_project(geo, primary2)
            primary_coords = None
            if primary_original:
                geom = primary_original.get("geometry") or {}
                coords = geom.get("coordinates") or []
                if geom.get("type") == "Point" and len(coords) >= 2:
                    try:
                        lon0, lat0 = float(coords[0]), float(coords[1])
                        if is_valid_wgs84(lon0, lat0):
                            primary_coords = [lon0, lat0]
                    except Exception:
                        pass
            if share_geometry and primary_coords is None:
                item = metadata.get(primary2) or {}
                lon0 = parse_number(item.get("lon"))
                lat0 = parse_number(item.get("lat"))
                if is_valid_wgs84(lon0, lat0):
                    primary_coords = [float(lon0), float(lat0)]
            if share_geometry and primary_coords is None:
                raise ValueError(
                    "圖資主工程目前沒有有效代表點，請先校正主工程代表點，再建立共用圖資。"
                )

            snapshot = group.setdefault("member_original_points_before_group", {})
            linked_ids = list(ids2)

            # 先處理所有既有 feature 的空間群組 metadata。
            for f in geo.get("features", []):
                props = f.get("properties") or {}
                pid = display_text(props.get("project_id"))
                if pid not in linked_ids:
                    continue
                if pid not in snapshot and display_text(props.get("role")) == "original_point":
                    g0 = f.get("geometry") or {}
                    c0 = g0.get("coordinates") or []
                    if g0.get("type") == "Point" and len(c0) >= 2:
                        try:
                            snapshot[pid] = [float(c0[0]), float(c0[1])]
                        except Exception:
                            pass

                props["spatial_project_id"] = spj_id
                props["spatial_primary_project_id"] = primary2
                props["spatial_linked_project_ids"] = linked_ids
                props["spatial_relation_type"] = relation_type
                props["spatial_share_mode"] = "shared" if share_geometry else "related_only"
                props["spatial_group_updated_at"] = now_iso()
                props["spatial_group_updated_by"] = display_text(editor)

                role = display_text(props.get("role"))
                if share_geometry and pid != primary2 and role != "original_point":
                    props["spatial_suppressed"] = True
                elif pid == primary2:
                    props["spatial_suppressed"] = False

            # 共用模式：每個 PRJ 都保留自己的 original_point，
            # 但幾何統一使用 primary，讓系統管理者可同步至各 Excel 管控列。
            if share_geometry:
                for pid in linked_ids:
                    original = original_feature_for_project(geo, pid)
                    item = metadata.get(pid) or {}
                    if original is None:
                        original = {
                            "type": "Feature",
                            "geometry": {"type": "Point", "coordinates": list(primary_coords)},
                            "properties": {
                                "geo_id": next_geo_id(reg),
                                "project_id": pid,
                                "role": "original_point",
                                "feature_active": True,
                                "project_active": True,
                                "created_at": now_iso(),
                            },
                        }
                        geo.setdefault("features", []).append(original)

                    original["geometry"] = {
                        "type": "Point",
                        "coordinates": list(primary_coords),
                    }
                    props = original.setdefault("properties", {})
                    props.update({
                        "project_id": pid,
                        "project_name": display_text(item.get("project_name")) or display_text(props.get("project_name")),
                        "county": display_text(item.get("county")) or display_text(props.get("county")),
                        "role": "original_point",
                        "source": "gis_representative",
                        "feature_active": True,
                        "project_active": True,
                        "coordinate_authority": "gis",
                        "coordinate_edit_source": "spatial_shared_group",
                        "coordinate_updated_at": now_iso(),
                        "coordinate_updated_by": display_text(editor),
                        "excel_sync_status": "pending",
                        "coord_source": "共用GIS代表點",
                        "coord_status": "共用圖資代表點；Excel待管理者同步",
                        "spatial_project_id": spj_id,
                        "spatial_primary_project_id": primary2,
                        "spatial_linked_project_ids": linked_ids,
                        "spatial_relation_type": relation_type,
                        "spatial_share_mode": "shared",
                        "spatial_shared_secondary": pid != primary2,
                        "spatial_suppressed": False,
                        "updated_at": now_iso(),
                    })

            group.update({
                "spatial_project_id": spj_id,
                "active": True,
                "primary_project_id": primary2,
                "linked_project_ids": linked_ids,
                "relation_type": relation_type,
                "share_mode": "shared" if share_geometry else "related_only",
                "member_info": member_info,
                "note": display_text(note),
                "updated_at": now_iso(),
                "updated_by": display_text(editor),
            })
            reg["spatial_project_groups"][spj_id] = group

            # 建立群組後，不再讓同一 member signature 被「非重複」舊判斷阻擋。
            signature = _spatial_group_signature_v3651(linked_ids)
            reg["duplicate_project_reviews"].pop(signature, None)

            return {
                self.cfg.geo_path: json_dump_bytes(geo),
                self.cfg.registry_path: json_dump_bytes(reg),
            }, {
                "spatial_project_id": spj_id,
                "primary_project_id": primary2,
                "linked_project_ids": linked_ids,
                "share_mode": group["share_mode"],
                "_new_geo": geo,
                "_new_registry": reg,
            }

        return self.store.atomic_update(
            [self.cfg.geo_path, self.cfg.registry_path],
            mutate,
            f"建立/更新共用空間工程群組：{primary}",
            clear_read_caches=False,
        )

    def confirm_duplicate_projects_distinct_v3651(
        self,
        project_ids: Sequence[str],
        note: str,
        editor: str,
    ) -> Dict[str, Any]:
        ids = sorted({display_text(x) for x in project_ids if display_text(x)})
        if len(ids) < 2:
            raise ValueError("至少需要 2 個工程ID。")
        signature = _spatial_group_signature_v3651(ids)

        def mutate(cur):
            reg = json_load_bytes(cur.get(self.cfg.registry_path), empty_registry())
            reviews = reg.setdefault("duplicate_project_reviews", {})
            reviews[signature] = {
                "project_ids": ids,
                "decision": "confirmed_distinct",
                "note": display_text(note),
                "reviewed_by": display_text(editor),
                "reviewed_at": now_iso(),
            }
            return {
                self.cfg.registry_path: json_dump_bytes(reg),
            }, {
                "signature": signature,
                "project_ids": ids,
                "_new_registry": reg,
            }

        return self.store.atomic_update(
            [self.cfg.registry_path],
            mutate,
            "確認同名工程不是同一實體工程",
            clear_read_caches=False,
        )

    def dissolve_spatial_project_group_v3651(
        self,
        spatial_project_id: str,
        editor: str,
    ) -> Dict[str, Any]:
        spj_id = display_text(spatial_project_id)
        if not spj_id:
            raise ValueError("缺少 SPJ-ID。")

        def mutate(cur):
            geo = json_load_bytes(cur.get(self.cfg.geo_path), empty_geojson())
            reg = json_load_bytes(cur.get(self.cfg.registry_path), empty_registry())
            groups = reg.setdefault("spatial_project_groups", {})
            group = groups.get(spj_id)
            if not isinstance(group, dict) or not group.get("active", True):
                raise ValueError(f"{spj_id} 目前不是有效群組。")

            members = {
                display_text(x)
                for x in group.get("linked_project_ids") or []
                if display_text(x)
            }
            for f in geo.get("features", []):
                props = f.get("properties") or {}
                pid = display_text(props.get("project_id"))
                if pid not in members:
                    continue
                for key in [
                    "spatial_project_id",
                    "spatial_primary_project_id",
                    "spatial_linked_project_ids",
                    "spatial_relation_type",
                    "spatial_share_mode",
                    "spatial_shared_secondary",
                    "spatial_suppressed",
                    "spatial_group_updated_at",
                    "spatial_group_updated_by",
                ]:
                    props.pop(key, None)

                # 解除群組不自動復原舊座標，避免把群組建立後的合法校正倒退。
                # 各 PRJ 保留當下座標，若需分開位置可再用代表點校正。
                if display_text(props.get("role")) == "original_point":
                    props["coordinate_authority"] = "gis"
                    props["coordinate_edit_source"] = "spatial_group_dissolved"
                    props["excel_sync_status"] = "pending"
                    props["coord_source"] = "GIS代表點"
                    props["coord_status"] = "已解除共用圖資；Excel待管理者同步"
                    props["coordinate_updated_at"] = now_iso()
                    props["coordinate_updated_by"] = display_text(editor)

            group["active"] = False
            group["dissolved_at"] = now_iso()
            group["dissolved_by"] = display_text(editor)

            return {
                self.cfg.geo_path: json_dump_bytes(geo),
                self.cfg.registry_path: json_dump_bytes(reg),
            }, {
                "spatial_project_id": spj_id,
                "linked_project_ids": sorted(members),
                "_new_geo": geo,
                "_new_registry": reg,
            }

        return self.store.atomic_update(
            [self.cfg.geo_path, self.cfg.registry_path],
            mutate,
            f"解除共用空間工程群組：{spj_id}",
            clear_read_caches=False,
        )

    def assign_new_projects(self, selected: Sequence[Tuple[str, str]]) -> Dict[str, Any]:
        """selected: [(row_key, expected_project_name), ...]"""
        if not selected:
            raise ValueError("請至少選擇一筆待編號工程。")

        def mutate(cur):
            excel_bytes = cur[self.cfg.excel_path]
            geo = json_load_bytes(cur[self.cfg.geo_path], empty_geojson())
            hist = json_load_bytes(cur[self.cfg.history_path], empty_history())
            reg = json_load_bytes(cur[self.cfg.registry_path], empty_registry())
            wb, rows, _ = scan_workbook(excel_bytes, self.cfg.excel_path, ensure_id_cols=True)
            sync_registry_with_existing_ids(reg, rows, geo, hist)
            assigned = []
            for row_key, name in selected:
                sheet, row = split_row_key(row_key)
                pid = next_root_id(reg)
                set_row_id(wb, sheet, row, name, pid)
                assigned.append((pid, name))
            new_excel = _save_wb_bytes(wb)
            _, rows2, _ = scan_workbook(new_excel, self.cfg.excel_path)
            sync_original_features(geo, reg, rows2)
            sync_registry_with_existing_ids(reg, rows2, geo, hist)
            changed = {
                self.cfg.excel_path: new_excel,
                self.cfg.geo_path: json_dump_bytes(geo),
                self.cfg.registry_path: json_dump_bytes(reg),
            }
            return changed, {"assigned": assigned}

        return self.store.atomic_update(self.paths, mutate, "新增核定工程並自動建立PRJ-ID")

    def import_historical_project_catalog(self, excel_bytes: bytes, source_name: str = "舊版管控表") -> Dict[str, Any]:
        """從舊版管控表只匯入工程識別資料到永久歷史目錄，不改 current.xlsx。"""
        if not excel_bytes:
            raise ValueError("沒有收到舊版管控表檔案。")
        try:
            _, old_rows, _ = scan_workbook(excel_bytes, source_name or "historical.xlsx", ensure_id_cols=False)
        except Exception as exc:
            raise ValueError(f"無法讀取舊版管控表：{exc}") from exc
        candidates = [p for p in old_rows if p.project_id and PROJECT_ID_RE.match(p.project_id)]
        if not candidates:
            raise ValueError("舊版管控表中找不到有效的『系統工程ID』資料。")

        def mutate(cur):
            reg = json_load_bytes(cur[self.cfg.registry_path], empty_registry())
            geo = json_load_bytes(cur[self.cfg.geo_path], empty_geojson())
            hist = json_load_bytes(cur[self.cfg.history_path], empty_history())
            # 先把目前系統已知資料補入 catalog，再匯入舊表。
            current_excel = cur.get(self.cfg.excel_path)
            current_rows: List[ProjectRow] = []
            if current_excel:
                try:
                    _, current_rows, _ = scan_workbook(current_excel, self.cfg.excel_path, ensure_id_cols=False)
                except Exception:
                    current_rows = []
            sync_registry_with_existing_ids(reg, current_rows, geo, hist)

            before = set(project_catalog_map(reg).keys())
            imported = 0
            for oldp in candidates:
                if _catalog_upsert_project(
                    reg, oldp.project_id, oldp.project_name,
                    sheet_name=oldp.sheet_name, county=oldp.county, source="historical_excel"
                ):
                    imported += 1
                rn = root_number(oldp.project_id)
                if rn is not None:
                    reg.setdefault("counters", {}).setdefault("PRJ", 0)
                    reg["counters"]["PRJ"] = max(int(reg["counters"].get("PRJ", 0)), rn)
                pc = infer_parent_and_child_no(oldp.project_id)
                if pc:
                    parent, child = pc
                    cc = reg.setdefault("child_counters", {})
                    cc[parent] = max(int(cc.get(parent, 0)), child)
            after = set(project_catalog_map(reg).keys())
            return {self.cfg.registry_path: json_dump_bytes(reg)}, {
                "rows_with_id": len(candidates),
                "catalog_total": len(after),
                "new_ids": len(after - before),
                "updated_entries": imported,
            }

        return self.store.atomic_update(
            [self.cfg.excel_path, self.cfg.geo_path, self.cfg.history_path, self.cfg.registry_path],
            mutate,
            f"匯入歷史工程目錄：{display_text(source_name) or '舊版管控表'}",
        )

    def split_project(self, parent_id: str, selected: Sequence[Tuple[str, str]]) -> Dict[str, Any]:
        if len(selected) < 2:
            raise ValueError("分標至少請選擇 2 筆新的工程列。")

        def mutate(cur):
            excel_bytes = cur[self.cfg.excel_path]
            geo = json_load_bytes(cur[self.cfg.geo_path], empty_geojson())
            hist = json_load_bytes(cur[self.cfg.history_path], empty_history())
            reg = json_load_bytes(cur[self.cfg.registry_path], empty_registry())
            wb, rows, _ = scan_workbook(excel_bytes, self.cfg.excel_path, ensure_id_cols=True)
            sync_registry_with_existing_ids(reg, rows, geo, hist)
            known = all_known_projects(rows, geo, hist, reg)
            eligible = eligible_historical_source_projects(rows, geo, hist, reg)
            if parent_id not in known:
                raise ValueError(f"找不到母工程 {parent_id}")
            if parent_id not in eligible:
                raise ValueError(
                    f"{parent_id} 目前不是可用的歷史母工程。"
                    "來源工程必須已不在 current.xlsx，且尚未完成過其他分併標。"
                )

            children = []
            for row_key, name in selected:
                child_id = next_child_id(reg, parent_id)
                sheet, row = split_row_key(row_key)
                set_row_id(wb, sheet, row, name, child_id)
                children.append((child_id, name))
                hist.setdefault("events", []).append(
                    {
                        "history_id": next_history_id(reg),
                        "event_type": "分標",
                        "from_project_id": parent_id,
                        "from_project_name": known.get(parent_id, ""),
                        "to_project_id": child_id,
                        "to_project_name": name,
                        "created_at": now_iso(),
                    }
                )

            set_project_inactive(geo, parent_id)
            new_excel = _save_wb_bytes(wb)
            _, rows2, _ = scan_workbook(new_excel, self.cfg.excel_path)
            sync_original_features(geo, reg, rows2)
            sync_registry_with_existing_ids(reg, rows2, geo, hist)
            changed = {
                self.cfg.excel_path: new_excel,
                self.cfg.geo_path: json_dump_bytes(geo),
                self.cfg.history_path: json_dump_bytes(hist),
                self.cfg.registry_path: json_dump_bytes(reg),
            }
            return changed, {"parent": parent_id, "children": children}

        return self.store.atomic_update(self.paths, mutate, f"工程分標：{parent_id}")

    def merge_projects(self, source_ids: Sequence[str], target: Tuple[str, str]) -> Dict[str, Any]:
        source_ids = list(dict.fromkeys(source_ids))
        if len(source_ids) < 2:
            raise ValueError("併標至少需要選擇 2 個來源工程。")

        def mutate(cur):
            excel_bytes = cur[self.cfg.excel_path]
            geo = json_load_bytes(cur[self.cfg.geo_path], empty_geojson())
            hist = json_load_bytes(cur[self.cfg.history_path], empty_history())
            reg = json_load_bytes(cur[self.cfg.registry_path], empty_registry())
            wb, rows, _ = scan_workbook(excel_bytes, self.cfg.excel_path, ensure_id_cols=True)
            sync_registry_with_existing_ids(reg, rows, geo, hist)
            known = all_known_projects(rows, geo, hist, reg)
            eligible = eligible_historical_source_projects(rows, geo, hist, reg)
            missing = [x for x in source_ids if x not in known]
            if missing:
                raise ValueError(f"找不到來源工程：{', '.join(missing)}")
            unavailable = [x for x in source_ids if x not in eligible]
            if unavailable:
                raise ValueError(
                    "以下來源工程目前不可再使用：" + "、".join(unavailable) +
                    "。來源工程必須已不在 current.xlsx，且尚未完成過其他分併標。"
                )

            row_key, target_name = target
            new_id = next_root_id(reg)
            sheet, row = split_row_key(row_key)
            set_row_id(wb, sheet, row, target_name, new_id)

            for sid in source_ids:
                hist.setdefault("events", []).append(
                    {
                        "history_id": next_history_id(reg),
                        "event_type": "併標",
                        "from_project_id": sid,
                        "from_project_name": known.get(sid, ""),
                        "to_project_id": new_id,
                        "to_project_name": target_name,
                        "created_at": now_iso(),
                    }
                )
                set_project_inactive(geo, sid)

            new_excel = _save_wb_bytes(wb)
            _, rows2, _ = scan_workbook(new_excel, self.cfg.excel_path)
            sync_original_features(geo, reg, rows2)
            sync_registry_with_existing_ids(reg, rows2, geo, hist)
            changed = {
                self.cfg.excel_path: new_excel,
                self.cfg.geo_path: json_dump_bytes(geo),
                self.cfg.history_path: json_dump_bytes(hist),
                self.cfg.registry_path: json_dump_bytes(reg),
            }
            return changed, {"new_id": new_id, "target_name": target_name, "sources": source_ids}

        return self.store.atomic_update(self.paths, mutate, "工程併標並建立新PRJ-ID")

    def restructure_projects(
        self,
        source_ids: Sequence[str],
        targets: Sequence[Tuple[str, str]],
    ) -> Dict[str, Any]:
        """V3.6.39：多對多分併標重組，例如 2標→3標、5標→3標。

        多來源沒有唯一母工程，因此每個新工程都建立新的根 PRJ-ID；
        沿革以單一事件保存整組來源與整組新工程。
        """
        source_ids = list(dict.fromkeys(display_text(x) for x in source_ids if display_text(x)))
        targets = list(dict.fromkeys((display_text(rk), display_text(name)) for rk, name in targets if display_text(rk)))
        if len(source_ids) < 2:
            raise ValueError("多對多重組至少需要 2 個來源工程。")
        if len(targets) < 2:
            raise ValueError("多對多重組至少需要 2 筆重組後的新工程。")

        def mutate(cur):
            excel_bytes = cur[self.cfg.excel_path]
            geo = json_load_bytes(cur[self.cfg.geo_path], empty_geojson())
            hist = json_load_bytes(cur[self.cfg.history_path], empty_history())
            reg = json_load_bytes(cur[self.cfg.registry_path], empty_registry())
            wb, rows, _ = scan_workbook(excel_bytes, self.cfg.excel_path, ensure_id_cols=True)
            sync_registry_with_existing_ids(reg, rows, geo, hist)

            known = all_known_projects(rows, geo, hist, reg)
            eligible = eligible_historical_source_projects(rows, geo, hist, reg)
            missing = [x for x in source_ids if x not in known]
            if missing:
                raise ValueError("找不到來源工程：" + "、".join(missing))
            unavailable = [x for x in source_ids if x not in eligible]
            if unavailable:
                raise ValueError(
                    "以下來源工程目前不可再使用：" + "、".join(unavailable) +
                    "。來源工程必須已不在 current.xlsx，且尚未完成過其他分併標。"
                )

            created: List[Tuple[str, str, str]] = []
            for row_key, target_name in targets:
                new_id = next_root_id(reg)
                sheet, row = split_row_key(row_key)
                set_row_id(wb, sheet, row, target_name, new_id)
                created.append((row_key, new_id, target_name))

            event = {
                "history_id": next_history_id(reg),
                "event_type": "分併標重組",
                "from_project_ids": list(source_ids),
                "from_project_names": [known.get(pid, "") for pid in source_ids],
                "to_project_ids": [new_id for _rk, new_id, _name in created],
                "to_project_names": [name for _rk, _new_id, name in created],
                "source_count": len(source_ids),
                "target_count": len(created),
                "created_at": now_iso(),
            }
            hist.setdefault("events", []).append(event)

            for sid in source_ids:
                set_project_inactive(geo, sid)

            new_excel = _save_wb_bytes(wb)
            _, rows2, _ = scan_workbook(new_excel, self.cfg.excel_path)
            sync_original_features(geo, reg, rows2)
            sync_registry_with_existing_ids(reg, rows2, geo, hist)
            changed = {
                self.cfg.excel_path: new_excel,
                self.cfg.geo_path: json_dump_bytes(geo),
                self.cfg.history_path: json_dump_bytes(hist),
                self.cfg.registry_path: json_dump_bytes(reg),
            }
            return changed, {
                "sources": list(source_ids),
                "targets": created,
                "history_id": event["history_id"],
            }

        return self.store.atomic_update(
            self.paths,
            mutate,
            f"工程分併標重組：{len(source_ids)}標→{len(targets)}標",
        )

    def save_manual_drawings(
        self,
        project_id: str,
        project_name: str,
        drawings: Sequence[Dict[str, Any]],
        geo_name: str = "",
        replace_manual: bool = False,
    ) -> Dict[str, Any]:
        allowed = {"Point", "LineString", "Polygon"}
        drawings2 = [d for d in drawings if geometry_type(d) in allowed]
        if not drawings2:
            raise ValueError("沒有可儲存的點、線或面。")

        def mutate(cur):
            geo = json_load_bytes(cur[self.cfg.geo_path], empty_geojson())
            hist = json_load_bytes(cur[self.cfg.history_path], empty_history())
            reg = json_load_bytes(cur[self.cfg.registry_path], empty_registry())
            excel_bytes = cur[self.cfg.excel_path]
            _, rows, _ = scan_workbook(excel_bytes, self.cfg.excel_path)
            sync_registry_with_existing_ids(reg, rows, geo, hist)
            known = workbook_to_project_map(rows)
            target_project_id = _spatial_primary_from_registry_v3651(reg, project_id)
            if target_project_id not in known:
                raise ValueError(f"目前工程資料庫找不到 {target_project_id}，請先同步工程資料。")
            if target_project_id != project_id:
                project_name = known[target_project_id].project_name

            if replace_manual:
                for f in geo.get("features", []):
                    p = f.get("properties") or {}
                    if p.get("project_id") == target_project_id and p.get("source") == "manual" and p.get("feature_active", True):
                        p["feature_active"] = False
                        p["superseded_at"] = now_iso()

            created = []
            for idx, drawing in enumerate(drawings2, 1):
                gid = next_geo_id(reg)
                name = geo_name.strip() if geo_name.strip() else f"人工圖資{idx}"
                if len(drawings2) > 1 and geo_name.strip():
                    name = f"{geo_name.strip()}-{idx}"
                feature = {
                    "type": "Feature",
                    "geometry": copy.deepcopy(drawing.get("geometry")),
                    "properties": {
                        "geo_id": gid,
                        "project_id": target_project_id,
                        "project_name": project_name,
                        "geo_name": name,
                        "role": "manual_geometry",
                        "source": "manual",
                        "feature_active": True,
                        "project_active": True,
                        "created_at": now_iso(),
                        "updated_at": now_iso(),
                    },
                }
                geo.setdefault("features", []).append(feature)
                created.append(gid)

            changed = {
                self.cfg.geo_path: json_dump_bytes(geo),
                self.cfg.registry_path: json_dump_bytes(reg),
            }
            return changed, {"created_geo_ids": created}

        return self.store.atomic_update(self.paths, mutate, f"更新工程圖資：{target_project_id}")

    def deactivate_manual_features(self, project_id: str, geo_ids: Sequence[str]) -> Dict[str, Any]:
        targets = set(geo_ids)
        if not targets:
            raise ValueError("請選擇要刪除的人工圖資。")

        def mutate(cur):
            geo = json_load_bytes(cur[self.cfg.geo_path], empty_geojson())
            changed_ids = []
            for f in geo.get("features", []):
                p = f.get("properties") or {}
                if p.get("project_id") == project_id and p.get("geo_id") in targets and p.get("source") == "manual":
                    p["feature_active"] = False
                    p["deleted_at"] = now_iso()
                    changed_ids.append(p.get("geo_id"))
            return {self.cfg.geo_path: json_dump_bytes(geo)}, {"deactivated": changed_ids}

        return self.store.atomic_update(self.paths, mutate, f"停用工程人工圖資：{project_id}")


    def update_project_representative_point_gis_only(
        self,
        project_id: str,
        lon: float,
        lat: float,
        editor: str = "",
        project_name: str = "",
        county: str = "",
        status: str = "",
    ) -> Dict[str, Any]:
        """V3.6.8：校正工程代表點時只更新 GIS，不寫 current.xlsx。

        GIS 代表點自此成為最新座標；current.xlsx 由系統管理者日後批次同步。
        若工程原先沒有 original_point，會同時使用 registry 建立新的 GEO-ID。
        """
        pid = display_text(project_id)
        if not pid:
            raise ValueError("缺少系統工程ID，無法儲存工程代表點。")
        lon = float(lon)
        lat = float(lat)
        if not is_valid_wgs84(lon, lat):
            raise ValueError("新座標不在臺灣及離島合理 WGS84 範圍。")

        def mutate(cur):
            geo = json_load_bytes(cur.get(self.cfg.geo_path), empty_geojson())
            reg = json_load_bytes(cur.get(self.cfg.registry_path), empty_registry())

            target_pid = _spatial_primary_from_registry_v3651(reg, pid)
            spj_id, spj_group = _active_spatial_group_for_project_v3651(reg, target_pid)
            shared_members = (
                [display_text(x) for x in spj_group.get("linked_project_ids") or [] if display_text(x)]
                if spj_id and _spatial_group_is_shared_v3651(spj_group)
                else [target_pid]
            )

            original = original_feature_for_project(geo, target_pid)
            created = False
            if original is None:
                original = {
                    "type": "Feature",
                    "geometry": {"type": "Point", "coordinates": [lon, lat]},
                    "properties": {
                        "geo_id": next_geo_id(reg),
                        "project_id": target_pid,
                        "role": "original_point",
                        "feature_active": True,
                        "project_active": True,
                        "created_at": now_iso(),
                    },
                }
                geo.setdefault("features", []).append(original)
                created = True
            else:
                original["geometry"] = {"type": "Point", "coordinates": [lon, lat]}

            props = original.setdefault("properties", {})
            props.update({
                "project_id": target_pid,
                "project_name": display_text(project_name) or display_text(props.get("project_name")),
                "county": display_text(county) or display_text(props.get("county")),
                "status_snapshot": display_text(status) or display_text(props.get("status_snapshot")),
                "role": "original_point",
                "source": "gis_representative",
                "feature_active": True,
                "project_active": True,
                "coordinate_authority": "gis",
                "coordinate_edit_source": "gis_map",
                "coordinate_updated_at": now_iso(),
                "coordinate_updated_by": display_text(editor) or "未填編輯者",
                "excel_sync_status": "pending",
                "coord_source": "GIS代表點",
                "coord_status": "GIS為最新座標；Excel待管理者同步",
                "updated_at": now_iso(),
            })

            # 共用圖資群組：主工程代表點一旦改變，所有 linked PRJ 的 GIS original_point
            # 一起更新，Excel 仍等待系統管理者批次同步。
            affected = []
            if spj_id and _spatial_group_is_shared_v3651(spj_group):
                linked_ids = list(shared_members)
                for member_pid in linked_ids:
                    mf = original_feature_for_project(geo, member_pid)
                    if mf is None:
                        mf = {
                            "type": "Feature",
                            "geometry": {"type": "Point", "coordinates": [lon, lat]},
                            "properties": {
                                "geo_id": next_geo_id(reg),
                                "project_id": member_pid,
                                "role": "original_point",
                                "feature_active": True,
                                "project_active": True,
                                "created_at": now_iso(),
                            },
                        }
                        geo.setdefault("features", []).append(mf)
                        created = True
                    mf["geometry"] = {"type": "Point", "coordinates": [lon, lat]}
                    mp = mf.setdefault("properties", {})
                    mp.update({
                        "source": "gis_representative",
                        "coordinate_authority": "gis",
                        "coordinate_edit_source": "spatial_shared_group",
                        "coordinate_updated_at": now_iso(),
                        "coordinate_updated_by": display_text(editor) or "未填編輯者",
                        "excel_sync_status": "pending",
                        "coord_source": "共用GIS代表點",
                        "coord_status": "共用圖資代表點；Excel待管理者同步",
                        "spatial_project_id": spj_id,
                        "spatial_primary_project_id": target_pid,
                        "spatial_linked_project_ids": linked_ids,
                        "spatial_relation_type": display_text(spj_group.get("relation_type")),
                        "spatial_share_mode": "shared",
                        "spatial_shared_secondary": member_pid != target_pid,
                        "updated_at": now_iso(),
                    })
                    affected.append(member_pid)
            else:
                affected = [target_pid]

            changed = {self.cfg.geo_path: json_dump_bytes(geo)}
            if created:
                changed[self.cfg.registry_path] = json_dump_bytes(reg)
            return changed, {
                "project_id": target_pid,
                "project_name": props.get("project_name", ""),
                "lon": lon,
                "lat": lat,
                "created": created,
                "affected_project_ids": affected,
                "_new_geo": geo,
            }

        paths = [self.cfg.geo_path, self.cfg.registry_path]
        return self.store.atomic_update(
            paths,
            mutate,
            f"校正工程GIS代表點：{pid}",
            clear_read_caches=False,
        )

    def sync_gis_representative_points_to_excel(
        self,
        project_ids: Sequence[str],
        editor: str = "",
    ) -> Dict[str, Any]:
        """V3.6.8：由系統管理者把 GIS 最新代表點批次同步至 current.xlsx。

        僅修改座標儲存格 XML，不以 OpenPyXL 重存整本活頁簿，以保護公式 cached values。
        """
        ids = []
        seen = set()
        for x in project_ids or []:
            pid = display_text(x)
            if pid and pid not in seen:
                ids.append(pid)
                seen.add(pid)
        if not ids:
            raise ValueError("目前沒有需要同步至 Excel 的工程代表點。")

        def mutate(cur):
            excel_bytes = cur.get(self.cfg.excel_path)
            if not excel_bytes:
                raise ValueError(f"找不到 {self.cfg.excel_path}")
            geo = json_load_bytes(cur.get(self.cfg.geo_path), empty_geojson())
            wb, rows, _ = scan_workbook(excel_bytes, self.cfg.excel_path, ensure_id_cols=False)
            id_map: Dict[str, List[ProjectRow]] = {}
            for project in rows:
                if project.project_id:
                    id_map.setdefault(project.project_id, []).append(project)

            patches: List[Dict[str, Any]] = []
            items: List[Dict[str, Any]] = []
            for pid in ids:
                matches = id_map.get(pid, [])
                if not matches:
                    raise ValueError(f"current.xlsx 找不到系統工程ID：{pid}")
                if len(matches) > 1:
                    raise ValueError(f"current.xlsx 發現重複系統工程ID：{pid}，已停止同步。")
                f = original_feature_for_project(geo, pid)
                geom = (f or {}).get("geometry") or {}
                coords = geom.get("coordinates") or []
                if geom.get("type") != "Point" or len(coords) < 2:
                    raise ValueError(f"{pid} 目前沒有可同步的 GIS 工程代表點。")
                lon, lat = float(coords[0]), float(coords[1])
                project = matches[0]
                point_patches, write_info = _project_coordinate_xml_patches(wb, project, lon, lat)
                patches.extend(point_patches)
                props = f.setdefault("properties", {})
                is_county_return = (
                    display_text(props.get("coordinate_edit_source"))
                    == "county_return_review"
                    or display_text(props.get("coordinate_origin"))
                    == "縣市回填核准"
                )
                props.update({
                    "coordinate_authority": "gis",
                    "excel_sync_status": "synced",
                    "excel_synced_at": now_iso(),
                    "excel_synced_by": display_text(editor) or "系統管理者",
                    "coord_source": (
                        "縣市回填核准"
                        if is_county_return
                        else "GIS代表點"
                    ),
                    "coord_status": (
                        "縣市回填已核准；GIS／Excel已同步"
                        if is_county_return
                        else "GIS／Excel已同步"
                    ),
                    "updated_at": now_iso(),
                })
                items.append({
                    "project_id": pid,
                    "project_name": project.project_name,
                    "sheet_name": project.sheet_name,
                    "county": project.county,
                    "lon": lon,
                    "lat": lat,
                    "written_headers": list(write_info.get("written_headers") or []),
                })

            new_excel = _patch_xlsx_numeric_cells(excel_bytes, patches)
            return {
                self.cfg.excel_path: new_excel,
                self.cfg.geo_path: json_dump_bytes(geo),
            }, {
                "count": len(items),
                "items": items,
                "_new_excel_bytes": new_excel,
                "_new_geo": geo,
            }

        return self.store.atomic_update(
            [self.cfg.excel_path, self.cfg.geo_path],
            mutate,
            f"系統管理者批次同步GIS代表點至Excel：{len(ids)}件",
            clear_read_caches=True,
        )


    def update_project_representative_points_batch(
        self,
        updates: Sequence[Dict[str, Any]],
        editor: str = "",
    ) -> Dict[str, Any]:
        """一次同步多件工程代表點到 current.xlsx 與 engineering_geo.geojson。

        V3.6.6 快速批次模式：
        - 整批只下載 current.xlsx / GEO 各一次。
        - 整本 Excel 只掃描一次、儲存一次。
        - 整批只產生一個 GitHub commit，避免每移動一點就觸發一次 Streamlit 更新。
        """
        normalized = []
        seen = set()
        for item in updates or []:
            pid = display_text((item or {}).get("project_id"))
            if not pid or pid in seen:
                continue
            lon = float((item or {}).get("lon"))
            lat = float((item or {}).get("lat"))
            if not is_valid_wgs84(lon, lat):
                raise ValueError(f"{pid} 的新座標不在臺灣及離島合理 WGS84 範圍。")
            normalized.append({"project_id": pid, "lon": lon, "lat": lat})
            seen.add(pid)
        if not normalized:
            raise ValueError("目前沒有待同步的工程代表點。")

        def mutate(cur):
            excel_bytes = cur.get(self.cfg.excel_path)
            if not excel_bytes:
                raise ValueError(f"找不到 {self.cfg.excel_path}")
            geo = json_load_bytes(cur.get(self.cfg.geo_path), empty_geojson())

            # 整批只掃描一次 current.xlsx。
            wb, rows, _ = scan_workbook(excel_bytes, self.cfg.excel_path)
            id_map: Dict[str, List[ProjectRow]] = {}
            for project in rows:
                if project.project_id:
                    id_map.setdefault(project.project_id, []).append(project)

            result_rows = []
            excel_cell_patches: List[Dict[str, Any]] = []
            for item in normalized:
                pid = item["project_id"]
                matches = id_map.get(pid, [])
                if not matches:
                    raise ValueError(f"current.xlsx 找不到系統工程ID：{pid}")
                if len(matches) > 1:
                    raise ValueError(f"current.xlsx 發現重複系統工程ID：{pid}，為避免錯寫已停止更新。")
                project = matches[0]
                lon, lat = item["lon"], item["lat"]
                point_patches, write_info = _project_coordinate_xml_patches(wb, project, lon, lat)
                excel_cell_patches.extend(point_patches)

                original = original_feature_for_project(geo, pid)
                if original is None:
                    original = {
                        "type": "Feature",
                        "geometry": {"type": "Point", "coordinates": [lon, lat]},
                        "properties": {
                            "geo_id": "",
                            "project_id": pid,
                            "source": "excel_original",
                            "role": "original_point",
                            "active": True,
                            "project_name": project.project_name,
                            "county": project.county,
                            "status": project.status,
                        },
                    }
                    geo.setdefault("features", []).append(original)
                else:
                    original["geometry"] = {"type": "Point", "coordinates": [lon, lat]}

                props = original.setdefault("properties", {})
                props.update({
                    "project_id": pid,
                    "project_name": project.project_name,
                    "county": project.county,
                    "status": project.status,
                    "source": props.get("source") or "excel_original",
                    "role": "original_point",
                    "active": True,
                    "coordinate_edit_source": "gis_map",
                    "coordinate_updated_at": now_iso(),
                    "coordinate_updated_by": display_text(editor) or "未填編輯者",
                })
                result_rows.append({
                    "project_id": pid,
                    "project_name": project.project_name,
                    "county": project.county,
                    "lon": lon,
                    "lat": lat,
                    "written_headers": list(write_info.get("written_headers") or []),
                })

            # 只修改座標儲存格 XML，不以 OpenPyXL 重存整本活頁簿；
            # 這樣可保留治理工程經費公式既有 cached values。
            new_excel = _patch_xlsx_numeric_cells(excel_bytes, excel_cell_patches)
            new_geo_bytes = json_dump_bytes(geo)
            changed = {
                self.cfg.excel_path: new_excel,
                self.cfg.geo_path: new_geo_bytes,
            }
            return changed, {
                "count": len(result_rows),
                "items": result_rows,
                "_new_excel_bytes": new_excel,
                "_new_geo": geo,
            }

        return self.store.atomic_update(
            [self.cfg.excel_path, self.cfg.geo_path],
            mutate,
            f"批次校正工程代表點並同步Excel：{len(normalized)}件",
            clear_read_caches=False,
        )

    def update_project_representative_point(
        self,
        project_id: str,
        lon: float,
        lat: float,
        editor: str = "",
    ) -> Dict[str, Any]:
        """相容舊呼叫；V3.6.8 起單件代表點只寫 GIS，不再自動寫 current.xlsx。"""
        return self.update_project_representative_point_gis_only(
            project_id, lon, lat, editor=editor
        )


# ============================================================
# 地圖顯示
# ============================================================
def active_project_edit_features(geo: Dict[str, Any], project_id: str) -> List[Dict[str, Any]]:
    """回傳工程目前有效、可由「工程圖資編輯」直接載入的人工／既有圖資。

    V3.6.27：舊版 GEO 的 role/source 欄位曾有多種寫法，若仍只接受
    role=manual_geometry 或 source=manual*，縣市預覽會漏掉早期已完成圖資。
    因此改以「同 project_id、有效、幾何為 Point/LineString/Polygon，且不是
    original_point」為主判斷；original_point 仍由代表點校正功能獨立管理。
    """
    out: List[Dict[str, Any]] = []
    pid = display_text(project_id)
    if not pid:
        return out
    pid = _spatial_primary_project_id_from_geo_v3651(geo, pid)
    for f in geo.get("features", []):
        if not isinstance(f, dict) or not f.get("geometry"):
            continue
        props = f.get("properties") or {}
        if display_text(props.get("project_id")) != pid:
            continue
        if not props.get("feature_active", True) or not props.get("project_active", True):
            continue
        if props.get("spatial_suppressed"):
            continue
        gt = geometry_type(f)
        if gt not in {"Point", "LineString", "Polygon"}:
            continue
        role = display_text(props.get("role")).lower()
        source = display_text(props.get("source")).lower()
        # 系統代表點永遠不混入人工工程圖資編輯器。
        if role == "original_point" or source in {"excel_auto", "gis_representative"}:
            continue
        out.append(f)
    return out


def active_manual_features(geo: Dict[str, Any], project_id: str) -> List[Dict[str, Any]]:
    """相容舊呼叫；V3.6.27 起統一使用較寬鬆的工程可編輯圖資判定。"""
    return active_project_edit_features(geo, project_id)


def original_feature(geo: Dict[str, Any], project_id: str) -> Optional[Dict[str, Any]]:
    project_id = _spatial_primary_project_id_from_geo_v3651(geo, project_id)
    f = original_feature_for_project(geo, project_id)
    if f and f.get("geometry") and (f.get("properties") or {}).get("feature_active", True):
        return f
    return None


def feature_center(feature: Dict[str, Any]) -> Optional[Tuple[float, float]]:
    g = feature.get("geometry") or {}
    t = g.get("type")
    c = g.get("coordinates")
    if not c:
        return None
    try:
        if t == "Point":
            return float(c[1]), float(c[0])
        pts = []
        if t == "LineString":
            pts = c
        elif t == "Polygon":
            pts = c[0]
        if pts:
            lons = [float(x[0]) for x in pts]
            lats = [float(x[1]) for x in pts]
            return sum(lats) / len(lats), sum(lons) / len(lons)
    except Exception:
        return None
    return None


def project_popup_html(p: ProjectRow) -> str:
    def esc(s: Any) -> str:
        return (
            display_text(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    return (
        f"<b>{esc(p.project_name)}</b><br>"
        f"工程ID：{esc(p.project_id)}<br>"
        f"工程類別：{esc(p.project_type)}<br>"
        f"縣市：{esc(p.county)}<br>"
        f"執行單位：{esc(p.unit)}<br>"
        f"執行情形：{esc(p.status)}"
    )


def add_geometry_to_map(m: folium.Map, f: Dict[str, Any], p: ProjectRow, point_cluster=None, reference: bool = False):
    g = f.get("geometry") or {}
    t = g.get("type")
    coords = g.get("coordinates")
    if not t or not coords:
        return

    color = "#666666" if reference else status_color(p.status)
    popup = folium.Popup(project_popup_html(p), max_width=420)
    tooltip = f"{p.project_name}｜{p.status or '未填'}"

    if t == "Point":
        lon, lat = coords
        marker = folium.CircleMarker(
            location=[lat, lon], radius=5 if reference else 7,
            color="#666666" if reference else "#000000", weight=2 if reference else 2.5,
            fill=True, fill_color=color, fill_opacity=.75 if reference else .95,
            tooltip=tooltip, popup=popup,
        )
        if point_cluster is not None and not reference:
            marker.add_to(point_cluster)
        else:
            marker.add_to(m)
    elif t in {"LineString", "Polygon"}:
        style = {
            "color": color,
            "weight": 5 if t == "LineString" else 3,
            "fillColor": color,
            "fillOpacity": 0.12 if reference else 0.28,
            "opacity": 0.45 if reference else 0.95,
        }
        folium.GeoJson(
            data={"type": "Feature", "geometry": g, "properties": {}},
            style_function=lambda _x, s=style: s,
            tooltip=tooltip,
            popup=popup,
        ).add_to(m)


def build_overview_map(
    rows: Sequence[ProjectRow],
    geo: Dict[str, Any],
    show_original_with_manual: bool = False,
) -> Tuple[folium.Map, Dict[str, int]]:
    m = folium.Map(location=TAIWAN_CENTER, zoom_start=TAIWAN_ZOOM, tiles="OpenStreetMap", control_scale=True)
    _add_leaflet_compat_css(m)
    cluster = MarkerCluster(name="工程點位").add_to(m)
    counts = {"displayed": 0, "no_geometry": 0, "manual_projects": 0}

    for p in _spatial_dedupe_rows_v3651(rows, geo):
        if not p.project_id:
            continue
        manuals = active_project_edit_features(geo, p.project_id)
        orig = original_feature(geo, p.project_id)
        if manuals:
            counts["manual_projects"] += 1
            for f in manuals:
                add_geometry_to_map(m, f, p, point_cluster=cluster)
                counts["displayed"] += 1
            if show_original_with_manual and orig:
                add_geometry_to_map(m, orig, p, point_cluster=None, reference=True)
                counts["displayed"] += 1
        elif orig:
            add_geometry_to_map(m, orig, p, point_cluster=cluster)
            counts["displayed"] += 1
        else:
            counts["no_geometry"] += 1

    folium.LayerControl(collapsed=True).add_to(m)
    AdaptiveVectorScale().add_to(m)
    return m, counts


def filter_rows(rows: Sequence[ProjectRow], keyword: str, sheets: Sequence[str], counties: Sequence[str], statuses: Sequence[str]) -> List[ProjectRow]:
    kw = norm_text(keyword)
    out = []
    for p in rows:
        if sheets and p.sheet_name not in sheets:
            continue
        if counties and p.county not in counties:
            continue
        if statuses and p.status not in statuses:
            continue
        if kw:
            hay = norm_text(" ".join([p.project_name, p.project_id, p.county, p.unit, p.address, p.status]))
            if kw not in hay:
                continue
        out.append(p)
    return out


# ============================================================
# 給既有「工程查詢 → 開啟查詢案件地圖」共用的 GIS 地圖工具
# ============================================================
def load_query_geo_snapshot() -> Dict[str, Any]:
    """讀取工程 GeoJSON；短時間切頁直接使用快取，正式儲存後會自動清除。"""
    settings = GitHubSettings.from_streamlit()
    cached = globals().get("_cached_query_geo_snapshot_v351")
    if cached is not None:
        return cached(
            settings.owner, settings.repo, settings.branch, settings.geo_path, settings.token
        )
    store = GitHubRepoStore(settings)
    commit, _ = store.get_head_commit()
    return json_load_bytes(store.read_file(settings.geo_path, ref=commit, allow_missing=True), empty_geojson())


def _query_feature_popup(rec: Dict[str, Any]) -> str:
    import html as _html
    name = _html.escape(display_text(rec.get("name")) or "未命名工程")
    pid = _html.escape(display_text(rec.get("project_id")))
    plan = _html.escape(display_text(rec.get("plan")))
    unit = _html.escape(display_text(rec.get("unit")))
    status = _html.escape(display_text(rec.get("status")))
    lines = [f"<b>{name}</b>"]
    if pid:
        lines.append(f"工程ID：{pid}")
    if plan:
        lines.append(f"計畫別：{plan}")
    if unit:
        lines.append(f"執行單位：{unit}")
    if status:
        lines.append(f"執行情形：{status}")
    return "<br>".join(lines)


def _add_query_label(m: folium.Map, lat: float, lon: float, name: str) -> None:
    import html as _html
    folium.Marker(
        location=[lat, lon],
        icon=folium.DivIcon(
            icon_size=(300, 40),
            icon_anchor=(-8, 18),
            html=(
                "<div style='font-size:12px;font-weight:600;color:#111827;"
                "background:rgba(255,255,255,.90);border:1px solid #d1d5db;"
                "border-radius:4px;padding:2px 5px;white-space:nowrap;"
                "box-shadow:0 1px 2px rgba(0,0,0,.12);'>"
                + _html.escape(name) + "</div>"
            ),
        ),
    ).add_to(m)


def build_query_results_map(records: Sequence[Dict[str, Any]]) -> Tuple[folium.Map, Dict[str, int]]:
    """
    將既有工程查詢結果畫成 OSM 地圖。
    有 PRJ-ID 且已有人工 Line/Polygon/Point 時優先顯示人工圖資；
    沒有人工圖資時顯示 GIS 原始點；若 GIS 尚未初始化則 fallback 到查詢結果既有座標。
    """
    try:
        geo = load_query_geo_snapshot()
    except Exception:
        geo = empty_geojson()

    prepared = []
    bounds: List[List[float]] = []
    shape_count = 0
    gis_project_count = 0
    fallback_count = 0

    for rec in records:
        pid = display_text(rec.get("project_id"))
        name = display_text(rec.get("name")) or "未命名工程"
        status = display_text(rec.get("status"))
        popup_html = _query_feature_popup(rec)

        features: List[Dict[str, Any]] = []
        if pid:
            manuals = active_manual_features(geo, pid)
            if manuals:
                features = manuals
            else:
                orig = original_feature(geo, pid)
                if orig and orig.get("geometry"):
                    features = [orig]

        if features:
            gis_project_count += 1
            prepared.append((rec, features, popup_html))
            for f in features:
                c = feature_center(f)
                if c:
                    bounds.append([float(c[0]), float(c[1])])
            continue

        lon = parse_number(rec.get("lon"))
        lat = parse_number(rec.get("lat"))
        if is_valid_wgs84(lon, lat):
            fallback_count += 1
            fallback = {
                "type": "Feature",
                "geometry": {"type": "Point", "coordinates": [lon, lat]},
                "properties": {"source": "query_fallback"},
            }
            prepared.append((rec, [fallback], popup_html))
            bounds.append([lat, lon])

    if bounds:
        avg_lat = sum(x[0] for x in bounds) / len(bounds)
        avg_lon = sum(x[1] for x in bounds) / len(bounds)
        zoom = 8 if len(bounds) > 1 else 15
        m = folium.Map(location=[avg_lat, avg_lon], zoom_start=zoom, tiles="OpenStreetMap", control_scale=True)
        _add_leaflet_compat_css(m)
    else:
        m = folium.Map(location=TAIWAN_CENTER, zoom_start=TAIWAN_ZOOM, tiles="OpenStreetMap", control_scale=True)
        _add_leaflet_compat_css(m)

    for rec, features, popup_html in prepared:
        name = display_text(rec.get("name")) or "未命名工程"
        color = status_color(rec.get("status"))
        label_done = False
        for f in features:
            geom = f.get("geometry") or {}
            gt = geom.get("type")
            coords = geom.get("coordinates")
            if gt == "Point" and isinstance(coords, list) and len(coords) >= 2:
                lon, lat = float(coords[0]), float(coords[1])
                folium.CircleMarker(
                    location=[lat, lon], radius=6, color="#000000", weight=2.5,
                    fill=True, fill_color=color, fill_opacity=0.9,
                    popup=folium.Popup(popup_html, max_width=380), tooltip=name,
                ).add_to(m)
                bounds.append([lat, lon])
            elif gt == "LineString" and coords:
                locs = [[float(y), float(x)] for x, y, *_ in coords]
                folium.PolyLine(
                    locations=locs, color="#000000", weight=9.5, opacity=0.96, interactive=False,
                ).add_to(m)
                folium.PolyLine(
                    locations=locs, color=color, weight=6, opacity=0.9,
                    popup=folium.Popup(popup_html, max_width=380), tooltip=name,
                ).add_to(m)
                bounds.extend(locs)
            elif gt == "Polygon" and coords and coords[0]:
                locs = [[float(y), float(x)] for x, y, *_ in coords[0]]
                folium.Polygon(
                    locations=locs, color="#000000", weight=4.5,
                    fill=True, fill_color=color, fill_opacity=0.25,
                    popup=folium.Popup(popup_html, max_width=380), tooltip=name,
                ).add_to(m)
                bounds.extend(locs)
            else:
                continue
            shape_count += 1
            if not label_done:
                c = feature_center(f)
                if c:
                    _add_query_label(m, float(c[0]), float(c[1]), name)
                    label_done = True

    if len(bounds) > 1:
        m.fit_bounds(bounds, padding=(20, 20))
    AdaptiveVectorScale().add_to(m)

    return m, {
        "projects": len(prepared),
        "shapes": shape_count,
        "gis_projects": gis_project_count,
        "fallback_projects": fallback_count,
    }


# ============================================================
# Streamlit 頁面
# ============================================================
def _row_option_label(p: ProjectRow) -> str:
    return f"{p.row_key}｜{p.project_name}｜{p.unit}"


def _project_option_label(p: ProjectRow) -> str:
    return f"{p.project_id}｜{p.project_name}｜{p.status or '未填'}"


def _known_option(pid: str, name: str) -> str:
    return f"{pid}｜{name}"


def _known_option_with_history(pid: str, name: str, current_ids: set) -> str:
    base = _known_option(pid, name or "（歷史工程名稱未記錄）")
    if pid not in current_ids:
        return base + "｜📚 歷史案件（目前管控表已移除）"
    return base


def _show_status_legend():
    parts = []
    for s in ["未發包", "招標中", "訂約中", "待開工", "施工中", "停工中", "落後", "已完工", "已解約", "已取消"]:
        parts.append(f'<span style="display:inline-block;margin-right:14px;"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{status_color(s)};margin-right:4px;"></span>{s}</span>')
    st.markdown("".join(parts), unsafe_allow_html=True)


def _gis_snapshot_key_v366(settings: GitHubSettings) -> Tuple[str, ...]:
    return (
        settings.owner, settings.repo, settings.branch, settings.excel_path,
        settings.geo_path, settings.history_path, settings.registry_path,
        _reach_path(), _project_draft_path(),
    )


def _get_gis_session_snapshot_v366(settings: GitHubSettings, force: bool = False) -> Dict[str, Any]:
    """V3.6.6：同一個瀏覽 Session 只載入一次 GIS 主資料。

    Streamlit 的 st.cache_data 每次 rerun 仍需序列化/複製大型 bytes/dict；
    這裡再加一層 Session 常駐，切縣市/點工程時直接取記憶體。
    """
    key = _gis_snapshot_key_v366(settings)
    snap = st.session_state.get("_gis_v366_snapshot")
    if not force and isinstance(snap, dict) and st.session_state.get("_gis_v366_snapshot_key") == key:
        return snap

    excel_bytes, geo, history, registry, reaches, project_drafts = _cached_gis_page_snapshot_v351(
        settings.owner, settings.repo, settings.branch, settings.excel_path,
        settings.geo_path, settings.history_path, settings.registry_path,
        _reach_path(), _project_draft_path(), settings.token,
    )
    rows = _cached_scan_rows_v351(excel_bytes, settings.excel_path)
    snap = {
        "excel_bytes": excel_bytes,
        "geo": geo,
        "history": history,
        "registry": registry,
        "reaches": reaches,
        "project_drafts": project_drafts,
        "rows": rows,
    }
    st.session_state["_gis_v366_snapshot"] = snap
    st.session_state["_gis_v366_snapshot_key"] = key
    return snap


def _pending_point_updates_v366() -> Dict[str, Dict[str, Any]]:
    obj = st.session_state.setdefault("gis_v366_pending_point_updates", {})
    if not isinstance(obj, dict):
        obj = {}
        st.session_state["gis_v366_pending_point_updates"] = obj
    return obj


def _apply_point_batch_to_session_snapshot_v366(result: Dict[str, Any]) -> None:
    """GitHub 批次同步成功後，直接更新本 Session 快照，不再重新下載整本 Excel/GIS。"""
    snap = st.session_state.get("_gis_v366_snapshot")
    if not isinstance(snap, dict):
        return
    if result.get("_new_excel_bytes") is not None:
        snap["excel_bytes"] = result["_new_excel_bytes"]
    if result.get("_new_geo") is not None:
        snap["geo"] = result["_new_geo"]
    by_id = {str(i.get("project_id")): i for i in (result.get("items") or []) if i.get("project_id")}
    for row in snap.get("rows") or []:
        item = by_id.get(row.project_id)
        if item:
            row.lon = float(item["lon"])
            row.lat = float(item["lat"])
            row.coord_source = "GIS校正"
            row.coord_status = "GIS校正後已同步Excel"
    st.session_state["_gis_v366_snapshot"] = snap


# V3.6.27：已移除早期重複的 render_engineering_gis_page 定義，避免後續修正誤改到不會被呼叫的舊介面。

# ============================================================
import math
from branca.element import Element
from gis_auth import (
    ROLE_LABELS,
    MODE_OPEN,
    MODE_LOGIN,
    MODE_LOCKED,
    load_security_config,
    identity_and_permissions,
    render_login_panel,
    render_admin_panel,
    admin_unlocked,
)

DEFAULT_REACH_PATH = "system_data/river_reaches.geojson"
DEFAULT_PROJECT_DRAFT_PATH = "system_data/project_geo_drafts.geojson"

# 詳細模式：依使用者確認的規則。
DETAIL_STATUS_COLORS = {
    "未發包": "#808080",
    "招標中": "#FFD54F",
    "訂約中": "#FFCC80",
    "待開工": "#FFCC80",
    "施工中": "#FFCC80",
    "停工中": "#F57C00",
    "落後": "#F57C00",
    "已完工": "#43A047",
    "已解約": "#111111",
    "已取消": "#111111",
}
DETAIL_DEFAULT_COLOR = "#808080"
REACH_PENDING_COLOR = "#D32F2F"
REACH_COMPLETED_DETAIL_COLOR = "#1B5E20"
REACH_UNASSIGNED_COLOR = "#1976D2"
SIMPLE_APPROVED_COLOR = "#FBC02D"
SIMPLE_COMPLETED_COLOR = "#43A047"

REACH_STATUS_UNASSIGNED = "未指定現況"
REACH_STATUSES = [REACH_STATUS_UNASSIGNED, "尚待治理", "已完成治理"]
REACH_BASIS_OPTIONS = ["現場已知", "歷史工程", "地方政府提供", "舊治理計畫", "航照／圖資判讀", "其他"]


def _reach_path() -> str:
    try:
        return str(st.secrets["github"].get("reach_path", DEFAULT_REACH_PATH))
    except Exception:
        return DEFAULT_REACH_PATH


def _project_draft_path() -> str:
    try:
        return str(st.secrets["github"].get("project_draft_path", DEFAULT_PROJECT_DRAFT_PATH))
    except Exception:
        return DEFAULT_PROJECT_DRAFT_PATH


def empty_project_drafts() -> Dict[str, Any]:
    return {"type": "FeatureCollection", "schema_version": 1, "features": []}


def empty_reaches() -> Dict[str, Any]:
    return {"type": "FeatureCollection", "schema_version": 1, "features": []}


def _detail_project_color(status: Any) -> str:
    s = display_text(status)
    if s in DETAIL_STATUS_COLORS:
        return DETAIL_STATUS_COLORS[s]
    for k, color in DETAIL_STATUS_COLORS.items():
        if k and k in s:
            return color
    return DETAIL_DEFAULT_COLOR


def _simple_project_category(status: Any) -> str:
    s = display_text(status)
    if "已完工" in s:
        return "已完成治理"
    if "已解約" in s or "已取消" in s:
        return "不顯示"
    return "已核定"


def _simple_project_color(status: Any) -> Optional[str]:
    cat = _simple_project_category(status)
    if cat == "已完成治理":
        return SIMPLE_COMPLETED_COLOR
    if cat == "已核定":
        return SIMPLE_APPROVED_COLOR
    return None


def _project_color(status: Any, display_mode: str) -> Optional[str]:
    return _simple_project_color(status) if display_mode == "簡易顯示" else _detail_project_color(status)


def _reach_color(status: str, display_mode: str) -> str:
    status_text = display_text(status)
    if status_text == "尚待治理":
        return REACH_PENDING_COLOR
    if status_text == "已完成治理":
        if display_mode == "簡易顯示":
            return SIMPLE_COMPLETED_COLOR
        return REACH_COMPLETED_DETAIL_COLOR
    # 空白、舊資料未知值及新版「未指定現況」一律維持藍色，避免誤顯示成已完成。
    return REACH_UNASSIGNED_COLOR


def _legend_html(display_mode: str) -> str:
    if display_mode == "簡易顯示":
        items = [
            (REACH_UNASSIGNED_COLOR, "未指定現況（待確認）"),
            (REACH_PENDING_COLOR, "尚待治理"),
            (SIMPLE_APPROVED_COLOR, "已核定"),
            (SIMPLE_COMPLETED_COLOR, "已完成治理"),
        ]
        title = "河道整治情形"
    else:
        items = [
            (REACH_UNASSIGNED_COLOR, "未指定現況（待確認）"),
            (REACH_PENDING_COLOR, "尚待治理"),
            (DETAIL_STATUS_COLORS["未發包"], "未發包"),
            (DETAIL_STATUS_COLORS["招標中"], "招標中"),
            (DETAIL_STATUS_COLORS["訂約中"], "訂約中／待開工"),
            (DETAIL_STATUS_COLORS["施工中"], "施工中"),
            (DETAIL_STATUS_COLORS["停工中"], "停工中／落後"),
            (DETAIL_STATUS_COLORS["已完工"], "已完工工程"),
            (REACH_COMPLETED_DETAIL_COLOR, "已完成治理（底圖）"),
            (DETAIL_STATUS_COLORS["已取消"], "已解約／已取消"),
        ]
        title = "詳細顯示"
    rows = "".join(
        f'<div style="display:flex;align-items:center;gap:7px;margin:3px 0;">'
        f'<span style="width:24px;height:5px;background:{c};display:inline-block;border-radius:3px"></span>'
        f'<span style="color:#111827 !important;-webkit-text-fill-color:#111827 !important">{label}</span></div>' for c, label in items
    )
    return f"""
    <div class="wra-gis-legend" style="position:fixed;top:76px;right:14px;z-index:9999;background:rgba(255,255,255,.96);border:1px solid #aaa;
                border-radius:7px;padding:9px 11px;font-size:12px;line-height:1.3;box-shadow:0 1px 4px rgba(0,0,0,.16);
                color:#111827 !important;-webkit-text-fill-color:#111827 !important;color-scheme:light;">
      <div style="font-weight:700;margin-bottom:5px;color:#111827 !important;-webkit-text-fill-color:#111827 !important">{title}</div>{rows}
    </div>
    """


def _add_leaflet_compat_css(m: folium.Map) -> None:
    """V3.6.16：跨瀏覽器固定 Leaflet 文字顏色，避免深色模式白字白底。"""
    m.get_root().html.add_child(Element(r"""
    <style>
      .wra-gis-legend, .wra-gis-legend * {
        color:#111827 !important; -webkit-text-fill-color:#111827 !important;
      }
      .leaflet-popup-content-wrapper, .leaflet-popup-tip, .leaflet-tooltip {
        background:#ffffff !important; color:#111827 !important; color-scheme:light !important;
      }
      .leaflet-popup-content, .leaflet-popup-content *, .leaflet-tooltip, .leaflet-tooltip * {
        color:#111827 !important; -webkit-text-fill-color:#111827 !important;
        opacity:1 !important; text-shadow:none !important;
      }
      .leaflet-control-layers, .leaflet-control-layers *, .leaflet-control-scale, .leaflet-control-scale * {
        color:#111827 !important; -webkit-text-fill-color:#111827 !important; color-scheme:light !important;
      }
    </style>
    """))




class AdaptiveVectorScale(MacroElement):
    """V3.6.58：依 Leaflet zoom 動態縮放點符號與線／面外框。

    只改『畫面符號尺寸』，不改 GeoJSON 幾何。縮到全縣／全臺時讓點與線更細小，
    放大後再逐步回到原尺寸，降低大量圖資彼此遮蔽。線／面外框採比點位更積極的
    縮小比例，避免全縣／全臺預覽時黑色外框糊成一團。
    """
    _template = Template(r"""
    {% macro script(this, kwargs) %}
    (function(){
      var map={{ this._parent.get_name() }};
      function factor(z){
        if(z<=7) return 0.34;
        if(z<=9) return 0.45;
        if(z<=11) return 0.58;
        if(z<=13) return 0.72;
        if(z<=15) return 0.88;
        return 1.0;
      }
      function lineFactor(z){
        if(z<=7) return 0.16;
        if(z<=9) return 0.24;
        if(z<=11) return 0.36;
        if(z<=13) return 0.55;
        if(z<=15) return 0.78;
        return 1.0;
      }
      function applyPath(layer, pointF, pathF){
        try{
          var isPoint = layer instanceof L.CircleMarker && !(layer instanceof L.Circle);
          if(isPoint){
            if(layer.__wraBaseRadius==null) layer.__wraBaseRadius=Number(layer.getRadius ? layer.getRadius() : 6) || 6;
            var rr=Math.max(2.2, layer.__wraBaseRadius*pointF);
            if(layer.setRadius) layer.setRadius(rr);
          }
          if(layer instanceof L.Path && layer.setStyle){
            if(layer.__wraBaseWeight==null) layer.__wraBaseWeight=Number(layer.options && layer.options.weight) || 2;
            var wf = isPoint ? pointF : pathF;
            var minW = isPoint ? (layer.__wraOutline ? 1.8 : 1.2) : 0.75;
            var ww=Math.max(minW, layer.__wraBaseWeight*wf);
            var opt={weight:ww};
            if(layer.options && layer.options.fill && layer.__wraBaseFillOpacity==null){
              layer.__wraBaseFillOpacity=Number(layer.options.fillOpacity || 0);
            }
            if(layer.__wraBaseFillOpacity!=null){
              opt.fillOpacity=Math.max(0.08, layer.__wraBaseFillOpacity*(0.58+0.42*wf));
            }
            layer.setStyle(opt);
          }
        }catch(e){}
      }
      function walk(layer, pointF, pathF){
        if(!layer) return;
        if(layer instanceof L.Path){ applyPath(layer,pointF,pathF); return; }
        if(layer.eachLayer){ try{ layer.eachLayer(function(ch){walk(ch,pointF,pathF);}); }catch(e){} }
      }
      function refresh(){
        var z=map.getZoom(), pointF=factor(z), pathF=lineFactor(z);
        try{ map.eachLayer(function(layer){walk(layer,pointF,pathF);}); }catch(e){}
      }
      map.on('zoomend', refresh);
      setTimeout(refresh, 0);
    })();
    {% endmacro %}
    """)

    def __init__(self):
        super().__init__()
        self._name = "AdaptiveVectorScale"


class PointCalibrationCursor(MacroElement):
    """V3.6.28 代表點校正游標：一般滑鼠箭頭，箭頭尖端有與原始代表點近似大小的紅點。"""
    _template = Template(r"""
    {% macro script(this, kwargs) %}
    (function(){
      var map={{ this._parent.get_name() }};
      var el=map.getContainer();
      if(!el) return;
      // 紅點半徑 8px，接近校正地圖原始代表點的 radius=8。
      // 紅點中心 (9,9) = 箭頭尖端 = CSS cursor hotspot (9,9)，
      // 因此使用者看到的紅點中心就是實際落點。
      var svg="<svg xmlns='http://www.w3.org/2000/svg' width='48' height='50' viewBox='0 0 48 50'>"+
              "<path d='M9 9 L9 37 L16 30 L23 45 L29 42 L22 28 L36 28 Z' fill='#ffffff' stroke='#111827' stroke-width='2.2' stroke-linejoin='round'/>"+
              "<circle cx='9' cy='9' r='8' fill='#ef4444' stroke='#ffffff' stroke-width='2.5'/>"+
              "</svg>";
      var uri='data:image/svg+xml;charset=UTF-8,'+encodeURIComponent(svg);
      var cur='url("'+uri+'") 9 9, crosshair';
      el.style.cursor=cur;
      try{
        el.querySelectorAll('.leaflet-interactive').forEach(function(x){x.style.cursor=cur;});
      }catch(e){}
    })();
    {% endmacro %}
    """)

    def __init__(self):
        super().__init__()
        self._name = "PointCalibrationCursor"


def _add_legend(m: folium.Map, display_mode: str) -> None:
    _add_leaflet_compat_css(m)
    m.get_root().html.add_child(Element(_legend_html(display_mode)))


def _line_weight(feature: Dict[str, Any], default: int = 6) -> int:
    try:
        return max(2, min(12, int((feature.get("properties") or {}).get("line_weight", default))))
    except Exception:
        return default


def _popup_escape(v: Any) -> str:
    import html as _html
    return _html.escape(display_text(v))



# ============================================================
# V3.6.44 座標品質檢核／批次 CRS 診斷／縣市複核表
# ============================================================
def _coord_audit_median(values: Sequence[float]) -> Optional[float]:
    vals = sorted(float(x) for x in values if x is not None)
    if not vals:
        return None
    n = len(vals)
    m = n // 2
    return vals[m] if n % 2 else (vals[m - 1] + vals[m]) / 2.0


def _coord_audit_raw_columns(hmap: Dict[str, int]) -> Dict[str, Optional[int]]:
    return {
        "wgs_x": find_col(
            hmap,
            [
                "google X座標-E", "google X座標", "googleX座標-E",
                "googleX座標", "WGS84經度", "經度",
            ],
        ),
        "wgs_y": find_col(
            hmap,
            [
                "google Y座標-N", "google Y座標", "googleY座標-N",
                "googleY座標", "WGS84緯度", "緯度",
            ],
        ),
        "twd_x": find_col(
            hmap,
            [
                "TWD_97\n經度-X座標", "TWD_97經度-X座標", "TWD_97\n經度",
                "TWD97經度-X座標", "TWD97經度", "X座標",
            ],
        ),
        "twd_y": find_col(
            hmap,
            [
                "TWD_97\n緯度-Y座標", "TWD_97緯度-Y座標", "TWD_97\n緯度",
                "TWD97緯度-Y座標", "TWD97緯度", "Y座標",
            ],
        ),
    }


def _coord_audit_pair(ws, row: int, cx: Optional[int], cy: Optional[int]) -> Tuple[Any, Any]:
    return (
        ws.cell(row, cx).value if cx else None,
        ws.cell(row, cy).value if cy else None,
    )


def _coord_audit_tm2_candidate(
    a: Any,
    b: Any,
    county: str,
    datum: str,
) -> Optional[Dict[str, Any]]:
    x = parse_number(a)
    y = parse_number(b)
    if x is None or y is None:
        return None

    swapped = False
    if 80_000 <= x <= 420_000 and 2_300_000 <= y <= 2_900_000:
        tx, ty = x, y
    elif 80_000 <= y <= 420_000 and 2_300_000 <= x <= 2_900_000:
        tx, ty = y, x
        swapped = True
    else:
        return None

    county_n = norm_text(county)
    prefer_119 = any(k in county_n for k in ["澎湖", "金門", "連江"])
    if str(datum).upper() == "TWD67":
        pairs = (
            [("3827", TRANSFORMER_67_119), ("3828", TRANSFORMER_67_121)]
            if prefer_119
            else [("3828", TRANSFORMER_67_121), ("3827", TRANSFORMER_67_119)]
        )
    else:
        pairs = (
            [("3825", TRANSFORMER_119), ("3826", TRANSFORMER_121)]
            if prefer_119
            else [("3826", TRANSFORMER_121), ("3825", TRANSFORMER_119)]
        )

    for epsg, tr in pairs:
        try:
            lon, lat = tr.transform(tx, ty)
            if is_valid_wgs84(lon, lat):
                return {
                    "lon": float(lon), "lat": float(lat),
                    "swapped": swapped, "epsg": epsg,
                    "datum": str(datum).upper(),
                }
        except Exception:
            pass
    return None


def _coord_audit_twd67_geo_candidate(a: Any, b: Any) -> Optional[Dict[str, Any]]:
    p = normalize_wgs84_pair(a, b)
    if not p:
        return None
    raw_lon, raw_lat, swapped = p
    try:
        lon, lat = TRANSFORMER_67_GEO.transform(raw_lon, raw_lat)
        if is_valid_wgs84(lon, lat):
            return {
                "lon": float(lon), "lat": float(lat),
                "swapped": swapped, "epsg": "3821",
                "datum": "TWD67-GEO",
            }
    except Exception:
        pass
    return None


def _coord_audit_gis_reference(
    geo: Dict[str, Any], project_id: str
) -> Optional[Tuple[float, float]]:
    if not project_id:
        return None
    f = original_feature(geo, project_id)
    if not f:
        return None
    prop = f.get("properties") or {}
    authoritative = bool(
        prop.get("coordinate_authority") == "gis"
        or prop.get("coordinate_edit_source") == "gis_map"
        or prop.get("source") == "gis_representative"
    )
    if not authoritative:
        return None
    g = f.get("geometry") or {}
    c = g.get("coordinates") or []
    if g.get("type") != "Point" or len(c) < 2:
        return None
    try:
        lon, lat = float(c[0]), float(c[1])
    except Exception:
        return None
    if not is_valid_wgs84(lon, lat):
        return None
    return lon, lat


def _coord_audit_distance(
    cand: Optional[Dict[str, Any]],
    ref: Optional[Tuple[float, float]],
) -> Optional[float]:
    if not cand or not ref:
        return None
    try:
        return _haversine_m(
            float(cand["lat"]), float(cand["lon"]),
            float(ref[1]), float(ref[0]),
        )
    except Exception:
        return None


def _coord_audit_current_distance(
    p: "ProjectRow",
    ref: Optional[Tuple[float, float]],
) -> Optional[float]:
    if not ref or p.lon is None or p.lat is None:
        return None
    try:
        return _haversine_m(float(p.lat), float(p.lon), float(ref[1]), float(ref[0]))
    except Exception:
        return None


def _coord_audit_batch_assessment(
    records: Sequence[Dict[str, Any]],
) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """依縣市分 projected / geographic 兩種資料型態做批次 datum 判斷。

    只有已存在人工 GIS 校正點的案件才當作參考樣本。
    至少 3 件才會給批次結論；否則只顯示樣本不足。
    """
    result: Dict[str, Dict[str, Dict[str, Any]]] = {}
    counties = sorted({display_text(r.get("county")) for r in records if display_text(r.get("county"))})
    for county in counties:
        crows = [r for r in records if display_text(r.get("county")) == county]
        result[county] = {}

        projected = [
            r for r in crows
            if r.get("d97_m") is not None and r.get("d67_tm2_m") is not None
        ]
        geographic = [
            r for r in crows
            if r.get("d_wgs_m") is not None and r.get("d67_geo_m") is not None
        ]

        for mode, samples, base_key, alt_key, base_name in [
            ("projected", projected, "d97_m", "d67_tm2_m", "TWD97"),
            ("geographic", geographic, "d_wgs_m", "d67_geo_m", "WGS84"),
        ]:
            base_med = _coord_audit_median([r[base_key] for r in samples])
            alt_med = _coord_audit_median([r[alt_key] for r in samples])
            n = len(samples)
            status = "樣本不足"
            detail = "至少需要 3 件已人工校正 GIS 代表點，才做整批判斷。"
            if n >= 3 and base_med is not None and alt_med is not None:
                improvement_to_67 = base_med - alt_med
                improvement_to_base = alt_med - base_med
                if (
                    improvement_to_67 >= 120
                    and alt_med <= max(120.0, base_med * 0.45)
                ):
                    status = "高度疑似TWD67"
                    detail = (
                        f"{n} 件參考樣本中，{base_name} 假設中位誤差約 {base_med:,.0f}m，"
                        f"TWD67 假設約 {alt_med:,.0f}m。"
                    )
                elif (
                    improvement_to_base >= 120
                    and base_med <= max(120.0, alt_med * 0.45)
                ):
                    status = f"{base_name}較符合"
                    detail = (
                        f"{n} 件參考樣本中，{base_name} 假設中位誤差約 {base_med:,.0f}m，"
                        f"TWD67 假設約 {alt_med:,.0f}m。"
                    )
                else:
                    status = "無明顯整批偏移"
                    detail = (
                        f"{n} 件參考樣本：{base_name} 中位誤差約 {base_med:,.0f}m；"
                        f"TWD67 約 {alt_med:,.0f}m，差異不足以安全判定。"
                    )
            result[county][mode] = {
                "status": status,
                "sample_count": n,
                "base_median_m": base_med,
                "twd67_median_m": alt_med,
                "detail": detail,
            }
    return result


def _coord_audit_cost_map(excel_bytes: bytes, rows: Sequence["ProjectRow"]) -> Dict[str, Optional[float]]:
    """只讀取 data_only cached value；不重算整本 Excel 公式，避免拖慢 GIS。"""
    out: Dict[str, Optional[float]] = {}
    try:
        wb = load_workbook(io.BytesIO(excel_bytes), data_only=True, read_only=True)
    except Exception:
        return out

    by_sheet: Dict[str, List["ProjectRow"]] = {}
    for p in rows:
        by_sheet.setdefault(p.sheet_name, []).append(p)

    for sheet_name, prows in by_sheet.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_row = find_header_row(ws)
        if not header_row:
            continue
        hmap = normalize_header_map(ws, header_row)

        total_col = find_col(
            hmap,
            [
                "總經費(千元)", "總經費（千元）", "總經費",
                "核定經費(千元)", "核定經費（千元）", "核定經費",
            ],
        )
        total_central_col = find_col(hmap, ["總中央款", "中央款合計", "中央補助合計"])
        total_local_col = find_col(hmap, ["總地方款", "地方款合計"])

        component_cols = [
            find_col(hmap, ["中央款工程費合計(含橋梁中央款)", "中央款工程費合計", "中央補助經費(千元)"]),
            find_col(hmap, ["地方款工程費合計(含橋梁地方款)", "地方款工程費合計"]),
            find_col(hmap, ["中央款用地費合計"]),
            find_col(hmap, ["地方款用地費合計"]),
        ]

        for p in prows:
            value = None
            if total_col:
                value = parse_number(ws.cell(p.row, total_col).value)
            if value is None and (total_central_col or total_local_col):
                parts = []
                for c in [total_central_col, total_local_col]:
                    if c:
                        v = parse_number(ws.cell(p.row, c).value)
                        if v is not None:
                            parts.append(v)
                if parts:
                    value = sum(parts)
            if value is None:
                parts = []
                for c in component_cols:
                    if c:
                        v = parse_number(ws.cell(p.row, c).value)
                        if v is not None:
                            parts.append(v)
                if parts:
                    value = sum(parts)
            out[p.row_key] = value
    try:
        wb.close()
    except Exception:
        pass
    return out


def _build_coordinate_audit_records(
    rows: Sequence["ProjectRow"],
    geo: Dict[str, Any],
    excel_bytes: bytes,
) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, Dict[str, Any]]]]:
    wb = _load_wb(excel_bytes, DEFAULT_EXCEL_PATH)
    costs = _coord_audit_cost_map(excel_bytes, rows)

    sheet_meta: Dict[str, Tuple[Any, Dict[str, int], Dict[str, Optional[int]]]] = {}
    for sheet_name in TARGET_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        hr = find_header_row(ws)
        if not hr:
            continue
        hmap = normalize_header_map(ws, hr)
        sheet_meta[sheet_name] = (ws, hmap, _coord_audit_raw_columns(hmap))

    records: List[Dict[str, Any]] = []
    for p in rows:
        if p.sheet_name not in sheet_meta:
            continue
        ws, _hmap, cols = sheet_meta[p.sheet_name]
        wgx, wgy = _coord_audit_pair(ws, p.row, cols.get("wgs_x"), cols.get("wgs_y"))
        tx, ty = _coord_audit_pair(ws, p.row, cols.get("twd_x"), cols.get("twd_y"))

        wgs_pair = normalize_wgs84_pair(wgx, wgy)
        wgs_cand = None
        if wgs_pair:
            wgs_cand = {
                "lon": float(wgs_pair[0]), "lat": float(wgs_pair[1]),
                "swapped": bool(wgs_pair[2]), "datum": "WGS84", "epsg": "4326",
            }

        p97 = _coord_audit_tm2_candidate(tx, ty, p.county, "TWD97")
        p67 = _coord_audit_tm2_candidate(tx, ty, p.county, "TWD67")
        # 若投影座標被填錯在 Google/WGS84 欄，也要能辨識。
        if not p97:
            p97 = _coord_audit_tm2_candidate(wgx, wgy, p.county, "TWD97")
        if not p67:
            p67 = _coord_audit_tm2_candidate(wgx, wgy, p.county, "TWD67")

        p67_geo = _coord_audit_twd67_geo_candidate(wgx, wgy) if wgs_cand else None
        ref = _coord_audit_gis_reference(geo, p.project_id)

        d_current = _coord_audit_current_distance(p, ref)
        d_wgs = _coord_audit_distance(wgs_cand, ref)
        d97 = _coord_audit_distance(p97, ref)
        d67_tm2 = _coord_audit_distance(p67, ref)
        d67_geo = _coord_audit_distance(p67_geo, ref)

        raw_has_any = any(
            parse_number(v) is not None for v in [wgx, wgy, tx, ty]
        )
        raw_pair_incomplete = (
            (parse_number(wgx) is None) != (parse_number(wgy) is None)
            or (parse_number(tx) is None) != (parse_number(ty) is None)
        )

        if wgs_cand:
            raw_source = "WGS84/Google欄位"
            raw_x, raw_y = wgx, wgy
        elif p97 or p67:
            if parse_number(tx) is not None and parse_number(ty) is not None:
                raw_source = "TWD X/Y欄位"
                raw_x, raw_y = tx, ty
            else:
                raw_source = "Google欄位疑似填入TM2座標"
                raw_x, raw_y = wgx, wgy
        elif parse_number(tx) is not None or parse_number(ty) is not None:
            raw_source = "TWD X/Y欄位"
            raw_x, raw_y = tx, ty
        else:
            raw_source = "WGS84/Google欄位"
            raw_x, raw_y = wgx, wgy

        diagnosis = ""
        action = ""
        confidence = ""
        suggested = None
        suggested_system = ""

        # 有人工 GIS 正確點時，優先用實際距離判斷，而不是只看數值格式。
        if ref:
            baseline = d_current
            alternatives = []
            if p67 and d67_tm2 is not None:
                alternatives.append(("TWD67 TM2", p67, d67_tm2))
            if p67_geo and d67_geo is not None:
                alternatives.append(("TWD67經緯度", p67_geo, d67_geo))
            alternatives.sort(key=lambda x: x[2])
            best67 = alternatives[0] if alternatives else None

            if (
                best67
                and baseline is not None
                and baseline - best67[2] >= 120
                and best67[2] <= max(120.0, baseline * 0.45)
            ):
                diagnosis = f"疑似{best67[0]}"
                action = "中央確認轉換"
                confidence = "高"
                suggested = best67[1]
                suggested_system = best67[0]
            elif baseline is not None and baseline <= 50:
                if p.coord_swapped:
                    diagnosis = "經緯度／X-Y顛倒但系統可判讀"
                    action = "中央確認欄位"
                    confidence = "高"
                else:
                    diagnosis = "與GIS校正點一致"
                    action = "無需處理"
                    confidence = "高"
            elif baseline is not None and baseline >= 300:
                diagnosis = "原座標與GIS校正點差異過大"
                action = "退回縣市重填"
                confidence = "高"
            elif p.coord_swapped:
                diagnosis = "經緯度／X-Y疑似顛倒"
                action = "中央確認欄位"
                confidence = "中"
            else:
                diagnosis = "需人工確認"
                action = "請縣市確認"
                confidence = "中"
        else:
            if not raw_has_any:
                diagnosis = "無座標"
                action = "退回縣市重填"
                confidence = "高"
            elif raw_pair_incomplete and not (wgs_cand or p97 or p67):
                diagnosis = "座標欄位不完整"
                action = "退回縣市重填"
                confidence = "高"
            elif wgs_cand:
                if bool(wgs_cand.get("swapped")):
                    diagnosis = "經緯度顛倒但系統可判讀"
                    action = "中央確認欄位"
                    confidence = "高"
                else:
                    diagnosis = "WGS84格式有效"
                    action = "無需處理"
                    confidence = "格式判斷"
            elif p97:
                diagnosis = "TM2投影座標；目前按TWD97判讀"
                action = "視需要抽查"
                confidence = "格式判斷"
            else:
                diagnosis = "無法判讀座標"
                action = "退回縣市重填"
                confidence = "高"

        rec = {
            "row_key": p.row_key,
            "project_id": p.project_id,
            "sheet_name": p.sheet_name,
            "county": p.county,
            "project_name": p.project_name,
            "project_content": p.project_content,
            "water_system": p.water_system,
            "status": p.status,
            "unit": p.unit,
            "address": p.address,
            "cost_k": costs.get(p.row_key),
            "raw_source": raw_source,
            "raw_x": raw_x,
            "raw_y": raw_y,
            "wgs_raw_x": wgx,
            "wgs_raw_y": wgy,
            "twd_raw_x": tx,
            "twd_raw_y": ty,
            "current_lon": p.lon,
            "current_lat": p.lat,
            "coord_source": p.coord_source,
            "coord_swapped": bool(p.coord_swapped),
            "gis_ref_lon": ref[0] if ref else None,
            "gis_ref_lat": ref[1] if ref else None,
            "has_gis_reference": bool(ref),
            "d_current_m": d_current,
            "d_wgs_m": d_wgs,
            "d97_m": d97,
            "d67_tm2_m": d67_tm2,
            "d67_geo_m": d67_geo,
            "twd97_candidate": p97,
            "twd67_tm2_candidate": p67,
            "twd67_geo_candidate": p67_geo,
            "diagnosis": diagnosis,
            "action": action,
            "confidence": confidence,
            "suggested_lon": suggested.get("lon") if suggested else None,
            "suggested_lat": suggested.get("lat") if suggested else None,
            "suggested_system": suggested_system,
        }
        records.append(rec)

    batch = _coord_audit_batch_assessment(records)

    # 把「整批偏移」的判斷補到沒有人工 GIS 參考點的同縣市案件。
    for r in records:
        if r.get("has_gis_reference"):
            continue
        county = display_text(r.get("county"))
        b = batch.get(county, {})
        projected_status = (b.get("projected") or {}).get("status")
        geographic_status = (b.get("geographic") or {}).get("status")

        if projected_status == "高度疑似TWD67" and r.get("twd67_tm2_candidate"):
            r["diagnosis"] = "疑似TWD67 TM2（依縣市批次）"
            r["action"] = "中央確認轉換"
            r["confidence"] = "批次推估"
            r["suggested_lon"] = r["twd67_tm2_candidate"].get("lon")
            r["suggested_lat"] = r["twd67_tm2_candidate"].get("lat")
            r["suggested_system"] = "TWD67 TM2"
        elif geographic_status == "高度疑似TWD67" and r.get("twd67_geo_candidate"):
            r["diagnosis"] = "疑似TWD67經緯度（依縣市批次）"
            r["action"] = "中央確認轉換"
            r["confidence"] = "批次推估"
            r["suggested_lon"] = r["twd67_geo_candidate"].get("lon")
            r["suggested_lat"] = r["twd67_geo_candidate"].get("lat")
            r["suggested_system"] = "TWD67經緯度"

    try:
        wb.close()
    except Exception:
        pass
    return records, batch


def _coord_audit_review_rows(
    records: Sequence[Dict[str, Any]],
    scope: str,
) -> List[Dict[str, Any]]:
    if scope == "只匯出建議退回縣市重填":
        return [r for r in records if r.get("action") == "退回縣市重填"]
    if scope == "匯出全部異常與待確認":
        return [
            r for r in records
            if r.get("action") not in {"無需處理"}
        ]
    return list(records)


def _build_county_coordinate_review_xlsx(
    county: str,
    records: Sequence[Dict[str, Any]],
) -> bytes:
    wb = Workbook()
    ws_note = wb.active
    ws_note.title = "填報說明"
    ws = wb.create_sheet("座標複核表")

    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    editable_fill = PatternFill("solid", fgColor="FFF2CC")
    warn_fill = PatternFill("solid", fgColor="FCE4D6")
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws_note["A1"] = f"{county} 工程座標複核填報說明"
    ws_note["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws_note["A1"].fill = title_fill
    ws_note.merge_cells("A1:F1")
    notes = [
        "一、本檔僅供座標複核，請勿新增、刪除或重新排序工程列。",
        "二、黃色欄位為縣市填報欄位；其餘欄位由中央提供參考。",
        "三、請先選擇『縣市填報座標系統』，再填入新 X/經度、Y/緯度。",
        "四、若採 WGS84，X 為經度、Y 為緯度；若採 TWD97/TWD67 TM2，X/Y 請依所選分帶填寫。",
        "五、如無法確認原座標系統，請選『不知道』並於備註說明。",
        "六、回傳時請保留『系統工程ID』，中央將以該 ID 對應工程，不以工程名稱比對。",
        "七、本階段回傳資料不會自動覆寫正式 GIS；仍須由中央審核後處理。",
    ]
    for i, note in enumerate(notes, start=3):
        ws_note.cell(i, 1).value = note
    ws_note.column_dimensions["A"].width = 110

    headers = [
        "系統工程ID", "工作表", "縣市", "工程名稱", "工程內容", "總經費(千元)",
        "原座標欄位來源", "原X/經度", "原Y/緯度", "系統檢核結果", "系統建議",
        "人工檢核結果", "人工測試座標系統", "人工檢核者", "人工檢核日期",
        "縣市填報座標系統", "縣市新X/經度", "縣市新Y/緯度", "縣市備註",
    ]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c)
        cell.value = h
        cell.font = Font(bold=True)
        cell.fill = editable_fill if c >= 16 else header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for rr, rec in enumerate(records, start=2):
        vals = [
            rec.get("project_id"), rec.get("sheet_name"), rec.get("county"),
            rec.get("project_name"), rec.get("project_content"), rec.get("cost_k"),
            rec.get("raw_source"), rec.get("raw_x"), rec.get("raw_y"),
            rec.get("diagnosis"), rec.get("action"),
            rec.get("manual_decision"), rec.get("manual_tested_crs"),
            rec.get("manual_reviewed_by"), rec.get("manual_reviewed_at"),
            "", "", "", "",
        ]
        for cc, value in enumerate(vals, start=1):
            cell = ws.cell(rr, cc)
            cell.value = value
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(cc in {4, 5, 10, 11, 12, 13, 14, 15, 19}))
            if cc >= 16:
                cell.fill = editable_fill
                cell.protection = Protection(locked=False)
            else:
                cell.protection = Protection(locked=True)
            if cc in {10, 11} and rec.get("action") in {"退回縣市重填", "中央確認轉換", "請縣市確認"}:
                cell.fill = warn_fill
        if isinstance(rec.get("cost_k"), (int, float)):
            ws.cell(rr, 6).number_format = '#,##0'
        for cc in [8, 9, 17, 18]:
            ws.cell(rr, cc).number_format = '0.00000000'

    if records:
        dv = DataValidation(
            type="list",
            formula1='"WGS84,TWD97 TM2 zone 121,TWD97 TM2 zone 119,TWD67 TM2 zone 121,TWD67 TM2 zone 119,不知道"',
            allow_blank=True,
        )
        ws.add_data_validation(dv)
        dv.add(f"P2:P{len(records)+1}")

    widths = {
        1: 17, 2: 14, 3: 10, 4: 38, 5: 55, 6: 16, 7: 22,
        8: 18, 9: 18, 10: 30, 11: 20, 12: 26, 13: 24,
        14: 18, 15: 22, 16: 24, 17: 18, 18: 18, 19: 40,
    }
    for c, width in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:S{max(1, len(records)+1)}"
    ws.row_dimensions[1].height = 34
    # 只鎖定中央提供欄位；黃色欄位可直接填寫。未設密碼，中央仍可自行解除保護。
    ws.protection.sheet = True
    ws.protection.sort = False
    ws.protection.autoFilter = False
    ws.protection.selectUnlockedCells = False

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _coord_audit_table_rows(records: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out = []
    for r in records:
        out.append({
            "系統工程ID": r.get("project_id"),
            "工程名稱": r.get("project_name"),
            "水系": r.get("water_system"),
            "原座標來源": r.get("raw_source"),
            "原X/經度": r.get("raw_x"),
            "原Y/緯度": r.get("raw_y"),
            "系統目前經度": None if r.get("current_lon") is None else round(float(r["current_lon"]), 7),
            "系統目前緯度": None if r.get("current_lat") is None else round(float(r["current_lat"]), 7),
            "檢核結果": r.get("diagnosis"),
            "建議處理": r.get("action"),
            "可信度": r.get("confidence"),
            "人工檢核結果": r.get("manual_decision"),
            "人工測試座標系統": r.get("manual_tested_crs"),
            "人工檢核者": r.get("manual_reviewed_by"),
            "人工檢核時間": r.get("manual_reviewed_at"),
            "與GIS校正點距離(m)": None if r.get("d_current_m") is None else round(float(r["d_current_m"]), 1),
            "TWD67候選經度": None if r.get("suggested_lon") is None else round(float(r["suggested_lon"]), 7),
            "TWD67候選緯度": None if r.get("suggested_lat") is None else round(float(r["suggested_lat"]), 7),
        })
    return out


def _render_coordinate_audit_map(
    county: str,
    records: Sequence[Dict[str, Any]],
    only_attention: bool = True,
) -> Optional[folium.Map]:
    rows = [
        r for r in records
        if display_text(r.get("county")) == county
        and (not only_attention or r.get("action") != "無需處理")
    ]
    rows = [
        r for r in rows
        if is_valid_wgs84(
            parse_number(r.get("current_lon")),
            parse_number(r.get("current_lat")),
        )
        or is_valid_wgs84(
            parse_number(r.get("suggested_lon")),
            parse_number(r.get("suggested_lat")),
        )
        or is_valid_wgs84(
            parse_number(r.get("gis_ref_lon")),
            parse_number(r.get("gis_ref_lat")),
        )
    ]
    if not rows:
        return None

    # 避免一次把過多診斷線塞進瀏覽器；異常優先，最多 400 件。
    priority = {
        "退回縣市重填": 0,
        "中央確認轉換": 1,
        "請縣市確認": 2,
        "中央確認欄位": 3,
        "視需要抽查": 4,
        "無需處理": 5,
    }
    rows = sorted(rows, key=lambda r: priority.get(r.get("action"), 9))[:400]

    pts = []
    for r in rows:
        for xk, yk in [
            ("current_lon", "current_lat"),
            ("suggested_lon", "suggested_lat"),
            ("gis_ref_lon", "gis_ref_lat"),
        ]:
            lon = parse_number(r.get(xk))
            lat = parse_number(r.get(yk))
            if is_valid_wgs84(lon, lat):
                pts.append([lat, lon])

    if not pts:
        return None
    m = folium.Map(
        location=[
            sum(x[0] for x in pts) / len(pts),
            sum(x[1] for x in pts) / len(pts),
        ],
        zoom_start=10,
        tiles="OpenStreetMap",
        control_scale=True,
    )
    fg_now = folium.FeatureGroup(name="目前系統解讀", show=True)
    fg_67 = folium.FeatureGroup(name="TWD67修正候選", show=True)
    fg_ref = folium.FeatureGroup(name="人工GIS校正參考", show=True)
    fg_link = folium.FeatureGroup(name="差異連線", show=True)

    for r in rows:
        name = display_text(r.get("project_name"))
        pid = display_text(r.get("project_id"))
        diag = display_text(r.get("diagnosis"))
        action = display_text(r.get("action"))
        popup = (
            f"<b>{_popup_escape(name)}</b><br>"
            f"{_popup_escape(pid)}<br>"
            f"檢核：{_popup_escape(diag)}<br>"
            f"建議：{_popup_escape(action)}"
        )

        lon = parse_number(r.get("current_lon"))
        lat = parse_number(r.get("current_lat"))
        if is_valid_wgs84(lon, lat):
            color = "#2E7D32" if action == "無需處理" else ("#D32F2F" if action == "退回縣市重填" else "#616161")
            folium.CircleMarker(
                [lat, lon], radius=6, color="#000000", weight=2,
                fill=True, fill_color=color, fill_opacity=0.9,
                tooltip=name, popup=folium.Popup(popup, max_width=420),
            ).add_to(fg_now)

        slon = parse_number(r.get("suggested_lon"))
        slat = parse_number(r.get("suggested_lat"))
        if is_valid_wgs84(slon, slat):
            folium.CircleMarker(
                [slat, slon], radius=7, color="#000000", weight=2,
                fill=True, fill_color="#FB8C00", fill_opacity=0.95,
                tooltip=f"TWD67候選｜{name}",
            ).add_to(fg_67)
            if is_valid_wgs84(lon, lat):
                folium.PolyLine(
                    [[lat, lon], [slat, slon]],
                    color="#FB8C00", weight=2, opacity=0.7, dash_array="6,6",
                ).add_to(fg_link)

        rlon = parse_number(r.get("gis_ref_lon"))
        rlat = parse_number(r.get("gis_ref_lat"))
        if is_valid_wgs84(rlon, rlat):
            folium.CircleMarker(
                [rlat, rlon], radius=7, color="#000000", weight=2.5,
                fill=True, fill_color="#1565C0", fill_opacity=0.95,
                tooltip=f"GIS已校正｜{name}",
            ).add_to(fg_ref)
            if is_valid_wgs84(lon, lat):
                folium.PolyLine(
                    [[lat, lon], [rlat, rlon]],
                    color="#1565C0", weight=2, opacity=0.65,
                ).add_to(fg_link)

    fg_now.add_to(m)
    fg_67.add_to(m)
    fg_ref.add_to(m)
    fg_link.add_to(m)
    folium.LayerControl(collapsed=False).add_to(m)
    try:
        m.fit_bounds(pts)
    except Exception:
        pass
    return m




def _coordinate_review_key(rec: Dict[str, Any]) -> str:
    return display_text(rec.get("project_id")) or display_text(rec.get("row_key"))


def _coordinate_review_for_record(
    rec: Dict[str, Any],
    registry: Dict[str, Any],
) -> Dict[str, Any]:
    key = _coordinate_review_key(rec)
    bucket = (registry or {}).get("coordinate_reviews") or {}
    value = bucket.get(key) if isinstance(bucket, dict) else None
    return value if isinstance(value, dict) else {}


def _load_selected_raw_coordinates_v3646(
    excel_bytes: bytes,
    selected_rows: Sequence["ProjectRow"],
) -> Dict[str, Dict[str, Any]]:
    """人工按下『載入原始座標』時才開 current.xlsx 一次。

    回傳同一案件的 WGS/Google 欄位與 TWD X/Y 欄位，供人工任意切換 CRS 測試。
    """
    if not selected_rows:
        return {}
    wanted = {p.row_key: p for p in selected_rows}
    out: Dict[str, Dict[str, Any]] = {}

    wb = _load_wb(excel_bytes, DEFAULT_EXCEL_PATH)
    by_sheet: Dict[str, List["ProjectRow"]] = {}
    for p in selected_rows:
        by_sheet.setdefault(p.sheet_name, []).append(p)

    for sheet_name, prows in by_sheet.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        hr = find_header_row(ws)
        if not hr:
            continue
        hmap = normalize_header_map(ws, hr)
        cols = _coord_audit_raw_columns(hmap)
        for p in prows:
            wgx, wgy = _coord_audit_pair(ws, p.row, cols.get("wgs_x"), cols.get("wgs_y"))
            tx, ty = _coord_audit_pair(ws, p.row, cols.get("twd_x"), cols.get("twd_y"))
            out[p.row_key] = {
                "row_key": p.row_key,
                "project_id": p.project_id,
                "project_name": p.project_name,
                "county": p.county,
                "wgs_x": wgx,
                "wgs_y": wgy,
                "twd_x": tx,
                "twd_y": ty,
            }
    try:
        wb.close()
    except Exception:
        pass
    return out


def _manual_coord_pair_v3646(
    raw: Dict[str, Any],
    pair_source: str,
    tested_crs: str,
) -> Tuple[Optional[float], Optional[float], str]:
    """選擇人工試轉要使用哪一組原始 X/Y。"""
    geo_mode = tested_crs.startswith("WGS84") or tested_crs.startswith("TWD67 Geographic")

    def pair(prefix: str):
        x = parse_number(raw.get(prefix + "_x"))
        y = parse_number(raw.get(prefix + "_y"))
        return (x, y) if x is not None and y is not None else None

    wgs = pair("wgs")
    twd = pair("twd")

    if pair_source == "WGS84 / Google 原始欄位":
        chosen = wgs
        label = "WGS84 / Google 原始欄位"
    elif pair_source == "TWD X/Y 原始欄位":
        chosen = twd
        label = "TWD X/Y 原始欄位"
    else:
        # 自動：經緯度 CRS 優先用 Google/WGS 欄；TM2 優先用 TWD X/Y。
        if geo_mode:
            chosen = wgs or twd
            label = "WGS84 / Google 原始欄位" if wgs else "TWD X/Y 原始欄位"
        else:
            chosen = twd or wgs
            label = "TWD X/Y 原始欄位" if twd else "WGS84 / Google 原始欄位"

    if not chosen:
        return None, None, label
    return float(chosen[0]), float(chosen[1]), label


def _manual_transform_candidate_v3646(
    raw: Dict[str, Any],
    pair_source: str,
    tested_crs: str,
) -> Dict[str, Any]:
    x, y, used_source = _manual_coord_pair_v3646(raw, pair_source, tested_crs)
    result = {
        "lon": None, "lat": None, "valid": False, "reason": "",
        "used_source": used_source, "raw_x": x, "raw_y": y,
        "tested_crs": tested_crs,
    }
    if x is None or y is None:
        result["reason"] = "所選原始欄位沒有完整 X/Y。"
        return result

    try:
        if tested_crs == "WGS84":
            lon, lat = x, y
        elif tested_crs == "WGS84（X/Y對調）":
            lon, lat = y, x
        elif tested_crs == "TWD97 TM2 zone 121":
            lon, lat = TRANSFORMER_121.transform(x, y)
        elif tested_crs == "TWD97 TM2 zone 119":
            lon, lat = TRANSFORMER_119.transform(x, y)
        elif tested_crs == "TWD67 TM2 zone 121":
            lon, lat = TRANSFORMER_67_121.transform(x, y)
        elif tested_crs == "TWD67 TM2 zone 119":
            lon, lat = TRANSFORMER_67_119.transform(x, y)
        elif tested_crs == "TWD67 Geographic":
            lon, lat = TRANSFORMER_67_GEO.transform(x, y)
        elif tested_crs == "TWD67 Geographic（X/Y對調）":
            lon, lat = TRANSFORMER_67_GEO.transform(y, x)
        else:
            result["reason"] = "未知座標系統。"
            return result

        if is_valid_wgs84(lon, lat):
            result["lon"] = float(lon)
            result["lat"] = float(lat)
            result["valid"] = True
        else:
            result["reason"] = "轉換後不在臺灣常用 WGS84 合理範圍。"
    except Exception as exc:
        result["reason"] = f"轉換失敗：{exc}"
    return result



def _manual_selection_marker_label_v3647(rec: Dict[str, Any]) -> str:
    return (
        f"{display_text(rec.get('project_name')) or '未命名工程'}｜"
        f"{display_text(rec.get('project_id')) or display_text(rec.get('row_key'))}"
    )



def _system_coord_issue_style_v3649(
    rec: Dict[str, Any],
) -> Tuple[str, str]:
    """人工檢核選點地圖的底色永遠代表『系統判定』。

    人工是否已檢核只用外框表示，避免人工結果把原本異常類型的顏色蓋掉。
    """
    action = display_text(rec.get("action"))
    diagnosis = display_text(rec.get("diagnosis"))

    if action == "退回縣市重填":
        return "#C62828", "系統：建議退回縣市重填"
    if action == "中央確認轉換":
        return "#7B1FA2", "系統：疑似座標系統問題／中央確認轉換"
    if action == "中央確認欄位":
        return "#EF6C00", "系統：疑似欄位或X/Y顛倒"
    if action == "請縣市確認":
        return "#F9A825", "系統：需縣市確認"
    if action == "視需要抽查":
        return "#1565C0", "系統：視需要抽查"
    if action == "無需處理":
        return "#9E9E9E", "系統：目前無需處理"

    if "TWD67" in diagnosis:
        return "#7B1FA2", "系統：疑似TWD67"
    if "顛倒" in diagnosis:
        return "#EF6C00", "系統：疑似X/Y或經緯度顛倒"
    if "異常" in diagnosis or "無法" in diagnosis or "無座標" in diagnosis:
        return "#C62828", "系統：座標異常"
    return "#757575", "系統：其他／尚未分類"


def _coordinate_issue_in_merged_export_v3649(
    rec: Dict[str, Any],
) -> bool:
    """『全部異常與待確認』的整併規則。

    優先順序：
    1. 人工已確認原座標可用 → 排除，即使系統原本曾判異常。
    2. 人工有其他結論（退回、CRS問題、待確認）→ 納入。
    3. 尚未人工判讀 → 只要系統不是『無需處理』就納入。
    """
    manual = display_text(rec.get("manual_decision"))
    action = display_text(rec.get("action"))

    if manual == "確認原座標可用／不需處理":
        return False
    if manual:
        return True
    return action != "無需處理"


def _coordinate_issue_source_v3649(
    rec: Dict[str, Any],
) -> str:
    manual = display_text(rec.get("manual_decision"))
    action = display_text(rec.get("action"))
    if manual:
        if action != "無需處理":
            return "系統建議＋人工判定"
        return "人工判定"
    return "系統建議"


def _coordinate_issue_export_fingerprint_v3649(
    records: Sequence[Dict[str, Any]],
) -> str:
    compact = []
    for rec in sorted(
        [dict(r) for r in records],
        key=lambda r: (
            display_text(r.get("project_id")),
            display_text(r.get("row_key")),
        ),
    ):
        compact.append([
            display_text(rec.get("project_id")),
            display_text(rec.get("row_key")),
            display_text(rec.get("action")),
            display_text(rec.get("manual_decision")),
            display_text(rec.get("manual_reviewed_at")),
        ])
    raw = json.dumps(compact, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return hashlib.sha1(raw).hexdigest()[:16]


def _merge_latest_manual_into_records_v3649(
    base_records: Sequence[Dict[str, Any]],
    latest_registry: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """將 GitHub 最新人工結論套回目前系統建議結果。"""
    return _apply_manual_reviews_to_records_v3646(base_records, latest_registry)


def _build_all_counties_merged_issue_records_v3649(
    rows: Sequence["ProjectRow"],
    geo: Dict[str, Any],
    latest_registry: Dict[str, Any],
) -> Tuple[List[Dict[str, Any]], int]:
    """顯式執行時才做：快速建立全縣市『系統建議＋人工判定』整併異常清單。

    不重新開 current.xlsx；只用 ProjectRow + GIS + 最新 registry。
    """
    merged: Dict[str, Dict[str, Any]] = {}
    counties = sorted({
        display_text(p.county) for p in rows if display_text(p.county)
    })

    for county in counties:
        county_rows = [p for p in rows if display_text(p.county) == county]
        fast_records, _batch = _build_coordinate_audit_records_fast(
            county_rows, geo, county
        )
        checked = _merge_latest_manual_into_records_v3649(
            fast_records, latest_registry
        )
        for rec in checked:
            if not _coordinate_issue_in_merged_export_v3649(rec):
                continue
            key = _coordinate_review_key(rec)
            if not key:
                continue
            item = dict(rec)
            item["issue_source"] = _coordinate_issue_source_v3649(item)
            merged[key] = item

    # 若 registry 中仍有已人工判定、但目前 current.xlsx 已找不到的歷史案件，
    # 只要人工不是「確認無需處理」，也保留在整併清單。
    all_manual_records, missing_current = _all_latest_coordinate_review_records_v3648(
        rows, latest_registry
    )
    for rec in all_manual_records:
        key = _coordinate_review_key(rec)
        if not key or key in merged:
            continue
        if _coordinate_issue_in_merged_export_v3649(rec):
            item = dict(rec)
            item["issue_source"] = _coordinate_issue_source_v3649(item)
            merged[key] = item

    result = list(merged.values())
    result.sort(
        key=lambda r: (
            display_text(r.get("county")),
            display_text(r.get("project_name")),
            display_text(r.get("project_id")),
        )
    )
    return result, missing_current



def _build_manual_selection_map_v3647(
    records: Sequence[Dict[str, Any]],
    selected_keys: Sequence[str],
    display_scope: str,
    memory_key: str,
) -> Tuple[Optional[folium.Map], List[Tuple[Dict[str, Any], float, float, str]], int]:
    """V3.6.49：人工座標檢核直接在地圖點選。

    底色固定代表系統異常類型：
    紅＝退回縣市、紫＝疑似CRS、橘＝欄位/X-Y、黃＝請縣市確認、
    藍＝視需要抽查、灰＝無需處理。
    人工已檢核改用綠色外框；選取中則用粗白框。
    """
    selected = {display_text(x) for x in selected_keys if display_text(x)}
    visible: List[Dict[str, Any]] = []
    unmappable = 0

    for rec in records:
        row_key = display_text(rec.get("row_key"))
        manual = display_text(rec.get("manual_decision"))
        action = display_text(rec.get("action"))

        if display_scope == "尚未人工檢核":
            include = not manual
        elif display_scope == "系統建議異常／待確認":
            include = action != "無需處理"
        else:
            include = True

        # 已選案件始終保留。
        include = include or row_key in selected
        if not include:
            continue

        lon = parse_number(rec.get("current_lon"))
        lat = parse_number(rec.get("current_lat"))
        if not is_valid_wgs84(lon, lat):
            unmappable += 1
            continue

        visible.append(rec)

    if not visible:
        return None, [], unmappable

    pts = []
    point_rows: List[Tuple[Dict[str, Any], float, float, str]] = []
    for rec in visible:
        lon = float(rec["current_lon"])
        lat = float(rec["current_lat"])
        pts.append([lat, lon])
        point_rows.append(
            (rec, lat, lon, _manual_selection_marker_label_v3647(rec))
        )

    m = folium.Map(
        location=[
            sum(p[0] for p in pts) / len(pts),
            sum(p[1] for p in pts) / len(pts),
        ],
        zoom_start=10,
        tiles="OpenStreetMap",
        control_scale=True,
    )

    for rec, lat, lon, label in point_rows:
        row_key = display_text(rec.get("row_key"))
        is_selected = row_key in selected
        manual = display_text(rec.get("manual_decision"))
        diagnosis = display_text(rec.get("diagnosis"))
        action = display_text(rec.get("action"))
        fill, style_label = _system_coord_issue_style_v3649(rec)

        # V3.6.49：人工檢核不再把異常底色蓋成綠色。
        # 白框＝目前已選；綠框＝已有人工檢核；黑框＝尚未人工檢核。
        if is_selected:
            outline = "#FFFFFF"
            radius = 10.5
            weight = 4.5
        elif manual:
            outline = "#00A152"
            radius = 8.0
            weight = 3.5
        else:
            outline = "#111111"
            radius = 7.0
            weight = 2.0

        popup = (
            f"<b>{_popup_escape(display_text(rec.get('project_name')))}</b><br>"
            f"{_popup_escape(display_text(rec.get('project_id')) or '待確認ID')}<br>"
            f"<b>{_popup_escape(style_label)}</b><br>"
            f"系統檢核：{_popup_escape(diagnosis)}<br>"
            f"系統處理建議：{_popup_escape(action)}<br>"
            f"人工結果：{_popup_escape(manual or '尚未人工檢核')}<br><br>"
            f"<b>{'已選取；再點一次可取消' if is_selected else '點一下加入人工檢核'}</b>"
        )

        folium.CircleMarker(
            location=[lat, lon],
            radius=radius,
            color=outline,
            weight=weight,
            fill=True,
            fill_color=fill,
            fill_opacity=0.98,
            tooltip=label,
            popup=folium.Popup(popup, max_width=450),
        ).add_to(m)

    if len(pts) > 1:
        try:
            m.fit_bounds(pts, padding=(20, 20))
        except Exception:
            pass
    elif len(pts) == 1:
        m.location = pts[0]
        m.options["zoom"] = 15

    BrowserViewportMemory(memory_key).add_to(m)

    legend = """
    <div style="
      position:fixed; left:14px; bottom:28px; z-index:9999;
      background:#fff; color:#111; padding:9px 11px; border:1px solid #888;
      border-radius:7px; font-size:12px; line-height:1.65;
      box-shadow:0 1px 5px rgba(0,0,0,.22);">
      <b>系統座標檢核分類</b><br>
      <span style="color:#C62828">●</span> 建議退回縣市　
      <span style="color:#7B1FA2">●</span> 疑似座標系統<br>
      <span style="color:#EF6C00">●</span> 欄位／X-Y確認　
      <span style="color:#F9A825">●</span> 請縣市確認<br>
      <span style="color:#1565C0">●</span> 視需要抽查　
      <span style="color:#9E9E9E">●</span> 無需處理<br>
      <span style="color:#00A152">◎</span> 綠框＝已有人工檢核　
      <span style="color:#777">◉</span> 粗白框＝目前已選
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend))
    return m, point_rows, unmappable



def _match_clicked_manual_record_v3647(
    state: Optional[Dict[str, Any]],
    points: Sequence[Tuple[Dict[str, Any], float, float, str]],
) -> Optional[Dict[str, Any]]:
    if not state:
        return None

    tooltip = display_text(state.get("last_object_clicked_tooltip"))
    if tooltip:
        for rec, _lat, _lon, label in points:
            if tooltip == label:
                return rec

    obj = state.get("last_object_clicked")
    if not isinstance(obj, dict):
        return None
    try:
        lat = float(obj.get("lat"))
        lon = float(obj.get("lng"))
    except Exception:
        return None

    best = None
    best_d = 999.0
    for rec, plat, plon, _label in points:
        d = (plat - lat) ** 2 + (plon - lon) ** 2
        if d < best_d:
            best_d = d
            best = rec

    if best is not None and best_d <= 0.0006 ** 2:
        return best
    return None


def _manual_review_map_v3646(
    selected_records: Sequence[Dict[str, Any]],
    candidates: Dict[str, Dict[str, Any]],
) -> Optional[folium.Map]:
    pts: List[List[float]] = []
    for rec in selected_records:
        for xk, yk in [
            ("current_lon", "current_lat"),
            ("gis_ref_lon", "gis_ref_lat"),
        ]:
            lon = parse_number(rec.get(xk))
            lat = parse_number(rec.get(yk))
            if is_valid_wgs84(lon, lat):
                pts.append([lat, lon])
        cand = candidates.get(display_text(rec.get("row_key"))) or {}
        lon = parse_number(cand.get("lon"))
        lat = parse_number(cand.get("lat"))
        if is_valid_wgs84(lon, lat):
            pts.append([lat, lon])

    if not pts:
        return None

    m = folium.Map(
        location=[
            sum(p[0] for p in pts) / len(pts),
            sum(p[1] for p in pts) / len(pts),
        ],
        zoom_start=11,
        tiles="OpenStreetMap",
        control_scale=True,
    )
    fg_current = folium.FeatureGroup(name="系統目前解析位置", show=True)
    fg_candidate = folium.FeatureGroup(name="人工試轉候選", show=True)
    fg_ref = folium.FeatureGroup(name="已人工校正GIS參考點", show=True)
    fg_link = folium.FeatureGroup(name="位置差異", show=True)

    for rec in selected_records:
        name = display_text(rec.get("project_name"))
        pid = display_text(rec.get("project_id"))
        row_key = display_text(rec.get("row_key"))
        cur_lon = parse_number(rec.get("current_lon"))
        cur_lat = parse_number(rec.get("current_lat"))
        ref_lon = parse_number(rec.get("gis_ref_lon"))
        ref_lat = parse_number(rec.get("gis_ref_lat"))
        cand = candidates.get(row_key) or {}
        cand_lon = parse_number(cand.get("lon"))
        cand_lat = parse_number(cand.get("lat"))

        popup = f"<b>{_popup_escape(name)}</b><br>{_popup_escape(pid)}"
        if is_valid_wgs84(cur_lon, cur_lat):
            folium.CircleMarker(
                [cur_lat, cur_lon], radius=7, color="#000000", weight=2,
                fill=True, fill_color="#616161", fill_opacity=0.95,
                tooltip=f"目前｜{name}", popup=folium.Popup(popup, max_width=420),
            ).add_to(fg_current)

        if is_valid_wgs84(cand_lon, cand_lat):
            folium.CircleMarker(
                [cand_lat, cand_lon], radius=8, color="#000000", weight=2,
                fill=True, fill_color="#FB8C00", fill_opacity=0.98,
                tooltip=f"人工試轉｜{name}",
            ).add_to(fg_candidate)
            if is_valid_wgs84(cur_lon, cur_lat):
                folium.PolyLine(
                    [[cur_lat, cur_lon], [cand_lat, cand_lon]],
                    color="#FB8C00", weight=2.5, opacity=0.8, dash_array="6,6",
                ).add_to(fg_link)

        if is_valid_wgs84(ref_lon, ref_lat):
            folium.CircleMarker(
                [ref_lat, ref_lon], radius=8, color="#000000", weight=2.5,
                fill=True, fill_color="#1565C0", fill_opacity=0.98,
                tooltip=f"GIS已校正｜{name}",
            ).add_to(fg_ref)
            if is_valid_wgs84(cand_lon, cand_lat):
                folium.PolyLine(
                    [[cand_lat, cand_lon], [ref_lat, ref_lon]],
                    color="#1565C0", weight=2, opacity=0.7,
                ).add_to(fg_link)

    fg_current.add_to(m)
    fg_candidate.add_to(m)
    fg_ref.add_to(m)
    fg_link.add_to(m)
    folium.LayerControl(collapsed=False).add_to(m)
    try:
        m.fit_bounds(pts)
    except Exception:
        pass
    return m



def _manual_review_map_v3648(
    selected_records: Sequence[Dict[str, Any]],
    candidates: Dict[str, Dict[str, Any]],
    tested_crs: str,
    show_current: bool = True,
) -> Optional[folium.Map]:
    """V3.6.48：讓切換 CRS 後的候選位置非常明顯。

    灰色小點＝目前解析位置
    橘色大點白框＝本次 CRS 試轉候選
    藍色＝人工 GIS 校正參考點
    橘線＝目前位置 → 候選位置，tooltip 顯示位移距離
    """
    pts: List[List[float]] = []
    project_points: Dict[str, List[List[float]]] = {}

    for rec in selected_records:
        row_key = display_text(rec.get("row_key"))
        ppts: List[List[float]] = []

        if show_current:
            lon = parse_number(rec.get("current_lon"))
            lat = parse_number(rec.get("current_lat"))
            if is_valid_wgs84(lon, lat):
                pts.append([lat, lon])
                ppts.append([lat, lon])

        rlon = parse_number(rec.get("gis_ref_lon"))
        rlat = parse_number(rec.get("gis_ref_lat"))
        if is_valid_wgs84(rlon, rlat):
            pts.append([rlat, rlon])
            ppts.append([rlat, rlon])

        cand = candidates.get(row_key) or {}
        clon = parse_number(cand.get("lon"))
        clat = parse_number(cand.get("lat"))
        if is_valid_wgs84(clon, clat):
            pts.append([clat, clon])
            ppts.append([clat, clon])

        project_points[row_key] = ppts

    if not pts:
        return None

    m = folium.Map(
        location=[
            sum(p[0] for p in pts) / len(pts),
            sum(p[1] for p in pts) / len(pts),
        ],
        zoom_start=12,
        tiles="OpenStreetMap",
        control_scale=True,
    )

    fg_current = folium.FeatureGroup(name="① 系統目前解析位置", show=show_current)
    fg_candidate = folium.FeatureGroup(name=f"② 本次試轉：{tested_crs}", show=True)
    fg_ref = folium.FeatureGroup(name="③ 已人工校正 GIS 參考點", show=True)
    fg_link = folium.FeatureGroup(name="目前位置 → 試轉候選", show=show_current)

    for rec in selected_records:
        name = display_text(rec.get("project_name"))
        pid = display_text(rec.get("project_id"))
        row_key = display_text(rec.get("row_key"))
        cur_lon = parse_number(rec.get("current_lon"))
        cur_lat = parse_number(rec.get("current_lat"))
        ref_lon = parse_number(rec.get("gis_ref_lon"))
        ref_lat = parse_number(rec.get("gis_ref_lat"))
        cand = candidates.get(row_key) or {}
        cand_lon = parse_number(cand.get("lon"))
        cand_lat = parse_number(cand.get("lat"))

        if show_current and is_valid_wgs84(cur_lon, cur_lat):
            folium.CircleMarker(
                [cur_lat, cur_lon],
                radius=5,
                color="#222222",
                weight=1.5,
                fill=True,
                fill_color="#777777",
                fill_opacity=0.86,
                tooltip=f"原位置｜{name}",
                popup=folium.Popup(
                    f"<b>{_popup_escape(name)}</b><br>"
                    f"{_popup_escape(pid)}<br>"
                    f"系統目前位置：{cur_lon:.7f}, {cur_lat:.7f}",
                    max_width=430,
                ),
            ).add_to(fg_current)

        displacement = None
        if (
            is_valid_wgs84(cur_lon, cur_lat)
            and is_valid_wgs84(cand_lon, cand_lat)
        ):
            displacement = _haversine_m(
                float(cur_lat), float(cur_lon),
                float(cand_lat), float(cand_lon),
            )

        if is_valid_wgs84(cand_lon, cand_lat):
            displacement_txt = (
                f"{displacement:,.0f} m" if displacement is not None else "無法計算"
            )
            # 大型橘色候選＋白框，切換 CRS 後最醒目
            folium.CircleMarker(
                [cand_lat, cand_lon],
                radius=11,
                color="#FFFFFF",
                weight=4,
                fill=True,
                fill_color="#FB8C00",
                fill_opacity=1.0,
                tooltip=f"🟠 {tested_crs}｜{name}｜位移 {displacement_txt}",
                popup=folium.Popup(
                    f"<b>{_popup_escape(name)}</b><br>"
                    f"{_popup_escape(pid)}<br>"
                    f"<b>本次試轉 CRS：{_popup_escape(tested_crs)}</b><br>"
                    f"候選位置：{cand_lon:.7f}, {cand_lat:.7f}<br>"
                    f"相對目前位置位移：{_popup_escape(displacement_txt)}",
                    max_width=460,
                ),
            ).add_to(fg_candidate)

            if show_current and is_valid_wgs84(cur_lon, cur_lat):
                folium.PolyLine(
                    [[cur_lat, cur_lon], [cand_lat, cand_lon]],
                    color="#FB8C00",
                    weight=5,
                    opacity=0.88,
                    dash_array="9,7",
                    tooltip=f"目前位置 → {tested_crs}：位移 {displacement_txt}",
                ).add_to(fg_link)

        if is_valid_wgs84(ref_lon, ref_lat):
            folium.CircleMarker(
                [ref_lat, ref_lon],
                radius=8,
                color="#FFFFFF",
                weight=3,
                fill=True,
                fill_color="#1565C0",
                fill_opacity=0.98,
                tooltip=f"🔵 GIS已校正｜{name}",
                popup=folium.Popup(
                    f"<b>{_popup_escape(name)}</b><br>"
                    f"已人工校正 GIS：{ref_lon:.7f}, {ref_lat:.7f}",
                    max_width=430,
                ),
            ).add_to(fg_ref)

    if show_current:
        fg_current.add_to(m)
    fg_candidate.add_to(m)
    fg_ref.add_to(m)
    if show_current:
        fg_link.add_to(m)
    folium.LayerControl(collapsed=False).add_to(m)

    # 讓目前使用中的 CRS 在地圖上也明確可見
    banner = f"""
    <div style="
      position:fixed; top:12px; left:50%; transform:translateX(-50%);
      z-index:9999; background:rgba(255,255,255,.96); color:#111;
      border:3px solid #FB8C00; border-radius:9px; padding:8px 14px;
      font-size:14px; font-weight:700; box-shadow:0 2px 8px rgba(0,0,0,.25);">
      🟠 本次人工試轉：{_popup_escape(tested_crs)}
    </div>
    """
    m.get_root().html.add_child(folium.Element(banner))

    try:
        if len(selected_records) == 1:
            # 單件檢核：只框住這件的原點/候選/GIS參考點，並允許放大至 18。
            row_key = display_text(selected_records[0].get("row_key"))
            one_pts = project_points.get(row_key) or []
            if len(one_pts) >= 2:
                m.fit_bounds(one_pts, padding=(45, 45), max_zoom=18)
            elif len(one_pts) == 1:
                m.location = one_pts[0]
                m.options["zoom"] = 17
        else:
            m.fit_bounds(pts, padding=(30, 30), max_zoom=15)
    except Exception:
        pass

    return m


def _latest_registry_only_v3648(
    svc: "EngineeringGISService",
) -> Dict[str, Any]:
    """只讀 GitHub 最新 id_registry.json，不重新載入 current.xlsx / GeoJSON。"""
    commit, _tree = svc.store.get_head_commit()
    raw = svc.store.read_file(
        svc.cfg.registry_path,
        ref=commit,
        allow_missing=True,
    )
    return json_load_bytes(raw, empty_registry())


def _all_latest_coordinate_review_records_v3648(
    rows: Sequence["ProjectRow"],
    registry: Dict[str, Any],
) -> Tuple[List[Dict[str, Any]], int]:
    """把所有使用者分批儲存的『每件最新人工結論』映射回目前工程資料。

    找不到 current.xlsx 對應列的舊案件仍保留在清單，惟工程內容/經費/原座標可能留白。
    """
    bucket = (registry or {}).get("coordinate_reviews") or {}
    if not isinstance(bucket, dict):
        return [], 0

    by_pid = {
        display_text(p.project_id): p
        for p in rows if display_text(p.project_id)
    }
    by_row = {
        display_text(p.row_key): p
        for p in rows if display_text(p.row_key)
    }

    records: List[Dict[str, Any]] = []
    missing_current = 0

    for _key, rv0 in bucket.items():
        if not isinstance(rv0, dict):
            continue
        rv = dict(rv0)
        pid = display_text(rv.get("project_id"))
        row_key = display_text(rv.get("row_key"))
        p = by_pid.get(pid) or by_row.get(row_key)

        if p is None:
            missing_current += 1
            records.append({
                "row_key": row_key,
                "project_id": pid,
                "sheet_name": "",
                "county": display_text(rv.get("county")),
                "project_name": display_text(rv.get("project_name")),
                "project_content": "",
                "water_system": "",
                "status": "",
                "unit": "",
                "address": "",
                "cost_k": None,
                "raw_source": "",
                "raw_x": None,
                "raw_y": None,
                "current_lon": None,
                "current_lat": None,
                "diagnosis": display_text(rv.get("system_suggestion")),
                "action": display_text(rv.get("system_action")),
                "manual_decision": display_text(rv.get("manual_decision")),
                "manual_tested_crs": display_text(rv.get("tested_crs")),
                "manual_reviewed_by": display_text(rv.get("reviewed_by")),
                "manual_reviewed_at": display_text(rv.get("reviewed_at")),
                "manual_note": display_text(rv.get("review_note")),
                "manual_candidate_lon": parse_number(rv.get("candidate_lon")),
                "manual_candidate_lat": parse_number(rv.get("candidate_lat")),
            })
            continue

        records.append({
            "row_key": p.row_key,
            "project_id": p.project_id,
            "sheet_name": p.sheet_name,
            "county": p.county,
            "project_name": p.project_name,
            "project_content": p.project_content,
            "water_system": p.water_system,
            "status": p.status,
            "unit": p.unit,
            "address": p.address,
            "cost_k": None,
            "raw_source": "",
            "raw_x": None,
            "raw_y": None,
            "current_lon": p.lon,
            "current_lat": p.lat,
            "diagnosis": display_text(rv.get("system_suggestion")),
            "action": display_text(rv.get("system_action")),
            "manual_decision": display_text(rv.get("manual_decision")),
            "manual_tested_crs": display_text(rv.get("tested_crs")),
            "manual_reviewed_by": display_text(rv.get("reviewed_by")),
            "manual_reviewed_at": display_text(rv.get("reviewed_at")),
            "manual_note": display_text(rv.get("review_note")),
            "manual_candidate_lon": parse_number(rv.get("candidate_lon")),
            "manual_candidate_lat": parse_number(rv.get("candidate_lat")),
        })

    records.sort(
        key=lambda r: (
            display_text(r.get("county")),
            display_text(r.get("project_name")),
            display_text(r.get("project_id")),
        )
    )
    return records, missing_current


def _coordinate_review_registry_fingerprint_v3648(
    registry: Dict[str, Any],
) -> str:
    bucket = (registry or {}).get("coordinate_reviews") or {}
    if not isinstance(bucket, dict):
        return "none"
    compact = []
    for key in sorted(bucket.keys()):
        v = bucket.get(key) or {}
        compact.append([
            key,
            display_text(v.get("manual_decision")),
            display_text(v.get("reviewed_at")),
            display_text(v.get("reviewed_by")),
        ])
    raw = json.dumps(compact, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return hashlib.sha1(raw).hexdigest()[:16]


def _build_coordinate_review_zip_v3648(
    grouped_records: Dict[str, Sequence[Dict[str, Any]]],
) -> bytes:
    """將已補齊資料的各縣市複核表包成單一 ZIP。"""
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for county in sorted(grouped_records.keys()):
            recs = list(grouped_records[county])
            if not recs:
                continue
            safe_county = re.sub(r'[\\/:*?"<>|]+', "_", county) or "縣市"
            xlsx = _build_county_coordinate_review_xlsx(county, recs)
            z.writestr(f"{safe_county}_工程座標複核表.xlsx", xlsx)
    return out.getvalue()



def _apply_manual_reviews_to_records_v3646(
    records: Sequence[Dict[str, Any]],
    registry: Dict[str, Any],
) -> List[Dict[str, Any]]:
    out = []
    for rec0 in records:
        rec = dict(rec0)
        review = _coordinate_review_for_record(rec, registry)
        rec["manual_decision"] = display_text(review.get("manual_decision"))
        rec["manual_tested_crs"] = display_text(review.get("tested_crs"))
        rec["manual_reviewed_by"] = display_text(review.get("reviewed_by"))
        rec["manual_reviewed_at"] = display_text(review.get("reviewed_at"))
        rec["manual_note"] = display_text(review.get("review_note"))
        out.append(rec)
    return out



# ============================================================
# V3.6.50 縣市座標回填匯入／中央審核
# ============================================================
def _normalize_county_return_crs_v3650(value: Any) -> str:
    raw = display_text(value)
    if not raw:
        return ""
    s = (
        raw.upper()
        .replace("　", "")
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace("ZONE", "")
        .replace("TM2", "")
    )
    if "不知道" in raw or "不確定" in raw:
        return "不知道"
    if "WGS84" in s or "EPSG4326" in s:
        return "WGS84"
    if "TWD97" in s or "EPSG3826" in s or "EPSG3825" in s:
        if "119" in s or "3825" in s:
            return "TWD97 TM2 zone 119"
        return "TWD97 TM2 zone 121"
    if "TWD67" in s or "EPSG3828" in s or "EPSG3827" in s:
        if "119" in s or "3827" in s:
            return "TWD67 TM2 zone 119"
        return "TWD67 TM2 zone 121"
    return raw


def _transform_county_return_coordinate_v3650(
    crs: Any,
    x: Any,
    y: Any,
) -> Dict[str, Any]:
    norm_crs = _normalize_county_return_crs_v3650(crs)
    xv = parse_number(x)
    yv = parse_number(y)
    out = {
        "crs": norm_crs,
        "raw_x": xv,
        "raw_y": yv,
        "lon": None,
        "lat": None,
        "valid": False,
        "swapped_suggestion": False,
        "message": "",
    }
    if xv is None or yv is None:
        out["message"] = "新 X/Y 座標未完整填寫。"
        return out
    if not norm_crs or norm_crs == "不知道":
        out["message"] = "縣市未確認座標系統，無法自動轉成 WGS84。"
        return out

    try:
        if norm_crs == "WGS84":
            lon, lat = float(xv), float(yv)
            if not is_valid_wgs84(lon, lat) and is_valid_wgs84(yv, xv):
                out["swapped_suggestion"] = True
                out["message"] = "WGS84 原 X/Y 無效，但 X/Y 對調後落在合理範圍，疑似填反。"
                return out
        elif norm_crs == "TWD97 TM2 zone 121":
            lon, lat = TRANSFORMER_121.transform(float(xv), float(yv))
        elif norm_crs == "TWD97 TM2 zone 119":
            lon, lat = TRANSFORMER_119.transform(float(xv), float(yv))
        elif norm_crs == "TWD67 TM2 zone 121":
            lon, lat = TRANSFORMER_67_121.transform(float(xv), float(yv))
        elif norm_crs == "TWD67 TM2 zone 119":
            lon, lat = TRANSFORMER_67_119.transform(float(xv), float(yv))
        else:
            out["message"] = f"目前不支援此座標系統：{norm_crs}"
            return out

        if not is_valid_wgs84(lon, lat):
            # TM2 也做一次 X/Y 對調提示，但不自動採用。
            if norm_crs != "WGS84":
                try:
                    if norm_crs == "TWD97 TM2 zone 121":
                        slon, slat = TRANSFORMER_121.transform(float(yv), float(xv))
                    elif norm_crs == "TWD97 TM2 zone 119":
                        slon, slat = TRANSFORMER_119.transform(float(yv), float(xv))
                    elif norm_crs == "TWD67 TM2 zone 121":
                        slon, slat = TRANSFORMER_67_121.transform(float(yv), float(xv))
                    else:
                        slon, slat = TRANSFORMER_67_119.transform(float(yv), float(xv))
                    if is_valid_wgs84(slon, slat):
                        out["swapped_suggestion"] = True
                        out["message"] = "原 X/Y 轉換無效，但 X/Y 對調後合理，疑似填反。"
                        return out
                except Exception:
                    pass
            out["message"] = "轉換後不在臺灣及離島合理 WGS84 範圍。"
            return out

        out.update({
            "lon": float(lon),
            "lat": float(lat),
            "valid": True,
            "message": "座標可轉換。",
        })
        return out
    except Exception as exc:
        out["message"] = f"座標轉換失敗：{exc}"
        return out


def _county_return_submission_hash_v3650(
    project_id: Any,
    crs: Any,
    x: Any,
    y: Any,
) -> str:
    raw = "|".join([
        display_text(project_id),
        _normalize_county_return_crs_v3650(crs),
        display_text(x),
        display_text(y),
    ])
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def _return_sheet_header_map_v3650(ws) -> Tuple[int, Dict[str, int]]:
    wanted = {
        norm_text("系統工程ID"),
        norm_text("縣市填報座標系統"),
        norm_text("縣市新X/經度"),
        norm_text("縣市新Y/緯度"),
    }
    for r in range(1, min(ws.max_row, 12) + 1):
        hmap = {}
        for c in range(1, ws.max_column + 1):
            txt = norm_text(ws.cell(r, c).value)
            if txt:
                hmap[txt] = c
        if norm_text("系統工程ID") in hmap and len(wanted.intersection(hmap)) >= 3:
            return r, hmap
    return 0, {}


def _parse_county_return_workbooks_v3650(
    files: Sequence[Tuple[str, bytes]],
    rows: Sequence["ProjectRow"],
    geo: Dict[str, Any],
    registry: Dict[str, Any],
) -> Tuple[List[Dict[str, Any]], List[str]]:
    id_map: Dict[str, List["ProjectRow"]] = {}
    for p in rows:
        if display_text(p.project_id):
            id_map.setdefault(display_text(p.project_id), []).append(p)

    latest_reviews = (registry or {}).get("coordinate_return_reviews") or {}
    if not isinstance(latest_reviews, dict):
        latest_reviews = {}

    records: List[Dict[str, Any]] = []
    global_messages: List[str] = []

    for file_name, data in files:
        try:
            wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
        except Exception as exc:
            global_messages.append(f"{file_name}：無法開啟 Excel（{exc}）")
            continue

        found_sheet = False
        for ws in wb.worksheets:
            header_row, hmap = _return_sheet_header_map_v3650(ws)
            if not header_row:
                continue
            found_sheet = True

            def col(name: str) -> Optional[int]:
                return hmap.get(norm_text(name))

            c_pid = col("系統工程ID")
            c_county = col("縣市")
            c_name = col("工程名稱")
            c_crs = col("縣市填報座標系統")
            c_x = col("縣市新X/經度")
            c_y = col("縣市新Y/緯度")
            c_note = col("縣市備註")

            for rr in range(header_row + 1, ws.max_row + 1):
                pid = display_text(ws.cell(rr, c_pid).value) if c_pid else ""
                crs_raw = display_text(ws.cell(rr, c_crs).value) if c_crs else ""
                x_raw = ws.cell(rr, c_x).value if c_x else None
                y_raw = ws.cell(rr, c_y).value if c_y else None
                note = display_text(ws.cell(rr, c_note).value) if c_note else ""

                # 完全空白列略過；只要 PID 或縣市有回填，就視為一筆。
                if not pid and not crs_raw and parse_number(x_raw) is None and parse_number(y_raw) is None:
                    continue

                workbook_county = display_text(
                    ws.cell(rr, c_county).value
                ) if c_county else ""
                workbook_name = display_text(
                    ws.cell(rr, c_name).value
                ) if c_name else ""

                transform = _transform_county_return_coordinate_v3650(
                    crs_raw, x_raw, y_raw
                )
                messages = []
                level = "可審核"
                reviewable = True

                matches = id_map.get(pid, [])
                p = matches[0] if len(matches) == 1 else None
                if not pid:
                    messages.append("缺少系統工程ID。")
                    level = "錯誤"
                    reviewable = False
                elif not matches:
                    messages.append("系統找不到此工程ID。")
                    level = "錯誤"
                    reviewable = False
                elif len(matches) > 1:
                    messages.append("目前 current.xlsx 有重複工程ID。")
                    level = "錯誤"
                    reviewable = False

                if not transform.get("valid"):
                    messages.append(display_text(transform.get("message")))
                    level = "錯誤"
                    reviewable = False

                current_county = display_text(p.county) if p else workbook_county
                current_name = display_text(p.project_name) if p else workbook_name
                if (
                    p
                    and workbook_county
                    and normalize_county_name(workbook_county)
                    and normalize_county_name(current_county)
                    and normalize_county_name(workbook_county) != normalize_county_name(current_county)
                ):
                    messages.append(
                        f"縣市欄與目前系統不一致：回填={workbook_county}、系統={current_county}。"
                    )
                    if level != "錯誤":
                        level = "警告"

                old_lon = old_lat = None
                old_source = ""
                if p:
                    f = original_feature_for_project(geo, p.project_id)
                    geom = (f or {}).get("geometry") or {}
                    coords = geom.get("coordinates") or []
                    if geom.get("type") == "Point" and len(coords) >= 2:
                        try:
                            glon, glat = float(coords[0]), float(coords[1])
                            if is_valid_wgs84(glon, glat):
                                old_lon, old_lat = glon, glat
                                old_source = "GIS代表點"
                        except Exception:
                            pass
                    if old_lon is None and p.lon is not None and p.lat is not None:
                        old_lon, old_lat = float(p.lon), float(p.lat)
                        old_source = "current.xlsx解析點"

                distance = None
                if (
                    transform.get("valid")
                    and old_lon is not None
                    and old_lat is not None
                ):
                    distance = _haversine_m(
                        float(old_lat), float(old_lon),
                        float(transform["lat"]), float(transform["lon"]),
                    )
                    if distance >= 2000:
                        messages.append(f"新點與目前GIS位置相差約 {distance:,.0f}m（高度警告）。")
                        if level != "錯誤":
                            level = "警告"
                    elif distance >= 500:
                        messages.append(f"新點與目前GIS位置相差約 {distance:,.0f}m（請特別確認）。")
                        if level != "錯誤":
                            level = "警告"
                    elif distance >= 100:
                        messages.append(f"新點與目前GIS位置相差約 {distance:,.0f}m。")

                submission_hash = _county_return_submission_hash_v3650(
                    pid, crs_raw, x_raw, y_raw
                )
                previous = latest_reviews.get(pid, {}) if pid else {}
                if not isinstance(previous, dict):
                    previous = {}
                same_submission = (
                    display_text(previous.get("submission_hash")) == submission_hash
                )
                previous_decision = (
                    display_text(previous.get("review_decision"))
                    if same_submission else ""
                )

                rec = {
                    "record_key": f"{file_name}::{ws.title}::{rr}",
                    "source_file": file_name,
                    "source_sheet": ws.title,
                    "source_row": rr,
                    "project_id": pid,
                    "row_key": p.row_key if p else "",
                    "project_name": current_name,
                    "county": current_county,
                    "status": p.status if p else "",
                    "workbook_project_name": workbook_name,
                    "workbook_county": workbook_county,
                    "returned_crs_raw": crs_raw,
                    "returned_crs": transform.get("crs"),
                    "raw_x": parse_number(x_raw),
                    "raw_y": parse_number(y_raw),
                    "candidate_lon": transform.get("lon"),
                    "candidate_lat": transform.get("lat"),
                    "county_note": note,
                    "old_lon": old_lon,
                    "old_lat": old_lat,
                    "old_source": old_source,
                    "distance_from_gis_m": distance,
                    "validation_status": level,
                    "validation_message": "；".join([m for m in messages if m]),
                    "reviewable": bool(reviewable),
                    "submission_hash": submission_hash,
                    "previous_decision": previous_decision,
                    "previous_reviewed_by": (
                        display_text(previous.get("reviewed_by"))
                        if same_submission else ""
                    ),
                    "previous_reviewed_at": (
                        display_text(previous.get("reviewed_at"))
                        if same_submission else ""
                    ),
                }
                records.append(rec)

        try:
            wb.close()
        except Exception:
            pass

        if not found_sheet:
            global_messages.append(
                f"{file_name}：找不到『系統工程ID／縣市填報座標系統／新X／新Y』欄位。"
            )

    # 多檔或同一檔重複 PID：全部標成錯誤，避免同批不知採哪一筆。
    by_pid: Dict[str, List[Dict[str, Any]]] = {}
    for rec in records:
        pid = display_text(rec.get("project_id"))
        if pid:
            by_pid.setdefault(pid, []).append(rec)

    for pid, items in by_pid.items():
        if len(items) <= 1:
            continue
        for rec in items:
            rec["validation_status"] = "錯誤"
            rec["reviewable"] = False
            msg = display_text(rec.get("validation_message"))
            extra = f"本次上傳中同一工程ID {pid} 出現 {len(items)} 筆，請先整理後再審核。"
            rec["validation_message"] = f"{msg}；{extra}" if msg else extra

    return records, global_messages


def _county_return_display_rows_v3650(
    records: Sequence[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    out = []
    for r in records:
        dist = r.get("distance_from_gis_m")
        out.append({
            "來源檔案": r.get("source_file"),
            "系統工程ID": r.get("project_id"),
            "縣市": r.get("county"),
            "工程名稱": r.get("project_name"),
            "縣市填報CRS": r.get("returned_crs"),
            "新X/經度": r.get("raw_x"),
            "新Y/緯度": r.get("raw_y"),
            "轉換後經度": None if r.get("candidate_lon") is None else round(float(r["candidate_lon"]), 7),
            "轉換後緯度": None if r.get("candidate_lat") is None else round(float(r["candidate_lat"]), 7),
            "與目前GIS位移(m)": None if dist is None else round(float(dist), 1),
            "自動檢核": r.get("validation_status"),
            "檢核訊息": r.get("validation_message"),
            "既有中央審核": r.get("previous_decision"),
            "縣市備註": r.get("county_note"),
        })
    return out


def _county_return_marker_label_v3650(rec: Dict[str, Any]) -> str:
    return (
        f"縣市回填｜{display_text(rec.get('project_name'))}｜"
        f"{display_text(rec.get('project_id'))}｜"
        f"{display_text(rec.get('submission_hash'))}"
    )


class CountyReturnViewportMemory(MacroElement):
    """點選回填點位後，同時恢復地圖視角與整個 Streamlit 頁面位置。"""

    _template = Template(r"""
    {% macro script(this, kwargs) %}
    (function() {
        var mapObj = {{ this._parent.get_name() }};
        var storageKey = {{ this.storage_key_json | safe }};
        var pageIntentKey = 'wra_gis_page_scroll_intent_v3659';
        function storageObj() {
            try { return window.parent.sessionStorage; } catch (e) {}
            try { return window.sessionStorage; } catch (e) {}
            return null;
        }
        function readViewport() {
            try {
                var raw = storageObj().getItem(storageKey);
                var view = raw ? JSON.parse(raw) : null;
                if (!view || typeof view.lat !== 'number' ||
                    typeof view.lng !== 'number' || typeof view.zoom !== 'number' ||
                    !isFinite(view.lat) || !isFinite(view.lng) || !isFinite(view.zoom)) return null;
                return view;
            } catch (e) { return null; }
        }
        function pageSnapshot() {
            try {
                var p = window.parent, frame = window.frameElement;
                var y = Number(p.scrollY || p.pageYOffset || 0);
                var top = frame ? Number(frame.getBoundingClientRect().top) : NaN;
                return {page_y:isFinite(y)?y:0, frame_top:isFinite(top)?top:null};
            } catch(e) { return {page_y:0,frame_top:null}; }
        }
        function saveViewport(markPageIntent) {
            try {
                var center = mapObj.getCenter();
                var page = pageSnapshot();
                var payload = {
                    lat: Number(center.lat), lng: Number(center.lng),
                    zoom: Number(mapObj.getZoom()), page_y:Number(page.page_y),
                    frame_top:page.frame_top, ts:Date.now()
                };
                storageObj().setItem(storageKey, JSON.stringify(payload));
                if (markPageIntent === true) {
                    storageObj().setItem(pageIntentKey, JSON.stringify({
                        y:Number(page.page_y), frame_top:page.frame_top,
                        view_key:storageKey, source:'county-return-map', ts:Date.now()
                    }));
                    storageObj().setItem('wra_gis_page_scroll_v3659', JSON.stringify({
                        y:Number(page.page_y), ts:Date.now()
                    }));
                }
            } catch (e) {}
        }
        function readPageIntent() {
            try {
                var raw = storageObj().getItem(pageIntentKey);
                var obj = raw ? JSON.parse(raw) : null;
                if (!obj || obj.view_key !== storageKey || !isFinite(obj.y) ||
                    !isFinite(obj.ts) || Date.now()-Number(obj.ts) >= 8000) return null;
                return obj;
            } catch(e) { return null; }
        }
        function restorePageOnce(intent) {
            try {
                var p=window.parent, frame=window.frameElement;
                var currentY=Number(p.scrollY||p.pageYOffset||0), targetY=Number(intent.y);
                if(frame && intent.frame_top!==null && isFinite(intent.frame_top)) {
                    var currentTop=Number(frame.getBoundingClientRect().top);
                    if(isFinite(currentTop)) targetY=currentY+currentTop-Number(intent.frame_top);
                }
                targetY=Math.max(0,targetY);
                try { p.scrollTo({top:targetY,left:0,behavior:'auto'}); }
                catch(e) { try { p.scrollTo(0,targetY); } catch(_e) {} }
            } catch(e) {}
        }
        function schedulePageRestore() {
            var intent=readPageIntent();
            if(!intent) return;
            [0,70,220,500,900,1450,2050].forEach(function(ms){
                setTimeout(function(){ restorePageOnce(intent); },ms);
            });
        }
        function saveInteraction(){ saveViewport(true); }
        var saved = readViewport();
        // 先恢復視角，再監聽移動；避免地圖初始 fitBounds 把上次位置蓋掉。
        if (saved) {
            try { mapObj.setView([saved.lat, saved.lng], saved.zoom, {animate: false}); } catch (e) {}
        }
        schedulePageRestore();
        mapObj.on('moveend', saveInteraction);
        // 沒有先平移／縮放也能點選，因此要在點位 click 時記下當前視角。
        mapObj.on('click', saveInteraction);
        mapObj.eachLayer(function(layer) {
            if (layer instanceof L.CircleMarker) layer.on('click', saveInteraction);
        });
        try {
            var container=mapObj.getContainer();
            if(container && container.addEventListener) container.addEventListener('pointerdown',saveInteraction,true);
        } catch(e) {}
        if(!saved) setTimeout(function(){ saveViewport(false); },0);
    })();
    {% endmacro %}
    """)

    def __init__(self, memory_key: str):
        super().__init__()
        self._name = "CountyReturnViewportMemory"
        self.storage_key_json = json.dumps(
            "wra_gis_county_return_view_v3656::" + memory_key,
            ensure_ascii=False,
        )


def _build_county_return_review_map_v3650(
    records: Sequence[Dict[str, Any]],
    selected_keys: Sequence[str],
    geo: Dict[str, Any],
    county_filter: str,
    scope: str,
    memory_key: Optional[str] = None,
) -> Tuple[Optional[folium.Map], List[Tuple[Dict[str, Any], float, float, str]]]:
    selected = {display_text(x) for x in selected_keys if display_text(x)}
    visible = []

    for r in records:
        if county_filter != "全部縣市" and display_text(r.get("county")) != county_filter:
            continue
        if scope == "只看尚未中央審核" and display_text(r.get("previous_decision")):
            continue
        if scope == "只看有警告/錯誤" and r.get("validation_status") == "可審核":
            continue
        if not r.get("reviewable"):
            continue
        lon = parse_number(r.get("candidate_lon"))
        lat = parse_number(r.get("candidate_lat"))
        if not is_valid_wgs84(lon, lat):
            continue
        visible.append(r)

    # 已選項目固定保留
    selected_map = {display_text(r.get("record_key")): r for r in records}
    for key in selected:
        rec = selected_map.get(key)
        if rec and rec not in visible and rec.get("reviewable"):
            lon = parse_number(rec.get("candidate_lon"))
            lat = parse_number(rec.get("candidate_lat"))
            if is_valid_wgs84(lon, lat):
                visible.append(rec)

    if not visible:
        return None, []

    pts = [[float(r["candidate_lat"]), float(r["candidate_lon"])] for r in visible]
    m = folium.Map(
        location=[
            sum(p[0] for p in pts) / len(pts),
            sum(p[1] for p in pts) / len(pts),
        ],
        zoom_start=10,
        tiles="OpenStreetMap",
        control_scale=True,
    )

    point_rows = []
    selected_pids = set()

    for r in visible:
        key = display_text(r.get("record_key"))
        pid = display_text(r.get("project_id"))
        name = display_text(r.get("project_name"))
        lon = float(r["candidate_lon"])
        lat = float(r["candidate_lat"])
        old_lon = parse_number(r.get("old_lon"))
        old_lat = parse_number(r.get("old_lat"))
        decision = display_text(r.get("previous_decision"))
        level = display_text(r.get("validation_status"))
        is_selected = key in selected

        if decision == "核准縣府新座標":
            fill = "#2E7D32"
        elif decision == "退回縣府再修正":
            fill = "#C62828"
        elif decision == "待確認":
            fill = "#F9A825"
        elif decision == "採用中央既有GIS點位":
            fill = "#1565C0"
        elif level == "警告":
            fill = "#EF6C00"
        else:
            fill = "#FB8C00"

        outline = "#FFFFFF" if is_selected else "#111111"
        weight = 4 if is_selected else 2
        radius = 10 if is_selected else 8

        dist = r.get("distance_from_gis_m")
        dist_txt = "無目前GIS參考點" if dist is None else f"{float(dist):,.0f}m"
        popup = (
            f"<b>{_popup_escape(name)}</b><br>"
            f"{_popup_escape(pid)}<br>"
            f"縣市填報 CRS：{_popup_escape(display_text(r.get('returned_crs')))}<br>"
            f"新位置：{lon:.7f}, {lat:.7f}<br>"
            f"與目前 GIS 位移：{_popup_escape(dist_txt)}<br>"
            f"自動檢核：{_popup_escape(level)}<br>"
            f"中央審核：{_popup_escape(decision or '尚未審核')}<br><br>"
            f"<b>{'已選取；再點一次取消' if is_selected else '點一下加入中央審核'}</b>"
        )

        label = _county_return_marker_label_v3650(r)
        folium.CircleMarker(
            [lat, lon],
            radius=radius,
            color=outline,
            weight=weight,
            fill=True,
            fill_color=fill,
            fill_opacity=0.98,
            tooltip=label,
            popup=folium.Popup(popup, max_width=480),
        ).add_to(m)
        point_rows.append((r, lat, lon, label))

        if is_valid_wgs84(old_lon, old_lat):
            folium.CircleMarker(
                [old_lat, old_lon],
                radius=5,
                color="#FFFFFF",
                weight=2,
                fill=True,
                fill_color="#546E7A",
                fill_opacity=0.9,
                tooltip=f"目前GIS｜{name}",
            ).add_to(m)
            folium.PolyLine(
                [[old_lat, old_lon], [lat, lon]],
                color="#FB8C00",
                weight=3,
                opacity=0.75,
                dash_array="7,6",
                tooltip=f"目前GIS → 縣市新點：{dist_txt}",
            ).add_to(m)

        if is_selected:
            selected_pids.add(pid)

    # 只有已選案件才疊加其工程人工線/面，避免全縣市地圖太重。
    for f in geo.get("features", []):
        prop = f.get("properties") or {}
        pid = display_text(prop.get("project_id"))
        if pid not in selected_pids:
            continue
        if not prop.get("feature_active", True) or not prop.get("project_active", True):
            continue
        if prop.get("spatial_suppressed"):
            continue
        geom = f.get("geometry") or {}
        if geom.get("type") not in {"LineString", "MultiLineString", "Polygon", "MultiPolygon"}:
            continue
        try:
            folium.GeoJson(
                f,
                name="已選工程既有圖資",
                style_function=lambda _x: {
                    "color": "#3949AB",
                    "weight": 5,
                    "opacity": 0.75,
                    "fillOpacity": 0.08,
                },
            ).add_to(m)
        except Exception:
            pass

    try:
        if len(pts) > 1:
            m.fit_bounds(pts, padding=(25, 25), max_zoom=16)
        else:
            m.location = pts[0]
            m.options["zoom"] = 16
    except Exception:
        pass

    CountyReturnViewportMemory(
        memory_key or f"{county_filter}::{scope}"
    ).add_to(m)

    legend = """
    <div style="
      position:fixed; left:14px; bottom:28px; z-index:9999;
      background:#fff; color:#111; padding:9px 11px; border:1px solid #888;
      border-radius:7px; font-size:12px; line-height:1.65;
      box-shadow:0 1px 5px rgba(0,0,0,.22);">
      <b>縣市回填中央審核</b><br>
      <span style="color:#FB8C00">●</span> 尚未審核　
      <span style="color:#EF6C00">●</span> 有警告<br>
      <span style="color:#2E7D32">●</span> 已核准　
      <span style="color:#C62828">●</span> 退回再修正<br>
      <span style="color:#F9A825">●</span> 待確認　
      <span style="color:#1565C0">●</span> 採中央GIS<br>
      <span style="color:#546E7A">●</span> 目前GIS位置
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend))
    return m, point_rows


def _match_county_return_clicked_v3650(
    state: Optional[Dict[str, Any]],
    points: Sequence[Tuple[Dict[str, Any], float, float, str]],
) -> Optional[Dict[str, Any]]:
    if not state:
        return None
    tooltip = display_text(state.get("last_object_clicked_tooltip"))
    if tooltip:
        for rec, _lat, _lon, label in points:
            if tooltip == label:
                return rec

    obj = state.get("last_object_clicked")
    if not isinstance(obj, dict):
        return None
    try:
        lat = float(obj.get("lat"))
        lon = float(obj.get("lng"))
    except Exception:
        return None
    best = None
    best_d = 999.0
    for rec, plat, plon, _label in points:
        d = (plat - lat) ** 2 + (plon - lon) ** 2
        if d < best_d:
            best_d = d
            best = rec
    if best is not None and best_d <= 0.0006 ** 2:
        return best
    return None


def _apply_county_return_save_to_session_v3650(
    result: Dict[str, Any],
) -> None:
    snap = st.session_state.get("_gis_v366_snapshot")
    if isinstance(snap, dict):
        if result.get("_new_geo") is not None:
            snap["geo"] = result["_new_geo"]
        if result.get("_new_registry") is not None:
            snap["registry"] = result["_new_registry"]
        st.session_state["_gis_v366_snapshot"] = snap

    fn = globals().get("_cached_query_geo_snapshot_v351")
    if fn is not None and hasattr(fn, "clear"):
        try:
            fn.clear()
        except Exception:
            pass



def _raw_active_features_for_project_v3651(
    geo: Dict[str, Any],
    project_id: str,
) -> List[Dict[str, Any]]:
    out = []
    pid = display_text(project_id)
    for f in geo.get("features", []):
        props = f.get("properties") or {}
        if display_text(props.get("project_id")) != pid:
            continue
        if not props.get("feature_active", True) or not props.get("project_active", True):
            continue
        if display_text(props.get("role")) == "original_point":
            continue
        if geometry_type(f) in {"Point", "LineString", "Polygon"}:
            out.append(f)
    return out


def _build_duplicate_candidate_map_v3651(
    candidate: Dict[str, Any],
    geo: Dict[str, Any],
    primary_project_id: str,
) -> Optional[folium.Map]:
    members = list(candidate.get("members") or [])
    if not members:
        return None

    palette = [
        "#1565C0", "#C62828", "#2E7D32", "#7B1FA2",
        "#EF6C00", "#00838F", "#AD1457", "#5D4037",
    ]
    pts = []
    member_pts = {}
    for p in members:
        pt = _raw_project_point_v3651(p, geo)
        if pt:
            member_pts[p.project_id] = pt
            pts.append([pt[0], pt[1]])

    if pts:
        center = [
            sum(x[0] for x in pts) / len(pts),
            sum(x[1] for x in pts) / len(pts),
        ]
        zoom = 14 if len(pts) <= 2 else 12
    else:
        center, zoom = TAIWAN_CENTER, TAIWAN_ZOOM

    m = folium.Map(
        location=center,
        zoom_start=zoom,
        tiles="OpenStreetMap",
        control_scale=True,
    )
    _add_leaflet_compat_css(m)

    primary_pt = member_pts.get(primary_project_id)
    for idx, p in enumerate(members):
        color = palette[idx % len(palette)]
        pt = member_pts.get(p.project_id)

        # 畫每筆 PRJ 自己原本的人工線/面，協助判斷是否其實是不同工區。
        for f in _raw_active_features_for_project_v3651(geo, p.project_id):
            _add_feature_shape(
                m, f, color,
                popup_html=(
                    f"<b>{_popup_escape(p.project_name)}</b><br>"
                    f"{_popup_escape(p.sheet_name)}｜{_popup_escape(p.project_id)}"
                ),
                tooltip=f"{p.sheet_name}｜{p.project_id}",
                opacity=.82,
                reference=False,
            )

        if not pt:
            continue
        lat, lon = pt
        is_primary = p.project_id == primary_project_id
        folium.CircleMarker(
            [lat, lon],
            radius=10 if is_primary else 8,
            color="#FFFFFF" if is_primary else "#111111",
            weight=4 if is_primary else 2,
            fill=True,
            fill_color=color,
            fill_opacity=1.0,
            tooltip=f"{'★ 圖資主工程｜' if is_primary else ''}{p.sheet_name}｜{p.project_id}",
            popup=folium.Popup(
                f"<b>{_popup_escape(p.project_name)}</b><br>"
                f"分頁：{_popup_escape(p.sheet_name)}<br>"
                f"PRJ-ID：{_popup_escape(p.project_id)}<br>"
                f"座標：{lon:.7f}, {lat:.7f}",
                max_width=430,
            ),
        ).add_to(m)

        if primary_pt and not is_primary:
            d = _haversine_m(primary_pt[0], primary_pt[1], lat, lon)
            folium.PolyLine(
                [[primary_pt[0], primary_pt[1]], [lat, lon]],
                color=color,
                weight=3,
                opacity=.75,
                dash_array="7,6",
                tooltip=f"與圖資主工程距離：約 {d:,.0f}m",
            ).add_to(m)

    if len(pts) > 1:
        try:
            m.fit_bounds(pts, padding=(35, 35), max_zoom=17)
        except Exception:
            pass
    elif len(pts) == 1:
        m.location = pts[0]
        m.options["zoom"] = 16

    banner = f"""
    <div style="
      position:fixed; top:12px; left:50%; transform:translateX(-50%);
      z-index:9999; background:rgba(255,255,255,.96); color:#111;
      border:2px solid #374151; border-radius:8px; padding:7px 12px;
      font-size:13px; font-weight:700; box-shadow:0 2px 7px rgba(0,0,0,.2);">
      重複工程人工確認｜{_popup_escape(candidate.get('project_name'))}
    </div>
    """
    m.get_root().html.add_child(folium.Element(banner))
    return m


_DUPLICATE_DECISIONS_V3657 = [
    "同一實體工程多計畫經費－共用圖資",
    "同一工程座標有誤－採圖資主工程位置",
    "同一整體工程但多位置－只建立關聯、不共用代表點",
    "同名但不同工程－不合併",
]


def _duplicate_member_label_v3657(p: "ProjectRow") -> str:
    return f"{p.sheet_name}｜{p.project_id}｜{p.project_name}"


def _duplicate_member_payload_v3657(
    members: Sequence["ProjectRow"],
    geo: Dict[str, Any],
) -> List[Dict[str, Any]]:
    payload = []
    for p in members:
        pt = _raw_project_point_v3651(p, geo)
        payload.append({
            "project_id": p.project_id,
            "project_name": p.project_name,
            "sheet_name": p.sheet_name,
            "county": p.county,
            "lon": None if pt is None else pt[1],
            "lat": None if pt is None else pt[0],
        })
    return payload


def _duplicate_decision_flags_v3657(decision: str) -> Tuple[bool, bool, bool]:
    is_shared = decision in {
        "同一實體工程多計畫經費－共用圖資",
        "同一工程座標有誤－採圖資主工程位置",
    }
    is_related_only = (
        decision == "同一整體工程但多位置－只建立關聯、不共用代表點"
    )
    is_distinct = decision == "同名但不同工程－不合併"
    return is_shared, is_related_only, is_distinct


def _save_duplicate_batch_v3657(
    svc: "EngineeringGISService",
    jobs: Sequence[Dict[str, Any]],
    editor_name: str,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, str]]]:
    """由同一次按鈕操作逐組儲存；每組仍建立自己的 SPJ，不會跨組誤併。"""
    saved: List[Dict[str, Any]] = []
    errors: List[Dict[str, str]] = []
    for job in jobs:
        label = display_text(job.get("label")) or "未命名候選"
        decision = display_text(job.get("decision"))
        members = list(job.get("members") or [])
        project_ids = [
            display_text(x.get("project_id"))
            for x in members if isinstance(x, dict) and display_text(x.get("project_id"))
        ]
        is_shared, _is_related_only, is_distinct = (
            _duplicate_decision_flags_v3657(decision)
        )
        try:
            if is_distinct:
                result = svc.confirm_duplicate_projects_distinct_v3651(
                    project_ids,
                    display_text(job.get("note")),
                    editor_name,
                )
            else:
                result = svc.save_spatial_project_group_v3651(
                    members,
                    display_text(job.get("primary_project_id")),
                    decision,
                    share_geometry=is_shared,
                    note=display_text(job.get("note")),
                    editor=editor_name,
                )
            saved.append({"label": label, "result": result})
        except Exception as exc:
            errors.append({"label": label, "error": str(exc)})
    return saved, errors


def _apply_spatial_group_result_v3651(result: Dict[str, Any]) -> None:
    snap = st.session_state.get("_gis_v366_snapshot")
    if isinstance(snap, dict):
        if result.get("_new_geo") is not None:
            snap["geo"] = result["_new_geo"]
        if result.get("_new_registry") is not None:
            snap["registry"] = result["_new_registry"]
        st.session_state["_gis_v366_snapshot"] = snap
    fn = globals().get("_cached_query_geo_snapshot_v351")
    if fn is not None and hasattr(fn, "clear"):
        try:
            fn.clear()
        except Exception:
            pass


def _render_duplicate_spatial_project_tab_v3651(
    svc: "EngineeringGISService",
    rows: Sequence["ProjectRow"],
    geo: Dict[str, Any],
    registry: Dict[str, Any],
    editor_name: str,
) -> None:
    st.subheader("🔗 重複工程與共用圖資整理")
    st.caption(
        "PRJ-ID 仍代表每一筆管控資料；SPJ-ID 代表實際空間工程。"
        "系統只提出疑似重複候選，不會因工程名稱相同就自動合併。"
    )

    groups = (registry or {}).get("spatial_project_groups") or {}
    active_groups = {
        k: v for k, v in groups.items()
        if isinstance(v, dict) and v.get("active", True)
    } if isinstance(groups, dict) else {}

    candidates = _duplicate_project_candidates_v3651(rows, geo, registry)
    high_n = sum(1 for x in candidates if x.get("risk") == "高度疑似")
    cross_n = sum(1 for x in candidates if x.get("cross_plan"))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("疑似重複群組", f"{len(candidates):,}")
    c2.metric("高度疑似", f"{high_n:,}")
    c3.metric("跨分頁／跨計畫候選", f"{cross_n:,}")
    c4.metric("已建立SPJ", f"{len(active_groups):,}")

    st.info(
        "建議先處理『治理工程＋前瞻治理工程』同名案件。"
        "若確認是同一實體工程，可共用一組 GIS；經費、核定屬性、進度等 Excel 管控資料不會合併。"
    )

    batch_notice = st.session_state.pop("gis_v3657_dup_batch_notice", None)
    if isinstance(batch_notice, dict):
        saved_n = int(batch_notice.get("saved", 0) or 0)
        failed = list(batch_notice.get("errors") or [])
        if saved_n:
            st.success(f"批次整理完成：成功 {saved_n:,} 組。")
        for item in failed:
            st.error(
                f"批次整理失敗｜{display_text(item.get('label'))}："
                f"{display_text(item.get('error'))}"
            )

    if candidates:
        st.markdown("#### 1. 系統提出的疑似重複候選")
        summary_rows = []
        for x in candidates:
            d = x.get("max_distance_m")
            summary_rows.append({
                "縣市": x.get("county"),
                "工程名稱": x.get("project_name"),
                "管控筆數": x.get("member_count"),
                "來源分頁": "、".join(x.get("sheets") or []),
                "最大點位差距(m)": None if d is None else round(float(d), 1),
                "系統判斷": x.get("risk"),
                "建議": x.get("suggestion"),
            })
        st.dataframe(summary_rows, hide_index=True, use_container_width=True, height=330)

        county_opts = ["全部縣市"] + sorted({
            display_text(x.get("county"))
            for x in candidates if display_text(x.get("county"))
        })
        county_filter = st.selectbox(
            "候選縣市",
            county_opts,
            key="gis_v3651_dup_county",
        )
        filtered = [
            x for x in candidates
            if county_filter == "全部縣市"
            or display_text(x.get("county")) == county_filter
        ]

        option_map = {}
        for x in filtered:
            dist = x.get("max_distance_m")
            dist_txt = "缺座標" if dist is None else f"{float(dist):,.0f}m"
            label = (
                f"{x.get('county')}｜{x.get('project_name')}｜{x.get('member_count')}筆｜"
                f"{'、'.join(x.get('sheets') or [])}｜最大差距 {dist_txt}"
            )
            option_map[label] = x

        st.markdown("#### 2. 多組、多工程批次整理")
        st.caption(
            "每個候選可勾選 2 筆以上工程，也可同時選取多個候選群組。"
            "最後只按一次儲存；各候選仍各自建立 SPJ，不會跨工程名稱合併。"
        )

        batch_key = "gis_v3657_dup_batch_candidates"
        available_labels = list(option_map.keys())
        old_batch_labels = [
            x for x in st.session_state.get(batch_key, [])
            if x in option_map
        ]
        if st.session_state.get(batch_key) != old_batch_labels:
            st.session_state[batch_key] = old_batch_labels

        quick1, quick2 = st.columns(2)
        with quick1:
            if st.button(
                "勾選點位完整的高度疑似候選",
                use_container_width=True,
                key="gis_v3657_dup_batch_select_high",
            ):
                st.session_state[batch_key] = [
                    label for label, item in option_map.items()
                    if item.get("max_distance_m") is not None
                    and float(item.get("max_distance_m")) <= 30
                    and all(
                        _raw_project_point_v3651(p, geo) is not None
                        for p in (item.get("members") or [])
                    )
                ][:20]
        with quick2:
            if st.button(
                "清除批次勾選",
                use_container_width=True,
                key="gis_v3657_dup_batch_clear",
            ):
                st.session_state[batch_key] = []

        batch_labels = st.multiselect(
            "勾選要一起整理的候選群組（一次最多 20 組）",
            available_labels,
            key=batch_key,
        )

        batch_jobs: List[Dict[str, Any]] = []
        batch_invalid: List[str] = []
        if len(batch_labels) > 20:
            batch_invalid.append("一次最多處理 20 組，請取消部分候選。")

        if batch_labels:
            batch_note = st.text_input(
                "本次批次整理共同備註",
                placeholder="例如：同名且點位相同，確認為跨計畫同一實體工程。",
                key="gis_v3657_dup_batch_note",
            )

            for batch_label in batch_labels:
                item = option_map[batch_label]
                candidate_id = display_text(item.get("candidate_id"))
                all_members = list(item.get("members") or [])
                all_member_options = {
                    _duplicate_member_label_v3657(p): p for p in all_members
                }
                member_key = f"gis_v3657_dup_batch_members_{candidate_id}"
                if member_key not in st.session_state:
                    st.session_state[member_key] = list(all_member_options.keys())
                else:
                    st.session_state[member_key] = [
                        x for x in st.session_state.get(member_key, [])
                        if x in all_member_options
                    ]

                with st.expander(
                    f"批次設定｜{batch_label}",
                    expanded=len(batch_labels) <= 3,
                ):
                    selected_member_labels = st.multiselect(
                        "本組要納入整理的工程（至少 2 筆）",
                        list(all_member_options.keys()),
                        key=member_key,
                    )
                    selected_members = [
                        all_member_options[x] for x in selected_member_labels
                    ]
                    if len(selected_members) < 2:
                        msg = f"{batch_label}：至少要勾選 2 筆工程。"
                        batch_invalid.append(msg)
                        st.error("請至少勾選 2 筆工程。")

                    selected_spjs = []
                    for p in selected_members:
                        spj0, _group0 = _active_spatial_group_for_project_v3651(
                            registry, p.project_id
                        )
                        if spj0 and spj0 not in selected_spjs:
                            selected_spjs.append(spj0)

                    existing_group = None
                    locked_primary = ""
                    if len(selected_spjs) > 1:
                        msg = f"{batch_label}：選取工程分屬多個既有 SPJ。"
                        batch_invalid.append(msg)
                        st.error(
                            "選取工程分屬多個既有 SPJ，請先解除或整理既有群組。"
                        )
                    elif len(selected_spjs) == 1:
                        existing_group = active_groups.get(selected_spjs[0])
                        if isinstance(existing_group, dict):
                            locked_primary = display_text(
                                existing_group.get("primary_project_id")
                            )
                            st.info(
                                f"將延伸既有 {selected_spjs[0]}；"
                                f"圖資主工程固定為 {locked_primary}。"
                            )

                    if isinstance(existing_group, dict):
                        if _spatial_group_is_shared_v3651(existing_group):
                            decision_options = _DUPLICATE_DECISIONS_V3657[:2]
                        else:
                            decision_options = [_DUPLICATE_DECISIONS_V3657[2]]
                    else:
                        decision_options = list(_DUPLICATE_DECISIONS_V3657)

                    decision_key = f"gis_v3657_dup_batch_decision_{candidate_id}"
                    if st.session_state.get(decision_key) not in decision_options:
                        st.session_state[decision_key] = decision_options[0]
                    batch_decision = st.selectbox(
                        "本組處理方式",
                        decision_options,
                        key=decision_key,
                    )
                    is_shared, _is_related_only, is_distinct = (
                        _duplicate_decision_flags_v3657(batch_decision)
                    )

                    primary_project_id = ""
                    if not is_distinct and selected_members:
                        selected_options = {
                            _duplicate_member_label_v3657(p): p
                            for p in selected_members
                        }
                        if locked_primary:
                            primary_project_id = locked_primary
                            if locked_primary not in {
                                p.project_id for p in selected_members
                            }:
                                msg = (
                                    f"{batch_label}：必須同時勾選既有圖資主工程 "
                                    f"{locked_primary}。"
                                )
                                batch_invalid.append(msg)
                                st.error(
                                    f"請在本組工程中勾選既有圖資主工程 {locked_primary}。"
                                )
                            else:
                                st.caption(f"圖資主工程：{locked_primary}（既有群組已鎖定）")
                        else:
                            primary_key = (
                                f"gis_v3657_dup_batch_primary_{candidate_id}"
                            )
                            primary_labels = list(selected_options.keys())
                            if st.session_state.get(primary_key) not in primary_labels:
                                st.session_state[primary_key] = primary_labels[0]
                            primary_label = st.selectbox(
                                "本組圖資主工程",
                                primary_labels,
                                key=primary_key,
                            )
                            primary_project_id = selected_options[
                                primary_label
                            ].project_id

                        if is_shared:
                            primary_obj = next(
                                (
                                    p for p in selected_members
                                    if p.project_id == primary_project_id
                                ),
                                None,
                            )
                            if (
                                primary_obj is None
                                or _raw_project_point_v3651(primary_obj, geo) is None
                            ):
                                msg = f"{batch_label}：圖資主工程沒有有效代表點。"
                                batch_invalid.append(msg)
                                st.error(
                                    "共用圖資的主工程必須有有效代表點，請改選其他主工程。"
                                )

                    dist = item.get("max_distance_m")
                    if is_shared and dist is not None and float(dist) > 30:
                        st.warning(
                            f"本組最大點位差距約 {float(dist):,.0f} 公尺，"
                            "請先用下方單組地圖確認後再批次儲存。"
                        )

                    if len(selected_members) >= 2:
                        batch_jobs.append({
                            "label": batch_label,
                            "decision": batch_decision,
                            "primary_project_id": primary_project_id,
                            "members": _duplicate_member_payload_v3657(
                                selected_members, geo
                            ),
                            "note": batch_note,
                        })

            confirm_payload = [
                {
                    "label": job.get("label"),
                    "decision": job.get("decision"),
                    "primary_project_id": job.get("primary_project_id"),
                    "project_ids": [
                        x.get("project_id") for x in (job.get("members") or [])
                    ],
                }
                for job in batch_jobs
            ]
            confirm_sig = hashlib.sha1(
                json.dumps(
                    confirm_payload,
                    ensure_ascii=False,
                    sort_keys=True,
                ).encode("utf-8")
            ).hexdigest()[:12]
            batch_confirm = st.checkbox(
                f"我已逐組確認以上 {len(batch_labels)} 組工程的成員、處理方式及圖資主工程",
                key=f"gis_v3657_dup_batch_confirm_{confirm_sig}",
            )
            for msg in dict.fromkeys(batch_invalid):
                st.warning(msg)

            batch_disabled = (
                not batch_labels
                or len(batch_labels) > 20
                or bool(batch_invalid)
                or len(batch_jobs) != len(batch_labels)
                or not batch_confirm
            )
            if st.button(
                f"💾 一次儲存 {len(batch_labels)} 組重複工程整理結果",
                type="primary",
                use_container_width=True,
                disabled=batch_disabled,
                key="gis_v3657_save_duplicate_batch",
            ):
                _mask = _saving_overlay(
                    f"正在批次儲存 {len(batch_jobs)} 組重複工程整理結果…"
                )
                try:
                    saved_items, error_items = _save_duplicate_batch_v3657(
                        svc, batch_jobs, editor_name
                    )
                finally:
                    _close_saving_overlay(_mask)

                for saved_item in saved_items:
                    result = saved_item.get("result") or {}
                    _apply_spatial_group_result_v3651(result)
                st.session_state["gis_v3657_dup_batch_notice"] = {
                    "saved": len(saved_items),
                    "errors": error_items,
                }
                st.rerun()
        else:
            st.info("請勾選要一起整理的候選群組；也可使用上方按鈕快速勾選高度疑似案件。")

        st.markdown("#### 3. 單組詳細確認")
        selected_label = st.selectbox(
            "選擇一組候選進行人工確認",
            ["請選擇候選"] + list(option_map.keys()),
            key="gis_v3651_dup_candidate",
        )
        candidate = option_map.get(selected_label)

        if candidate:
            members = list(candidate.get("members") or [])
            member_rows = []
            for p in members:
                pt = _raw_project_point_v3651(p, geo)
                spj0, group0 = _active_spatial_group_for_project_v3651(
                    registry, p.project_id
                )
                member_rows.append({
                    "PRJ-ID": p.project_id,
                    "分頁": p.sheet_name,
                    "工程名稱": p.project_name,
                    "縣市": p.county,
                    "執行情形": p.status,
                    "經度": None if not pt else round(float(pt[1]), 7),
                    "緯度": None if not pt else round(float(pt[0]), 7),
                    "既有SPJ": spj0,
                })
            st.dataframe(member_rows, hide_index=True, use_container_width=True)

            existing_spjs = candidate.get("existing_spj_ids") or []
            existing_group = None
            locked_primary = ""
            if len(existing_spjs) == 1:
                existing_group = active_groups.get(existing_spjs[0])
                if isinstance(existing_group, dict):
                    locked_primary = display_text(existing_group.get("primary_project_id"))
                    st.info(
                        f"這組候選已有部分工程屬於 {existing_spjs[0]}；"
                        "若要加入新成員，會延伸既有群組，圖資主工程不可更換。"
                    )
            elif len(existing_spjs) > 1:
                st.error(
                    "這組候選的成員目前分屬多個 SPJ，不能直接合併。"
                    "請先在下方『既有 SPJ 群組』檢查並整理。"
                )

            member_options = {
                f"{p.sheet_name}｜{p.project_id}｜{p.project_name}": p
                for p in members
            }
            default_primary_label = next(
                (
                    label for label, p in member_options.items()
                    if p.project_id == locked_primary
                ),
                list(member_options.keys())[0],
            )
            if locked_primary:
                primary_label = default_primary_label
                st.text_input(
                    "圖資主工程（既有SPJ已鎖定）",
                    value=primary_label,
                    disabled=True,
                    key="gis_v3651_locked_primary",
                )
            else:
                primary_label = st.radio(
                    "選擇『圖資主工程』",
                    list(member_options.keys()),
                    index=list(member_options.keys()).index(default_primary_label),
                    key="gis_v3651_dup_primary",
                )
            primary_p = member_options[primary_label]

            cmap = _build_duplicate_candidate_map_v3651(
                candidate, geo, primary_p.project_id
            )
            if cmap is not None:
                st_folium(
                    cmap,
                    width=None,
                    height=590,
                    returned_objects=[],
                    key=(
                        f"gis_v3651_dup_map_{candidate.get('candidate_id')}_"
                        f"{primary_p.project_id}"
                    ),
                )
            else:
                st.warning("這組候選目前沒有任何有效點位，請依工程資料人工判斷。")

            st.markdown("##### 單組人工判定")
            decision = st.radio(
                "這組同名工程應如何處理？",
                _DUPLICATE_DECISIONS_V3657,
                key="gis_v3651_dup_decision",
            )
            note = st.text_area(
                "整理備註",
                placeholder=(
                    "例如：治理工程及前瞻治理工程為同一標案，經費分屬兩計畫；"
                    "或兩筆為不同工區，確認不是重複工程。"
                ),
                key="gis_v3651_dup_note",
            )

            is_shared, is_related_only, is_distinct = (
                _duplicate_decision_flags_v3657(decision)
            )

            if is_shared:
                st.warning(
                    f"確認後會建立 SPJ，共 {len(members)} 筆 PRJ 共用「{primary_p.project_id}」圖資。"
                    "其他 PRJ 的人工線/面不會刪除，但會暫時隱藏；"
                    "代表點統一後，Excel 仍由系統管理者日後批次同步。"
                )
                confirm = st.checkbox(
                    "我已確認這些 PRJ 是同一實體工程，並同意使用上述圖資主工程",
                    key="gis_v3651_dup_confirm_shared",
                )
            elif is_related_only:
                st.info(
                    "只建立 SPJ 關聯，不會隱藏任何圖資，也不會改變任何代表點。"
                    "適合同一整體工程但確實有多個施工位置。"
                )
                confirm = True
            elif is_distinct:
                st.info(
                    "系統會記錄這組 PRJ 已人工確認為不同工程；"
                    "只要成員組合不變，之後不再列入重複候選。"
                )
                confirm = True
            else:
                confirm = False

            disabled = len(existing_spjs) > 1 or not confirm
            if st.button(
                "💾 儲存重複工程整理結果",
                type="primary",
                use_container_width=True,
                disabled=disabled,
                key="gis_v3651_save_duplicate_review",
            ):
                _mask = _saving_overlay("正在儲存重複工程與共用圖資整理結果…")
                try:
                    if is_distinct:
                        result = svc.confirm_duplicate_projects_distinct_v3651(
                            [p.project_id for p in members],
                            note,
                            editor_name,
                        )
                    else:
                        payload = _duplicate_member_payload_v3657(members, geo)
                        result = svc.save_spatial_project_group_v3651(
                            payload,
                            primary_p.project_id,
                            decision,
                            share_geometry=is_shared,
                            note=note,
                            editor=editor_name,
                        )
                finally:
                    _close_saving_overlay(_mask)

                _apply_spatial_group_result_v3651(result)
                if is_distinct:
                    st.session_state["gis_v3651_dup_notice"] = (
                        "已記錄為『同名但不同工程』，後續不再列為相同成員的重複候選。"
                    )
                else:
                    st.session_state["gis_v3651_dup_notice"] = (
                        f"已建立/更新 {result.get('spatial_project_id')}；"
                        f"圖資主工程 {result.get('primary_project_id')}。"
                    )
                st.rerun()
    else:
        st.success("目前沒有尚待人工確認的同名重複工程候選。")

    notice = st.session_state.pop("gis_v3651_dup_notice", "")
    if notice:
        st.success(notice)

    st.markdown("#### 4. 既有 SPJ 空間工程群組")
    active_groups = {
        k: v for k, v in ((registry or {}).get("spatial_project_groups") or {}).items()
        if isinstance(v, dict) and v.get("active", True)
    }
    if not active_groups:
        st.caption("目前尚未建立任何 SPJ 群組。")
        return

    group_rows = []
    for spj_id, g in sorted(active_groups.items()):
        linked = [display_text(x) for x in g.get("linked_project_ids") or [] if display_text(x)]
        info = g.get("member_info") or {}
        names = sorted({
            display_text((info.get(pid) or {}).get("project_name"))
            for pid in linked
            if display_text((info.get(pid) or {}).get("project_name"))
        })
        group_rows.append({
            "SPJ-ID": spj_id,
            "圖資主工程": g.get("primary_project_id"),
            "關聯PRJ筆數": len(linked),
            "工程名稱": "、".join(names),
            "關係": g.get("relation_type"),
            "圖資模式": "共用圖資" if g.get("share_mode") == "shared" else "只建立關聯",
            "最後更新者": g.get("updated_by") or g.get("created_by"),
            "最後更新時間": g.get("updated_at") or g.get("created_at"),
        })
    st.dataframe(group_rows, hide_index=True, use_container_width=True)

    with st.expander("⚠️ 解除既有 SPJ 群組", expanded=False):
        st.warning(
            "解除群組不會自動把座標還原成建立群組前的位置，"
            "避免誤把後續合法校正倒退。解除後各 PRJ 保留當下代表點；"
            "若要分開位置，再到工程圖資編輯逐筆校正。"
        )
        spj_choice = st.selectbox(
            "選擇要解除的 SPJ",
            list(sorted(active_groups.keys())),
            key="gis_v3651_dissolve_spj",
        )
        g = active_groups[spj_choice]
        st.caption(
            f"主工程：{g.get('primary_project_id')}｜"
            f"關聯：{'、'.join(g.get('linked_project_ids') or [])}"
        )
        confirm_dissolve = st.checkbox(
            "我確認要解除這個 SPJ 共用/關聯關係",
            key="gis_v3651_confirm_dissolve",
        )
        if st.button(
            "解除 SPJ 群組",
            disabled=not confirm_dissolve,
            use_container_width=True,
            key="gis_v3651_dissolve_btn",
        ):
            _mask = _saving_overlay(f"正在解除 {spj_choice}…")
            try:
                result = svc.dissolve_spatial_project_group_v3651(
                    spj_choice, editor_name
                )
            finally:
                _close_saving_overlay(_mask)
            _apply_spatial_group_result_v3651(result)
            st.session_state["gis_v3651_dup_notice"] = f"已解除 {spj_choice}。"
            st.rerun()



def _render_county_coordinate_return_review_tab(
    svc: "EngineeringGISService",
    rows: Sequence["ProjectRow"],
    geo: Dict[str, Any],
    registry: Dict[str, Any],
    editor_name: str,
) -> None:
    st.subheader("📥 縣市座標回填審核")
    st.caption(
        "縣市回傳 Excel 後，先由中央在此確認。"
        "只有『核准縣府新座標』會更新 GIS 代表點；current.xlsx 不會在這裡修改。"
        "最後仍由系統管理者使用既有『GIS代表點同步至 current.xlsx』功能一次同步。"
    )

    uploaded = st.file_uploader(
        "上傳縣市回傳的座標複核 Excel（可一次多檔）",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        key="gis_v3650_county_return_upload",
    )

    if not uploaded:
        st.info(
            "請上傳縣市填回的複核表。系統會以『系統工程ID』對應，"
            "不使用工程名稱作為主鍵。"
        )
        return

    payloads = [(f.name, f.getvalue()) for f in uploaded]
    upload_sig = hashlib.sha1(
        b"||".join(
            [name.encode("utf-8") + b":" + hashlib.sha1(data).digest()
             for name, data in payloads]
        )
    ).hexdigest()[:16]

    parse_cache = st.session_state.setdefault(
        "gis_v3650_county_return_parse_cache", {}
    )
    parsed = parse_cache.get(upload_sig)

    if st.button(
        "🔎 解析回填檔並建立中央審核清單",
        type="primary",
        use_container_width=True,
        key="gis_v3650_parse_return_files",
    ):
        _mask = _saving_overlay("正在解析縣市回填座標…")
        try:
            records, messages = _parse_county_return_workbooks_v3650(
                payloads, rows, geo, registry
            )
            parsed = {
                "records": records,
                "messages": messages,
            }
            parse_cache[upload_sig] = parsed
        finally:
            _close_saving_overlay(_mask)

    if parsed is None:
        st.info("檔案已選取，請按「解析回填檔並建立中央審核清單」。")
        return

    records = list(parsed.get("records") or [])
    messages = list(parsed.get("messages") or [])
    for msg in messages:
        st.warning(msg)

    if not records:
        st.error("目前上傳檔案沒有可辨識的縣市回填資料。")
        return

    total = len(records)
    reviewable = sum(1 for r in records if r.get("reviewable"))
    warning_n = sum(1 for r in records if r.get("validation_status") == "警告")
    error_n = sum(1 for r in records if r.get("validation_status") == "錯誤")
    approved_prev = sum(
        1 for r in records if r.get("previous_decision") == "核准縣府新座標"
    )
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("回填筆數", f"{total:,}")
    c2.metric("可進入地圖審核", f"{reviewable:,}")
    c3.metric("警告／錯誤", f"{warning_n + error_n:,}")
    c4.metric("本次座標已曾核准", f"{approved_prev:,}")

    st.markdown("#### 1. 自動檢核結果")
    st.dataframe(
        _county_return_display_rows_v3650(records),
        use_container_width=True,
        hide_index=True,
        height=380,
    )
    if error_n:
        st.warning(
            f"有 {error_n:,} 筆不可直接核准。常見原因包含：ID不存在、重複ID、"
            "CRS未填、X/Y不完整、座標疑似填反或轉換後不在合理範圍。"
        )

    st.markdown("#### 2. 地圖確認縣市新點位")
    counties = sorted({
        display_text(r.get("county"))
        for r in records if display_text(r.get("county"))
    })
    county_filter = st.selectbox(
        "地圖縣市",
        ["全部縣市"] + counties,
        key="gis_v3650_return_county_filter",
    )
    scope = st.radio(
        "地圖顯示",
        ["只看尚未中央審核", "只看有警告/錯誤", "全部可定位案件"],
        horizontal=True,
        key="gis_v3650_return_scope",
    )

    selection_key = f"gis_v3650_return_selection::{upload_sig}"
    selected_keys = [
        display_text(x)
        for x in st.session_state.get(selection_key, [])
        if display_text(x)
    ]
    valid_keys = {display_text(r.get("record_key")) for r in records}
    selected_keys = [x for x in selected_keys if x in valid_keys]
    st.session_state[selection_key] = selected_keys

    map_epoch_key = (
        f"gis_v3656_return_map_epoch::{upload_sig}::{county_filter}::{scope}"
    )
    map_epoch = int(st.session_state.get(map_epoch_key, 0) or 0)
    return_map, points = _build_county_return_review_map_v3650(
        records,
        selected_keys,
        geo,
        county_filter,
        scope,
        memory_key=f"{upload_sig}::{county_filter}::{scope}",
    )
    if return_map is None:
        st.info("目前篩選沒有可在地圖審核的有效縣市新點。")
        map_state = None
    else:
        map_state = st_folium(
            return_map,
            width=None,
            height=620,
            returned_objects=[
                "last_object_clicked",
                "last_object_clicked_tooltip",
            ],
            key=(
                f"gis_v3656_return_map_{upload_sig}_"
                f"{county_filter}_{scope}_{map_epoch}"
            ),
        )

    clicked = _match_county_return_clicked_v3650(map_state, points)
    if clicked is not None:
        key = display_text(clicked.get("record_key"))
        updated = list(selected_keys)
        if key in updated:
            updated.remove(key)
            st.session_state["gis_v3650_return_notice"] = (
                f"已取消：{display_text(clicked.get('project_name'))}"
            )
        else:
            if len(updated) >= 50:
                st.session_state["gis_v3650_return_notice"] = (
                    "一次批次中央審核最多 50 件，請先儲存或取消部分選取。"
                )
            else:
                updated.append(key)
                st.session_state["gis_v3650_return_notice"] = (
                    f"已加入：{display_text(clicked.get('project_name'))}"
                )
        st.session_state[selection_key] = updated
        # 每次點選都給地圖新的元件 key，避免達到 50 件上限時重播同一個點擊。
        st.session_state[map_epoch_key] = map_epoch + 1
        st.rerun()

    notice = st.session_state.pop("gis_v3650_return_notice", "")
    if notice:
        st.info(notice)

    selected_set = set(st.session_state.get(selection_key, []))
    selected_records = [
        r for r in records
        if display_text(r.get("record_key")) in selected_set
    ]

    m1, m2 = st.columns([1, 2])
    m1.metric("目前已選", f"{len(selected_records)} 件")
    with m2:
        if st.button(
            "🧹 清除全部選取",
            disabled=not selected_records,
            use_container_width=True,
            key="gis_v3650_return_clear_selection",
        ):
            st.session_state[selection_key] = []
            st.rerun()

    if selected_records:
        st.dataframe(
            [
                {
                    "系統工程ID": r.get("project_id"),
                    "工程名稱": r.get("project_name"),
                    "縣市填報CRS": r.get("returned_crs"),
                    "新經度": round(float(r["candidate_lon"]), 7),
                    "新緯度": round(float(r["candidate_lat"]), 7),
                    "與目前GIS位移(m)": (
                        None
                        if r.get("distance_from_gis_m") is None
                        else round(float(r["distance_from_gis_m"]), 1)
                    ),
                    "自動檢核": r.get("validation_status"),
                    "既有中央審核": r.get("previous_decision") or "尚未審核",
                }
                for r in selected_records
            ],
            hide_index=True,
            use_container_width=True,
        )

        st.markdown("#### 3. 中央人工審核")
        decision = st.radio(
            "本次選取案件要如何處理",
            [
                "核准縣府新座標",
                "退回縣府再修正",
                "待確認",
                "採用中央既有GIS點位",
            ],
            horizontal=True,
            key="gis_v3650_return_decision",
        )
        review_note = st.text_area(
            "中央審核備註",
            placeholder=(
                "例如：縣府新點落於正確排水線旁，核准；"
                "或新點落在錯誤鄉鎮，退回縣府再確認。"
            ),
            key="gis_v3650_return_review_note",
        )

        selected_warning = [
            r for r in selected_records
            if r.get("validation_status") == "警告"
        ]
        high_shift = [
            r for r in selected_records
            if r.get("distance_from_gis_m") is not None
            and float(r["distance_from_gis_m"]) >= 2000
        ]
        approve_confirmed = True
        if decision == "核准縣府新座標" and (selected_warning or high_shift):
            st.warning(
                f"本批有 {len(selected_warning):,} 件自動檢核警告，"
                f"其中 {len(high_shift):,} 件與目前 GIS 位移達 2 公里以上。"
                "系統不會阻止人工核准，但必須再次確認。"
            )
            approve_confirmed = st.checkbox(
                "我已在地圖檢查上述警告案件，仍確認採用縣府新座標",
                key="gis_v3650_return_approve_warning_confirm",
            )

        if decision == "採用中央既有GIS點位":
            no_gis = [
                r for r in selected_records
                if not is_valid_wgs84(
                    parse_number(r.get("old_lon")),
                    parse_number(r.get("old_lat")),
                )
            ]
            if no_gis:
                st.warning(
                    f"選取中有 {len(no_gis):,} 件目前沒有有效中央 GIS 點位，"
                    "不能整批選擇『採用中央既有GIS點位』。"
                )

        disabled = False
        if decision == "核准縣府新座標" and not approve_confirmed:
            disabled = True
        if decision == "採用中央既有GIS點位" and any(
            not is_valid_wgs84(
                parse_number(r.get("old_lon")),
                parse_number(r.get("old_lat")),
            )
            for r in selected_records
        ):
            disabled = True

        st.info(
            "核准縣府新座標後，只會更新 GIS 代表點並標記『Excel待系統管理者同步』；"
            "此處不會直接修改 current.xlsx。"
        )

        if st.button(
            f"💾 儲存 {len(selected_records)} 件中央審核結果",
            type="primary",
            use_container_width=True,
            disabled=disabled,
            key="gis_v3650_save_return_reviews",
        ):
            payload = []
            for r in selected_records:
                payload.append({
                    "project_id": r.get("project_id"),
                    "row_key": r.get("row_key"),
                    "project_name": r.get("project_name"),
                    "county": r.get("county"),
                    "status": r.get("status"),
                    "source_file": r.get("source_file"),
                    "submission_hash": r.get("submission_hash"),
                    "returned_crs": r.get("returned_crs"),
                    "raw_x": r.get("raw_x"),
                    "raw_y": r.get("raw_y"),
                    "candidate_lon": r.get("candidate_lon"),
                    "candidate_lat": r.get("candidate_lat"),
                    "county_note": r.get("county_note"),
                    "validation_status": r.get("validation_status"),
                    "validation_message": r.get("validation_message"),
                    "distance_from_gis_m": r.get("distance_from_gis_m"),
                    "review_decision": decision,
                    "review_note": review_note,
                })

            _mask = _saving_overlay("正在儲存縣市座標回填中央審核結果…")
            try:
                result = svc.save_county_coordinate_return_reviews(
                    payload, editor_name
                )
            finally:
                _close_saving_overlay(_mask)

            _apply_county_return_save_to_session_v3650(result)
            st.session_state[selection_key] = []
            st.session_state.pop(
                "gis_v3650_county_return_parse_cache", None
            )
            st.session_state["gis_v3650_return_saved_notice"] = (
                f"已儲存 {result.get('count', 0):,} 件中央審核；"
                f"其中 {result.get('approved_count', 0):,} 件縣府新座標已更新 GIS。"
                "current.xlsx 尚未修改，等待系統管理者同步。"
            )
            st.rerun()
    else:
        st.info("請直接在上方地圖點選要中央確認的縣市新點位。")

    saved_notice = st.session_state.pop("gis_v3650_return_saved_notice", "")
    if saved_notice:
        st.success(saved_notice)

    st.markdown("#### 4. 審核結果與 Excel 同步狀態")
    latest = (registry or {}).get("coordinate_return_reviews") or {}
    if isinstance(latest, dict) and latest:
        latest_rows = []
        for _pid, rv in latest.items():
            if not isinstance(rv, dict):
                continue
            latest_rows.append({
                "系統工程ID": rv.get("project_id"),
                "縣市": rv.get("county"),
                "工程名稱": rv.get("project_name"),
                "中央審核": rv.get("review_decision"),
                "審核者": rv.get("reviewed_by"),
                "審核時間": rv.get("reviewed_at"),
                "GIS已套用": "是" if rv.get("gis_applied") else "否",
                "Excel同步": (
                    "待系統管理者同步"
                    if rv.get("excel_sync_status") == "pending"
                    else ""
                ),
            })
        if latest_rows:
            latest_rows.sort(
                key=lambda x: display_text(x.get("審核時間")),
                reverse=True,
            )
            st.dataframe(
                latest_rows[:500],
                hide_index=True,
                use_container_width=True,
                height=360,
            )

    st.caption(
        "系統管理者不需要在此頁操作 Excel。核准後的 GIS／Excel 差異會自動出現在"
        "「系統管理 → GIS代表點同步至 current.xlsx」既有清單中。"
    )



def _coord_audit_fast_raw_and_twd67_candidate(
    p: "ProjectRow",
) -> Tuple[str, Optional[float], Optional[float], Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    """V3.6.45：直接利用 scan_workbook 已解析的 ProjectRow，不重新開 current.xlsx。

    - WGS84/Google：目前 lon/lat 就是原始經緯度（若曾顛倒，ProjectRow 已記錄）。
      同時可把這組數字視為 TWD67 Geographic 測試候選。
    - TWD97/EPSG:3826 或 3825：先把目前 WGS84 反算回 TWD97 TM2，
      再把同一組 X/Y 當作 TWD67 TM2 解讀，得到 TWD67 候選。
    """
    if p.lon is None or p.lat is None:
        return display_text(p.coord_source) or "無有效座標", None, None, None, None

    src = display_text(p.coord_source)
    lon = float(p.lon)
    lat = float(p.lat)

    twd67_tm2 = None
    twd67_geo = None
    raw_x = lon
    raw_y = lat
    raw_source = src or "系統已解析座標"

    try:
        if "3826" in src:
            x, y = TRANSFORMER_TO_121.transform(lon, lat)
            raw_x, raw_y = float(x), float(y)
            raw_source = "TWD97 TM2 zone 121"
            lon67, lat67 = TRANSFORMER_67_121.transform(raw_x, raw_y)
            if is_valid_wgs84(lon67, lat67):
                twd67_tm2 = {
                    "lon": float(lon67), "lat": float(lat67),
                    "datum": "TWD67", "epsg": "3828", "swapped": bool(p.coord_swapped),
                }
        elif "3825" in src:
            x, y = TRANSFORMER_TO_119.transform(lon, lat)
            raw_x, raw_y = float(x), float(y)
            raw_source = "TWD97 TM2 zone 119"
            lon67, lat67 = TRANSFORMER_67_119.transform(raw_x, raw_y)
            if is_valid_wgs84(lon67, lat67):
                twd67_tm2 = {
                    "lon": float(lon67), "lat": float(lat67),
                    "datum": "TWD67", "epsg": "3827", "swapped": bool(p.coord_swapped),
                }
        elif "WGS84" in src or "Google" in src:
            raw_source = "WGS84/Google"
            # 若原始數字其實是 TWD67 Geographic，將同一組 lon/lat 轉成 WGS84 候選。
            lon67, lat67 = TRANSFORMER_67_GEO.transform(lon, lat)
            if is_valid_wgs84(lon67, lat67):
                twd67_geo = {
                    "lon": float(lon67), "lat": float(lat67),
                    "datum": "TWD67-GEO", "epsg": "3821", "swapped": bool(p.coord_swapped),
                }
    except Exception:
        pass

    return raw_source, raw_x, raw_y, twd67_tm2, twd67_geo


def _build_coordinate_audit_records_fast(
    rows: Sequence["ProjectRow"],
    geo: Dict[str, Any],
    county: str,
) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, Dict[str, Any]]]]:
    """只分析單一縣市，且不重新開 current.xlsx。"""
    crows = [p for p in rows if display_text(p.county) == display_text(county)]
    records: List[Dict[str, Any]] = []

    for p in crows:
        raw_source, raw_x, raw_y, p67_tm2, p67_geo = _coord_audit_fast_raw_and_twd67_candidate(p)
        ref = _coord_audit_gis_reference(geo, p.project_id)

        d_current = _coord_audit_current_distance(p, ref)
        d67_tm2 = _coord_audit_distance(p67_tm2, ref)
        d67_geo = _coord_audit_distance(p67_geo, ref)

        diagnosis = ""
        action = ""
        confidence = ""
        suggested = None
        suggested_system = ""

        if ref:
            alternatives = []
            if p67_tm2 and d67_tm2 is not None:
                alternatives.append(("TWD67 TM2", p67_tm2, d67_tm2))
            if p67_geo and d67_geo is not None:
                alternatives.append(("TWD67經緯度", p67_geo, d67_geo))
            alternatives.sort(key=lambda x: x[2])
            best67 = alternatives[0] if alternatives else None

            if (
                best67
                and d_current is not None
                and d_current - best67[2] >= 120
                and best67[2] <= max(120.0, d_current * 0.45)
            ):
                diagnosis = f"疑似{best67[0]}"
                action = "中央確認轉換"
                confidence = "高"
                suggested = best67[1]
                suggested_system = best67[0]
            elif d_current is not None and d_current <= 50:
                if p.coord_swapped:
                    diagnosis = "經緯度／X-Y顛倒但系統可判讀"
                    action = "中央確認欄位"
                    confidence = "高"
                else:
                    diagnosis = "與GIS校正點一致"
                    action = "無需處理"
                    confidence = "高"
            elif d_current is not None and d_current >= 300:
                diagnosis = "原座標與GIS校正點差異過大"
                action = "退回縣市重填"
                confidence = "高"
            else:
                diagnosis = "需人工確認"
                action = "請縣市確認"
                confidence = "中"
        else:
            if p.lon is None or p.lat is None:
                diagnosis = "無有效座標／需讀原始欄位確認"
                action = "退回縣市重填"
                confidence = "高"
            elif p.coord_swapped:
                diagnosis = "經緯度／X-Y顛倒但系統可判讀"
                action = "中央確認欄位"
                confidence = "高"
            elif "WGS84" in display_text(p.coord_source) or "Google" in display_text(p.coord_source):
                diagnosis = "WGS84格式有效"
                action = "無需處理"
                confidence = "格式判斷"
            elif "TWD97" in display_text(p.coord_source):
                diagnosis = "TWD97格式有效"
                action = "視需要抽查"
                confidence = "格式判斷"
            else:
                diagnosis = "需人工確認"
                action = "請縣市確認"
                confidence = "中"

        records.append({
            "row_key": p.row_key,
            "project_id": p.project_id,
            "sheet_name": p.sheet_name,
            "county": p.county,
            "project_name": p.project_name,
            "project_content": p.project_content,
            "water_system": p.water_system,
            "status": p.status,
            "unit": p.unit,
            "address": p.address,
            "cost_k": None,  # V3.6.45：初檢不讀經費，匯出時才補
            "raw_source": raw_source,
            "raw_x": raw_x,
            "raw_y": raw_y,
            "current_lon": p.lon,
            "current_lat": p.lat,
            "coord_source": p.coord_source,
            "coord_swapped": bool(p.coord_swapped),
            "gis_ref_lon": ref[0] if ref else None,
            "gis_ref_lat": ref[1] if ref else None,
            "has_gis_reference": bool(ref),
            "d_current_m": d_current,
            "d_wgs_m": d_current if ("WGS84" in display_text(p.coord_source) or "Google" in display_text(p.coord_source)) else None,
            "d97_m": d_current if "TWD97" in display_text(p.coord_source) else None,
            "d67_tm2_m": d67_tm2,
            "d67_geo_m": d67_geo,
            "twd97_candidate": None,
            "twd67_tm2_candidate": p67_tm2,
            "twd67_geo_candidate": p67_geo,
            "diagnosis": diagnosis,
            "action": action,
            "confidence": confidence,
            "suggested_lon": suggested.get("lon") if suggested else None,
            "suggested_lat": suggested.get("lat") if suggested else None,
            "suggested_system": suggested_system,
        })

    # 批次診斷只對本縣市這一小批做。
    batch = _coord_audit_batch_assessment(records)

    # 套用同縣市的整批 TWD67 推估。
    b = batch.get(display_text(county), {})
    projected_status = (b.get("projected") or {}).get("status")
    geographic_status = (b.get("geographic") or {}).get("status")

    for r in records:
        if r.get("has_gis_reference"):
            continue
        if projected_status == "高度疑似TWD67" and r.get("twd67_tm2_candidate"):
            r["diagnosis"] = "疑似TWD67 TM2（依縣市批次）"
            r["action"] = "中央確認轉換"
            r["confidence"] = "批次推估"
            r["suggested_lon"] = r["twd67_tm2_candidate"].get("lon")
            r["suggested_lat"] = r["twd67_tm2_candidate"].get("lat")
            r["suggested_system"] = "TWD67 TM2"
        elif geographic_status == "高度疑似TWD67" and r.get("twd67_geo_candidate"):
            r["diagnosis"] = "疑似TWD67經緯度（依縣市批次）"
            r["action"] = "中央確認轉換"
            r["confidence"] = "批次推估"
            r["suggested_lon"] = r["twd67_geo_candidate"].get("lon")
            r["suggested_lat"] = r["twd67_geo_candidate"].get("lat")
            r["suggested_system"] = "TWD67經緯度"

    return records, batch


def _enrich_coordinate_review_records_for_export(
    excel_bytes: bytes,
    rows: Sequence["ProjectRow"],
    records: Sequence[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """只有按下『產製縣市複核表』時才打開 Excel 一次，補原始座標與經費。"""
    if not records:
        return []

    wanted_keys = {display_text(r.get("row_key")) for r in records}
    selected_rows = [p for p in rows if p.row_key in wanted_keys]
    row_map = {p.row_key: p for p in selected_rows}
    result = [dict(r) for r in records]

    # 一次開 workbook，同時讀原始座標與經費，不再開第二本 data_only workbook。
    try:
        wb = load_workbook(
            io.BytesIO(excel_bytes),
            data_only=True,
            read_only=False,
            keep_vba=DEFAULT_EXCEL_PATH.lower().endswith(".xlsm"),
        )
    except Exception:
        return result

    by_sheet: Dict[str, List[Dict[str, Any]]] = {}
    for rec in result:
        p = row_map.get(display_text(rec.get("row_key")))
        if p:
            by_sheet.setdefault(p.sheet_name, []).append(rec)

    for sheet_name, recs in by_sheet.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        hr = find_header_row(ws)
        if not hr:
            continue
        hmap = normalize_header_map(ws, hr)
        coord_cols = _coord_audit_raw_columns(hmap)

        total_col = find_col(
            hmap,
            [
                "總經費(千元)", "總經費（千元）", "總經費",
                "核定經費(千元)", "核定經費（千元）", "核定經費",
            ],
        )
        total_central_col = find_col(hmap, ["總中央款", "中央款合計", "中央補助合計"])
        total_local_col = find_col(hmap, ["總地方款", "地方款合計"])
        component_cols = [
            find_col(hmap, ["中央款工程費合計(含橋梁中央款)", "中央款工程費合計", "中央補助經費(千元)"]),
            find_col(hmap, ["地方款工程費合計(含橋梁地方款)", "地方款工程費合計"]),
            find_col(hmap, ["中央款用地費合計"]),
            find_col(hmap, ["地方款用地費合計"]),
        ]

        for rec in recs:
            p = row_map.get(display_text(rec.get("row_key")))
            if not p:
                continue

            wgx, wgy = _coord_audit_pair(ws, p.row, coord_cols.get("wgs_x"), coord_cols.get("wgs_y"))
            tx, ty = _coord_audit_pair(ws, p.row, coord_cols.get("twd_x"), coord_cols.get("twd_y"))

            if parse_number(wgx) is not None or parse_number(wgy) is not None:
                rec["raw_source"] = "WGS84/Google欄位"
                rec["raw_x"] = wgx
                rec["raw_y"] = wgy
            elif parse_number(tx) is not None or parse_number(ty) is not None:
                rec["raw_source"] = "TWD X/Y欄位"
                rec["raw_x"] = tx
                rec["raw_y"] = ty

            value = None
            if total_col:
                value = parse_number(ws.cell(p.row, total_col).value)
            if value is None and (total_central_col or total_local_col):
                parts = []
                for c in [total_central_col, total_local_col]:
                    if c:
                        v = parse_number(ws.cell(p.row, c).value)
                        if v is not None:
                            parts.append(v)
                if parts:
                    value = sum(parts)
            if value is None:
                parts = []
                for c in component_cols:
                    if c:
                        v = parse_number(ws.cell(p.row, c).value)
                        if v is not None:
                            parts.append(v)
                if parts:
                    value = sum(parts)
            rec["cost_k"] = value

    try:
        wb.close()
    except Exception:
        pass
    return result


@st.cache_data(show_spinner=False, ttl=1800)
def _cached_coordinate_audit_v3644(
    excel_bytes: bytes,
    geo_json_text: str,
) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, Dict[str, Any]]]]:
    # cache key 只依 Excel bytes + GeoJSON 文字；切縣市不重新掃整本 workbook。
    _, rows, _ = scan_workbook(excel_bytes, DEFAULT_EXCEL_PATH, ensure_id_cols=False)
    geo = json.loads(geo_json_text) if geo_json_text else empty_geojson()
    return _build_coordinate_audit_records(rows, geo, excel_bytes)


def _render_coordinate_quality_tab(
    svc: "EngineeringGISService",
    rows: Sequence["ProjectRow"],
    geo: Dict[str, Any],
    registry: Dict[str, Any],
    excel_bytes: bytes,
    editor_name: str,
) -> None:
    st.subheader("📍 座標品質檢核與人工審核")
    st.caption(
        "系統檢核只作為『建議』，不是最後判定。"
        "人工可選一件或多件工程，試套不同座標系統比較落點，再決定是否為座標系統問題或退回縣市重填。"
    )

    counties = sorted({
        display_text(p.county) for p in rows if display_text(p.county)
    })
    if not counties:
        st.info("目前工程資料沒有可辨識的縣市。")
        return

    county = st.selectbox(
        "1. 選擇要檢核的縣市",
        counties,
        key="gis_v3646_coord_county",
    )
    county_rows = [p for p in rows if display_text(p.county) == county]
    st.caption(f"{county}目前共有 {len(county_rows):,} 件工程。")

    cache_root = st.session_state.setdefault("gis_v3645_coord_results", {})
    cached = cache_root.get(county)

    c1, c2 = st.columns([1, 1])
    with c1:
        start_clicked = st.button(
            "▶️ 取得系統檢核建議",
            type="primary",
            use_container_width=True,
            key=f"gis_v3646_start_{county}",
        )
    with c2:
        refresh_clicked = st.button(
            "🔄 重新計算系統建議",
            use_container_width=True,
            disabled=cached is None,
            key=f"gis_v3646_refresh_{county}",
        )

    if start_clicked or refresh_clicked:
        with st.spinner(f"正在快速分析 {county} 座標…"):
            t0 = time.perf_counter()
            records, batch = _build_coordinate_audit_records_fast(
                county_rows, geo, county
            )
            cache_root[county] = {
                "records": records,
                "batch": batch,
                "elapsed": time.perf_counter() - t0,
                "created_at": now_iso(),
            }
            cached = cache_root[county]

    if cached is None:
        st.info("請先按「取得系統檢核建議」。系統不會自動把建議當成最終處置。")
        return

    base_records = list(cached.get("records") or [])
    batch = dict(cached.get("batch") or {})
    elapsed = float(cached.get("elapsed") or 0.0)
    records = _apply_manual_reviews_to_records_v3646(base_records, registry)

    st.success(
        f"系統建議已完成：{county} {len(records):,} 件，約 {elapsed:.2f} 秒。"
        "下方人工檢核結果才是承辦人最後判斷。"
    )

    total = len(records)
    refs = sum(1 for r in records if r.get("has_gis_reference"))
    reviewed = sum(1 for r in records if display_text(r.get("manual_decision")))
    return_manual = sum(
        1 for r in records if r.get("manual_decision") == "退回縣市政府重填"
    )
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("本縣市工程", f"{total:,}")
    m2.metric("人工GIS參考點", f"{refs:,}")
    m3.metric("已人工檢核", f"{reviewed:,}")
    m4.metric("人工判定退回", f"{return_manual:,}")

    st.markdown("#### 2. 系統批次診斷建議")
    b = batch.get(county, {})
    for label, mode in [
        ("TM2 X/Y（TWD97 vs TWD67）", "projected"),
        ("經緯度（WGS84 vs TWD67 Geographic）", "geographic"),
    ]:
        info = b.get(mode) or {}
        status = info.get("status", "樣本不足")
        n = int(info.get("sample_count", 0) or 0)
        detail = display_text(info.get("detail"))
        if status == "高度疑似TWD67":
            st.warning(f"**系統建議｜{label}：{status}**｜參考樣本 {n} 件\n\n{detail}")
        elif "較符合" in status:
            st.success(f"**系統建議｜{label}：{status}**｜參考樣本 {n} 件\n\n{detail}")
        else:
            st.info(f"**系統建議｜{label}：{status}**｜參考樣本 {n} 件\n\n{detail}")

    st.caption(
        "以上只提供人工判讀方向。即使系統顯示『高度疑似TWD67』，也不會直接修改正式座標。"
    )

    st.markdown("#### 3. 人工座標檢核工作台")
    st.caption(
        "直接在地圖點工程：點一次加入、再點一次取消。"
        "可連續點選多件後一起試轉相同座標系統。"
    )

    selection_key = f"gis_v3647_manual_map_selection::{county}"
    selected_keys = [
        display_text(x)
        for x in st.session_state.get(selection_key, [])
        if display_text(x)
    ]
    valid_row_keys = {display_text(r.get("row_key")) for r in records}
    selected_keys = [x for x in selected_keys if x in valid_row_keys]
    st.session_state[selection_key] = selected_keys

    map_scope = st.radio(
        "地圖顯示",
        ["尚未人工檢核", "系統建議異常／待確認", "全部可定位案件"],
        horizontal=True,
        key=f"gis_v3647_manual_map_scope_{county}",
    )

    map_epoch_key = f"gis_v3647_manual_map_epoch::{county}"
    map_epoch = int(st.session_state.get(map_epoch_key, 0) or 0)
    select_map, selectable_points, unmappable_count = _build_manual_selection_map_v3647(
        records,
        selected_keys,
        map_scope,
        memory_key=f"coord_manual_select::{county}",
    )

    if select_map is None:
        st.info("目前篩選條件沒有可定位的工程點位。")
        map_state = None
    else:
        map_state = st_folium(
            select_map,
            width=None,
            height=600,
            returned_objects=["last_object_clicked", "last_object_clicked_tooltip"],
            key=f"gis_v3647_manual_select_map_{county}_{map_scope}_{map_epoch}",
        )

    clicked_rec = _match_clicked_manual_record_v3647(map_state, selectable_points)
    if clicked_rec is not None:
        clicked_key = display_text(clicked_rec.get("row_key"))
        updated = list(selected_keys)
        if clicked_key in updated:
            updated.remove(clicked_key)
            st.session_state["gis_v3647_selection_notice"] = (
                f"已取消：{display_text(clicked_rec.get('project_name'))}"
            )
        else:
            if len(updated) >= 30:
                st.session_state["gis_v3647_selection_notice"] = (
                    "一次人工試轉最多 30 件；請先取消部分已選工程。"
                )
            else:
                updated.append(clicked_key)
                st.session_state["gis_v3647_selection_notice"] = (
                    f"已加入：{display_text(clicked_rec.get('project_name'))}"
                )
        st.session_state[selection_key] = updated
        st.session_state[map_epoch_key] = map_epoch + 1
        st.rerun()

    notice = st.session_state.pop("gis_v3647_selection_notice", "")
    if notice:
        st.info(notice)

    selected_set = set(st.session_state.get(selection_key, []))
    selected_records = [
        r for r in records if display_text(r.get("row_key")) in selected_set
    ]

    a1, a2, a3 = st.columns([1, 1, 2])
    a1.metric("目前已選", f"{len(selected_records)} 件")
    with a2:
        if st.button(
            "🧹 清除全部選取",
            use_container_width=True,
            disabled=not selected_records,
            key=f"gis_v3647_clear_selected_{county}",
        ):
            st.session_state[selection_key] = []
            st.session_state[map_epoch_key] = map_epoch + 1
            st.rerun()
    with a3:
        if unmappable_count:
            st.caption(
                f"另有 {unmappable_count:,} 件沒有有效地圖座標，無法用地圖點選；"
                "這類案件可在後面的縣市複核表直接處理。"
            )

    if selected_records:
        st.dataframe(
            [
                {
                    "工程名稱": r.get("project_name"),
                    "系統工程ID": r.get("project_id"),
                    "系統建議": r.get("diagnosis"),
                    "人工結果": r.get("manual_decision") or "尚未人工檢核",
                }
                for r in selected_records
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("尚未選取工程，請直接在上方地圖點選要人工確認的點位。")

    if len(selected_records) > 30:
        st.warning("為維持地圖操作順暢，人工試轉一次最多 30 件；請取消部分工程。")
        return

    raw_cache = st.session_state.setdefault("gis_v3646_manual_raw_cache", {})
    selected_sig = "|".join(sorted(display_text(r.get("row_key")) for r in selected_records))
    raw_key = f"{county}::{selected_sig}"

    if selected_records:
        st.caption(
            "要測試不同 CRS，需要讀取原始 X/Y。只有按下方按鈕時才開 current.xlsx 一次，"
            "之後反覆切換座標系統不會再讀 Excel。"
        )
        if st.button(
            f"📥 載入這 {len(selected_records)} 件原始座標供人工試轉",
            type="primary",
            use_container_width=True,
            key=f"gis_v3646_load_raw_{county}",
        ):
            selected_row_keys = {display_text(r.get("row_key")) for r in selected_records}
            selected_project_rows = [p for p in county_rows if p.row_key in selected_row_keys]
            _mask = _saving_overlay("正在載入選取案件原始座標…")
            try:
                t0 = time.perf_counter()
                raw_cache[raw_key] = {
                    "rows": _load_selected_raw_coordinates_v3646(
                        excel_bytes, selected_project_rows
                    ),
                    "elapsed": time.perf_counter() - t0,
                }
            finally:
                _close_saving_overlay(_mask)

    raw_loaded = raw_cache.get(raw_key) if selected_records else None
    if selected_records and raw_loaded:
        raw_rows = raw_loaded.get("rows") or {}
        st.success(
            f"已載入 {len(raw_rows):,} 件原始座標，約 {float(raw_loaded.get('elapsed') or 0):.1f} 秒。"
            "現在切換下列座標系統只做記憶體內轉換。"
        )

        pair_source = st.radio(
            "使用哪一組原始欄位",
            ["自動選擇", "WGS84 / Google 原始欄位", "TWD X/Y 原始欄位"],
            horizontal=True,
            key=f"gis_v3646_pair_source_{county}",
        )
        tested_crs = st.selectbox(
            "人工測試座標系統",
            [
                "WGS84",
                "WGS84（X/Y對調）",
                "TWD97 TM2 zone 121",
                "TWD97 TM2 zone 119",
                "TWD67 TM2 zone 121",
                "TWD67 TM2 zone 119",
                "TWD67 Geographic",
                "TWD67 Geographic（X/Y對調）",
            ],
            key=f"gis_v3646_tested_crs_{county}",
        )

        candidates: Dict[str, Dict[str, Any]] = {}
        preview_rows = []
        valid_count = 0
        for rec in selected_records:
            row_key = display_text(rec.get("row_key"))
            raw = raw_rows.get(row_key) or {}
            cand = _manual_transform_candidate_v3646(raw, pair_source, tested_crs)
            candidates[row_key] = cand
            if cand.get("valid"):
                valid_count += 1

            d_to_gis = None
            if cand.get("valid") and rec.get("gis_ref_lon") is not None and rec.get("gis_ref_lat") is not None:
                d_to_gis = _haversine_m(
                    float(cand["lat"]), float(cand["lon"]),
                    float(rec["gis_ref_lat"]), float(rec["gis_ref_lon"]),
                )

            d_from_current = None
            if (
                cand.get("valid")
                and is_valid_wgs84(
                    parse_number(rec.get("current_lon")),
                    parse_number(rec.get("current_lat")),
                )
            ):
                d_from_current = _haversine_m(
                    float(rec["current_lat"]), float(rec["current_lon"]),
                    float(cand["lat"]), float(cand["lon"]),
                )

            preview_rows.append({
                "系統工程ID": rec.get("project_id"),
                "工程名稱": rec.get("project_name"),
                "系統建議": rec.get("diagnosis"),
                "實際使用原始欄位": cand.get("used_source"),
                "原X": cand.get("raw_x"),
                "原Y": cand.get("raw_y"),
                "人工測試CRS": tested_crs,
                "候選經度": None if not cand.get("valid") else round(float(cand["lon"]), 7),
                "候選緯度": None if not cand.get("valid") else round(float(cand["lat"]), 7),
                "相對目前位置位移(m)": None if d_from_current is None else round(float(d_from_current), 1),
                "與人工GIS參考點距離(m)": None if d_to_gis is None else round(float(d_to_gis), 1),
                "轉換狀態": "可顯示" if cand.get("valid") else cand.get("reason"),
            })

        st.success(
            f"🟠 已依「{tested_crs}」重新計算並重繪候選位置："
            f"{len(selected_records)} 件中 {valid_count} 件有效。"
        )
        st.caption(
            "灰色小點＝目前解析位置；橘色大點＋白框＝本次試轉候選；"
            "藍色＝人工GIS校正參考點。表格新增『相對目前位置位移(m)』。"
        )
        st.dataframe(preview_rows, hide_index=True, use_container_width=True)

        hide_current_preview = st.checkbox(
            "👁️ 隱藏灰色原位置，只看本次候選＋GIS參考點",
            value=False,
            key=f"gis_v3648_hide_current_{county}_{selected_sig}",
        )

        mmap = _manual_review_map_v3648(
            selected_records,
            candidates,
            tested_crs=tested_crs,
            show_current=not hide_current_preview,
        )
        if mmap is not None:
            st_folium(
                mmap,
                width=None,
                height=650,
                returned_objects=[],
                key=(
                    f"gis_v3648_manual_map_{county}_{tested_crs}_"
                    f"{pair_source}_{selected_sig}_{hide_current_preview}"
                ),
            )
        else:
            st.info("本次選取案件沒有可顯示的有效候選位置。")

        st.markdown("##### 人工最後判定")
        manual_decision = st.radio(
            "本次選取案件要如何處理",
            [
                "確認為座標系統問題－採本次測試結果",
                "退回縣市政府重填",
                "確認原座標可用／不需處理",
                "待後續確認",
            ],
            key=f"gis_v3646_manual_decision_{county}",
        )
        note = st.text_area(
            "人工檢核備註",
            placeholder="例如：TWD67 zone 121 轉換後與既有排水線位置吻合；或現場位置仍無法判斷，退回縣市確認。",
            key=f"gis_v3646_manual_note_{county}",
        )

        if manual_decision == "確認為座標系統問題－採本次測試結果" and valid_count != len(selected_records):
            st.warning("部分案件沒有有效的試轉候選點，不能把整批直接判定為採本次測試結果。請縮小選取或改測其他座標系統。")
            save_disabled = True
        else:
            save_disabled = False

        st.warning(
            "這裡的「儲存人工檢核結論」只會寫入審核紀錄，"
            "不會把候選座標直接覆蓋到 GIS 或 current.xlsx。"
        )
        if st.button(
            f"💾 儲存 {len(selected_records)} 件人工檢核結論",
            type="primary",
            use_container_width=True,
            disabled=save_disabled,
            key=f"gis_v3646_save_review_{county}",
        ):
            payload = []
            for rec in selected_records:
                row_key = display_text(rec.get("row_key"))
                cand = candidates.get(row_key) or {}
                payload.append({
                    "project_id": rec.get("project_id"),
                    "row_key": row_key,
                    "project_name": rec.get("project_name"),
                    "county": rec.get("county"),
                    "system_suggestion": rec.get("diagnosis"),
                    "system_action": rec.get("action"),
                    "manual_decision": manual_decision,
                    "raw_pair_source": cand.get("used_source") or pair_source,
                    "tested_crs": tested_crs if manual_decision == "確認為座標系統問題－採本次測試結果" else "",
                    "candidate_lon": cand.get("lon") if manual_decision == "確認為座標系統問題－採本次測試結果" else None,
                    "candidate_lat": cand.get("lat") if manual_decision == "確認為座標系統問題－採本次測試結果" else None,
                    "review_note": note,
                })

            _mask = _saving_overlay("正在儲存人工座標檢核結論…")
            try:
                res = svc.save_coordinate_reviews(payload, editor_name)
            finally:
                _close_saving_overlay(_mask)

            # 同步目前 session 裡的 registry，避免剛存完還看舊值。
            bucket = registry.setdefault("coordinate_reviews", {})
            for ent in res.get("reviews", []):
                key = display_text(ent.get("project_id")) or display_text(ent.get("row_key"))
                if key:
                    bucket[key] = ent

            hist_bucket = registry.setdefault("coordinate_review_history", [])
            if isinstance(hist_bucket, list):
                hist_bucket.extend(res.get("history_entries", []))

            st.success(
                f"已儲存 {res.get('count', 0)} 件人工檢核結論；正式 GIS 與 current.xlsx 均未修改。"
            )
            st.rerun()

    elif selected_records:
        st.info("請先載入這批案件的原始座標，再進行人工座標系統試轉。")

    st.markdown("#### 4. 檢核明細")
    scope = st.radio(
        "明細顯示",
        ["只看尚未人工檢核", "只看人工判定退回縣市", "只看異常／待確認", "全部案件"],
        horizontal=True,
        key=f"gis_v3646_coord_table_scope_{county}",
    )
    if scope == "只看尚未人工檢核":
        shown = [r for r in records if not display_text(r.get("manual_decision"))]
    elif scope == "只看人工判定退回縣市":
        shown = [r for r in records if r.get("manual_decision") == "退回縣市政府重填"]
    elif scope == "只看異常／待確認":
        shown = [
            r for r in records
            if _coordinate_issue_in_merged_export_v3649(r)
        ]
    else:
        shown = records

    st.dataframe(
        _coord_audit_table_rows(shown),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("#### 5. 產製縣市座標複核表")
    export_scope = st.radio(
        "匯出範圍",
        [
            "只匯出人工判定退回縣市",
            "只匯出系統建議退回縣市重填",
            "匯出全部異常與待確認",
            "匯出本縣市全部案件",
        ],
        horizontal=True,
        key=f"gis_v3646_coord_export_scope_{county}",
    )
    if export_scope == "只匯出人工判定退回縣市":
        export_rows = [
            r for r in records if r.get("manual_decision") == "退回縣市政府重填"
        ]
    elif export_scope == "只匯出系統建議退回縣市重填":
        export_rows = [r for r in records if r.get("action") == "退回縣市重填"]
    elif export_scope == "匯出全部異常與待確認":
        export_rows = [
            r for r in records
            if _coordinate_issue_in_merged_export_v3649(r)
        ]
    else:
        export_rows = list(records)

    if export_scope == "匯出全部異常與待確認":
        sys_only_n = sum(
            1 for r in export_rows
            if _coordinate_issue_source_v3649(r) == "系統建議"
        )
        manual_n = sum(
            1 for r in export_rows
            if display_text(r.get("manual_decision"))
        )
        st.caption(
            f"目前整併共 {len(export_rows):,} 件：其中尚未人工覆核的系統異常/待確認 "
            f"{sys_only_n:,} 件；已有人工判定且仍需處理/待確認 {manual_n:,} 件。"
            "人工若已確認『原座標可用／不需處理』，會從整併異常清單排除。"
        )
    else:
        st.caption(
            f"目前範圍共 {len(export_rows):,} 件。"
            "正式退回縣市時仍可使用「只匯出人工判定退回縣市」。"
        )

    export_cache = st.session_state.setdefault("gis_v3645_coord_export_files", {})
    export_fp = _coordinate_issue_export_fingerprint_v3649(export_rows)
    export_key = f"v3649::{county}::{export_scope}::{export_fp}"

    if st.button(
        "📄 產製縣市座標複核表",
        disabled=not export_rows,
        use_container_width=True,
        key=f"gis_v3646_build_export_{county}",
    ):
        _mask = _saving_overlay("正在產製縣市座標複核表…")
        try:
            t0 = time.perf_counter()

            # 產檔當下再讀 GitHub 最新 registry，納入其他使用者剛儲存的人工結論。
            latest_for_export = _latest_registry_only_v3648(svc)
            latest_records = _merge_latest_manual_into_records_v3649(
                base_records, latest_for_export
            )

            if export_scope == "只匯出人工判定退回縣市":
                export_rows_now = [
                    r for r in latest_records
                    if r.get("manual_decision") == "退回縣市政府重填"
                ]
            elif export_scope == "只匯出系統建議退回縣市重填":
                export_rows_now = [
                    r for r in latest_records
                    if r.get("action") == "退回縣市重填"
                ]
            elif export_scope == "匯出全部異常與待確認":
                export_rows_now = [
                    r for r in latest_records
                    if _coordinate_issue_in_merged_export_v3649(r)
                ]
            else:
                export_rows_now = list(latest_records)

            enriched = _enrich_coordinate_review_records_for_export(
                excel_bytes, county_rows, export_rows_now
            )

            latest_reviews = (
                latest_for_export.get("coordinate_reviews") or {}
            ) if isinstance(latest_for_export, dict) else {}
            for rec in enriched:
                rv = latest_reviews.get(
                    _coordinate_review_key(rec), {}
                ) if isinstance(latest_reviews, dict) else {}
                if isinstance(rv, dict):
                    rec["manual_decision"] = display_text(rv.get("manual_decision"))
                    rec["manual_tested_crs"] = display_text(rv.get("tested_crs"))
                    rec["manual_reviewed_by"] = display_text(rv.get("reviewed_by"))
                    rec["manual_reviewed_at"] = display_text(rv.get("reviewed_at"))

            xlsx = _build_county_coordinate_review_xlsx(county, enriched)
            export_cache[export_key] = {
                "bytes": xlsx,
                "elapsed": time.perf_counter() - t0,
            }
        finally:
            _close_saving_overlay(_mask)

    ready = export_cache.get(export_key)
    if ready:
        safe_county = re.sub(r'[\\/:*?"<>|]+', "_", county) or "縣市"
        date_txt = datetime.now().strftime("%Y%m%d")
        st.success(
            f"複核表已產製完成，約 {float(ready.get('elapsed') or 0):.1f} 秒。"
        )
        st.download_button(
            "📤 下載縣市座標複核表",
            data=ready["bytes"],
            file_name=f"{safe_county}_工程座標複核表_{date_txt}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"gis_v3646_download_coord_review_{county}",
        )

    st.markdown("#### 6. 📦 所有使用者人工檢核彙整與縣市輸出")
    st.caption(
        "不同使用者可以分批、分天儲存。這裡會另外只讀 GitHub 最新 id_registry.json，"
        "把所有人的『每件最新人工結論』彙整起來，不需要重新下載 current.xlsx 或 GeoJSON。"
    )

    agg_cache_key = "gis_v3648_latest_registry_for_export"
    latest_registry = st.session_state.get(agg_cache_key)

    if st.button(
        "🔄 讀取所有使用者最新人工檢核結果",
        use_container_width=True,
        key="gis_v3648_refresh_all_reviews",
    ):
        _mask = _saving_overlay("正在讀取所有使用者最新人工檢核結果…")
        try:
            latest_registry = _latest_registry_only_v3648(svc)
            st.session_state[agg_cache_key] = latest_registry
        finally:
            _close_saving_overlay(_mask)

    if not isinstance(latest_registry, dict):
        st.info(
            "要正式彙整輸出前，請先按「讀取所有使用者最新人工檢核結果」。"
            "這可以避免目前瀏覽 Session 尚未看到其他人剛儲存的結果。"
        )
    else:
        all_review_records, missing_current = _all_latest_coordinate_review_records_v3648(
            rows, latest_registry
        )
        latest_count = len(all_review_records)
        return_records = [
            r for r in all_review_records
            if r.get("manual_decision") == "退回縣市政府重填"
        ]
        crs_records = [
            r for r in all_review_records
            if r.get("manual_decision") == "確認為座標系統問題－採本次測試結果"
        ]
        pending_records = [
            r for r in all_review_records
            if r.get("manual_decision") == "待後續確認"
        ]
        ok_records = [
            r for r in all_review_records
            if r.get("manual_decision") == "確認原座標可用／不需處理"
        ]

        q1, q2, q3, q4 = st.columns(4)
        q1.metric("累計最新人工結論", f"{latest_count:,}")
        q2.metric("退回縣市重填", f"{len(return_records):,}")
        q3.metric("中央確認CRS問題", f"{len(crs_records):,}")
        q4.metric("待後續確認", f"{len(pending_records):,}")

        st.caption(
            f"另有 {len(ok_records):,} 件人工確認原座標可用／不需處理。"
            "正式縣市複核表只會使用『退回縣市政府重填』，"
            "不會把中央已確認可自行轉換或仍待確認的案件一起送出。"
        )
        if missing_current:
            st.warning(
                f"有 {missing_current:,} 件人工檢核紀錄目前已找不到 current.xlsx 對應列；"
                "仍保留人工紀錄，但若其被判定退回，輸出時工程內容／經費／原座標可能留白。"
            )

        st.markdown("##### 系統建議＋人工判定整併異常清單")
        st.caption(
            "若要一次取得『系統判定異常/視需要抽查/待確認』加上『人工判定仍有問題』，"
            "請建立整併清單。人工已確認原座標可用的案件會排除。"
            "這個動作只做快速座標分析，不重新開 current.xlsx。"
        )

        merged_cache = st.session_state.setdefault(
            "gis_v3649_all_merged_issue_cache", {}
        )
        latest_fp = _coordinate_review_registry_fingerprint_v3648(latest_registry)
        merged_cache_key = f"{latest_fp}::{len(rows)}"

        if st.button(
            "🧩 建立全縣市整併異常與待確認清單",
            type="primary",
            use_container_width=True,
            key="gis_v3649_build_all_merged_issues",
        ):
            _mask = _saving_overlay("正在整併系統建議與所有人工判定…")
            try:
                merged_records_all, merged_missing = (
                    _build_all_counties_merged_issue_records_v3649(
                        rows, geo, latest_registry
                    )
                )
                merged_cache[merged_cache_key] = {
                    "records": merged_records_all,
                    "missing_current": merged_missing,
                }
            finally:
                _close_saving_overlay(_mask)

        merged_ready = merged_cache.get(merged_cache_key)
        if merged_ready:
            merged_records_all = list(merged_ready.get("records") or [])
            merged_missing = int(merged_ready.get("missing_current") or 0)

            merged_system_only = [
                r for r in merged_records_all
                if _coordinate_issue_source_v3649(r) == "系統建議"
            ]
            merged_manual = [
                r for r in merged_records_all
                if display_text(r.get("manual_decision"))
            ]
            merged_counties = sorted({
                display_text(r.get("county"))
                for r in merged_records_all
                if display_text(r.get("county"))
            })

            u1, u2, u3, u4 = st.columns(4)
            u1.metric("整併異常/待確認", f"{len(merged_records_all):,}")
            u2.metric("僅系統建議", f"{len(merged_system_only):,}")
            u3.metric("含人工判定", f"{len(merged_manual):,}")
            u4.metric("涉及縣市", f"{len(merged_counties):,}")

            if merged_missing:
                st.caption(
                    f"其中有 {merged_missing:,} 件人工紀錄目前找不到 current.xlsx 對應列；"
                    "仍會保留在整併名單，但可供縣市填寫的工程內容可能不完整。"
                )

            merged_summary = []
            for cty in merged_counties:
                crows = [
                    r for r in merged_records_all
                    if display_text(r.get("county")) == cty
                ]
                merged_summary.append({
                    "縣市": cty,
                    "全部異常與待確認": len(crows),
                    "僅系統建議": sum(
                        1 for r in crows
                        if _coordinate_issue_source_v3649(r) == "系統建議"
                    ),
                    "已有人工判定": sum(
                        1 for r in crows
                        if display_text(r.get("manual_decision"))
                    ),
                })
            if merged_summary:
                st.dataframe(
                    merged_summary,
                    use_container_width=True,
                    hide_index=True,
                )

            merged_export_cache = st.session_state.setdefault(
                "gis_v3649_merged_export_files", {}
            )
            merged_fp = _coordinate_issue_export_fingerprint_v3649(
                merged_records_all
            )

            if merged_counties:
                merged_cty = st.selectbox(
                    "整併清單：選擇單一縣市輸出",
                    merged_counties,
                    key="gis_v3649_merged_county",
                )
                merged_single = [
                    r for r in merged_records_all
                    if display_text(r.get("county")) == merged_cty
                ]
                merged_single_key = f"single::{merged_fp}::{merged_cty}"

                if st.button(
                    f"📄 產製 {merged_cty} 全部異常與待確認（系統＋人工）",
                    use_container_width=True,
                    key="gis_v3649_build_merged_single",
                ):
                    _mask = _saving_overlay(
                        f"正在產製 {merged_cty} 整併異常複核表…"
                    )
                    try:
                        enriched = _enrich_coordinate_review_records_for_export(
                            excel_bytes, rows, merged_single
                        )
                        merged_export_cache[merged_single_key] = (
                            _build_county_coordinate_review_xlsx(
                                merged_cty, enriched
                            )
                        )
                    finally:
                        _close_saving_overlay(_mask)

                if merged_single_key in merged_export_cache:
                    date_txt = datetime.now().strftime("%Y%m%d")
                    safe_cty = re.sub(
                        r'[\\/:*?"<>|]+', "_", merged_cty
                    ) or "縣市"
                    st.download_button(
                        f"📤 下載 {merged_cty} 整併異常複核表",
                        data=merged_export_cache[merged_single_key],
                        file_name=(
                            f"{safe_cty}_全部異常與待確認_"
                            f"系統加人工_{date_txt}.xlsx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        use_container_width=True,
                        key="gis_v3649_download_merged_single",
                    )

                merged_zip_key = f"zip::{merged_fp}"
                if st.button(
                    f"📦 一次產製全部 {len(merged_counties)} 個縣市整併 ZIP",
                    use_container_width=True,
                    key="gis_v3649_build_merged_zip",
                ):
                    _mask = _saving_overlay(
                        "正在產製全縣市『系統建議＋人工判定』整併異常 ZIP…"
                    )
                    try:
                        # 只在真正產檔時開 current.xlsx 一次。
                        enriched_all = _enrich_coordinate_review_records_for_export(
                            excel_bytes, rows, merged_records_all
                        )
                        grouped: Dict[str, List[Dict[str, Any]]] = {}
                        for rec in enriched_all:
                            cty = display_text(rec.get("county"))
                            if cty:
                                grouped.setdefault(cty, []).append(rec)
                        merged_export_cache[merged_zip_key] = (
                            _build_coordinate_review_zip_v3648(grouped)
                        )
                    finally:
                        _close_saving_overlay(_mask)

                if merged_zip_key in merged_export_cache:
                    date_txt = datetime.now().strftime("%Y%m%d")
                    st.download_button(
                        "⬇️ 下載全部縣市整併異常複核表 ZIP",
                        data=merged_export_cache[merged_zip_key],
                        file_name=(
                            f"各縣市_全部異常與待確認_"
                            f"系統加人工_{date_txt}.zip"
                        ),
                        mime="application/zip",
                        use_container_width=True,
                        key="gis_v3649_download_merged_zip",
                    )

        county_summary = []
        return_counties = sorted({
            display_text(r.get("county"))
            for r in return_records if display_text(r.get("county"))
        })
        for cty in return_counties:
            rows_cty = [
                r for r in return_records
                if display_text(r.get("county")) == cty
            ]
            reviewers = sorted({
                display_text(r.get("manual_reviewed_by"))
                for r in rows_cty if display_text(r.get("manual_reviewed_by"))
            })
            county_summary.append({
                "縣市": cty,
                "人工判定需重填": len(rows_cty),
                "參與檢核者": "、".join(reviewers),
                "最晚檢核時間": max(
                    [display_text(r.get("manual_reviewed_at")) for r in rows_cty]
                    or [""]
                ),
            })

        if county_summary:
            st.dataframe(county_summary, use_container_width=True, hide_index=True)
        else:
            st.success("目前沒有任何人工最後判定為「退回縣市政府重填」的案件。")

        # 可檢視歷程，但輸出判斷只採每件最新結論。
        with st.expander("🕘 查看人工檢核歷程（同一工程可有多次判斷）", expanded=False):
            hist = latest_registry.get("coordinate_review_history") or []
            if isinstance(hist, list) and hist:
                hist_rows = []
                for h in reversed(hist[-500:]):
                    if not isinstance(h, dict):
                        continue
                    hist_rows.append({
                        "檢核時間": h.get("reviewed_at"),
                        "檢核者": h.get("reviewed_by"),
                        "縣市": h.get("county"),
                        "工程名稱": h.get("project_name"),
                        "系統工程ID": h.get("project_id"),
                        "人工結論": h.get("manual_decision"),
                        "測試CRS": h.get("tested_crs"),
                        "備註": h.get("review_note"),
                    })
                st.dataframe(hist_rows, use_container_width=True, hide_index=True)
                if len(hist) > 500:
                    st.caption(f"歷程共 {len(hist):,} 筆，畫面先顯示最近 500 筆。")
            else:
                st.caption(
                    "目前尚無 V3.6.48 之後新增的歷程；V3.6.47以前只保存每件最新結論。"
                )

        if return_counties:
            st.markdown("##### 僅人工判定退回：依縣市產製正式複核檔")
            selected_export_county = st.selectbox(
                "選擇單一縣市",
                return_counties,
                key="gis_v3648_aggregate_county",
            )

            registry_fp = _coordinate_review_registry_fingerprint_v3648(
                latest_registry
            )
            agg_export_cache = st.session_state.setdefault(
                "gis_v3648_aggregate_export_cache", {}
            )

            single_records = [
                r for r in return_records
                if display_text(r.get("county")) == selected_export_county
            ]
            single_key = f"single::{registry_fp}::{selected_export_county}"

            if st.button(
                f"📄 產製 {selected_export_county} 全部人工退回案件",
                use_container_width=True,
                key="gis_v3648_build_single_county",
            ):
                _mask = _saving_overlay(
                    f"正在產製 {selected_export_county} 座標複核表…"
                )
                try:
                    enriched = _enrich_coordinate_review_records_for_export(
                        excel_bytes, rows, single_records
                    )
                    xlsx = _build_county_coordinate_review_xlsx(
                        selected_export_county, enriched
                    )
                    agg_export_cache[single_key] = xlsx
                finally:
                    _close_saving_overlay(_mask)

            if single_key in agg_export_cache:
                date_txt = datetime.now().strftime("%Y%m%d")
                safe_cty = re.sub(
                    r'[\\/:*?"<>|]+', "_", selected_export_county
                ) or "縣市"
                st.download_button(
                    f"📤 下載 {selected_export_county} 複核表",
                    data=agg_export_cache[single_key],
                    file_name=f"{safe_cty}_工程座標複核表_{date_txt}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="gis_v3648_download_single_county",
                )

            all_key = f"allzip::{registry_fp}::{len(return_records)}"
            if st.button(
                f"📦 一次產製全部 {len(return_counties)} 個縣市 ZIP",
                type="primary",
                use_container_width=True,
                key="gis_v3648_build_all_counties_zip",
            ):
                _mask = _saving_overlay(
                    "正在一次產製所有縣市座標複核表…"
                )
                try:
                    # 只開 current.xlsx 一次，把所有需退回案件的原座標與經費一起補齊。
                    enriched_all = _enrich_coordinate_review_records_for_export(
                        excel_bytes, rows, return_records
                    )
                    grouped: Dict[str, List[Dict[str, Any]]] = {}
                    for rec in enriched_all:
                        cty = display_text(rec.get("county"))
                        if cty:
                            grouped.setdefault(cty, []).append(rec)
                    agg_export_cache[all_key] = _build_coordinate_review_zip_v3648(
                        grouped
                    )
                finally:
                    _close_saving_overlay(_mask)

            if all_key in agg_export_cache:
                date_txt = datetime.now().strftime("%Y%m%d")
                st.download_button(
                    "⬇️ 下載全部縣市座標複核表 ZIP",
                    data=agg_export_cache[all_key],
                    file_name=f"各縣市_工程座標複核表_{date_txt}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="gis_v3648_download_all_zip",
                )

    st.divider()
    st.info(
        "V3.6.49 提供兩種輸出："
        "①『全部異常與待確認』＝系統異常/視需要抽查＋人工仍需處理案件整併；"
        "人工已確認無需處理者排除。"
        "②『人工判定退回縣市』＝只含人工明確決定要退回的案件。"
        "本版仍不會自動修改正式 GIS 或 current.xlsx。"
    )



def _project_popup_v35(p: ProjectRow, feature: Optional[Dict[str, Any]] = None) -> str:
    fp = (feature or {}).get("properties") or {}
    kw = display_text(fp.get("keywords"))
    geo_name = display_text(fp.get("geo_name"))
    extra = ""
    if geo_name:
        extra += f"<br>圖資名稱：{_popup_escape(geo_name)}"
    if kw:
        extra += f"<br>圖資關鍵字：{_popup_escape(kw)}"
    spj = display_text(fp.get("spatial_project_id"))
    linked = [display_text(x) for x in (fp.get("spatial_linked_project_ids") or []) if display_text(x)]
    if spj:
        extra += f"<br>空間工程ID：{_popup_escape(spj)}"
    if len(linked) > 1:
        extra += f"<br>共用圖資管控紀錄：{len(linked)}筆（{_popup_escape('、'.join(linked))}）"
    return (
        f"<b>{_popup_escape(p.project_name)}</b><br>"
        f"系統工程ID：{_popup_escape(p.project_id)}<br>"
        f"分頁：{_popup_escape(p.sheet_name)}<br>"
        f"縣市：{_popup_escape(p.county)}<br>"
        f"執行單位：{_popup_escape(p.unit)}<br>"
        f"執行情形：{_popup_escape(p.status)}"
        f"{extra}"
    )


def _reach_popup(f: Dict[str, Any]) -> str:
    p = f.get("properties") or {}
    kw = display_text(p.get("keywords"))
    return (
        f"<b>{_popup_escape(p.get('water_name') or p.get('reach_name') or '治理現況底圖')}</b><br>"
        f"治理狀態：{_popup_escape(p.get('status'))}<br>"
        f"縣市：{_popup_escape(p.get('county'))}<br>"
        f"治理依據：{_popup_escape(p.get('basis'))}<br>"
        f"備註：{_popup_escape(p.get('notes'))}"
        + (f"<br>圖資關鍵字：{_popup_escape(kw)}" if kw else "")
    )


def _add_feature_shape(m: folium.Map, feature: Dict[str, Any], color: str, popup_html: str, tooltip: str = "", opacity: float = .92, reference: bool = False) -> bool:
    """繪製圖資。V3.6.24：正式點／線／面一律加黑色外框；OSM/灰色參考層維持原樣。"""
    geom = feature.get("geometry") or {}
    gt = geom.get("type")
    coords = geom.get("coordinates")
    weight = 2 if reference else _line_weight(feature, 6)
    dash = "6,6" if reference else None
    try:
        if gt == "Point" and isinstance(coords, list) and len(coords) >= 2:
            lon, lat = float(coords[0]), float(coords[1])
            folium.CircleMarker(
                [lat, lon], radius=5 if reference else 7,
                color="#000000", weight=2 if reference else 2.5,
                fill=True, fill_color=color, fill_opacity=.75 if reference else .95,
                popup=None if reference else folium.Popup(popup_html, max_width=420),
                tooltip=tooltip or None,
            ).add_to(m)
            return True
        if gt == "LineString" and coords:
            locs = [[float(y), float(x)] for x, y, *_ in coords]
            casing = folium.PolyLine(
                locs, color="#000000", weight=weight + (2.5 if reference else 3.5), opacity=min(1.0, opacity + .05),
                interactive=False,
            )
            casing.add_to(m)
            folium.PolyLine(
                locs, color=color, weight=weight, opacity=opacity,
                dash_array=dash,
                popup=None if reference else folium.Popup(popup_html, max_width=420),
                tooltip=tooltip or None,
            ).add_to(m)
            return True
        if gt == "Polygon" and coords and coords[0]:
            locs = [[float(y), float(x)] for x, y, *_ in coords[0]]
            folium.Polygon(
                locs,
                color="#000000",
                weight=max(2.5, min(6, weight + 1)) if reference else max(2.5, min(7, weight + 1.5)),
                fill=True, fill_color=color, fill_opacity=.12 if reference else .25,
                popup=None if reference else folium.Popup(popup_html, max_width=420),
                tooltip=tooltip or None,
            ).add_to(m)
            return True
    except Exception:
        return False
    return False


def active_reaches(reaches: Dict[str, Any]) -> List[Dict[str, Any]]:
    out = []
    for f in reaches.get("features", []):
        p = f.get("properties") or {}
        if bool(p.get("active", True)) and geometry_type(f) == "LineString":
            out.append(f)
    return out


def _bootstrap_reaches(store: GitHubRepoStore) -> None:
    path = _reach_path()
    def mutate(cur):
        if cur.get(path) is None:
            return {path: json_dump_bytes(empty_reaches())}, {"created": True}
        return {}, {"created": False}
    store.atomic_update([path], mutate, "初始化治理現況底圖")


def _load_reaches(store: GitHubRepoStore) -> Dict[str, Any]:
    commit, _ = store.get_head_commit()
    return json_load_bytes(store.read_file(_reach_path(), ref=commit, allow_missing=True), empty_reaches())


def _next_rch_id(reg: Dict[str, Any]) -> str:
    counters = reg.setdefault("counters", {})
    counters["RCH"] = int(counters.get("RCH", 0) or 0) + 1
    return f"RCH-{counters['RCH']:07d}"


def save_reach_feature(
    svc: EngineeringGISService,
    drawing: Dict[str, Any],
    status: str,
    water_name: str,
    reach_name: str,
    county: str,
    basis: str,
    notes: str,
    line_weight: int,
    editor: str,
    reach_id: str = "",
) -> Dict[str, Any]:
    if geometry_type(drawing) != "LineString":
        raise ValueError("治理現況底圖目前只接受『線』。請用折線工具繪製河段。")
    path = _reach_path()
    paths = [path, svc.cfg.registry_path]

    def mutate(cur):
        reaches = json_load_bytes(cur.get(path), empty_reaches())
        reg = json_load_bytes(cur.get(svc.cfg.registry_path), empty_registry())
        reg.setdefault("counters", {}).setdefault("RCH", 0)
        rid = reach_id.strip()
        target = None
        if rid:
            for f in reaches.get("features", []):
                if (f.get("properties") or {}).get("reach_id") == rid:
                    target = f
                    break
            if target is None:
                raise ValueError(f"找不到治理現況底圖 {rid}")
        else:
            rid = _next_rch_id(reg)
            target = {"type": "Feature", "geometry": None, "properties": {}}
            reaches.setdefault("features", []).append(target)
            target["properties"]["created_at"] = now_iso()

        target["geometry"] = copy.deepcopy(drawing.get("geometry"))
        prop = target.setdefault("properties", {})
        prop.update({
            "reach_id": rid,
            "status": status,
            "water_name": water_name.strip(),
            "reach_name": reach_name.strip(),
            "county": county.strip(),
            "basis": basis,
            "notes": notes.strip(),
            "line_weight": int(line_weight),
            "source": "manual",
            "active": True,
            "updated_at": now_iso(),
            "updated_by": editor,
        })
        changed = {path: json_dump_bytes(reaches), svc.cfg.registry_path: json_dump_bytes(reg)}
        return changed, {"reach_id": rid}

    action = "更新" if reach_id else "新增"
    return svc.store.atomic_update(paths, mutate, f"{action}治理現況底圖：{water_name or reach_name or reach_id}｜{editor}")


def save_reach_features_batch(
    svc: EngineeringGISService,
    drawings: Sequence[Dict[str, Any]],
    water_name: str,
    reach_name: str,
    county: str,
    basis: str,
    notes: str,
    keywords: str,
    line_weight: int,
    editor: str,
    existing_reach_id: str = "",
    existing_reach_ids: Optional[Sequence[str]] = None,
) -> Dict[str, Any]:
    """一次儲存治理現況底圖的多個線段，且每段可有自己的治理狀態。

    V3.6.60 起允許先把藍色線以「未指定現況」儲存；之後重新載入該 RCH，
    再逐段補登「尚待治理」或「已完成治理」。每一段仍是獨立 RCH feature。
    """
    lines = [d for d in _normalize_drawings(drawings) if geometry_type(d) == "LineString"]
    if not lines:
        raise ValueError("目前沒有可儲存的治理現況線。")

    prepared = []
    for d in lines:
        props = d.get("properties") or {}
        status = str(props.get("reach_status_draft") or props.get("status") or "").strip()
        # 舊草稿的未指定狀態是空字串；正式寫入時改存明確值，方便後續查詢與補登。
        if status not in REACH_STATUSES:
            status = REACH_STATUS_UNASSIGNED
        prepared.append({
            "drawing": d,
            "status": status,
            "draft_reach_id": str(props.get("reach_id_draft") or props.get("reach_id") or "").strip(),
        })

    path = _reach_path()
    paths = [path, svc.cfg.registry_path]

    def mutate(cur):
        reaches = json_load_bytes(cur.get(path), empty_reaches())
        reg = json_load_bytes(cur.get(svc.cfg.registry_path), empty_registry())
        reg.setdefault("counters", {}).setdefault("RCH", 0)
        active_by_id = {}
        for f in reaches.get("features", []):
            rp = f.get("properties") or {}
            rid = str(rp.get("reach_id") or "")
            if rid:
                active_by_id[rid] = f

        original_ids = []
        for x in list(existing_reach_ids or []) + ([existing_reach_id] if existing_reach_id else []):
            rid0 = display_text(x)
            if rid0 and rid0 not in original_ids:
                original_ids.append(rid0)
        used_ids = set()
        created_ids = []
        updated_ids = []
        deactivated_ids = []

        # 單一既有 RCH 的相容邏輯：若使用者重畫導致 properties 沒保留 ID，
        # 第一段仍沿用該 RCH。多段共同編輯時則完全依各 drawing 的 reach_id_draft，
        # 避免誤把其中一段硬指定成別段 ID。
        if len(original_ids) == 1 and not any(x["draft_reach_id"] == original_ids[0] for x in prepared):
            prepared[0]["draft_reach_id"] = original_ids[0]

        for item in prepared:
            d = item["drawing"]
            status = item["status"]
            rid = item["draft_reach_id"]
            # 截斷後兩段可能從同一原始線繼承相同 ID；只允許第一段沿用，其他段建立新 ID。
            if rid and rid in used_ids:
                rid = ""
            target = active_by_id.get(rid) if rid else None
            if target is None:
                rid = _next_rch_id(reg)
                target = {"type": "Feature", "geometry": None, "properties": {"created_at": now_iso()}}
                reaches.setdefault("features", []).append(target)
                active_by_id[rid] = target
                created_ids.append(rid)
            else:
                updated_ids.append(rid)
            used_ids.add(rid)

            target["geometry"] = copy.deepcopy(d.get("geometry"))
            prop = target.setdefault("properties", {})
            prop.update({
                "reach_id": rid,
                "status": status,
                "water_name": water_name.strip(),
                "reach_name": reach_name.strip(),
                "county": county.strip(),
                "basis": basis,
                "notes": notes.strip(),
                "keywords": normalize_keywords(keywords),
                "line_weight": int(line_weight),
                "source": "manual",
                "active": True,
                "updated_at": now_iso(),
                "updated_by": editor,
            })

        # V3.6.25：多段既有 RCH 一起編輯／合併時，所有原始 RCH 都納入版本集合。
        # 若某個原始 RCH 已不在儲存後的 used_ids（例如兩段合併成一段），就停用舊線，
        # 避免總覽同時殘留合併前與合併後圖資。
        for original_id in original_ids:
            if original_id and original_id not in used_ids and original_id in active_by_id:
                oldp = active_by_id[original_id].setdefault("properties", {})
                oldp["active"] = False
                oldp["updated_at"] = now_iso()
                oldp["updated_by"] = editor
                deactivated_ids.append(original_id)

        changed = {path: json_dump_bytes(reaches), svc.cfg.registry_path: json_dump_bytes(reg)}
        return changed, {
            "created_reach_ids": created_ids,
            "updated_reach_ids": updated_ids,
            "deactivated_reach_ids": deactivated_ids,
            "saved_count": len(prepared),
            "unassigned_count": sum(
                1 for item in prepared if item["status"] == REACH_STATUS_UNASSIGNED
            ),
        }

    return svc.store.atomic_update(
        paths, mutate,
        f"儲存治理現況分段：{water_name or reach_name or existing_reach_id or '多段既有河段'}｜{len(prepared)}段｜{editor}"
    )


def deactivate_reach(svc: EngineeringGISService, reach_id: str, editor: str) -> Dict[str, Any]:
    path = _reach_path()
    def mutate(cur):
        reaches = json_load_bytes(cur.get(path), empty_reaches())
        found = False
        for f in reaches.get("features", []):
            p = f.get("properties") or {}
            if p.get("reach_id") == reach_id and p.get("active", True):
                p["active"] = False
                p["deleted_at"] = now_iso()
                p["deleted_by"] = editor
                found = True
        if not found:
            raise ValueError("找不到可停用的治理現況底圖。")
        return {path: json_dump_bytes(reaches)}, {"reach_id": reach_id}
    return svc.store.atomic_update([path], mutate, f"停用治理現況底圖：{reach_id}｜{editor}")



def _bootstrap_project_drafts(store: GitHubRepoStore) -> None:
    path = _project_draft_path()
    def mutate(cur):
        if cur.get(path) is None:
            return {path: json_dump_bytes(empty_project_drafts())}, {"created": True}
        return {}, {"created": False}
    store.atomic_update([path], mutate, "初始化待確認ID工程圖資草稿")


def _load_project_drafts(store: GitHubRepoStore) -> Dict[str, Any]:
    commit, _ = store.get_head_commit()
    return json_load_bytes(store.read_file(_project_draft_path(), ref=commit, allow_missing=True), empty_project_drafts())


def active_project_drafts(drafts: Dict[str, Any], row_key: str) -> List[Dict[str, Any]]:
    out = []
    for f in drafts.get("features", []):
        p = f.get("properties") or {}
        if p.get("row_key") == row_key and p.get("active", True):
            out.append(f)
    return out


def save_pending_project_drafts(
    svc: EngineeringGISService,
    row: ProjectRow,
    drawings: Sequence[Dict[str, Any]],
    geo_name: str,
    keywords: str,
    line_weight: int,
    editor: str,
) -> Dict[str, Any]:
    import uuid
    drawings2 = [copy.deepcopy(d) for d in drawings if geometry_type(d) in {"Point", "LineString", "Polygon"}]
    if not drawings2:
        raise ValueError("沒有可儲存的點、線或面草稿。")
    path = _project_draft_path()
    def mutate(cur):
        drafts = json_load_bytes(cur.get(path), empty_project_drafts())
        for f in drafts.get("features", []):
            p = f.get("properties") or {}
            if p.get("row_key") == row.row_key and p.get("active", True):
                p["active"] = False
                p["superseded_at"] = now_iso()
                p["superseded_by"] = editor
        ids = []
        for i, d in enumerate(drawings2, 1):
            did = "DRF-" + uuid.uuid4().hex[:12].upper()
            ids.append(did)
            nm = geo_name.strip() or f"待ID圖資草稿{i}"
            drafts.setdefault("features", []).append({
                "type": "Feature",
                "geometry": copy.deepcopy(d.get("geometry")),
                "properties": {
                    "draft_id": did,
                    "row_key": row.row_key,
                    "sheet_name": row.sheet_name,
                    "project_name": row.project_name,
                    "county": row.county,
                    "unit": row.unit,
                    "geo_name": nm,
                    "keywords": normalize_keywords(keywords),
                    "line_weight": int(line_weight),
                    "active": True,
                    "created_at": now_iso(),
                    "updated_by": editor,
                },
            })
        return {path: json_dump_bytes(drafts)}, {"draft_ids": ids}
    return svc.store.atomic_update([path], mutate, f"儲存待確認ID工程圖資草稿：{row.project_name}｜{editor}")


def promote_project_drafts(svc: EngineeringGISService, row_to_pid: Dict[str, Tuple[str, str]], editor: str) -> Dict[str, Any]:
    """row_to_pid: row_key -> (project_id, project_name). 將待ID草稿轉為正式 GEO。"""
    if not row_to_pid:
        return {"promoted": 0}
    dpath = _project_draft_path()
    paths = [dpath, svc.cfg.geo_path, svc.cfg.registry_path]
    def mutate(cur):
        drafts = json_load_bytes(cur.get(dpath), empty_project_drafts())
        geo = json_load_bytes(cur.get(svc.cfg.geo_path), empty_geojson())
        reg = json_load_bytes(cur.get(svc.cfg.registry_path), empty_registry())
        count = 0
        for f in drafts.get("features", []):
            p = f.get("properties") or {}
            rk = p.get("row_key")
            if rk not in row_to_pid or not p.get("active", True):
                continue
            pid, name = row_to_pid[rk]
            gid = next_geo_id(reg)
            nf = {
                "type": "Feature",
                "geometry": copy.deepcopy(f.get("geometry")),
                "properties": {
                    "geo_id": gid,
                    "project_id": pid,
                    "project_name": name,
                    "geo_name": p.get("geo_name") or "分併標前草稿",
                    "keywords": normalize_keywords(p.get("keywords")),
                    "role": "manual_geometry",
                    "source": "manual_draft_promoted",
                    "feature_active": True,
                    "project_active": True,
                    "line_weight": int(p.get("line_weight", 6) or 6),
                    "created_at": now_iso(),
                    "updated_at": now_iso(),
                    "updated_by": editor,
                    "from_draft_id": p.get("draft_id"),
                },
            }
            geo.setdefault("features", []).append(nf)
            p["active"] = False
            p["promoted_at"] = now_iso()
            p["promoted_to_project_id"] = pid
            p["promoted_to_geo_id"] = gid
            count += 1
        return {
            dpath: json_dump_bytes(drafts),
            svc.cfg.geo_path: json_dump_bytes(geo),
            svc.cfg.registry_path: json_dump_bytes(reg),
        }, {"promoted": count}
    return svc.store.atomic_update(paths, mutate, f"正式化待ID工程圖資草稿｜{editor}")


def save_project_drawings_v35(
    svc: EngineeringGISService,
    project_id: str,
    project_name: str,
    drawings: Sequence[Dict[str, Any]],
    geo_name: str,
    keywords: str,
    line_weight: int,
    editor: str,
    replace_manual: bool,
) -> Dict[str, Any]:
    allowed = {"Point", "LineString", "Polygon"}
    drawings2 = [copy.deepcopy(d) for d in drawings if geometry_type(d) in allowed]
    # V3.6.36：允許「原本有正式人工圖資，但本次草稿已刪成 0 筆」作為正式清空。
    # 空草稿只有在 replace_manual=True 時才有意義；新增模式／其他流程仍拒絕空內容。
    if not drawings2 and not replace_manual:
        raise ValueError("沒有可儲存的點、線或面。")

    def mutate(cur):
        geo = json_load_bytes(cur[svc.cfg.geo_path], empty_geojson())
        hist = json_load_bytes(cur[svc.cfg.history_path], empty_history())
        reg = json_load_bytes(cur[svc.cfg.registry_path], empty_registry())
        excel_bytes = cur[svc.cfg.excel_path]
        _, rows, _ = scan_workbook(excel_bytes, svc.cfg.excel_path)
        sync_registry_with_existing_ids(reg, rows, geo, hist)
        known = workbook_to_project_map(rows)
        if project_id not in known:
            raise ValueError(f"目前工程資料庫找不到 {project_id}。")

        deactivated = []
        if replace_manual:
            # V3.6.27：直接編輯器儲存時，以目前畫面為該工程最新人工圖資。
            # 舊版資料的 source/role 可能不同，因此不能只停用 source==manual。
            for f in geo.get("features", []):
                props0 = f.get("properties") or {}
                if display_text(props0.get("project_id")) != display_text(project_id):
                    continue
                if not props0.get("feature_active", True) or not props0.get("project_active", True):
                    continue
                gt0 = geometry_type(f)
                if gt0 not in {"Point", "LineString", "Polygon"}:
                    continue
                role0 = display_text(props0.get("role")).lower()
                source0 = display_text(props0.get("source")).lower()
                if role0 == "original_point" or source0 in {"excel_auto", "gis_representative"}:
                    continue
                props0["feature_active"] = False
                props0["superseded_at"] = now_iso()
                props0["superseded_by"] = editor
                if props0.get("geo_id"):
                    deactivated.append(str(props0.get("geo_id")))

        # V3.6.36：若目前草稿為空，代表使用者要正式清除既有人工圖資。
        # 若 GitHub 最新版本也已經沒有可停用圖資，則停止，避免產生無意義 commit。
        if not drawings2:
            if not deactivated:
                raise ValueError("目前沒有可正式清除的工程人工圖資。")
            return {
                svc.cfg.geo_path: json_dump_bytes(geo),
                svc.cfg.registry_path: json_dump_bytes(reg),
            }, {
                "created_geo_ids": [],
                "deactivated_geo_ids": deactivated,
                "cleared_all": True,
            }

        created = []
        for idx, d in enumerate(drawings2, 1):
            gid = next_geo_id(reg)
            nm = geo_name.strip() or f"人工圖資{idx}"
            if len(drawings2) > 1 and geo_name.strip():
                nm = f"{geo_name.strip()}-{idx}"
            feature = {
                "type": "Feature",
                "geometry": copy.deepcopy(d.get("geometry")),
                "properties": {
                    "geo_id": gid,
                    "project_id": project_id,
                    "project_name": project_name,
                    "geo_name": nm,
                    "keywords": normalize_keywords(keywords),
                    "role": "manual_geometry",
                    "source": "manual",
                    "feature_active": True,
                    "project_active": True,
                    "line_weight": int(line_weight),
                    "created_at": now_iso(),
                    "updated_at": now_iso(),
                    "updated_by": editor,
                },
            }
            geo.setdefault("features", []).append(feature)
            created.append(gid)
        return {
            svc.cfg.geo_path: json_dump_bytes(geo),
            svc.cfg.registry_path: json_dump_bytes(reg),
        }, {
            "created_geo_ids": created,
            "deactivated_geo_ids": deactivated,
            "cleared_all": False,
        }

    return svc.store.atomic_update(svc.paths, mutate, f"更新工程圖資：{project_id}｜{editor}")


def _normalize_drawings(drawings: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out = []
    for d in drawings or []:
        if not isinstance(d, dict):
            continue
        gt = geometry_type(d)
        if gt not in {"Point", "LineString", "Polygon"}:
            continue
        out.append({"type": "Feature", "geometry": copy.deepcopy(d.get("geometry")), "properties": copy.deepcopy(d.get("properties") or {})})
    return out


def _drawing_signature(drawings: Sequence[Dict[str, Any]]) -> str:
    # 幾何以外，治理現況底圖的「逐段治理狀態」也必須納入草稿版本判斷。
    # 否則只改尚待治理／已完成治理、但線形沒有改變時，Undo/Redo 會誤判成沒有變更。
    prop_keys = [
        "reach_status_draft", "reach_id_draft",
        "edit_operation", "edit_operation_at", "split_visual_group", "split_role",
        "split_point_lat", "split_point_lon", "split_history",
        "extract_start_lat", "extract_start_lon", "extract_end_lat", "extract_end_lon",
    ]
    clean = []
    for d in _normalize_drawings(drawings):
        props = d.get("properties") or {}
        clean.append({
            "type": "Feature",
            "geometry": d.get("geometry"),
            "properties": {k: props.get(k) for k in prop_keys if k in props},
        })
    return json.dumps(clean, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _draft_init(key: str, baseline: Sequence[Dict[str, Any]]) -> None:
    state_key = f"gis_draft::{key}"
    baseline2 = _normalize_drawings(baseline)
    if state_key not in st.session_state:
        st.session_state[state_key] = {
            "baseline": copy.deepcopy(baseline2),
            "snapshots": [copy.deepcopy(baseline2)],
            "index": 0,
            # V3.6.17：一般瀏覽器編輯事件不得改變 st_folium key。
            # 只有 Undo/Redo 這類「Python 主動要求重畫」才增加 render_epoch。
            "render_epoch": 0,
        }
    else:
        # 相容使用者從舊版升級後尚留在 Session 的草稿狀態。
        st.session_state[state_key].setdefault("render_epoch", 0)


def _draft_state(key: str) -> Dict[str, Any]:
    return st.session_state[f"gis_draft::{key}"]


def _draft_current(key: str) -> List[Dict[str, Any]]:
    s = _draft_state(key)
    return copy.deepcopy(s["snapshots"][s["index"]])


def _draft_push(key: str, drawings: Sequence[Dict[str, Any]]) -> bool:
    """把瀏覽器回傳的圖形推進草稿歷程；有實際變更才回傳 True。

    V3.6.18：回傳布林值可讓 st_folium on_change 與主流程 fallback
    共用同一套去重判斷，避免同一個瀏覽器事件被推進兩次。
    """
    s = _draft_state(key)
    new = _normalize_drawings(drawings)
    cur = s["snapshots"][s["index"]]
    if _drawing_signature(new) == _drawing_signature(cur):
        return False
    s["snapshots"] = s["snapshots"][: s["index"] + 1]
    s["snapshots"].append(copy.deepcopy(new))
    if len(s["snapshots"]) > 51:  # baseline + 50 次操作
        s["snapshots"] = s["snapshots"][-51:]
    s["index"] = len(s["snapshots"]) - 1
    return True


def _sync_folium_component_to_draft(component_key: str, draft_key: str) -> None:
    """在 st_folium 的 on_change callback 階段先同步 all_drawings。

    streamlit-folium 會先把最新 component value 寫入 st.session_state[key]，
    再執行 on_change。利用這個時機先更新 Python 草稿，下一輪 rerun 建地圖前
    就能拿到最新線段，避免「第一次操作被吃掉、第二次才出現」。
    """
    payload = st.session_state.get(component_key)
    if not isinstance(payload, dict):
        return
    drawings = payload.get("all_drawings")
    if drawings is None:
        return
    try:
        _draft_push(draft_key, drawings)
    except Exception:
        return


def _draft_undo(key: str) -> None:
    s = _draft_state(key)
    if s["index"] > 0:
        s["index"] -= 1
        # Undo 是後端主動切換快照，需要強制 Leaflet 重新載入該快照。
        s["render_epoch"] = int(s.get("render_epoch", 0) or 0) + 1


def _draft_redo(key: str) -> None:
    s = _draft_state(key)
    if s["index"] < len(s["snapshots"]) - 1:
        s["index"] += 1
        # Redo 同理，只在這裡改變元件 key；一般畫線事件不改 key。
        s["render_epoch"] = int(s.get("render_epoch", 0) or 0) + 1




def _draft_clear(key: str) -> None:
    st.session_state.pop(f"gis_draft::{key}", None)


@st.dialog("確認刪除治理現況底圖")
def _confirm_delete_reach_dialog(svc: EngineeringGISService, reach_id: str, editor: str):
    st.warning(f"確定要停用 {reach_id} 嗎？")
    st.caption("此操作會從目前地圖隱藏該河段；GitHub commit 歷史仍保留舊版本，但網站不提供歷史恢復入口。")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("取消", key=f"cancel_del_{reach_id}", use_container_width=True):
            st.rerun()
    with c2:
        if st.button("確定停用", key=f"confirm_del_{reach_id}", type="primary", use_container_width=True):
            try:
                _save_mask = _saving_overlay("正在停用治理現況河段…")
                try:
                    deactivate_reach(svc, reach_id, editor)
                finally:
                    _close_saving_overlay(_save_mask)
                st.success("已停用。")
                st.rerun()
            except Exception as exc:
                st.error(str(exc))




@st.dialog("確認正式清除工程圖資")
def _confirm_clear_project_drawings_dialog(
    svc: EngineeringGISService,
    project_id: str,
    project_name: str,
    geo_name: str,
    keywords: str,
    line_weight: int,
    editor: str,
    draft_key: str,
):
    st.warning(f"確定要正式清除「{project_name}」目前所有人工工程圖資嗎？")
    st.caption(
        "目前編輯草稿已沒有任何點、線或範圍。確認後，系統會把這個工程原本已儲存的人工 GIS 圖資全部停用；"
        "工程原始代表點不會刪除，GitHub commit 歷史仍會保留舊版本。"
    )
    c1, c2 = st.columns(2)
    with c1:
        if st.button("取消", key=f"cancel_clear_project::{project_id}", use_container_width=True):
            st.rerun()
    with c2:
        if st.button(
            "🗑️ 確定正式清除",
            key=f"confirm_clear_project::{project_id}",
            type="primary",
            use_container_width=True,
        ):
            try:
                _save_mask = _saving_overlay("正在正式清除工程圖資…")
                try:
                    res = save_project_drawings_v35(
                        svc, project_id, project_name, [], geo_name, keywords, line_weight, editor,
                        replace_manual=True,
                    )
                finally:
                    _close_saving_overlay(_save_mask)
                count = len(res.get("deactivated_geo_ids") or [])
                _draft_clear(draft_key)
                st.session_state["gis_project_clear_notice"] = (
                    f"{project_name} 已正式清除 {count} 筆人工工程圖資；工程原始代表點仍保留。"
                )
                st.rerun()
            except Exception as exc:
                st.error(f"正式清除失敗：{exc}")



def _latest_line_edit(drawings: Sequence[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """取得本次草稿最近一次線段處理的視覺提示資訊。

    edit_operation_* 僅存在 Streamlit Session 草稿；正式儲存工程／治理圖資時
    save_project_drawings_v35 / save_reach_feature 會重新建立正式 properties，
    因此這些標記不會寫進 GitHub 正式圖資。
    """
    candidates: List[Dict[str, Any]] = []
    for d in _normalize_drawings(drawings):
        if geometry_type(d) != "LineString":
            continue
        props = d.get("properties") or {}
        op = str(props.get("edit_operation") or "")
        at = str(props.get("edit_operation_at") or "")
        if op in {"split", "extract_interval", "merge"} and at:
            candidates.append({"drawing": d, "props": props, "operation": op, "at": at})
    if not candidates:
        return None
    latest_at = max(x["at"] for x in candidates)
    latest = [x for x in candidates if x["at"] == latest_at]
    if not latest:
        return None
    op = latest[0]["operation"]
    group = str(latest[0]["props"].get("split_visual_group") or "")
    if op == "split" and group:
        same_group = []
        for d in _normalize_drawings(drawings):
            props = d.get("properties") or {}
            if props.get("edit_operation") == "split" and str(props.get("split_visual_group") or "") == group:
                same_group.append({"drawing": d, "props": props, "operation": "split", "at": str(props.get("edit_operation_at") or "")})
        latest = same_group or latest
    return {"operation": op, "at": latest_at, "items": latest, "group": group}


def _line_midpoint_latlon(drawing: Dict[str, Any]) -> Optional[Tuple[float, float]]:
    coords = ((drawing.get("geometry") or {}).get("coordinates") or [])
    if len(coords) < 2:
        return None
    # 以節點間平面距離近似尋找沿線中點；標籤用途不需做投影轉換。
    segs: List[float] = []
    total = 0.0
    for i in range(len(coords) - 1):
        try:
            x1, y1 = float(coords[i][0]), float(coords[i][1])
            x2, y2 = float(coords[i + 1][0]), float(coords[i + 1][1])
        except Exception:
            segs.append(0.0)
            continue
        dist = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
        segs.append(dist)
        total += dist
    if total <= 0:
        try:
            x, y = coords[len(coords) // 2][:2]
            return float(y), float(x)
        except Exception:
            return None
    target = total / 2.0
    acc = 0.0
    for i, dist in enumerate(segs):
        if acc + dist >= target and dist > 0:
            ratio = (target - acc) / dist
            x1, y1 = float(coords[i][0]), float(coords[i][1])
            x2, y2 = float(coords[i + 1][0]), float(coords[i + 1][1])
            return y1 + (y2 - y1) * ratio, x1 + (x2 - x1) * ratio
        acc += dist
    try:
        x, y = coords[-1][:2]
        return float(y), float(x)
    except Exception:
        return None


def _edit_badge(text: str, bg: str = "#1D4ED8", fg: str = "#FFFFFF") -> folium.DivIcon:
    safe = str(text).replace("<", "&lt;").replace(">", "&gt;")
    return folium.DivIcon(
        html=(
            f'<div style="transform:translate(-50%,-50%);white-space:nowrap;'
            f'background:{bg};color:{fg};border:2px solid #ffffff;border-radius:999px;'
            f'box-shadow:0 1px 5px rgba(0,0,0,.45);padding:3px 7px;font-size:12px;'
            f'font-weight:800;pointer-events:none;">{safe}</div>'
        ),
        icon_size=(1, 1),
        icon_anchor=(0, 0),
    )


def _scissor_icon(label: str = "✂") -> folium.DivIcon:
    safe = str(label).replace("<", "&lt;").replace(">", "&gt;")
    return folium.DivIcon(
        html=(
            '<div style="transform:translate(-50%,-50%);width:34px;height:34px;'
            'display:flex;align-items:center;justify-content:center;background:#ffffff;'
            'border:3px solid #DC2626;border-radius:50%;box-shadow:0 2px 7px rgba(0,0,0,.45);'
            f'font-size:17px;font-weight:900;color:#B91C1C;pointer-events:none;">{safe}</div>'
        ),
        icon_size=(1, 1),
        icon_anchor=(0, 0),
    )



def _split_history_points(drawings: Sequence[Dict[str, Any]]) -> List[Tuple[float, float]]:
    """蒐集本次草稿所有截斷點；支援舊 split_point_* 與 V3.6.23 split_history。"""
    pts: List[Tuple[float, float]] = []
    seen = set()
    for d in _normalize_drawings(drawings):
        if geometry_type(d) != "LineString":
            continue
        props = d.get("properties") or {}
        history = props.get("split_history") or []
        if isinstance(history, list):
            for item in history:
                try:
                    lat = float((item or {}).get("lat")); lon = float((item or {}).get("lon"))
                    key = (round(lat, 8), round(lon, 8))
                    if key not in seen:
                        seen.add(key); pts.append((lat, lon))
                except Exception:
                    continue
        if props.get("edit_operation") == "split":
            try:
                lat = float(props.get("split_point_lat")); lon = float(props.get("split_point_lon"))
                key = (round(lat, 8), round(lon, 8))
                if key not in seen:
                    seen.add(key); pts.append((lat, lon))
            except Exception:
                pass
    return pts


def _add_edit_visual_overlays(map_obj: folium.Map, drawings: Sequence[Dict[str, Any]]) -> None:
    """在地圖上加編輯提示。V3.6.23 會保留本次草稿『所有』截斷剪刀；A/B 標籤只標最近一次。"""
    for lat, lon in _split_history_points(drawings):
        folium.Marker([lat, lon], icon=_scissor_icon("✂"), tooltip="截斷位置（本次草稿）").add_to(map_obj)

    latest = _latest_line_edit(drawings)
    if not latest:
        return
    op = latest["operation"]
    items = latest["items"]

    if op == "split":
        roles: Dict[str, Dict[str, Any]] = {}
        split_lat = split_lon = None
        for item in items:
            props = item["props"]
            role = str(props.get("split_role") or "")
            if role in {"A", "B"}:
                roles[role] = item["drawing"]
            try:
                split_lat = float(props.get("split_point_lat"))
                split_lon = float(props.get("split_point_lon"))
            except Exception:
                pass
        # 只有 A/B 兩段都仍存在時才顯示截斷完成提示，避免刪除其中一段後留下假標記。
        if "A" not in roles or "B" not in roles:
            return
        for role in ("A", "B"):
            mid = _line_midpoint_latlon(roles[role])
            if mid:
                label = "A段｜實線" if role == "A" else "B段｜虛線"
                folium.Marker(mid, icon=_edit_badge(label), tooltip=f"截斷後 {role} 段").add_to(map_obj)

    elif op == "extract_interval":
        item = items[0]
        props = item["props"]
        pts = []
        for prefix, label in (("extract_start", "①✂"), ("extract_end", "②✂")):
            try:
                lat = float(props.get(f"{prefix}_lat"))
                lon = float(props.get(f"{prefix}_lon"))
                pts.append((lat, lon, label))
            except Exception:
                continue
        for lat, lon, label in pts:
            folium.Marker([lat, lon], icon=_scissor_icon(label), tooltip="截取邊界（僅編輯提示）").add_to(map_obj)
        mid = _line_midpoint_latlon(item["drawing"])
        if mid:
            folium.Marker(mid, icon=_edit_badge("保留區間", bg="#0F766E"), tooltip="截取後保留的線段").add_to(map_obj)


def _render_latest_edit_notice(drawings: Sequence[Dict[str, Any]]) -> None:
    latest = _latest_line_edit(drawings)
    if not latest:
        return
    op = latest["operation"]
    if op == "split":
        roles = {str((x.get("props") or {}).get("split_role") or "") for x in latest["items"]}
        if {"A", "B"}.issubset(roles):
            st.success("✂️ 截斷完成：紅色 ✂ 是截斷點；A段為藍色實線，B段為藍色虛線。這些標示只供本次編輯辨識，正式儲存後不會留在圖資中。")
    elif op == "extract_interval":
        st.success("✂️ 截取完成：①✂ 與 ②✂ 是兩個截取邊界，中間標示『保留區間』。這些標示只供本次編輯辨識。")
    elif op == "merge":
        st.success("🔗 合併完成：兩條相鄰線段已合併。若結果不符預期，可按下方「↶ 上一步」。")


class _AttachFeatureProperties(MacroElement):
    """把 Python 草稿 properties 掛回 Leaflet path layer。

    Folium.PolyLine 預設不會把任意 properties 放進 layer.feature；但治理現況
    的逐段狀態、既有 RCH-ID 與截斷提示都需要跟著 all_drawings 回到 Streamlit。
    """
    _template = Template(r"""
    {% macro script(this, kwargs) %}
    try {
        var layer = {{ this._parent.get_name() }};
        layer.feature = layer.feature || {type:'Feature', properties:{}, geometry:null};
        layer.feature.properties = {{ this.properties_json | safe }};
    } catch(e) {}
    {% endmacro %}
    """)

    def __init__(self, properties: Dict[str, Any]):
        super().__init__()
        self._name = "AttachFeatureProperties"
        self.properties_json = json.dumps(properties or {}, ensure_ascii=False)


def _feature_group_from_drawings(drawings: Sequence[Dict[str, Any]], name: str = "編輯中圖資", reach_status_edit: bool = False) -> folium.FeatureGroup:
    fg = folium.FeatureGroup(name=name, show=True)
    normalized = _normalize_drawings(drawings)
    latest = _latest_line_edit(normalized)
    latest_at = str((latest or {}).get("at") or "")
    for d in normalized:
        geom = d.get("geometry") or {}
        gt = geom.get("type")
        c = geom.get("coordinates")
        props = d.get("properties") or {}
        try:
            if gt == "Point" and c:
                folium.CircleMarker(
                    [float(c[1]), float(c[0])], radius=7, color="#000000", weight=2.5,
                    fill=True, fill_color="#1976D2", fill_opacity=.95, tooltip="可編輯點位"
                ).add_to(fg)
            elif gt == "LineString" and c:
                line_color = "#1976D2"
                tooltip = None
                if reach_status_edit:
                    draft_status = str(props.get("reach_status_draft") or props.get("status") or "").strip()
                    if draft_status == "尚待治理":
                        line_color = REACH_PENDING_COLOR
                        tooltip = "尚待治理"
                    elif draft_status == "已完成治理":
                        line_color = REACH_COMPLETED_DETAIL_COLOR
                        tooltip = "已完成治理"
                    else:
                        line_color = "#1976D2"
                        tooltip = "未指定現況（待確認）"
                kwargs: Dict[str, Any] = {"color": line_color, "weight": 6, "opacity": 0.95}
                # 只把『最近一次截斷』的 B 段畫成虛線；避免歷次截斷全部變成虛線造成混亂。
                if (
                    latest_at
                    and str(props.get("edit_operation_at") or "") == latest_at
                    and props.get("edit_operation") == "split"
                    and props.get("split_role") == "B"
                ):
                    kwargs["dash_array"] = "12 8"
                line = folium.PolyLine([[float(y), float(x)] for x, y, *_ in c], tooltip=tooltip, **kwargs)
                line.add_child(_AttachFeatureProperties(props))
                line.add_to(fg)
            elif gt == "Polygon" and c and c[0]:
                folium.Polygon([[float(y), float(x)] for x, y, *_ in c[0]], color="#000000", weight=4.5, fill=True, fill_color="#1976D2", fill_opacity=.18).add_to(fg)
        except Exception:
            continue
    return fg


def _center_for_features(features: Sequence[Dict[str, Any]]) -> Tuple[List[float], int]:
    centers = [feature_center(f) for f in features]
    centers = [c for c in centers if c]
    if not centers:
        return TAIWAN_CENTER, TAIWAN_ZOOM
    return [float(centers[0][0]), float(centers[0][1])], 15


def _render_draft_controls(key: str) -> None:
    s = _draft_state(key)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("↶ 上一步", key=f"undo_{key}", disabled=s["index"] <= 0, use_container_width=True):
            _draft_undo(key)
            st.rerun()
    with c2:
        if st.button("↷ 下一步", key=f"redo_{key}", disabled=s["index"] >= len(s["snapshots"]) - 1, use_container_width=True):
            _draft_redo(key)
            st.rerun()
    st.caption(f"本次暫存：第 {s['index']} 步／最多 50 步。Undo/Redo 只存在目前瀏覽 Session，不寫入 GitHub。")




class BrowserViewportMemory(MacroElement):
    """V3.6.59：同時記住地圖視角及該地圖在 Streamlit 頁面的可視位置。

    必須先恢復再掛上 moveend 監聽，否則地圖重建時的初始 fitBounds 可能反過來
    覆蓋使用者剛才的視角。點選標記與繪圖事件會在 Streamlit rerun 前，同步保存
    iframe 相對於瀏覽器視窗的位置，重建後據此恢復整個頁面的垂直捲動位置。
    """

    _template = Template(r"""
    {% macro script(this, kwargs) %}
    (function() {
        var mapObj = {{ this._parent.get_name() }};
        var storageKey = {{ this.storage_key_json | safe }};
        var pageIntentKey = 'wra_gis_page_scroll_intent_v3659';
        if (!storageKey) return;
        function storageObj() {
            try { if (window.parent && window.parent.sessionStorage) return window.parent.sessionStorage; } catch(e) {}
            return window.sessionStorage;
        }
        function safeGet() {
            try {
                var raw = storageObj().getItem(storageKey);
                if (!raw) return null;
                var obj = JSON.parse(raw);
                if (!obj || !isFinite(obj.lat) || !isFinite(obj.lng) || !isFinite(obj.zoom)) return null;
                return obj;
            } catch(e) { return null; }
        }
        function pageSnapshot() {
            try {
                var p = window.parent;
                var frame = window.frameElement;
                var y = Number(p.scrollY || p.pageYOffset || 0);
                var top = frame ? Number(frame.getBoundingClientRect().top) : NaN;
                return {
                    page_y: isFinite(y) ? y : 0,
                    frame_top: isFinite(top) ? top : null
                };
            } catch(e) { return {page_y:0, frame_top:null}; }
        }
        function safeSave(markPageIntent) {
            try {
                var c = mapObj.getCenter(), z = mapObj.getZoom();
                var page = pageSnapshot();
                var payload = {
                    lat:Number(c.lat), lng:Number(c.lng), zoom:Number(z), ts:Date.now(),
                    page_y:Number(page.page_y), frame_top:page.frame_top
                };
                storageObj().setItem(storageKey, JSON.stringify(payload));
                if (markPageIntent === true) {
                    storageObj().setItem(pageIntentKey, JSON.stringify({
                        y:Number(page.page_y), frame_top:page.frame_top,
                        view_key:storageKey, source:'folium-map', ts:Date.now()
                    }));
                    try {
                        storageObj().setItem('wra_gis_page_scroll_v3659', JSON.stringify({
                            y:Number(page.page_y), ts:Date.now()
                        }));
                    } catch(_e) {}
                }
            } catch(e) {}
        }
        function readPageIntent() {
            try {
                var raw = storageObj().getItem(pageIntentKey);
                var obj = raw ? JSON.parse(raw) : null;
                if (!obj || obj.view_key !== storageKey || !isFinite(obj.y) ||
                    !isFinite(obj.ts) || Date.now()-Number(obj.ts) >= 8000) return null;
                return obj;
            } catch(e) { return null; }
        }
        var pageRestoreScheduled = false;
        function restorePageOnce(intent) {
            try {
                var p = window.parent;
                var currentY = Number(p.scrollY || p.pageYOffset || 0);
                var targetY = Number(intent.y);
                var frame = window.frameElement;
                if (frame && intent.frame_top !== null && isFinite(intent.frame_top)) {
                    var currentTop = Number(frame.getBoundingClientRect().top);
                    if (isFinite(currentTop)) {
                        targetY = currentY + currentTop - Number(intent.frame_top);
                    }
                }
                targetY = Math.max(0, targetY);
                try { p.scrollTo({top:targetY,left:0,behavior:'auto'}); }
                catch(e) { try { p.scrollTo(0,targetY); } catch(_e) {} }
            } catch(e) {}
        }
        function schedulePageRestore() {
            if (pageRestoreScheduled) return false;
            var intent = readPageIntent();
            if (!intent) return false;
            pageRestoreScheduled = true;
            [0,70,220,500,900,1450,2050].forEach(function(ms) {
                setTimeout(function(){ restorePageOnce(intent); }, ms);
            });
            return true;
        }
        function saveInteraction() { safeSave(true); }
        var saved = safeGet();
        // 先恢復視角，再開始監聽；避免初始 fitBounds/moveend 覆寫既有記憶。
        if (saved) {
            try { mapObj.setView([saved.lat, saved.lng], saved.zoom, {animate:false}); } catch(e) {}
        }
        schedulePageRestore();
        // 讓編輯工具在送出 draw 事件前主動保存地圖視角與頁面錨點。
        mapObj.__wraSaveViewport = saveInteraction;
        mapObj.__wraRestorePageAnchor = schedulePageRestore;
        mapObj.on('moveend', saveInteraction);
        mapObj.on('click', saveInteraction);
        mapObj.on('draw:created', saveInteraction);
        mapObj.on('draw:edited', saveInteraction);
        mapObj.on('draw:deleted', saveInteraction);
        function attachClickSaver(layer) {
            try {
                if (layer instanceof L.CircleMarker || layer instanceof L.Marker) {
                    if (!layer.__wraViewportSaverBound) {
                        layer.__wraViewportSaverBound = true;
                        layer.on('click', saveInteraction);
                    }
                }
                if (layer.eachLayer) layer.eachLayer(attachClickSaver);
            } catch(e) {}
        }
        try { mapObj.eachLayer(attachClickSaver); } catch(e) {}
        mapObj.on('layeradd', function(ev){ attachClickSaver(ev && ev.layer); });
        // capture 階段先記住視角，確保點位 click 導致元件重建前已寫入。
        try {
            var container = mapObj.getContainer();
            if (container && container.addEventListener) {
                container.addEventListener('pointerdown', saveInteraction, true);
            }
        } catch(e) {}
        if (!saved) setTimeout(function(){ safeSave(false); }, 0);
    })();
    {% endmacro %}
    """)

    def __init__(self, memory_key: str):
        super().__init__()
        self._name = "BrowserViewportMemory"
        self.storage_key_json = json.dumps(
            "wra_gis_view_v3618::" + str(memory_key or "default"), ensure_ascii=False
        )


class VisibleDrawToolbar(MacroElement):
    """固定顯示的中文 Leaflet Draw 工具列。

    除既有畫線／畫點／畫範圍／節點編輯／刪除外，加入線段工具：
    - 截斷線：在線上點一下，把 LineString 分成兩段，並以紅色✂＋A/B段清楚標示。
    - 截取區間：在同一條線上點兩下，只保留兩點之間的區間。
    - 合併相鄰線段：依序點兩條線，若最近端點在容許距離內即合併。
    - 治理現況模式：先畫／複製／截斷；藍線可先存為未指定現況，再補登治理狀態。

    這些操作只改本次地圖 FeatureGroup，透過 draw:edited 交回
    streamlit-folium；真正寫入 GitHub 仍由頁面下方「儲存工程圖資」控制。
    """

    _template = Template(r"""
    {% macro script(this, kwargs) %}
    (function() {
        var mapObj = {{ this._parent.get_name() }};
        var editGroup = {{ this.feature_group_name }};
        var activeEdit = null;
        var activeDelete = null;
        var activeDraw = null;

        // 線段工具狀態 -------------------------------------------------
        var lineToolMode = null;       // split | extract | merge | overlap | status_pending | status_completed | null
        var selectedLine = null;
        var extractFirst = null;
        var mergeFirstLine = null;
        var tempLayer = L.layerGroup().addTo(mapObj);
        // resultLayer 只顯示本次操作的剪刀／A段B段等視覺提示，不屬於 editable FeatureGroup。
        var resultLayer = L.layerGroup().addTo(mapObj);
        var toolStatus = null;
        var toolButtons = {};
        var baseLineColor = '#3388ff';
        var pendingLineColor = '#D32F2F';
        var completedLineColor = '#1B5E20';
        var selectedColor = '#d97706';
        var mergeToleranceM = {{ this.merge_tolerance_m }};
        var outlinePaneName='wra-edit-outline-pane';
        try {
            if(!mapObj.getPane(outlinePaneName)) {
                var opane=mapObj.createPane(outlinePaneName);
                opane.style.zIndex='395';
                opane.style.pointerEvents='none';
            }
        } catch(e) {}
        var outlineGroup=L.layerGroup().addTo(mapObj);
        var toolMemoryKey = {{ this.tool_memory_key_json | safe }};
        var editorModeMemoryKey = toolMemoryKey ? (toolMemoryKey + '::editor_mode') : '';
        var parentScrollMemoryKey = toolMemoryKey ? (toolMemoryKey + '::parent_scroll') : '';
        // V3.6.37：自訂線段工具採短暫批次同步，避免每點一下就立即造成 Streamlit rerun。
        var changedTimer = null;
        var changedPending = false;

        function parentStorage() {
            try { if (window.parent && window.parent.sessionStorage) return window.parent.sessionStorage; } catch(e) {}
            return window.sessionStorage;
        }
        function saveMapViewport() {
            try {
                if (typeof mapObj.__wraSaveViewport === 'function') mapObj.__wraSaveViewport();
            } catch(e) {}
        }
        function saveParentScroll() {
            // V3.6.59：繪圖／節點／刪除／線段處理觸發 rerun 前，
            // 同步保存地圖視角及 iframe 在頁面可視範圍中的位置。
            saveMapViewport();
            if (!parentScrollMemoryKey) return;
            try {
                var p = window.parent;
                var y = Number(p.scrollY || p.pageYOffset || 0);
                var frame = window.frameElement;
                var frameTop = frame ? Number(frame.getBoundingClientRect().top) : NaN;
                var payload = {
                    y:(isFinite(y)?y:0),
                    frame_top:(isFinite(frameTop)?frameTop:null),
                    ts:Date.now()
                };
                parentStorage().setItem(parentScrollMemoryKey, JSON.stringify(payload));
                try { parentStorage().setItem('wra_gis_page_scroll_v3659', JSON.stringify({y:payload.y,ts:payload.ts})); } catch(_e) {}
            } catch(e) {}
        }
        function restoreParentScroll() {
            // BrowserViewportMemory 已掌握這張地圖的精確 iframe 錨點時，由它統一恢復。
            try {
                if (typeof mapObj.__wraRestorePageAnchor === 'function') {
                    mapObj.__wraRestorePageAnchor();
                    return;
                }
            } catch(e) {}
            if (!parentScrollMemoryKey) return;
            try {
                var raw = parentStorage().getItem(parentScrollMemoryKey);
                var saved = null;
                try { saved = raw ? JSON.parse(raw) : null; } catch(_e) {}
                if (!saved || !isFinite(saved.y)) {
                    var legacyY = Number(raw);
                    if (!isFinite(legacyY)) return;
                    saved = {y:legacyY,frame_top:null};
                }
                var p = window.parent;
                [90, 260, 620, 1050].forEach(function(ms){
                    setTimeout(function(){
                        var targetY = Number(saved.y);
                        try {
                            var frame = window.frameElement;
                            if (frame && saved.frame_top !== null && isFinite(saved.frame_top)) {
                                var currentY = Number(p.scrollY || p.pageYOffset || 0);
                                var currentTop = Number(frame.getBoundingClientRect().top);
                                if (isFinite(currentY) && isFinite(currentTop)) {
                                    targetY = currentY + currentTop - Number(saved.frame_top);
                                }
                            }
                        } catch(e) {}
                        targetY = Math.max(0,targetY);
                        try { p.scrollTo({top:targetY,left:0,behavior:'auto'}); }
                        catch(e){ try { p.scrollTo(0,targetY); } catch(_e){} }
                    }, ms);
                });
            } catch(e) {}
        }
        function rememberEditorMode(mode) {
            if (!editorModeMemoryKey) return;
            try {
                if (mode) parentStorage().setItem(editorModeMemoryKey, String(mode));
                else parentStorage().removeItem(editorModeMemoryKey);
            } catch(e) {}
        }
        function rememberedEditorMode() {
            if (!editorModeMemoryKey) return '';
            try { return String(parentStorage().getItem(editorModeMemoryKey) || ''); }
            catch(e) { return ''; }
        }

        function rememberToolMode(mode) {
            if (!toolMemoryKey) return;
            try {
                if (mode) parentStorage().setItem(toolMemoryKey, String(mode));
                else parentStorage().removeItem(toolMemoryKey);
            } catch(e) {}
        }
        function rememberedToolMode() {
            if (!toolMemoryKey) return '';
            try { return String(parentStorage().getItem(toolMemoryKey) || ''); }
            catch(e) { return ''; }
        }

        function stopTool(tool, saveFirst) {
            if (!tool) return null;
            try { if (saveFirst && tool.save) tool.save(); } catch(e) {}
            try { if (tool.disable) tool.disable(); } catch(e) {}
            return null;
        }

        function finishLeafletTools() {
            activeEdit = stopTool(activeEdit, true);
            activeDelete = stopTool(activeDelete, true);
            activeDraw = stopTool(activeDraw, false);
        }

        function deepClone(obj) {
            try { return JSON.parse(JSON.stringify(obj || {})); }
            catch(e) { return {}; }
        }

        function layerProperties(layer) {
            try {
                if (layer && layer.feature && layer.feature.properties) {
                    return deepClone(layer.feature.properties);
                }
            } catch(e) {}
            return {};
        }

        function clearEditVisualProps(props) {
            var p = deepClone(props || {});
            [
                'split_visual_group','split_role','split_point_lat','split_point_lon',
                'extract_start_lat','extract_start_lon','extract_end_lat','extract_end_lon'
            ].forEach(function(k){ try { delete p[k]; } catch(e) {} });
            return p;
        }

        function makeBadgeIcon(text, bg) {
            var safe = String(text || '').replace(/</g,'&lt;').replace(/>/g,'&gt;');
            return L.divIcon({
                className:'', iconSize:[1,1], iconAnchor:[0,0],
                html:'<div style="transform:translate(-50%,-50%);white-space:nowrap;background:' + (bg || '#1D4ED8') + ';color:#fff;border:2px solid #fff;border-radius:999px;box-shadow:0 1px 5px rgba(0,0,0,.45);padding:3px 7px;font-size:12px;font-weight:800;pointer-events:none;">' + safe + '</div>'
            });
        }

        function makeScissorIcon(text) {
            var safe = String(text || '✂').replace(/</g,'&lt;').replace(/>/g,'&gt;');
            return L.divIcon({
                className:'', iconSize:[1,1], iconAnchor:[0,0],
                html:'<div style="transform:translate(-50%,-50%);width:34px;height:34px;display:flex;align-items:center;justify-content:center;background:#fff;border:3px solid #DC2626;border-radius:50%;box-shadow:0 2px 7px rgba(0,0,0,.45);font-size:17px;font-weight:900;color:#B91C1C;pointer-events:none;">' + safe + '</div>'
            });
        }

        function approximateLineCenter(layer) {
            try {
                var ll = flatLatLngs(layer);
                if (!ll || !ll.length) return null;
                if (ll.length === 1) return ll[0];
                var total = 0, segs = [];
                for (var i=0; i<ll.length-1; i++) {
                    var d = mapObj.distance(ll[i], ll[i+1]);
                    segs.push(d); total += d;
                }
                if (total <= 0) return ll[Math.floor(ll.length/2)];
                var target = total / 2, acc = 0;
                for (var j=0; j<segs.length; j++) {
                    if (acc + segs[j] >= target && segs[j] > 0) {
                        var r = (target - acc) / segs[j];
                        return L.latLng(
                            ll[j].lat + (ll[j+1].lat - ll[j].lat) * r,
                            ll[j].lng + (ll[j+1].lng - ll[j].lng) * r
                        );
                    }
                    acc += segs[j];
                }
                return ll[ll.length-1];
            } catch(e) { return null; }
        }

        function clearResultVisuals() { try { resultLayer.clearLayers(); } catch(e) {} }

        function splitHistoryFromProps(props) {
            var arr=[];
            try {
                var src=(props || {}).split_history || [];
                if (Array.isArray(src)) src.forEach(function(x){
                    var lat=Number(x && x.lat), lon=Number(x && x.lon);
                    if (isFinite(lat) && isFinite(lon)) arr.push({lat:lat,lon:lon});
                });
            } catch(e) {}
            return arr;
        }
        function mergedSplitHistory(propsA, propsB) {
            // V3.6.38：先合併兩邊的截斷歷史；真正合併時再移除已被消除的接合點剪刀。
            var out=[], seen={};
            function addOne(x){
                var lat=Number(x && x.lat), lon=Number(x && x.lon);
                if(!isFinite(lat)||!isFinite(lon)) return;
                var k=lat.toFixed(8)+'|'+lon.toFixed(8);
                if(!seen[k]){ seen[k]=true; out.push({lat:lat,lon:lon}); }
            }
            [propsA || {}, propsB || {}].forEach(function(p){
                splitHistoryFromProps(p).forEach(addOne);
                if(p.edit_operation==='split') addOne({lat:p.split_point_lat, lon:p.split_point_lon});
            });
            return out;
        }

        function collectAllSplitHistory() {
            var out=[], seen={};
            try {
                editGroup.eachLayer(function(layer){
                    var p=layerProperties(layer);
                    splitHistoryFromProps(p).forEach(function(x){
                        var k=x.lat.toFixed(8)+'|'+x.lon.toFixed(8);
                        if(!seen[k]){seen[k]=true;out.push(x);}
                    });
                    if(p.edit_operation==='split'){
                        var lat=Number(p.split_point_lat), lon=Number(p.split_point_lon);
                        if(isFinite(lat)&&isFinite(lon)){
                            var k=lat.toFixed(8)+'|'+lon.toFixed(8);
                            if(!seen[k]){seen[k]=true;out.push({lat:lat,lon:lon});}
                        }
                    }
                });
            }catch(e){}
            return out;
        }
        function renderAllScissors() {
            try {
                collectAllSplitHistory().forEach(function(x){
                    L.marker([x.lat,x.lon], {icon:makeScissorIcon('✂'), interactive:false}).addTo(resultLayer);
                });
            } catch(e) {}
        }

        function showSplitVisual(splitPoint, a, b) {
            clearResultVisuals();
            renderAllScissors();
            try {
                var ca = approximateLineCenter(a), cb = approximateLineCenter(b);
                if (ca) L.marker(ca, {icon:makeBadgeIcon('A段｜實線','#1D4ED8'), interactive:false}).addTo(resultLayer);
                if (cb) L.marker(cb, {icon:makeBadgeIcon('B段｜虛線','#1D4ED8'), interactive:false}).addTo(resultLayer);
            } catch(e) {}
        }

        function showExtractVisual(p1, p2, line) {
            clearResultVisuals();
            renderAllScissors();
            try {
                L.marker(p1, {icon:makeScissorIcon('①✂'), interactive:false}).addTo(resultLayer);
                L.marker(p2, {icon:makeScissorIcon('②✂'), interactive:false}).addTo(resultLayer);
                var c = approximateLineCenter(line);
                if (c) L.marker(c, {icon:makeBadgeIcon('保留區間','#0F766E'), interactive:false}).addTo(resultLayer);
            } catch(e) {}
        }

        function statusColorForProps(props) {
            var s = String((props || {}).reach_status_draft || (props || {}).status || '');
            if (s === '尚待治理') return pendingLineColor;
            if (s === '已完成治理') return completedLineColor;
            return baseLineColor;
        }

        function makeEditableLine(latlngs, props) {
            var p = deepClone(props || {});
            var ln = L.polyline(latlngs, {
                color: statusColorForProps(p),
                weight: {{ this.line_weight }},
                opacity: 0.95
            });
            ln.feature = {
                type: 'Feature',
                properties: p,
                geometry: null
            };
            return ln;
        }

        function isLineLayer(layer) {
            return !!(layer && layer.getLatLngs && !(layer instanceof L.Polygon));
        }

        function flatLatLngs(layer) {
            if (!isLineLayer(layer)) return [];
            var ll = layer.getLatLngs() || [];
            // LineString 應是一層陣列；若 Leaflet 回巢狀，取第一條。
            while (Array.isArray(ll) && ll.length && Array.isArray(ll[0])) ll = ll[0];
            return ll || [];
        }

        function removeOutline(layer){
            try { if(layer && layer.__wraOutline){ outlineGroup.removeLayer(layer.__wraOutline); layer.__wraOutline=null; } } catch(e) {}
        }
        function syncOutline(layer, selected){
            try {
                if(!layer) return;
                var white=!!selected && (layer instanceof L.CircleMarker) && !(layer instanceof L.Circle);
                var col=white ? '#FFFFFF' : '#000000';
                if(layer instanceof L.CircleMarker && !(layer instanceof L.Circle)){
                    if(!layer.__wraOutline){
                        layer.__wraOutline=L.circleMarker(layer.getLatLng(), {pane:outlinePaneName,radius:(layer.getRadius?layer.getRadius():7)+2.5,color:col,weight:3,fill:false,interactive:false}).addTo(outlineGroup);
                    } else {
                        layer.__wraOutline.setLatLng(layer.getLatLng());
                        layer.__wraOutline.setRadius((layer.getRadius?layer.getRadius():7)+2.5);
                        layer.__wraOutline.setStyle({color:col,weight:3});
                    }
                    return;
                }
                if(layer instanceof L.Polygon){
                    if(!layer.__wraOutline){
                        layer.__wraOutline=L.polygon(layer.getLatLngs(), {pane:outlinePaneName,color:col,weight:Math.max(5,{{ this.line_weight }}+2.5),fill:false,interactive:false,opacity:1}).addTo(outlineGroup);
                    } else { layer.__wraOutline.setLatLngs(layer.getLatLngs()); layer.__wraOutline.setStyle({color:col}); }
                    return;
                }
                if(isLineLayer(layer)){
                    if(!layer.__wraOutline){
                        layer.__wraOutline=L.polyline(layer.getLatLngs(), {pane:outlinePaneName,color:col,weight:{{ this.line_weight }}+3.5,interactive:false,opacity:1}).addTo(outlineGroup);
                    } else { layer.__wraOutline.setLatLngs(layer.getLatLngs()); layer.__wraOutline.setStyle({color:col,weight:{{ this.line_weight }}+3.5}); }
                }
            }catch(e){}
        }
        function syncAllOutlines(){ try{ editGroup.eachLayer(function(l){syncOutline(l,false);}); }catch(e){} }

        function setLineStyle(layer, selected) {
            try {
                if (layer && layer.setStyle) {
                    var props = layerProperties(layer);
                    var dash = (props.edit_operation === 'split' && props.split_role === 'B') ? '12 8' : null;
                    layer.setStyle({
                        color: selected ? selectedColor : statusColorForProps(props),
                        weight: selected ? Math.max({{ this.line_weight }} + 2, 8) : {{ this.line_weight }},
                        opacity: 0.95,
                        dashArray: dash
                    });
                    syncOutline(layer, !!selected);
                }
            } catch(e) {}
        }

        function setReachStatus(layer, reachStatus) {
            if (!layer || !isLineLayer(layer)) return false;
            var props = layerProperties(layer);
            props.reach_status_draft = reachStatus;
            layer.feature = layer.feature || {type:'Feature', properties:{}, geometry:null};
            layer.feature.properties = deepClone(props);
            setLineStyle(layer, false);
            // 治理狀態模式跨 rerun 保留；按一次紅/綠後可連續點多段。
            rememberToolMode(lineToolMode);
            setStatus((reachStatus === '尚待治理' ? '🔴 ' : '🟢 ') + '已設定此線段；可直接繼續點其他線段，或切換另一個治理狀態。', false);
            scheduleChanged(850);
            return true;
        }

        function resetHighlights() {
            try {
                if (selectedLine) setLineStyle(selectedLine, false);
                if (mergeFirstLine && mergeFirstLine !== selectedLine) setLineStyle(mergeFirstLine, false);
            } catch(e) {}
            selectedLine = null;
            mergeFirstLine = null;
        }

        function setStatus(msg, isError) {
            if (!toolStatus) return;
            toolStatus.textContent = msg || '';
            toolStatus.style.color = isError ? '#b91c1c' : '#475569';
        }

        function resetButtonStyles() {
            Object.keys(toolButtons).forEach(function(k) {
                var b = toolButtons[k];
                if (!b) return;
                b.style.background = '#ffffff';
                b.style.fontWeight = '400';
            });
        }

        function cancelLineTool(message) {
            lineToolMode = null;
            extractFirst = null;
            tempLayer.clearLayers();
            resetHighlights();
            resetButtonStyles();
            rememberToolMode('');
            if (message) setStatus(message, false);
        }

        function activateLineTool(mode, button, message, quietBroadcast) {
            saveParentScroll();
            rememberEditorMode('');
            finishLeafletTools();
            cancelLineTool('');
            clearResultVisuals();
            renderAllScissors();
            lineToolMode = mode;
            rememberToolMode(mode);
            if (button) {
                button.style.background = '#fef3c7';
                button.style.fontWeight = '700';
            }
            setStatus(message, false);
            if (!quietBroadcast) {
                try { mapObj.fire('wra:editor-tool-start', {mode:mode}); } catch(e) {}
            }
        }

        function fireChangedNow() {
            if (changedTimer) { try { clearTimeout(changedTimer); } catch(e) {} changedTimer = null; }
            changedPending = false;
            saveParentScroll();
            try { mapObj.fire('draw:edited', {layers: editGroup}); } catch(e) {}
        }

        function scheduleChanged(delayMs) {
            // 多個連續狀態指定／線段操作只在停止操作一小段時間後送回 Python，
            // 可明顯減少 iframe/整頁反覆重建；最後送出的仍是整個 editable FeatureGroup。
            changedPending = true;
            if (changedTimer) { try { clearTimeout(changedTimer); } catch(e) {} }
            var wait = Math.max(120, Number(delayMs || 450));
            changedTimer = setTimeout(function(){ fireChangedNow(); }, wait);
        }

        function flushChanged() {
            if (changedPending) fireChangedNow();
        }

        function latLngEquals(a, b) {
            if (!a || !b) return false;
            return Math.abs(a.lat - b.lat) < 1e-11 && Math.abs(a.lng - b.lng) < 1e-11;
        }

        function appendUnique(arr, p) {
            if (!p) return;
            if (!arr.length || !latLngEquals(arr[arr.length - 1], p)) {
                arr.push(L.latLng(p.lat, p.lng));
            }
        }

        function projectionOnLine(layer, clickLatLng) {
            var latlngs = flatLatLngs(layer);
            if (!latlngs || latlngs.length < 2) return null;
            var clickPt = mapObj.latLngToLayerPoint(clickLatLng);
            var best = null;
            var cumulative = 0;
            for (var i = 0; i < latlngs.length - 1; i++) {
                var a = mapObj.latLngToLayerPoint(latlngs[i]);
                var b = mapObj.latLngToLayerPoint(latlngs[i + 1]);
                var vx = b.x - a.x, vy = b.y - a.y;
                var wx = clickPt.x - a.x, wy = clickPt.y - a.y;
                var len2 = vx * vx + vy * vy;
                var t = len2 > 0 ? (wx * vx + wy * vy) / len2 : 0;
                t = Math.max(0, Math.min(1, t));
                var qx = a.x + t * vx, qy = a.y + t * vy;
                var dx = clickPt.x - qx, dy = clickPt.y - qy;
                var d2 = dx * dx + dy * dy;
                var segLen = Math.sqrt(len2);
                if (!best || d2 < best.d2) {
                    best = {
                        d2: d2,
                        segIndex: i,
                        t: t,
                        latlng: mapObj.layerPointToLatLng(L.point(qx, qy)),
                        alongPx: cumulative + t * segLen
                    };
                }
                cumulative += segLen;
            }
            if (best) best.totalPx = cumulative;
            return best;
        }

        function validInteriorProjection(prj) {
            if (!prj || !prj.totalPx || prj.totalPx <= 0) return false;
            // 避免剛好切在整條線的起終點，產生零長度線段。
            var edge = Math.max(1.0, prj.totalPx * 0.001);
            return prj.alongPx > edge && prj.alongPx < (prj.totalPx - edge);
        }

        function removeAndAdd(oldLayers, newLayers) {
            (oldLayers || []).forEach(function(layer) {
                try { editGroup.removeLayer(layer); } catch(e) {}
            });
            (newLayers || []).forEach(function(layer) {
                try { editGroup.addLayer(layer); } catch(e) {}
            });
        }

        function splitLineAt(layer, clickLatLng) {
            var latlngs = flatLatLngs(layer);
            var prj = projectionOnLine(layer, clickLatLng);
            if (!validInteriorProjection(prj)) {
                setStatus('截斷位置太靠近線段端點，請在線段中間再點一次。', true);
                return false;
            }
            var left = [], right = [];
            for (var i = 0; i <= prj.segIndex; i++) appendUnique(left, latlngs[i]);
            appendUnique(left, prj.latlng);
            appendUnique(right, prj.latlng);
            for (var j = prj.segIndex + 1; j < latlngs.length; j++) appendUnique(right, latlngs[j]);
            if (left.length < 2 || right.length < 2) {
                setStatus('截斷失敗：切點造成其中一段沒有足夠節點。', true);
                return false;
            }
            var opAt = (new Date()).toISOString();
            var groupId = 'split_' + Date.now() + '_' + Math.floor(Math.random() * 1000000);
            var baseProps = clearEditVisualProps(layerProperties(layer));
            var propsA = deepClone(baseProps), propsB = deepClone(baseProps);
            [propsA, propsB].forEach(function(p){
                p.edit_operation = 'split';
                p.edit_operation_at = opAt;
                p.split_visual_group = groupId;
                p.split_point_lat = Number(prj.latlng.lat);
                p.split_point_lon = Number(prj.latlng.lng);
            });
            var hist = splitHistoryFromProps(layerProperties(layer));
            hist.push({lat:Number(prj.latlng.lat), lon:Number(prj.latlng.lng)});
            // 去重，避免同一截斷點被重複記錄。
            var histSeen={}, hist2=[];
            hist.forEach(function(x){
                var k=Number(x.lat).toFixed(8)+'|'+Number(x.lon).toFixed(8);
                if(!histSeen[k]){histSeen[k]=true;hist2.push({lat:Number(x.lat),lon:Number(x.lon)});}
            });
            propsA.split_history = deepClone(hist2);
            propsB.split_history = deepClone(hist2);
            propsA.split_role = 'A';
            propsB.split_role = 'B';
            // 既有治理河段截斷後，只讓 A 段沿用原 RCH-ID；B 段存檔時建立新 RCH-ID。
            if (propsB.reach_id_draft) propsB.reach_id_draft = '';
            if (propsB.reach_id) propsB.reach_id = '';
            var a = makeEditableLine(left, propsA);
            var b = makeEditableLine(right, propsB);
            try { a.setStyle({dashArray:null}); b.setStyle({dashArray:'12 8'}); } catch(e) {}
            removeAndAdd([layer], [a, b]);
            showSplitVisual(prj.latlng, a, b);
            // 保持截斷模式，rerun 後也會恢復；可連續截斷多個位置。
            selectedLine = null; extractFirst = null; mergeFirstLine = null; tempLayer.clearLayers();
            rememberToolMode('split');
            setStatus('✂️ 截斷完成：剪刀會保留；可直接繼續截斷其他位置，或切換工具。', false);
            scheduleChanged(420);
            return true;
        }

        function extractBetween(layer, p1, p2) {
            if (!p1 || !p2) return false;
            if (Math.abs(p1.alongPx - p2.alongPx) < 2.0) {
                setStatus('兩個截取位置太接近，請把第二點選遠一些。', true);
                return false;
            }
            if (p1.alongPx > p2.alongPx) {
                var tmp = p1; p1 = p2; p2 = tmp;
            }
            var latlngs = flatLatLngs(layer);
            var out = [];
            appendUnique(out, p1.latlng);
            // p1 位於 segment i；p2 位於 segment k。保留中間原始節點。
            for (var i = p1.segIndex + 1; i <= p2.segIndex && i < latlngs.length; i++) {
                appendUnique(out, latlngs[i]);
            }
            appendUnique(out, p2.latlng);
            if (out.length < 2) {
                setStatus('截取失敗：兩個位置之間沒有有效線段。', true);
                return false;
            }
            var props = clearEditVisualProps(layerProperties(layer));
            props.edit_operation = 'extract_interval';
            props.edit_operation_at = (new Date()).toISOString();
            props.extract_start_lat = Number(p1.latlng.lat);
            props.extract_start_lon = Number(p1.latlng.lng);
            props.extract_end_lat = Number(p2.latlng.lat);
            props.extract_end_lon = Number(p2.latlng.lng);
            var only = makeEditableLine(out, props);
            removeAndAdd([layer], [only]);
            showExtractVisual(p1.latlng, p2.latlng, only);
            selectedLine = null; extractFirst = null; mergeFirstLine = null; tempLayer.clearLayers();
            rememberToolMode('extract');
            setStatus('✂️ 截取完成：可直接繼續截取其他線段，或切換工具。', false);
            scheduleChanged(420);
            return true;
        }

        function endpointCandidates(a, b) {
            var aa = flatLatLngs(a), bb = flatLatLngs(b);
            if (aa.length < 2 || bb.length < 2) return [];
            return [
                {aEnd:'start', bEnd:'start', dist:mapObj.distance(aa[0], bb[0])},
                {aEnd:'start', bEnd:'end',   dist:mapObj.distance(aa[0], bb[bb.length-1])},
                {aEnd:'end',   bEnd:'start', dist:mapObj.distance(aa[aa.length-1], bb[0])},
                {aEnd:'end',   bEnd:'end',   dist:mapObj.distance(aa[aa.length-1], bb[bb.length-1])}
            ].sort(function(x,y){ return x.dist-y.dist; });
        }

        function mergeLines(a, b) {
            var candidates = endpointCandidates(a, b);
            if (!candidates.length) {
                setStatus('合併失敗：選取圖形不是有效線段。', true);
                return false;
            }
            var best = candidates[0];
            if (best.dist > mergeToleranceM) {
                setStatus('兩線最近端點相距 ' + best.dist.toFixed(1) + 'm，超過 ' + mergeToleranceM + 'm；請選真正相鄰的線段。', true);
                return false;
            }
            var aa = flatLatLngs(a).slice();
            var bb = flatLatLngs(b).slice();
            // 將 a 的接合端調到尾端，b 的接合端調到前端。
            if (best.aEnd === 'start') aa.reverse();
            if (best.bEnd === 'end') bb.reverse();
            var out = [];
            aa.forEach(function(p){ appendUnique(out, p); });
            bb.forEach(function(p){ appendUnique(out, p); });
            if (out.length < 2) {
                setStatus('合併失敗：線段節點不足。', true);
                return false;
            }
            var rawPropsA = layerProperties(a);
            var propsB = layerProperties(b);
            var props = clearEditVisualProps(rawPropsA);
            // 合併兩段線時保留其他仍有效的截斷點，但移除這次已被合併掉的接合點剪刀。
            var mergedHistory = mergedSplitHistory(rawPropsA, propsB);
            var joinA = aa.length ? aa[aa.length - 1] : null;
            var joinB = bb.length ? bb[0] : null;
            var joinScissorToleranceM = 1.5;
            mergedHistory = mergedHistory.filter(function(x){
                try {
                    var p = L.latLng(Number(x.lat), Number(x.lon));
                    var da = joinA ? mapObj.distance(p, joinA) : Infinity;
                    var db = joinB ? mapObj.distance(p, joinB) : Infinity;
                    // 剪刀若位在本次兩線的接合端，即代表該截斷界線已因合併而不存在。
                    return Math.min(da, db) > joinScissorToleranceM;
                } catch(e) { return true; }
            });
            if (mergedHistory.length) props.split_history = deepClone(mergedHistory);
            else { try { delete props.split_history; } catch(e) {} }
            if (!props.reach_id_draft && propsB.reach_id_draft) props.reach_id_draft = propsB.reach_id_draft;
            if (!props.reach_id && propsB.reach_id) props.reach_id = propsB.reach_id;
            // 若兩段治理狀態不同，合併後回到未指定，要求使用者重新確認狀態。
            var sa = String(props.reach_status_draft || props.status || '');
            var sb = String(propsB.reach_status_draft || propsB.status || '');
            if (sa && sb && sa !== sb) {
                delete props.reach_status_draft;
                delete props.status;
            }
            props.edit_operation = 'merge';
            props.edit_operation_at = (new Date()).toISOString();
            props.merge_endpoint_distance_m = Number(best.dist.toFixed(3));
            var merged = makeEditableLine(out, props);
            removeAndAdd([a, b], [merged]);
            clearResultVisuals();
            renderAllScissors();
            selectedLine = null; extractFirst = null; mergeFirstLine = null; tempLayer.clearLayers();
            rememberToolMode('merge');
            setStatus('🔗 已合併兩條相鄰線段（接合端距離 ' + best.dist.toFixed(1) + 'm）；接合處剪刀已移除，其他仍有效的剪刀會保留。可直接繼續合併其他線段。', false);
            scheduleChanged(420);
            return true;
        }

        function pointSegmentDistancePx(p, a, b) {
            var dx=b.x-a.x, dy=b.y-a.y;
            if (dx===0 && dy===0) {
                var ex=p.x-a.x, ey=p.y-a.y; return Math.sqrt(ex*ex+ey*ey);
            }
            var t=((p.x-a.x)*dx+(p.y-a.y)*dy)/(dx*dx+dy*dy);
            t=Math.max(0,Math.min(1,t));
            var qx=a.x+t*dx, qy=a.y+t*dy;
            var ux=p.x-qx, uy=p.y-qy; return Math.sqrt(ux*ux+uy*uy);
        }
        function layerDistancePx(layer, latlng) {
            try {
                var pts=flatLatLngs(layer);
                if (!pts || pts.length<2 || !latlng) return Infinity;
                var p=mapObj.latLngToLayerPoint(latlng), best=Infinity;
                for (var i=1;i<pts.length;i++) {
                    var a=mapObj.latLngToLayerPoint(pts[i-1]);
                    var b=mapObj.latLngToLayerPoint(pts[i]);
                    best=Math.min(best,pointSegmentDistancePx(p,a,b));
                }
                return best;
            } catch(e) { return Infinity; }
        }
        function overlappingEditableLines(latlng, tolerancePx) {
            var out=[];
            try {
                editGroup.eachLayer(function(l){
                    if(!isLineLayer(l)) return;
                    var d=layerDistancePx(l,latlng);
                    if(isFinite(d) && d<=Number(tolerancePx||10)) out.push({layer:l,dist:d,id:L.stamp(l)});
                });
            } catch(e) {}
            out.sort(function(a,b){ if(Math.abs(a.dist-b.dist)>0.01) return a.dist-b.dist; return a.id-b.id; });
            return out;
        }
        function cycleOverlappingLine(clickedLayer, latlng) {
            var candidates=overlappingEditableLines(latlng,12);
            if(candidates.length<=1) {
                setStatus('此位置只偵測到 1 條可編輯線，沒有被遮住的下層線。', false);
                try { if(clickedLayer && clickedLayer.bringToFront) clickedLayer.bringToFront(); } catch(e) {}
                resetHighlights();
                if(clickedLayer) setLineStyle(clickedLayer,true);
                return;
            }
            var clickedId=L.stamp(clickedLayer), idx=-1;
            for(var i=0;i<candidates.length;i++) if(candidates[i].id===clickedId){idx=i;break;}
            // 目前點得到的是最上層線；第一次就切到另一條，之後每點一次繼續輪替。
            var target=candidates[(idx+1+candidates.length)%candidates.length].layer;
            resetHighlights();
            try { if(target.bringToFront) target.bringToFront(); } catch(e) {}
            setLineStyle(target,true);
            setStatus('🧭 此位置共有 ' + candidates.length + ' 條重疊線；已把下一條移到最上層並以橘色標示。可再點一次繼續切換，找到目標後再選截斷／合併／治理狀態或節點編輯。', false);
        }

        function handleEditableLineClick(layer, ev) {
            if (!lineToolMode) return;
            try {
                if (ev && ev.originalEvent) L.DomEvent.stopPropagation(ev.originalEvent);
            } catch(e) {}

            if (lineToolMode === 'overlap') {
                cycleOverlappingLine(layer, ev ? ev.latlng : null);
                return;
            }
            if (lineToolMode === 'status_pending') {
                setReachStatus(layer, '尚待治理');
                return;
            }
            if (lineToolMode === 'status_completed') {
                setReachStatus(layer, '已完成治理');
                return;
            }

            if (lineToolMode === 'split') {
                splitLineAt(layer, ev.latlng);
                return;
            }

            if (lineToolMode === 'extract') {
                if (!selectedLine) {
                    selectedLine = layer;
                    setLineStyle(layer, true);
                    extractFirst = projectionOnLine(layer, ev.latlng);
                    if (!extractFirst) {
                        cancelLineTool('找不到有效截取位置。');
                        return;
                    }
                    tempLayer.clearLayers();
                    L.circleMarker(extractFirst.latlng, {
                        radius:6, color:'#dc2626', weight:2, fill:true, fillColor:'#ffffff', fillOpacity:1
                    }).addTo(tempLayer);
                    setStatus('已選第一個截取點。請在同一條橘色線上點第二個位置；最後只保留兩點之間。', false);
                    return;
                }
                if (layer !== selectedLine) {
                    setStatus('截取區間的第二點必須在同一條橘色線上。', true);
                    return;
                }
                var second = projectionOnLine(layer, ev.latlng);
                extractBetween(layer, extractFirst, second);
                return;
            }

            if (lineToolMode === 'merge') {
                if (!mergeFirstLine) {
                    mergeFirstLine = layer;
                    setLineStyle(layer, true);
                    setStatus('已選第一條線。請再點第二條相鄰藍線；最近端點需在 ' + mergeToleranceM + 'm 內。', false);
                    return;
                }
                if (layer === mergeFirstLine) {
                    setStatus('請選另一條相鄰線段，不要重複點同一條。', true);
                    return;
                }
                mergeLines(mergeFirstLine, layer);
            }
        }

        function attachLineHandler(layer) {
            if (!isLineLayer(layer) || layer.__wraLineToolBound) return;
            layer.__wraLineToolBound = true;
            layer.on('click', function(ev){ handleEditableLineClick(layer, ev); });
        }

        try {
            editGroup.eachLayer(function(layer){ attachLineHandler(layer); syncOutline(layer,false); });
            editGroup.on('layeradd', function(ev){ saveMapViewport(); attachLineHandler(ev.layer); syncOutline(ev.layer,false); });
            editGroup.on('layerremove', function(ev){ saveMapViewport(); removeOutline(ev.layer); });
            mapObj.on('draw:edited', function(){ saveParentScroll(); setTimeout(syncAllOutlines,0); });
            mapObj.on('draw:created', function(){ saveParentScroll(); setTimeout(syncAllOutlines,0); });
            mapObj.on('draw:deleted', function(){ saveParentScroll(); setTimeout(syncAllOutlines,0); });
            setTimeout(restoreParentScroll, 40);
        } catch(e) {}

        // OSM 點選搜尋開始時，明確關閉上一個線段工具，避免 listener/mode 殘留。
        try {
            mapObj.on('wra:osm-pick-start', function(){
                saveParentScroll();
                rememberEditorMode('');
                finishLeafletTools();
                cancelLineTool('已切換為 OSM 搜尋模式。');
            });
            mapObj.on('wra:editor-tool-start', function(ev){
                if (ev && ev.mode === 'osm-copy') {
                    saveParentScroll();
                    rememberEditorMode('');
                    finishLeafletTools();
                    cancelLineTool('已切換為 OSM 複製模式。');
                }
            });
        } catch(e) {}

        var Ctrl = L.Control.extend({
            options:{position:'topleft'},
            onAdd:function(map) {
                var box = L.DomUtil.create('div','');
                box.style.background='rgba(255,255,255,.97)';
                box.style.border='1px solid #94a3b8';
                box.style.borderRadius='6px';
                box.style.padding='6px';
                box.style.marginTop='6px';
                box.style.minWidth='146px';
                box.style.maxWidth='190px';
                box.style.boxShadow='0 1px 5px rgba(0,0,0,.28)';
                box.style.fontSize='12px';
                box.style.lineHeight='1.2';
                box.style.color='#0f172a';
                box.style.webkitTextFillColor='#0f172a';
                box.style.maxHeight='70%';
                box.style.overflowY='auto';
                L.DomEvent.disableClickPropagation(box);
                L.DomEvent.disableScrollPropagation(box);

                var title=L.DomUtil.create('div','',box);
                title.textContent='✏️ 圖資編輯工具';
                title.style.fontWeight='700';
                title.style.marginBottom='5px';
                title.style.textAlign='center';
                title.style.color='#0f172a';

                function btn(label, handler) {
                    var b=L.DomUtil.create('button','',box);
                    b.type='button';
                    b.textContent=label;
                    b.style.display='block';
                    b.style.width='100%';
                    b.style.margin='3px 0';
                    b.style.padding='5px 6px';
                    b.style.border='1px solid #cbd5e1';
                    b.style.borderRadius='4px';
                    b.style.background='#ffffff';
                    b.style.color='#0f172a';
                    b.style.cursor='pointer';
                    b.style.fontSize='12px';
                    b.onclick=function(ev){ L.DomEvent.stop(ev); handler(b); };
                    return b;
                }

                var drawLineBtn = btn('＋ 畫線', function(){
                    saveParentScroll();
                    rememberEditorMode('draw_line');
                    cancelLineTool('');
                    finishLeafletTools();
                    try { mapObj.fire('wra:editor-tool-start', {mode:'draw'}); } catch(e) {}
                    if (L.Draw && L.Draw.Polyline) {
                        activeDraw=new L.Draw.Polyline(mapObj,{shapeOptions:{color:'#3388ff',weight:{{ this.line_weight }},opacity:.95}});
                        activeDraw.enable();
                    }
                });

                {% if this.allow_marker %}
                var drawPointBtn = btn('＋ 畫點', function(){
                    saveParentScroll();
                    rememberEditorMode('draw_point');
                    cancelLineTool('');
                    finishLeafletTools();
                    try { mapObj.fire('wra:editor-tool-start', {mode:'draw'}); } catch(e) {}
                    if (L.Draw && L.Draw.CircleMarker) {
                        activeDraw=new L.Draw.CircleMarker(mapObj,{shapeOptions:{radius:7,color:'#000000',weight:2.5,fill:true,fillColor:'#3388ff',fillOpacity:.95}});
                        activeDraw.enable();
                    } else if (L.Draw && L.Draw.Marker) {
                        activeDraw=new L.Draw.Marker(mapObj,{});
                        activeDraw.enable();
                    }
                });
                {% endif %}

                {% if this.allow_polygon %}
                var drawPolygonBtn = btn('＋ 畫範圍', function(){
                    saveParentScroll();
                    rememberEditorMode('draw_polygon');
                    cancelLineTool('');
                    finishLeafletTools();
                    try { mapObj.fire('wra:editor-tool-start', {mode:'draw'}); } catch(e) {}
                    if (L.Draw && L.Draw.Polygon) {
                        activeDraw=new L.Draw.Polygon(mapObj,{shapeOptions:{color:'#3388ff',weight:4,fillOpacity:.18}});
                        activeDraw.enable();
                    }
                });
                {% endif %}

                var editBtn=btn('✏️ 編輯節點', function(b){
                    saveParentScroll();
                    cancelLineTool('');
                    activeDelete=stopTool(activeDelete,true);
                    activeDraw=stopTool(activeDraw,false);
                    if (!activeEdit) {
                        rememberEditorMode('edit_nodes');
                        if (!editGroup || !editGroup.getLayers || editGroup.getLayers().length===0) {
                            b.textContent='⚠️ 尚無可編輯圖形';
                            setTimeout(function(){b.textContent='✏️ 編輯節點';},1200);
                            return;
                        }
                        activeEdit=new L.EditToolbar.Edit(mapObj,{featureGroup:editGroup,selectedPathOptions:{maintainColor:true}});
                        activeEdit.enable();
                        b.textContent='✅ 完成節點編輯';
                        b.style.background='#dcfce7';
                    } else {
                        // 先清除記憶，再 save；save 會觸發 draw:edited / rerun，避免回來後又自動進入編輯模式。
                        rememberEditorMode('');
                        saveParentScroll();
                        try { if (activeEdit.save) activeEdit.save(); } catch(e) {}
                        try { activeEdit.disable(); } catch(e) {}
                        activeEdit=null;
                        b.textContent='✏️ 編輯節點';
                        b.style.background='#ffffff';
                        // activeEdit.save() 本身已觸發 draw:edited；不要再送第二次事件，避免連續 rerun。
                    }
                });

                var delBtn=btn('🗑 刪除圖形', function(b){
                    saveParentScroll();
                    cancelLineTool('');
                    activeEdit=stopTool(activeEdit,true);
                    activeDraw=stopTool(activeDraw,false);
                    editBtn.textContent='✏️ 編輯節點';
                    editBtn.style.background='#ffffff';
                    if (!activeDelete) {
                        rememberEditorMode('delete_shapes');
                        if (!editGroup || !editGroup.getLayers || editGroup.getLayers().length===0) {
                            b.textContent='⚠️ 尚無可刪除圖形';
                            setTimeout(function(){b.textContent='🗑 刪除圖形';},1200);
                            return;
                        }
                        activeDelete=new L.EditToolbar.Delete(mapObj,{featureGroup:editGroup});
                        activeDelete.enable();
                        b.textContent='✅ 完成刪除';
                        b.style.background='#fee2e2';
                    } else {
                        rememberEditorMode('');
                        saveParentScroll();
                        try { if (activeDelete.save) activeDelete.save(); } catch(e) {}
                        try { activeDelete.disable(); } catch(e) {}
                        activeDelete=null;
                        b.textContent='🗑 刪除圖形';
                        b.style.background='#ffffff';
                        // activeDelete.save() 本身已觸發 draw:deleted；不要再送第二次事件。
                    }
                });

                var sep=L.DomUtil.create('div','',box);
                sep.style.borderTop='1px solid #e2e8f0';
                sep.style.margin='6px 0 4px 0';

                var lineTitle=L.DomUtil.create('div','',box);
                lineTitle.textContent='線段處理';
                lineTitle.style.fontWeight='700';
                lineTitle.style.fontSize='11px';
                lineTitle.style.color='#334155';
                lineTitle.style.marginBottom='2px';

                toolButtons.overlap = btn('🧭 選取重疊線', function(b){
                    activateLineTool('overlap', b, '若多條線重疊，請在重疊位置反覆點擊；系統會逐條移到最上層並以橘色標示。找到目標後再切換其他編輯工具。');
                });
                toolButtons.split = btn('✂️ 截斷線', function(b){
                    activateLineTool('split', b, '請直接點藍線上的截斷位置；完成後會變成兩段。');
                });
                toolButtons.extract = btn('✂️ 截取區間', function(b){
                    activateLineTool('extract', b, '請在同一條藍線上依序點兩個位置；最後只保留兩點之間。');
                });
                toolButtons.merge = btn('🔗 合併相鄰線段', function(b){
                    activateLineTool('merge', b, '請依序點兩條相鄰線；最近端點需在 ' + mergeToleranceM + 'm 內。');
                });

                {% if this.allow_reach_status %}
                var statusSep=L.DomUtil.create('div','',box);
                statusSep.style.borderTop='1px solid #e2e8f0';
                statusSep.style.margin='6px 0 4px 0';
                var statusTitle=L.DomUtil.create('div','',box);
                statusTitle.textContent='治理狀態（截斷後逐段指定）';
                statusTitle.style.fontWeight='700';
                statusTitle.style.fontSize='11px';
                statusTitle.style.color='#334155';
                statusTitle.style.marginBottom='2px';
                toolButtons.status_pending = btn('🔴 設為尚待治理', function(b){
                    activateLineTool('status_pending', b, '請點選要設為「尚待治理」的線段。');
                });
                toolButtons.status_completed = btn('🟢 設為已完成治理', function(b){
                    activateLineTool('status_completed', b, '請點選要設為「已完成治理」的線段。');
                });
                {% endif %}

                btn('✖ 取消線段工具', function(){
                    cancelLineTool('已取消線段處理工具。');
                    clearResultVisuals();
                });

                toolStatus=L.DomUtil.create('div','',box);
                {% if this.allow_reach_status %}
                toolStatus.textContent='先畫／複製整條線，再截斷；重疊線可用「🧭 選取重疊線」逐條切換。藍色＝未指定現況（可先儲存）、紅色＝尚待治理、綠色＝已完成治理。';
                {% else %}
                toolStatus.textContent='OSM 灰線點一下會複製成藍線；重疊藍線可用「🧭 選取重疊線」切換；截斷後會顯示紅色✂。';
                {% endif %}
                toolStatus.style.marginTop='5px';
                toolStatus.style.color='#475569';
                toolStatus.style.whiteSpace='normal';
                toolStatus.style.fontSize='11px';
                toolStatus.style.lineHeight='1.35';
                setTimeout(function(){
                    var remembered = rememberedToolMode();
                    if (remembered === 'overlap' && toolButtons.overlap) {
                        activateLineTool('overlap', toolButtons.overlap, '🧭 重疊線選取模式已恢復；可反覆點重疊位置切換下層線。', true);
                    } else if (remembered === 'split' && toolButtons.split) {
                        activateLineTool('split', toolButtons.split, '✂️ 截斷模式已恢復；可直接繼續點線段。', true);
                    } else if (remembered === 'extract' && toolButtons.extract) {
                        activateLineTool('extract', toolButtons.extract, '✂️ 截取模式已恢復；請重新選第一個截取點。', true);
                    } else if (remembered === 'merge' && toolButtons.merge) {
                        activateLineTool('merge', toolButtons.merge, '🔗 合併模式已恢復；請重新選第一條線。', true);
                    } else if (remembered === 'status_pending' && toolButtons.status_pending) {
                        activateLineTool('status_pending', toolButtons.status_pending, '🔴 尚待治理模式已恢復；可直接繼續點線段。', true);
                    } else if (remembered === 'status_completed' && toolButtons.status_completed) {
                        activateLineTool('status_completed', toolButtons.status_completed, '🟢 已完成治理模式已恢復；可直接繼續點線段。', true);
                    }
                    // V3.6.33：Draw/Edit 事件造成 Streamlit rerun 後，自動恢復剛才的工具，
                    // 讓畫點／線／面可以連續繪製，不必每畫一筆重新按一次。
                    var emode = rememberedEditorMode();
                    try {
                        if (emode === 'draw_line' && drawLineBtn) drawLineBtn.click();
                        {% if this.allow_marker %}
                        else if (emode === 'draw_point' && drawPointBtn) drawPointBtn.click();
                        {% endif %}
                        {% if this.allow_polygon %}
                        else if (emode === 'draw_polygon' && drawPolygonBtn) drawPolygonBtn.click();
                        {% endif %}
                        else if (emode === 'edit_nodes' && editBtn) editBtn.click();
                        else if (emode === 'delete_shapes' && delBtn) delBtn.click();
                    } catch(e) {}
                    restoreParentScroll();
                }, 160);
                return box;
            }
        });
        (new Ctrl()).addTo(mapObj);
    })();
    {% endmacro %}
    """)

    def __init__(self, feature_group: folium.FeatureGroup, line_weight: int = 6,
                 allow_marker: bool = True, allow_polygon: bool = True,
                 merge_tolerance_m: int = 20, allow_reach_status: bool = False,
                 tool_memory_key: str = ""):
        super().__init__()
        self._name = "VisibleDrawToolbar"
        self.feature_group_name = feature_group.get_name()
        self.line_weight = int(line_weight)
        self.allow_marker = bool(allow_marker)
        self.allow_polygon = bool(allow_polygon)
        self.merge_tolerance_m = int(merge_tolerance_m)
        self.allow_reach_status = bool(allow_reach_status)
        self.tool_memory_key_json = json.dumps(
            "wra_gis_tool_v3618::" + str(tool_memory_key or feature_group.get_name()),
            ensure_ascii=False,
        )


class BrowserOSMReferenceControl(MacroElement):
    """在使用者瀏覽器端直接查詢 Overpass 的安全小範圍 OSM 參考線工具。

    V3.6.15：
    - 工程圖資編輯：以工程代表點為中心載入附近 OSM。
    - 治理現況底圖：按「點選位置搜尋附近 OSM」後在地圖點一下。
    - 治理現況底圖可依目前縣市＋水路名稱關鍵字搜尋 OSM。
    - 搜尋半徑只允許 20 / 50 / 100 公尺，預設 50m，底層最大強制 100m。
    - OSM 候選線只存在目前瀏覽器地圖；點灰色虛線後才複製成藍色可編輯草稿。
    - V3.6.43：複製候選線後可從該線兩端各 80m 做「單層延伸」，單次最多 20 條，
      不自動遞迴，避免整個縣市河網查詢拖慢 Streamlit。
    """

    _template = Template(r"""
    {% macro script(this, kwargs) %}
    (function() {
        var mapObj = {{ this._parent.get_name() }};
        var refLayer = L.featureGroup().addTo(mapObj);
        var indicatorLayer = L.featureGroup().addTo(mapObj);
        var busy = false;
        var clickToSelect = {{ this.click_to_select_js }};
        var countyName = {{ this.county_json | safe }};
        var maxCandidates = 100;
        var extensionRadiusM = 80;
        var extensionLimit = 20;
        var osmMemoryKey = {{ this.memory_key_json | safe }};
        var pendingPickHandler = null;
        var osmBusyOverlayId = 'wra-osm-browser-busy-mask';
        var lastExtendAnchor = null;

        function showOsmBusy(message, note) {
            try {
                var p=window.parent, d=p.document;
                var old=d.getElementById(osmBusyOverlayId); if(old) old.remove();
                if(!d.getElementById('wra-osm-busy-style')) {
                    var style=d.createElement('style'); style.id='wra-osm-busy-style';
                    style.textContent='@keyframes wraOsmBusySpin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}';
                    d.head.appendChild(style);
                }
                var mask=d.createElement('div'); mask.id=osmBusyOverlayId;
                mask.style.position='fixed'; mask.style.inset='0'; mask.style.zIndex='2147483645';
                mask.style.background='rgba(15,23,42,.43)'; mask.style.backdropFilter='blur(1.2px)';
                mask.style.webkitBackdropFilter='blur(1.2px)'; mask.style.display='flex';
                mask.style.alignItems='center'; mask.style.justifyContent='center';
                mask.style.pointerEvents='all'; mask.style.cursor='wait';
                var card=d.createElement('div'); card.style.minWidth='min(430px,86vw)';
                card.style.maxWidth='86vw'; card.style.padding='26px 30px'; card.style.borderRadius='16px';
                card.style.background='#fff'; card.style.boxShadow='0 18px 55px rgba(0,0,0,.28)';
                card.style.textAlign='center'; card.style.fontFamily='sans-serif'; card.style.color='#111827';
                var sp=d.createElement('div'); sp.style.width='42px'; sp.style.height='42px';
                sp.style.margin='0 auto 14px'; sp.style.border='5px solid #e5e7eb';
                sp.style.borderTopColor='#0284c7'; sp.style.borderRadius='50%';
                sp.style.animation='wraOsmBusySpin .8s linear infinite';
                var title=d.createElement('div'); title.textContent=String(message||'正在載入 OSM 圖資…');
                title.style.fontSize='22px'; title.style.fontWeight='800'; title.style.marginBottom='8px';
                var desc=d.createElement('div'); desc.textContent=String(note||'OSM 公開服務查詢中，請稍候，不是當機。');
                desc.style.fontSize='15px'; desc.style.lineHeight='1.55'; desc.style.color='#4b5563';
                desc.style.whiteSpace='pre-line';
                card.appendChild(sp); card.appendChild(title); card.appendChild(desc); mask.appendChild(card); d.body.appendChild(mask);
            } catch(e) {}
        }
        function hideOsmBusy() {
            try { var d=window.parent.document, el=d.getElementById(osmBusyOverlayId); if(el) el.remove(); } catch(e) {}
        }

        function osmStorageKey() { return osmMemoryKey ? ('wra_gis_osm_v3618::' + osmMemoryKey) : ''; }
        function osmStorage() { try { if (window.parent && window.parent.sessionStorage) return window.parent.sessionStorage; } catch(e) {} return window.sessionStorage; }
        function saveOsmMemory(data, meta) {
            var k = osmStorageKey(); if (!k) return;
            try { osmStorage().setItem(k, JSON.stringify({data:data, meta:meta || {}, ts:Date.now()})); } catch(e) {}
        }
        function anchorStorageKey() {
            var k = osmStorageKey();
            return k ? (k + '::extend-anchor') : '';
        }
        function saveExtendAnchor(anchor) {
            lastExtendAnchor = anchor || null;
            var k = anchorStorageKey(); if (!k) return;
            try {
                if (anchor) osmStorage().setItem(k, JSON.stringify(anchor));
                else osmStorage().removeItem(k);
            } catch(e) {}
        }
        function loadExtendAnchor() {
            if (lastExtendAnchor) return lastExtendAnchor;
            var k = anchorStorageKey(); if (!k) return null;
            try {
                var raw = osmStorage().getItem(k);
                if (!raw) return null;
                var obj = JSON.parse(raw);
                if (!obj || !Array.isArray(obj.start) || !Array.isArray(obj.end)) return null;
                lastExtendAnchor = obj;
                return obj;
            } catch(e) { return null; }
        }
        function mergeOsmMemory(data, meta) {
            var old = loadOsmMemory();
            var merged = {elements:[]};
            var seen = {};
            var sources = [];
            if (old && old.data && Array.isArray(old.data.elements)) sources.push(old.data.elements);
            if (data && Array.isArray(data.elements)) sources.push(data.elements);
            for (var s=0; s<sources.length; s++) {
                for (var i=0; i<sources[s].length; i++) {
                    var el = sources[s][i] || {};
                    var key = String(el.type || 'way') + ':' + String(el.id || '');
                    if (!el.id || seen[key]) continue;
                    seen[key] = true;
                    merged.elements.push(el);
                    if (merged.elements.length >= maxCandidates) break;
                }
                if (merged.elements.length >= maxCandidates) break;
            }
            saveOsmMemory(merged, meta || ((old && old.meta) || {}));
            return merged;
        }
        function clearOsmMemory() {
            var k = osmStorageKey();
            try { if (k) osmStorage().removeItem(k); } catch(e) {}
            saveExtendAnchor(null);
        }
        function loadOsmMemory() {
            var k = osmStorageKey(); if (!k) return null;
            try {
                var raw = osmStorage().getItem(k);
                if (!raw) return null;
                var obj = JSON.parse(raw);
                if (!obj || !obj.data) return null;
                if (obj.ts && (Date.now() - Number(obj.ts)) > 2*60*60*1000) {
                    osmStorage().removeItem(k); return null;
                }
                return obj;
            } catch(e) { return null; }
        }
        function cancelPendingPick(message) {
            if (pendingPickHandler) {
                try { mapObj.off('click', pendingPickHandler); } catch(e) {}
                pendingPickHandler = null;
            }
            try { mapObj.getContainer().style.cursor = ''; } catch(e) {}
            if (message) setStatus(message, false);
        }

        function esc(s) {
            return String(s == null ? '' : s)
                .replace(/&/g, '&amp;').replace(/</g, '&lt;')
                .replace(/>/g, '&gt;').replace(/\"/g, '&quot;')
                .replace(/'/g, '&#039;');
        }

        function regexEscape(s) {
            return String(s == null ? '' : s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        }

        function overpassQuoted(s) {
            return String(s == null ? '' : s).replace(/\\/g, '\\\\').replace(/\"/g, '\\"');
        }

        function bboxFor(lat, lon, radiusM) {
            var dLat = radiusM / 111320.0;
            var cosLat = Math.cos(lat * Math.PI / 180.0);
            var dLon = radiusM / (111320.0 * Math.max(0.15, cosLat));
            return [lat-dLat, lon-dLon, lat+dLat, lon+dLon];
        }

        function kindOf(tags) {
            tags = tags || {};
            if (tags.man_made === 'embankment') return 'embankment';
            if (tags.man_made === 'dyke') return 'dyke';
            if (tags.barrier === 'retaining_wall') return 'retaining_wall';
            if (tags.waterway) return 'waterway=' + tags.waterway;
            if (tags.natural === 'water') return 'river water boundary';
            return 'OSM way';
        }

        function setStatus(msg, isError) {
            var el = document.getElementById('{{ this.status_id }}');
            if (!el) return;
            el.textContent = msg;
            el.style.color = isError ? '#b91c1c' : '#334155';
            el.style.fontWeight = isError ? '600' : '400';
        }

        async function fetchWithTimeout(url, query, ms) {
            var ctl = new AbortController();
            var timer = setTimeout(function(){ ctl.abort(); }, ms);
            try {
                var resp = await fetch(url, {
                    method: 'POST',
                    headers: {'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'},
                    body: 'data=' + encodeURIComponent(query),
                    signal: ctl.signal
                });
                if (!resp.ok) throw new Error('HTTP ' + resp.status);
                return await resp.json();
            } finally {
                clearTimeout(timer);
            }
        }

        async function runOverpass(query) {
            var endpoints = [
                'https://overpass.private.coffee/api/interpreter',
                'https://overpass-api.de/api/interpreter',
                'https://overpass.kumi.systems/api/interpreter'
            ];
            var data = null;
            var errs = [];
            for (var i=0; i<endpoints.length; i++) {
                try {
                    data = await fetchWithTimeout(endpoints[i], query, 15000);
                    if (data && Array.isArray(data.elements)) break;
                } catch (err) {
                    errs.push(endpoints[i] + ': ' + (err && err.message ? err.message : String(err)));
                    data = null;
                }
            }
            if (!data) console.warn('WRA GIS browser Overpass errors', errs);
            return data;
        }

        function editableHasOsmId(osmId) {
            var found = false;
            try {
                var editableGroup = {{ this.feature_group_name }};
                if (editableGroup && editableGroup.eachLayer) {
                    editableGroup.eachLayer(function(layer){
                        var p = (layer && layer.feature && layer.feature.properties) || {};
                        if (String(p.osm_id || '') === String(osmId || '')) found = true;
                    });
                }
            } catch(e) {}
            return found;
        }

        function refLayerHasOsmId(osmId) {
            var found = false;
            try {
                refLayer.eachLayer(function(layer){
                    if (String(layer && layer._wraOsmId || '') === String(osmId || '')) found = true;
                });
            } catch(e) {}
            return found;
        }

        function addCandidate(el) {
            if (editableHasOsmId(el && el.id) || refLayerHasOsmId(el && el.id)) return false;
            var geom = el.geometry || [];
            if (geom.length < 2) return false;
            var latlngs = geom.filter(function(pt){ return pt.lat != null && pt.lon != null; })
                              .map(function(pt){ return [Number(pt.lat), Number(pt.lon)]; });
            if (latlngs.length < 2) return false;
            var tags = el.tags || {};
            var kind = kindOf(tags);
            var nm = tags.name || tags['name:zh-Hant'] || tags['name:zh'] || tags.local_name || '未命名';
            var ref = L.polyline(latlngs, {
                color:'#64748b', weight:3, opacity:0.90,
                dashArray:'7,6', bubblingMouseEvents:true
            });
            ref._wraOsmId = String(el.id || '');
            ref.bindTooltip('OSM way ' + esc(el.id) + '｜' + esc(kind) + '｜' + esc(nm), {sticky:true});
            ref.on('mouseover', function(){ this.setStyle({weight:6, color:'#334155'}); });
            ref.on('mouseout', function(){ this.setStyle({weight:3, color:'#64748b'}); });
            (function(elCopy, latlngsCopy, kindCopy, nameCopy, tagsCopy) {
                ref.on('click', function(ev) {
                    if (ev && ev.originalEvent) L.DomEvent.stopPropagation(ev.originalEvent);
                    if (latlngsCopy && latlngsCopy.length >= 2) {
                        saveExtendAnchor({
                            osm_id:String(elCopy.id || ''),
                            osm_name:String(nameCopy || ''),
                            start:[Number(latlngsCopy[0][0]), Number(latlngsCopy[0][1])],
                            end:[Number(latlngsCopy[latlngsCopy.length-1][0]), Number(latlngsCopy[latlngsCopy.length-1][1])],
                            ts:Date.now()
                        });
                    }
                    var copyLine = L.polyline(latlngsCopy, {
                        color:'#3388ff', weight:{{ this.line_weight }}, opacity:0.95
                    });
                    copyLine.feature = {
                        type:'Feature',
                        properties:{
                            source:'OpenStreetMap', osm_type:'way', osm_id:elCopy.id,
                            osm_kind:kindCopy, osm_name:nameCopy, osm_tags:tagsCopy,
                            copied_at:(new Date()).toISOString()
                        },
                        geometry:null
                    };
                    var editableGroup = {{ this.feature_group_name }};
                    if (editableGroup && editableGroup.addLayer) {
                        try { refLayer.removeLayer(ref); } catch(e) {}
                        try { mapObj.fire('wra:editor-tool-start', {mode:'osm-copy'}); } catch(e) {}
                        mapObj.fire('draw:created', {layer:copyLine, layerType:'polyline'});
                        setStatus(
                            '已複製 OSM way ' + elCopy.id + ' 成藍色可編輯草稿；'
                            + '如河段尚未接完，可按「🧭 延伸相連 OSM 河段」從此線兩端各向外找下一層。',
                            false
                        );
                    } else {
                        setStatus('找不到可編輯圖層，請重新整理此頁後再試。', true);
                    }
                });
            })(el, latlngs, kind, nm, tags);
            ref.addTo(refLayer);
            return true;
        }

        function renderCandidates(data, contextLabel) {
            refLayer.clearLayers();
            var seen = {};
            var count = 0;
            var truncated = false;
            var elements = (data && data.elements) || [];
            for (var j=0; j<elements.length; j++) {
                var el = elements[j];
                var key = String(el.type || 'way') + ':' + String(el.id || '');
                if (seen[key]) continue;
                seen[key] = true;
                if (count >= maxCandidates) { truncated = true; break; }
                if (addCandidate(el)) count += 1;
            }
            if (count === 0) {
                setStatus(contextLabel + '沒有找到符合條件的 OSM 河道／排水線。', false);
            } else if (truncated) {
                setStatus('候選線超過安全上限，只顯示前 ' + maxCandidates + ' 條；請縮小範圍或輸入更完整名稱。', true);
            } else {
                setStatus(contextLabel + '找到 ' + count + ' 條。點灰色虛線即可複製成藍色草稿。', false);
            }
            return count;
        }

        function appendCandidates(data, contextLabel, limit) {
            var count = 0;
            var elements = (data && data.elements) || [];
            var safeLimit = Math.max(1, Math.min(extensionLimit, Number(limit) || extensionLimit));
            for (var j=0; j<elements.length; j++) {
                if (count >= safeLimit) break;
                if (addCandidate(elements[j])) count += 1;
            }
            if (count === 0) {
                setStatus(contextLabel + '沒有找到新的相連 OSM 河段；可改用點選位置附近 OSM。', false);
            } else {
                setStatus(
                    contextLabel + '新增 ' + count + ' 條候選線（單次最多 ' + safeLimit
                    + ' 條）。需要更遠時可先複製其中一段，再按一次延伸。',
                    false
                );
            }
            return count;
        }

        async function extendConnected() {
            if (busy) return;
            var anchor = loadExtendAnchor();
            if (!anchor || !Array.isArray(anchor.start) || !Array.isArray(anchor.end)) {
                setStatus('請先點一條灰色 OSM 候選線複製成藍色草稿，再使用「延伸相連 OSM 河段」。', true);
                return;
            }
            busy = true;
            cancelPendingPick('');
            indicatorLayer.clearLayers();
            try {
                var a = anchor.start, b = anchor.end;
                L.circle([Number(a[0]), Number(a[1])], {
                    radius:extensionRadiusM, color:'#7C3AED', weight:2, opacity:.85,
                    fillColor:'#C4B5FD', fillOpacity:.08, dashArray:'5,5', interactive:false
                }).addTo(indicatorLayer);
                L.circle([Number(b[0]), Number(b[1])], {
                    radius:extensionRadiusM, color:'#7C3AED', weight:2, opacity:.85,
                    fillColor:'#C4B5FD', fillOpacity:.08, dashArray:'5,5', interactive:false
                }).addTo(indicatorLayer);

                setStatus(
                    '正在從目前 OSM 線兩端各搜尋 ' + extensionRadiusM
                    + 'm 內下一層相連河段；不會自動遞迴。',
                    false
                );
                showOsmBusy(
                    '正在延伸相連 OSM 河段…',
                    '只搜尋目前線段兩個端點附近各 ' + extensionRadiusM
                    + 'm，單次最多新增 ' + extensionLimit + ' 條候選線\n'
                    + '採單層延伸，不會載入整個縣市河網'
                );

                var q = '[out:json][timeout:10];(' +
                    'way(around:' + extensionRadiusM + ',' + Number(a[0]).toFixed(7) + ',' + Number(a[1]).toFixed(7) + ')["waterway"~"^(river|stream|canal|drain|ditch|riverbank)$"];' +
                    'way(around:' + extensionRadiusM + ',' + Number(b[0]).toFixed(7) + ',' + Number(b[1]).toFixed(7) + ')["waterway"~"^(river|stream|canal|drain|ditch|riverbank)$"];' +
                    'way(around:' + extensionRadiusM + ',' + Number(a[0]).toFixed(7) + ',' + Number(a[1]).toFixed(7) + ')["man_made"~"^(embankment|dyke)$"];' +
                    'way(around:' + extensionRadiusM + ',' + Number(b[0]).toFixed(7) + ',' + Number(b[1]).toFixed(7) + ')["man_made"~"^(embankment|dyke)$"];' +
                    'way(around:' + extensionRadiusM + ',' + Number(a[0]).toFixed(7) + ',' + Number(a[1]).toFixed(7) + ')["barrier"="retaining_wall"];' +
                    'way(around:' + extensionRadiusM + ',' + Number(b[0]).toFixed(7) + ',' + Number(b[1]).toFixed(7) + ')["barrier"="retaining_wall"];' +
                    ');out tags geom qt;';
                var data = await runOverpass(q);
                if (!data) {
                    setStatus('相連河段延伸暫時查詢失敗；既有 OSM 候選線與藍色草稿均不受影響。', true);
                    return;
                }
                var count = appendCandidates(data, '相連河段延伸：', extensionLimit);
                if (count > 0) {
                    mergeOsmMemory(data, {
                        type:'extended',
                        anchor_osm_id:String(anchor.osm_id || ''),
                        anchor_osm_name:String(anchor.osm_name || ''),
                        radius:Number(extensionRadiusM)
                    });
                }
            } finally {
                busy = false;
                hideOsmBusy();
            }
        }

        function showSearchArea(lat, lon, radius) {
            indicatorLayer.clearLayers();
            L.circle([lat, lon], {
                radius: radius, color:'#2563EB', weight:2, opacity:0.9,
                fillColor:'#60A5FA', fillOpacity:0.10, dashArray:'5,5', interactive:false
            }).addTo(indicatorLayer);
            L.circleMarker([lat, lon], {
                radius:5, color:'#1D4ED8', weight:2,
                fillColor:'#FFFFFF', fillOpacity:1, interactive:false
            }).addTo(indicatorLayer);
        }

        async function loadRefsAt(lat, lon) {
            if (busy) return;
            busy = true;
            refLayer.clearLayers();
            var radius = Math.min(100, Math.max(20, Number({{ this.radius }}) || 50));
            showSearchArea(lat, lon, radius);
            setStatus('正在載入點位附近 ' + radius + 'm 的 OSM 參考線…', false);
            showOsmBusy(
                '正在載入 OSM 圖資…',
                '正在查詢點選位置附近 ' + radius + 'm 的 OpenStreetMap 公開資料\n請稍候，不是當機畫面'
            );
            try {
                var b = bboxFor(lat, lon, radius);
                var s=b[0].toFixed(7), w=b[1].toFixed(7), n=b[2].toFixed(7), e=b[3].toFixed(7);
                var q = '[out:json][timeout:10];(' +
                  'way["man_made"="embankment"]('+s+','+w+','+n+','+e+');' +
                  'way["man_made"="dyke"]('+s+','+w+','+n+','+e+');' +
                  'way["barrier"="retaining_wall"]('+s+','+w+','+n+','+e+');' +
                  'way["waterway"~"^(river|stream|canal|drain|ditch|riverbank)$"]('+s+','+w+','+n+','+e+');' +
                  'way["natural"="water"]["water"="river"]('+s+','+w+','+n+','+e+');' +
                  ');out tags geom qt;';
                var data = await runOverpass(q);
                if (!data) {
                    setStatus('OSM 公開服務目前無法連線，請稍後再試；既有草稿不受影響。', true);
                    return;
                }
                renderCandidates(data, '此 ' + radius + 'm 範圍');
                saveOsmMemory(data, {type:'nearby', lat:Number(lat), lon:Number(lon), radius:Number(radius)});
            } finally {
                busy = false;
                hideOsmBusy();
            }
        }

        async function searchByCountyName(keyword) {
            if (busy) return;
            try { mapObj.fire('wra:osm-pick-start', {type:'name'}); } catch(e) {}
            cancelPendingPick('');
            keyword = String(keyword || '').trim();
            if (!countyName) {
                setStatus('請先在上方選擇縣市，再使用縣內名稱搜尋。', true);
                return;
            }
            if (keyword.length < 2) {
                setStatus('水路名稱請至少輸入 2 個字，避免搜尋範圍過大。', true);
                return;
            }
            busy = true;
            refLayer.clearLayers();
            indicatorLayer.clearLayers();
            setStatus('正在「' + countyName + '」內搜尋名稱含「' + keyword + '」的 OSM 水路…', false);
            showOsmBusy(
                '正在載入 OSM 圖資…',
                '正在「' + countyName + '」內搜尋名稱含「' + keyword + '」的 OpenStreetMap 水路\n請稍候，不是當機畫面'
            );

            try {
            var countyAlt = countyName.indexOf('臺') >= 0 ? countyName.replace(/臺/g, '台') : countyName.replace(/台/g, '臺');
            var countyRegex = '^(' + regexEscape(countyName);
            if (countyAlt && countyAlt !== countyName) countyRegex += '|' + regexEscape(countyAlt);
            countyRegex += ')$';
            var kw = regexEscape(keyword);
            countyRegex = overpassQuoted(countyRegex);
            kw = overpassQuoted(kw);

            var areaPart = 'area["boundary"="administrative"]["name"~"' + countyRegex + '"]->.searchArea;';
            var q = '[out:json][timeout:12];' + areaPart + '(' +
                'way(area.searchArea)["waterway"~"^(river|stream|canal|drain|ditch|riverbank)$"]["name"~"' + kw + '",i];' +
                'way(area.searchArea)["waterway"~"^(river|stream|canal|drain|ditch|riverbank)$"]["name:zh"~"' + kw + '",i];' +
                'way(area.searchArea)["waterway"~"^(river|stream|canal|drain|ditch|riverbank)$"]["name:zh-Hant"~"' + kw + '",i];' +
                'way(area.searchArea)["man_made"~"^(embankment|dyke)$"]["name"~"' + kw + '",i];' +
                'way(area.searchArea)["barrier"="retaining_wall"]["name"~"' + kw + '",i];' +
                ');out tags geom qt;';
            var data = await runOverpass(q);
            if (!data) {
                setStatus('OSM 公開服務目前無法連線，請稍後再試；既有草稿不受影響。', true);
                return;
            }
            var count = renderCandidates(data, countyName + '內名稱搜尋：');
            saveOsmMemory(data, {type:'name', keyword:String(keyword), county:String(countyName)});
            if (count > 0) {
                try { mapObj.fitBounds(refLayer.getBounds(), {padding:[30,30], maxZoom:16}); } catch(e) {}
            }
            } finally {
                busy = false;
                hideOsmBusy();
            }
        }

        function startPickMode() {
            if (busy) return;
            cancelPendingPick('');
            try { mapObj.fire('wra:osm-pick-start', {}); } catch(e) {}
            setStatus('請在地圖上點一下搜尋中心；只搜尋該點附近 {{ this.radius }}m。', false);
            mapObj.getContainer().style.cursor = 'crosshair';
            pendingPickHandler = function(ev) {
                cancelPendingPick('');
                if (!ev || !ev.latlng) return;
                loadRefsAt(Number(ev.latlng.lat), Number(ev.latlng.lng));
            };
            mapObj.on('click', pendingPickHandler);
        }

        try { mapObj.on('wra:editor-tool-start', function(){ cancelPendingPick(''); }); } catch(e) {}

        var Ctrl = L.Control.extend({
            options:{position:'bottomright'},
            onAdd:function(map) {
                var box = L.DomUtil.create('div','leaflet-bar');
                box.style.background='#FFFDF2';
                box.style.padding='9px';
                box.style.minWidth='250px';
                box.style.maxWidth='315px';
                box.style.fontSize='12px';
                box.style.lineHeight='1.35';
                box.style.border='2px solid #EAB308';
                box.style.borderRadius='7px';
                box.style.boxShadow='0 2px 8px rgba(0,0,0,.35)';
                box.style.zIndex='10000';
                box.style.color='#111827';
                box.style.webkitTextFillColor='#111827';
                box.style.maxHeight='52%';
                box.style.overflowY='auto';
                box.style.marginBottom='18px';
                L.DomEvent.disableClickPropagation(box);
                L.DomEvent.disableScrollPropagation(box);

                var title = L.DomUtil.create('div','',box);
                title.textContent = clickToSelect ? '🌐 OSM 參考水路搜尋' : '🌐 OSM 參考線';
                title.style.fontWeight='700';
                title.style.fontSize='13px';
                title.style.marginBottom='6px';
                title.style.color='#713F12';

                var btn = L.DomUtil.create('button','',box);
                btn.type='button';
                btn.textContent = clickToSelect ? '📍 點選位置搜尋附近 OSM' : '🌐 載入工程附近 OSM';
                btn.style.width='100%';
                btn.style.cursor='pointer';
                btn.style.padding='7px 8px';
                btn.style.border='1px solid #D97706';
                btn.style.borderRadius='4px';
                btn.style.background='#FEF3C7';
                btn.style.fontWeight='600';
                btn.style.color='#111827';
                btn.style.webkitTextFillColor='#111827';
                btn.onclick=function(ev){
                    L.DomEvent.stop(ev);
                    if (clickToSelect) startPickMode();
                    else {
                        try { mapObj.fire('wra:osm-pick-start', {type:'project-nearby'}); } catch(e) {}
                        cancelPendingPick('');
                        loadRefsAt(Number({{ this.lat }}), Number({{ this.lon }}));
                    }
                };

                // V3.6.36：工程圖資編輯除了工程代表點附近，也可在長河段任意點選搜尋中心。
                // 兩種方式共用同一個安全半徑（20 / 50 / 100m），避免載入過大範圍。
                if (!clickToSelect) {
                    var pickBtn = L.DomUtil.create('button','',box);
                    pickBtn.type='button';
                    pickBtn.textContent='📍 載入點選位置附近 OSM';
                    pickBtn.style.width='100%';
                    pickBtn.style.cursor='pointer';
                    pickBtn.style.marginTop='6px';
                    pickBtn.style.padding='7px 8px';
                    pickBtn.style.border='1px solid #0284C7';
                    pickBtn.style.borderRadius='4px';
                    pickBtn.style.background='#E0F2FE';
                    pickBtn.style.fontWeight='600';
                    pickBtn.style.color='#0F172A';
                    pickBtn.style.webkitTextFillColor='#0F172A';
                    pickBtn.onclick=function(ev){
                        L.DomEvent.stop(ev);
                        startPickMode();
                    };
                }

                if (clickToSelect) {
                    var sep = L.DomUtil.create('div','',box);
                    sep.style.borderTop='1px solid #E5E7EB';
                    sep.style.margin='8px 0 6px';

                    var label = L.DomUtil.create('div','',box);
                    label.textContent = countyName ? ('🔎 ' + countyName + '內搜尋水路名稱') : '🔎 縣內搜尋水路名稱';
                    label.style.fontWeight='600';
                    label.style.marginBottom='4px';
                    label.style.color='#111827';
                    label.style.webkitTextFillColor='#111827';

                    var nameInput = L.DomUtil.create('input','',box);
                    nameInput.type='text';
                    nameInput.placeholder='例如：朴子溪、荷苞嶼';
                    nameInput.style.width='100%';
                    nameInput.style.boxSizing='border-box';
                    nameInput.style.padding='6px';
                    nameInput.style.border='1px solid #CBD5E1';
                    nameInput.style.borderRadius='4px';
                    nameInput.style.background='white';
                    nameInput.style.color='#111827';
                    nameInput.style.webkitTextFillColor='#111827';

                    var searchBtn = L.DomUtil.create('button','',box);
                    searchBtn.type='button';
                    searchBtn.textContent='🔎 縣內搜尋水路名稱 OSM';
                    searchBtn.style.width='100%';
                    searchBtn.style.cursor='pointer';
                    searchBtn.style.marginTop='5px';
                    searchBtn.style.padding='6px 8px';
                    searchBtn.style.border='1px solid #0284C7';
                    searchBtn.style.borderRadius='4px';
                    searchBtn.style.background='#E0F2FE';
                    searchBtn.style.color='#0F172A';
                    searchBtn.style.webkitTextFillColor='#0F172A';
                    searchBtn.onclick=function(ev){
                        L.DomEvent.stop(ev);
                        searchByCountyName(nameInput.value);
                    };
                    nameInput.addEventListener('keydown', function(ev){
                        if (ev.key === 'Enter') { ev.preventDefault(); searchByCountyName(nameInput.value); }
                    });
                }

                var extendBtn = L.DomUtil.create('button','',box);
                extendBtn.type='button';
                extendBtn.textContent='🧭 延伸相連 OSM 河段（1層）';
                extendBtn.style.width='100%';
                extendBtn.style.cursor='pointer';
                extendBtn.style.marginTop='6px';
                extendBtn.style.padding='6px 8px';
                extendBtn.style.border='1px solid #7C3AED';
                extendBtn.style.borderRadius='4px';
                extendBtn.style.background='#F3E8FF';
                extendBtn.style.color='#4C1D95';
                extendBtn.style.webkitTextFillColor='#4C1D95';
                extendBtn.title='先點一條灰色 OSM 候選線複製；之後只從該線兩端各搜尋 80m，單次最多 20 條，不自動遞迴。';
                extendBtn.onclick=function(ev){
                    L.DomEvent.stop(ev);
                    extendConnected();
                };

                var extendHelp = L.DomUtil.create('div','',box);
                extendHelp.textContent='延伸只查目前線段兩端各 80m、單次最多 20 條，不會自動追完整河網。';
                extendHelp.style.marginTop='3px';
                extendHelp.style.fontSize='10.5px';
                extendHelp.style.color='#6B7280';

                var clearBtn = L.DomUtil.create('button','',box);
                clearBtn.type='button';
                clearBtn.textContent='🧹 清除灰色 OSM 參考線';
                clearBtn.style.width='100%';
                clearBtn.style.cursor='pointer';
                clearBtn.style.marginTop='6px';
                clearBtn.style.padding='4px 8px';
                clearBtn.style.border='1px solid #E2E8F0';
                clearBtn.style.borderRadius='4px';
                clearBtn.style.background='white';
                clearBtn.style.color='#111827';
                clearBtn.style.webkitTextFillColor='#111827';
                clearBtn.onclick=function(ev){
                    L.DomEvent.stop(ev);
                    refLayer.clearLayers();
                    indicatorLayer.clearLayers();
                    clearOsmMemory();
                    cancelPendingPick('');
                    setStatus('OSM 候選線已清除；已複製的藍色草稿不受影響。', false);
                };

                var status = L.DomUtil.create('div','',box);
                status.id='{{ this.status_id }}';
                status.textContent = clickToSelect
                    ? '附近搜尋半徑 {{ this.radius }}m；也可用目前縣市＋水路名稱搜尋。複製一段後可用 80m 單層延伸。'
                    : '搜尋半徑 {{ this.radius }}m；可用工程代表點或任意位置搜尋，複製一段後可用 80m 單層延伸。';
                status.style.marginTop='6px';
                status.style.whiteSpace='normal';
                status.style.color='#475569';
                status.style.fontSize='11px';
                setTimeout(function(){
                    var mem = loadOsmMemory();
                    if (!mem || !mem.data) return;
                    if (mem.meta && mem.meta.type === 'nearby' && isFinite(mem.meta.lat) && isFinite(mem.meta.lon)) {
                        showSearchArea(Number(mem.meta.lat), Number(mem.meta.lon), Number(mem.meta.radius || {{ this.radius }}));
                    }
                    var restored = renderCandidates(mem.data, '已恢復上次 OSM 搜尋：');
                    if (restored > 0) setStatus('已恢復上次 OSM 候選線；已複製的線不會重複顯示。', false);
                }, 180);
                return box;
            }
        });
        (new Ctrl()).addTo(mapObj);
    })();
    {% endmacro %}
    """)

    def __init__(self, lat: float, lon: float, radius: int = 50, line_weight: int = 6,
                 feature_group: Optional[folium.FeatureGroup] = None,
                 click_to_select: bool = False, county: str = "", memory_key: str = ""):
        super().__init__()
        self._name = "BrowserOSMReferenceControl"
        self.lat = float(lat)
        self.lon = float(lon)
        r = int(radius)
        self.radius = r if r in (20, 50, 100) else 50
        self.line_weight = int(line_weight)
        self.feature_group_name = feature_group.get_name() if feature_group is not None else "drawnItems"
        self.click_to_select = bool(click_to_select)
        self.click_to_select_js = "true" if self.click_to_select else "false"
        self.county = str(county or "").strip()
        self.county_json = json.dumps(self.county, ensure_ascii=False)
        self.memory_key_json = json.dumps(str(memory_key or self.county or "default"), ensure_ascii=False)
        self.status_id = f"wra_osm_status_{int(time.time()*1000)%100000000}"

@st.cache_data(ttl=300, show_spinner=False)
def _fetch_osm_reference_lines(lat: float, lon: float, radius_m: int = 50) -> List[Dict[str, Any]]:
    """載入附近 OSM 線狀參考圖資。

    搜尋半徑只允許 20 / 50 / 100 公尺；即使外部傳入其他數值，
    也會回到安全預設 50 公尺，避免意外的大範圍 Overpass 查詢。

    公開 Overpass 服務偶爾會因負載、維護或網路路由暫時拒絕連線，
    因此不要綁死單一 endpoint；依序嘗試多個公開/官方備援節點。
    成功結果暫存 5 分鐘，避免使用者重複點擊造成不必要的公共 API 負載。
    """
    radius_m = int(radius_m)
    if radius_m not in (20, 50, 100):
        radius_m = 50

    query = f"""
    [out:json][timeout:20];
    (
      way(around:{int(radius_m)},{lat},{lon})[man_made=embankment];
      way(around:{int(radius_m)},{lat},{lon})[man_made=dyke];
      way(around:{int(radius_m)},{lat},{lon})[barrier=retaining_wall];
      way(around:{int(radius_m)},{lat},{lon})[waterway=river];
      way(around:{int(radius_m)},{lat},{lon})[waterway=stream];
    );
    out tags geom;
    """

    endpoints = [
        "https://overpass-api.de/api/interpreter",
        "https://overpass.private.coffee/api/interpreter",
        # Overpass 官方文件指出，在主入口異常時可直接使用個別主機作為 workaround。
        "https://gall.openstreetmap.de/api/interpreter",
        "https://lambert.openstreetmap.de/api/interpreter",
    ]
    headers = {
        "User-Agent": "WRA-GIS/3.5.4 Streamlit Overpass Client",
        "Accept": "application/json",
    }

    errors: List[str] = []
    data: Optional[Dict[str, Any]] = None
    for endpoint in endpoints:
        try:
            r = requests.post(
                endpoint,
                data={"data": query},
                headers=headers,
                timeout=(6, 24),
            )
            # 429 / 5xx 通常只是該公共節點暫時繁忙，直接換下一個。
            if r.status_code == 429 or 500 <= r.status_code <= 599:
                errors.append(f"{endpoint}: HTTP {r.status_code}")
                continue
            r.raise_for_status()
            payload = r.json()
            if isinstance(payload, dict):
                data = payload
                break
            errors.append(f"{endpoint}: 回傳格式不是 JSON 物件")
        except (requests.RequestException, ValueError) as exc:
            errors.append(f"{endpoint}: {exc}")
            continue

    if data is None:
        summary = "；".join(errors[-4:])
        raise RuntimeError(
            "目前所有 OSM Overpass 公開服務皆暫時無法連線。"
            "請稍後再按一次「載入附近 OSM 參考線」；您的工程圖資不會受影響。"
            + (f"\n技術資訊：{summary}" if summary else "")
        )

    out: List[Dict[str, Any]] = []
    for e in data.get("elements", []):
        geom = e.get("geometry") or []
        if len(geom) < 2:
            continue
        coords = [[float(x["lon"]), float(x["lat"])] for x in geom if "lon" in x and "lat" in x]
        if len(coords) < 2:
            continue
        tags = e.get("tags") or {}
        kind = tags.get("man_made") or tags.get("barrier") or tags.get("waterway") or "way"
        name = tags.get("name") or tags.get("local_name") or ""
        out.append({
            "osm_id": e.get("id"), "kind": kind, "name": name, "tags": tags,
            "feature": {
                "type": "Feature",
                "geometry": {"type": "LineString", "coordinates": coords},
                "properties": {"osm_id": e.get("id"), "source": "OpenStreetMap"},
            },
        })
    return out[:100]




def _feature_all_latlon_points(feature: Dict[str, Any]) -> List[List[float]]:
    """取出 GeoJSON feature 所有座標點，供河道整治總覽做精準 fit_bounds。

    舊版只用圖形中心點估算 bounds，長河道／長排水搜尋時可能無法把整段完整放入視窗。
    這裡直接走訪 Point / LineString / Polygon / Multi* 的所有座標。
    """
    geom = (feature or {}).get("geometry") or {}
    coords = geom.get("coordinates")
    out: List[List[float]] = []

    def walk(node: Any) -> None:
        if not isinstance(node, (list, tuple)):
            return
        if len(node) >= 2 and isinstance(node[0], (int, float)) and isinstance(node[1], (int, float)):
            try:
                lon, lat = float(node[0]), float(node[1])
                if -180 <= lon <= 180 and -90 <= lat <= 90:
                    out.append([lat, lon])
            except Exception:
                pass
            return
        for child in node:
            walk(child)

    walk(coords)
    return out


def _overview_legend_items(display_mode: str) -> Tuple[str, List[Tuple[str, str]]]:
    if display_mode == "簡易顯示":
        return "河道整治情形", [
            (REACH_UNASSIGNED_COLOR, "未指定現況（待確認）"),
            (REACH_PENDING_COLOR, "尚待治理"),
            (SIMPLE_APPROVED_COLOR, "已核定"),
            (SIMPLE_COMPLETED_COLOR, "已完成治理"),
        ]
    return "詳細顯示", [
        (REACH_UNASSIGNED_COLOR, "未指定現況（待確認）"),
        (REACH_PENDING_COLOR, "尚待治理"),
        (DETAIL_STATUS_COLORS["未發包"], "未發包"),
        (DETAIL_STATUS_COLORS["招標中"], "招標中"),
        (DETAIL_STATUS_COLORS["訂約中"], "訂約中／待開工"),
        (DETAIL_STATUS_COLORS["施工中"], "施工中"),
        (DETAIL_STATUS_COLORS["停工中"], "停工中／落後"),
        (DETAIL_STATUS_COLORS["已完工"], "已完工工程"),
        (REACH_COMPLETED_DETAIL_COLOR, "已完成治理（底圖）"),
        (DETAIL_STATUS_COLORS["已取消"], "已解約／已取消"),
    ]


class OverviewPresentationControl(MacroElement):
    """V3.6.24 河道整治總覽輸出控制。

    不再使用 html2canvas 截 Leaflet DOM。輸出時直接把『目前載入的 OSM tiles』與
    『目前可見 Leaflet Point / Line / Polygon』用同一個 Leaflet 投影座標重畫到 Canvas，
    再畫標題、圖例、比例尺、指北箭頭及 OSM attribution。這可避免不同瀏覽器的 CSS
    transform / DPI 造成點位與線面錯位。
    """
    _template = Template(r"""
    {% macro script(this, kwargs) %}
    (function(){
      var map = {{ this._parent.get_name() }};
      var mapEl = map.getContainer();
      if (!mapEl) return;

      var titleText = {{ this.title_json|safe }};
      var titlePosition = {{ this.title_position_json|safe }};
      var titlePx = {{ this.title_px }};
      var legendTitle = {{ this.legend_title_json|safe }};
      var legendItems = {{ this.legend_items_json|safe }};
      var exportWidth = {{ this.export_width }};

      function stopEvents(el){
        L.DomEvent.disableClickPropagation(el);
        L.DomEvent.disableScrollPropagation(el);
      }

      // 螢幕上的單行圖名。
      if (titleText) {
        var titleCtl = L.control({position: titlePosition});
        titleCtl.onAdd = function(){
          var d = L.DomUtil.create('div','wra-overview-title');
          d.textContent = titleText;
          d.style.background = 'rgba(255,255,255,.90)';
          d.style.border = '1px solid rgba(17,24,39,.55)';
          d.style.borderRadius = '7px';
          d.style.padding = '7px 11px';
          d.style.fontWeight = '700';
          d.style.fontSize = titlePx + 'px';
          d.style.lineHeight = '1.25';
          d.style.whiteSpace = 'nowrap';
          d.style.color = '#111827';
          d.style.webkitTextFillColor = '#111827';
          d.style.boxShadow = '0 1px 4px rgba(0,0,0,.18)';
          d.style.maxWidth = '70vw';
          d.style.overflow = 'hidden';
          d.style.textOverflow = 'ellipsis';
          stopEvents(d);
          return d;
        };
        titleCtl.addTo(map);
      }

      // 螢幕上的圖例。
      var legendCtl = L.control({position:'topright'});
      legendCtl.onAdd = function(){
        var d = L.DomUtil.create('div','wra-overview-legend');
        d.style.background='rgba(255,255,255,.94)';
        d.style.border='1px solid #9CA3AF';
        d.style.borderRadius='7px';
        d.style.padding='8px 10px';
        d.style.fontSize='12px';
        d.style.lineHeight='1.25';
        d.style.color='#111827';
        d.style.webkitTextFillColor='#111827';
        d.style.boxShadow='0 1px 4px rgba(0,0,0,.16)';
        var h=document.createElement('div');
        h.textContent=legendTitle; h.style.fontWeight='700'; h.style.marginBottom='5px'; d.appendChild(h);
        legendItems.forEach(function(it){
          var row=document.createElement('div');
          row.style.display='flex'; row.style.alignItems='center'; row.style.gap='7px'; row.style.margin='3px 0';
          var sw=document.createElement('span');
          sw.style.width='24px'; sw.style.height='5px'; sw.style.display='inline-block'; sw.style.borderRadius='3px';
          sw.style.background=it[0]; sw.style.border='1px solid #111827';
          var tx=document.createElement('span'); tx.textContent=it[1]; tx.style.color='#111827'; tx.style.webkitTextFillColor='#111827';
          row.appendChild(sw); row.appendChild(tx); d.appendChild(row);
        });
        stopEvents(d); return d;
      };
      legendCtl.addTo(map);

      L.control.scale({position:'bottomleft', metric:true, imperial:false, maxWidth:120}).addTo(map);
      var northCtl = L.control({position:'bottomleft'});
      northCtl.onAdd = function(){
        var d=L.DomUtil.create('div','wra-north-arrow');
        d.innerHTML='<div style="font-weight:800;font-size:12px;line-height:1;text-align:center">N</div>'+ 
                    '<div style="font-size:31px;line-height:.85;text-align:center;font-family:Arial,sans-serif">↑</div>';
        d.style.width='34px'; d.style.padding='4px 2px'; d.style.background='rgba(255,255,255,.86)';
        d.style.border='1px solid #9CA3AF'; d.style.borderRadius='6px'; d.style.color='#111827';
        d.style.webkitTextFillColor='#111827'; d.style.boxShadow='0 1px 3px rgba(0,0,0,.13)';
        stopEvents(d); return d;
      };
      northCtl.addTo(map);

      function safeFileName(s){
        var t=(s || '河道整治總覽').trim();
        t=t.replace(/[\\/:*?"<>|]/g,'_');
        return t || '河道整治總覽';
      }
      function eachVisibleVectorLayer(cb){
        var seen=[];
        function hasSeen(x){ return seen.indexOf(x)>=0; }
        function walk(layer){
          if(!layer || hasSeen(layer)) return; seen.push(layer);
          if(layer instanceof L.Path){
            try{ if(map.hasLayer(layer)) cb(layer); }catch(e){}
            return;
          }
          if(layer.eachLayer){ try{ layer.eachLayer(walk); }catch(e){} }
        }
        try{ map.eachLayer(walk); }catch(e){}
      }
      function dash(v){
        if(!v) return [];
        if(Array.isArray(v)) return v.map(Number).filter(isFinite);
        return String(v).split(/[ ,]+/).map(Number).filter(function(x){return isFinite(x)&&x>=0;});
      }
      function waitTilesReady(maxWait){
        var started=Date.now();
        return new Promise(function(resolve){
          function check(){
            var waiting=false;
            map.eachLayer(function(layer){
              if(!(layer instanceof L.TileLayer) || !map.hasLayer(layer)) return;
              var tiles=layer._tiles||{};
              Object.keys(tiles).forEach(function(k){
                var img=tiles[k]&&tiles[k].el;
                if(img && (!img.complete || !img.naturalWidth)) waiting=true;
              });
            });
            if(!waiting || Date.now()-started>=maxWait){ resolve(); return; }
            setTimeout(check,120);
          }
          check();
        });
      }
      function drawTiles(ctx){
        var currentZoom=map.getZoom();
        var origin=map.getPixelOrigin();
        map.eachLayer(function(layer){
          if(!(layer instanceof L.TileLayer) || !map.hasLayer(layer)) return;
          var ts=layer.getTileSize ? layer.getTileSize() : L.point(256,256);
          var opacity=Number(layer.options && layer.options.opacity!=null ? layer.options.opacity : 1);
          var tiles=layer._tiles||{};
          Object.keys(tiles).forEach(function(k){
            var t=tiles[k]||{}, img=t.el, c=t.coords;
            if(!img || !c || !img.complete || !img.naturalWidth) return;
            var z=(c.z==null ? (layer._tileZoom==null ? currentZoom : layer._tileZoom) : c.z);
            var sc=Math.pow(2,currentZoom-z);
            var x=c.x*ts.x*sc-origin.x;
            var y=c.y*ts.y*sc-origin.y;
            var w=ts.x*sc, h=ts.y*sc;
            try{
              ctx.save(); ctx.globalAlpha=opacity; ctx.drawImage(img,x,y,w,h); ctx.restore();
            }catch(e){}
          });
        });
      }
      function traceLatLngs(ctx, latlngs, closePath){
        if(!latlngs || !latlngs.length) return false;
        if(latlngs[0] && typeof latlngs[0].lat==='number'){
          var first=true;
          latlngs.forEach(function(ll){
            var p=map.latLngToContainerPoint(ll);
            if(first){ctx.moveTo(p.x,p.y); first=false;} else ctx.lineTo(p.x,p.y);
          });
          if(closePath) ctx.closePath();
          return true;
        }
        var ok=false;
        latlngs.forEach(function(ch){ if(traceLatLngs(ctx,ch,closePath)) ok=true; });
        return ok;
      }
      function drawVectors(ctx){
        eachVisibleVectorLayer(function(layer){
          try{
            var o=layer.options||{};
            if(layer instanceof L.CircleMarker && !(layer instanceof L.Circle)){
              var p=map.latLngToContainerPoint(layer.getLatLng());
              var r=Number(layer.getRadius ? layer.getRadius() : o.radius)||5;
              ctx.beginPath(); ctx.arc(p.x,p.y,r,0,Math.PI*2);
              if(o.fill!==false){ ctx.save(); ctx.globalAlpha=Number(o.fillOpacity==null?1:o.fillOpacity); ctx.fillStyle=o.fillColor||o.color||'#3388ff'; ctx.fill(); ctx.restore(); }
              ctx.save(); ctx.globalAlpha=Number(o.opacity==null?1:o.opacity); ctx.strokeStyle=o.color||'#000000'; ctx.lineWidth=Number(o.weight||2); ctx.stroke(); ctx.restore();
              return;
            }
            if(layer instanceof L.Polygon){
              ctx.save(); ctx.beginPath(); traceLatLngs(ctx,layer.getLatLngs(),true);
              if(o.fill!==false){ ctx.globalAlpha=Number(o.fillOpacity==null?0.2:o.fillOpacity); ctx.fillStyle=o.fillColor||o.color||'#3388ff'; try{ctx.fill('evenodd');}catch(e){ctx.fill();} }
              ctx.globalAlpha=Number(o.opacity==null?1:o.opacity); ctx.strokeStyle=o.color||'#3388ff'; ctx.lineWidth=Number(o.weight||2); ctx.lineJoin='round'; ctx.lineCap='round'; ctx.setLineDash(dash(o.dashArray)); ctx.stroke(); ctx.restore();
              return;
            }
            if(layer instanceof L.Polyline){
              ctx.save(); ctx.beginPath(); traceLatLngs(ctx,layer.getLatLngs(),false);
              ctx.globalAlpha=Number(o.opacity==null?1:o.opacity); ctx.strokeStyle=o.color||'#3388ff'; ctx.lineWidth=Number(o.weight||3); ctx.lineJoin='round'; ctx.lineCap='round'; ctx.setLineDash(dash(o.dashArray)); ctx.stroke(); ctx.restore();
            }
          }catch(e){}
        });
      }
      function roundedRect(ctx,x,y,w,h,r,fill,stroke){
        r=Math.max(0,Math.min(r,Math.min(w,h)/2));
        ctx.beginPath(); ctx.moveTo(x+r,y); ctx.arcTo(x+w,y,x+w,y+h,r); ctx.arcTo(x+w,y+h,x,y+h,r); ctx.arcTo(x,y+h,x,y,r); ctx.arcTo(x,y,x+w,y,r); ctx.closePath();
        if(fill){ctx.fillStyle=fill;ctx.fill();} if(stroke){ctx.strokeStyle=stroke;ctx.lineWidth=1;ctx.stroke();}
      }
      function textWidth(ctx,text,font){ ctx.font=font; return ctx.measureText(text).width; }
      function titleBox(ctx,W,H){
        if(!titleText) return null;
        var font='700 '+titlePx+'px Arial, "Noto Sans TC", sans-serif';
        var padX=11,padY=8, h=titlePx*1.35+padY*2, w=Math.min(W*0.72,textWidth(ctx,titleText,font)+padX*2);
        var x=12,y=12;
        if(titlePosition.indexOf('right')>=0) x=W-w-12;
        if(titlePosition.indexOf('bottom')>=0) y=H-h-12;
        roundedRect(ctx,x,y,w,h,7,'rgba(255,255,255,.92)','#6B7280');
        ctx.font=font; ctx.fillStyle='#111827'; ctx.textBaseline='middle'; ctx.textAlign='left';
        ctx.save(); ctx.beginPath(); ctx.rect(x+padX,y,w-padX*2,h); ctx.clip(); ctx.fillText(titleText,x+padX,y+h/2); ctx.restore();
        return {x:x,y:y,w:w,h:h};
      }
      function legendBox(ctx,W,H,tbox){
        var font='12px Arial, "Noto Sans TC", sans-serif', bold='700 12px Arial, "Noto Sans TC", sans-serif';
        var max=0; legendItems.forEach(function(it){max=Math.max(max,textWidth(ctx,it[1],font));});
        max=Math.max(max,textWidth(ctx,legendTitle,bold));
        var w=Math.min(255,Math.max(150,max+52)), rowH=20, h=30+legendItems.length*rowH, x=W-w-12, y=12;
        if(tbox && titlePosition==='topright') y=tbox.y+tbox.h+8;
        roundedRect(ctx,x,y,w,h,7,'rgba(255,255,255,.94)','#9CA3AF');
        ctx.fillStyle='#111827';ctx.textAlign='left';ctx.textBaseline='middle';ctx.font=bold;ctx.fillText(legendTitle,x+10,y+15);
        legendItems.forEach(function(it,i){
          var cy=y+32+i*rowH;
          ctx.strokeStyle='#111827';ctx.lineWidth=6;ctx.beginPath();ctx.moveTo(x+10,cy);ctx.lineTo(x+34,cy);ctx.stroke();
          ctx.strokeStyle=it[0];ctx.lineWidth=3.3;ctx.beginPath();ctx.moveTo(x+10,cy);ctx.lineTo(x+34,cy);ctx.stroke();
          ctx.fillStyle='#111827';ctx.font=font;ctx.fillText(it[1],x+43,cy);
        });
        return {x:x,y:y,w:w,h:h};
      }
      function niceScale(maxMeters){
        if(!isFinite(maxMeters)||maxMeters<=0) return 1;
        var exp=Math.pow(10,Math.floor(Math.log10(maxMeters))), vals=[1,2,5], best=exp;
        vals.forEach(function(v){var n=v*exp;if(n<=maxMeters)best=Math.max(best,n);});
        if(best>maxMeters) best=maxMeters;
        return best;
      }
      function scaleAndNorth(ctx,W,H,tbox){
        var y=H-23, startX=14;
        var a=map.containerPointToLatLng([0,Math.max(0,mapEl.clientHeight/2)]);
        var b=map.containerPointToLatLng([100,Math.max(0,mapEl.clientHeight/2)]);
        var m100=map.distance(a,b), maxM=m100*1.2, meters=niceScale(maxM), px=meters/(m100/100);
        ctx.strokeStyle='#111827';ctx.lineWidth=3;ctx.beginPath();ctx.moveTo(startX,y);ctx.lineTo(startX+px,y);ctx.stroke();
        ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(startX,y-4);ctx.lineTo(startX,y+4);ctx.moveTo(startX+px,y-4);ctx.lineTo(startX+px,y+4);ctx.stroke();
        ctx.font='11px Arial, sans-serif';ctx.fillStyle='#111827';ctx.textAlign='center';ctx.textBaseline='bottom';
        var label=meters>=1000?(meters/1000)+' km':Math.round(meters)+' m';ctx.fillText(label,startX+px/2,y-5);
        var nx=18,ny=H-84;
        if(tbox && titlePosition==='bottomleft') ny=Math.max(45,tbox.y-58);
        ctx.fillStyle='rgba(255,255,255,.88)';ctx.strokeStyle='#9CA3AF';ctx.lineWidth=1;roundedRect(ctx,nx,ny,34,49,6,'rgba(255,255,255,.88)','#9CA3AF');
        ctx.fillStyle='#111827';ctx.font='700 12px Arial';ctx.textAlign='center';ctx.textBaseline='top';ctx.fillText('N',nx+17,ny+4);
        ctx.font='30px Arial';ctx.fillText('↑',nx+17,ny+14);
      }
      function attribution(ctx,W,H){
        var txt='© OpenStreetMap contributors';
        ctx.font='10px Arial, sans-serif';ctx.textAlign='right';ctx.textBaseline='bottom';
        var w=ctx.measureText(txt).width+10;
        ctx.fillStyle='rgba(255,255,255,.82)';ctx.fillRect(W-w-4,H-17,w,15);
        ctx.fillStyle='#374151';ctx.fillText(txt,W-8,H-5);
      }
      function renderExport(){
        map.stop();
        return waitTilesReady(3000).then(function(){
          var W=Math.max(1,mapEl.clientWidth), H=Math.max(1,mapEl.clientHeight);
          var scale=Math.max(1,exportWidth/W), outH=Math.round(H*scale);
          var canvas=document.createElement('canvas'); canvas.width=Math.round(W*scale); canvas.height=outH;
          var ctx=canvas.getContext('2d');ctx.fillStyle='#ffffff';ctx.fillRect(0,0,canvas.width,canvas.height);
          ctx.save();ctx.scale(scale,scale);
          drawTiles(ctx);drawVectors(ctx);
          var tb=titleBox(ctx,W,H);legendBox(ctx,W,H,tb);scaleAndNorth(ctx,W,H,tb);attribution(ctx,W,H);
          ctx.restore();
          return canvas;
        });
      }

      {% if this.enable_png %}
      var exportCtl=L.control({position:'bottomright'});
      exportCtl.onAdd=function(){
        var wrap=L.DomUtil.create('div','wra-export-control');
        var btn=L.DomUtil.create('button','',wrap);btn.type='button';btn.textContent='🖼️ 輸出目前視窗 PNG';
        btn.style.background='#FFFFFF';btn.style.border='1px solid #6B7280';btn.style.borderRadius='6px';
        btn.style.padding='7px 10px';btn.style.fontWeight='700';btn.style.cursor='pointer';btn.style.color='#111827';
        btn.style.webkitTextFillColor='#111827';btn.style.boxShadow='0 1px 4px rgba(0,0,0,.16)';stopEvents(wrap);
        btn.onclick=function(ev){
          ev.preventDefault();ev.stopPropagation();var old=btn.textContent;btn.disabled=true;btn.textContent='⏳ 正在產製 PNG…';
          renderExport().then(function(canvas){
            var url=canvas.toDataURL('image/png');var a=document.createElement('a');a.download=safeFileName(titleText)+'.png';a.href=url;
            document.body.appendChild(a);a.click();a.remove();
          }).catch(function(err){
            alert('輸出 PNG 失敗。請確認 OSM 圖磚已載入完成，或重新整理地圖後再試。\\n'+err);
          }).finally(function(){btn.disabled=false;btn.textContent=old;});
        };
        return wrap;
      };
      exportCtl.addTo(map);
      {% endif %}
    })();
    {% endmacro %}
    """)

    def __init__(
        self,
        title: str,
        title_position: str,
        title_px: int,
        display_mode: str,
        export_width: int = 1600,
        enable_png: bool = False,
    ):
        super().__init__()
        self._name = "OverviewPresentationControl"
        legend_title, legend_items = _overview_legend_items(display_mode)
        self.title_json = json.dumps(display_text(title), ensure_ascii=False)
        self.title_position_json = json.dumps(title_position)
        self.title_px = int(title_px)
        self.legend_title_json = json.dumps(legend_title, ensure_ascii=False)
        self.legend_items_json = json.dumps(legend_items, ensure_ascii=False)
        self.export_width = max(1200, min(3000, int(export_width or 1600)))
        # V3.6.58：PNG 疊圖仍無法穩定對準，先保留底層程式但不顯示操作入口。
        self.enable_png = bool(enable_png)


def _overview_position_code(label: str) -> str:
    return {
        "左上": "topleft",
        "右上": "topright",
        "左下": "bottomleft",
        "右下": "bottomright",
    }.get(display_text(label), "topright")


def _overview_title_px(label: str) -> int:
    return {"小": 18, "中": 24, "大": 32}.get(display_text(label), 24)

def build_overview_map_v35(
    rows: Sequence[ProjectRow],
    geo: Dict[str, Any],
    reaches: Dict[str, Any],
    display_mode: str,
    basemap_opacity: float = 0.35,
    export_title: str = "",
    title_position: str = "右上",
    title_size: str = "中",
    show_points: bool = True,
    show_lines: bool = True,
    show_polygons: bool = True,
    hide_original_when_shape: bool = True,
    project_feature_filter: Optional[Dict[str, Optional[set]]] = None,
    export_width: int = 1600,
) -> Tuple[folium.Map, Dict[str, int]]:
    """V3.6.24 河道整治總覽（黑框＋縮放自適應＋Canvas原生PNG）。

    project_feature_filter：pid -> None 表示該工程全部人工圖資；pid -> set(GEO-ID)
    表示只顯示關鍵字命中的人工圖資。原始代表點不受 GEO-ID set 限制，但仍受
    show_points / hide_original_when_shape 控制。
    """
    display_rows = _spatial_dedupe_rows_v3651(rows, geo)

    def type_visible(f: Dict[str, Any]) -> bool:
        gt = geometry_type(f)
        if gt == "Point":
            return bool(show_points)
        if gt == "LineString":
            return bool(show_lines)
        if gt == "Polygon":
            return bool(show_polygons)
        return False

    def manuals_for_project(pid: str) -> List[Dict[str, Any]]:
        fs = active_manual_features(geo, pid)
        if project_feature_filter is not None and pid in project_feature_filter:
            allowed = project_feature_filter.get(pid)
            if allowed is not None:
                fs = [f for f in fs if display_text((f.get("properties") or {}).get("geo_id")) in allowed]
        return fs

    def project_features(p: ProjectRow) -> List[Dict[str, Any]]:
        manuals_all = manuals_for_project(p.project_id)
        visible_manuals = [f for f in manuals_all if type_visible(f)]
        has_line_or_polygon = any(geometry_type(f) in {"LineString", "Polygon"} for f in manuals_all)
        out = list(visible_manuals)
        orig = original_feature(geo, p.project_id)
        if show_points and orig:
            # 若工程完全沒有人工圖資，原始點是必要的 fallback。
            # 若已有人工點／線／面，使用者可決定是否仍保留原始代表點；
            # 「隱藏已有線／範圍案件原始點」只針對有 Line/Polygon 的案件。
            if not manuals_all:
                out.append(orig)
            elif has_line_or_polygon:
                if not hide_original_when_shape:
                    out.append(orig)
            else:
                # 只有人工 Point 時，不重複疊加 Excel/GIS 原始代表點。
                pass
        return out

    bounds: List[List[float]] = []
    for f in active_reaches(reaches):
        if type_visible(f):
            bounds.extend(_feature_all_latlon_points(f))
    for p in display_rows:
        if not p.project_id or _project_color(p.status, display_mode) is None:
            continue
        for f in project_features(p):
            bounds.extend(_feature_all_latlon_points(f))

    if bounds:
        center = [sum(x[0] for x in bounds) / len(bounds), sum(x[1] for x in bounds) / len(bounds)]
        initial_zoom = 15 if len(bounds) == 1 else 9
    else:
        center = TAIWAN_CENTER
        initial_zoom = TAIWAN_ZOOM

    opacity = max(0.30, min(1.00, float(basemap_opacity)))
    m = folium.Map(location=center, zoom_start=initial_zoom, tiles=None, control_scale=False)
    folium.TileLayer(
        tiles="OpenStreetMap",
        name="OpenStreetMap",
        opacity=opacity,
        control=False,
        show=True,
        cross_origin="anonymous",
    ).add_to(m)
    _add_leaflet_compat_css(m)
    m.get_root().html.add_child(Element(r"""
    <style>
      .wra-overview-title,.wra-overview-title *,
      .wra-overview-legend,.wra-overview-legend *,
      .wra-north-arrow,.wra-north-arrow *,
      .wra-export-control,.wra-export-control * {
        color:#111827 !important; -webkit-text-fill-color:#111827 !important; color-scheme:light !important;
      }
      .wra-overview-legend { max-height:310px; overflow:auto; }
      @media (max-width:760px) {
        .wra-overview-title { max-width:55vw !important; white-space:normal !important; }
        .wra-overview-legend { font-size:11px !important; max-height:230px; }
      }
    </style>
    """))

    rch_count = 0
    type_counts = {"Point": 0, "LineString": 0, "Polygon": 0}
    for f in active_reaches(reaches):
        if not type_visible(f):
            continue
        p = f.get("properties") or {}
        color = _reach_color(str(p.get("status", "")), display_mode)
        if _add_feature_shape(
            m, f, color, _reach_popup(f),
            tooltip=str(p.get("water_name") or p.get("reach_name") or "治理現況底圖")
        ):
            rch_count += 1
            type_counts[geometry_type(f)] = type_counts.get(geometry_type(f), 0) + 1

    project_shapes = 0
    project_count = 0
    original_points = 0
    for p in display_rows:
        if not p.project_id:
            continue
        color = _project_color(p.status, display_mode)
        if color is None:
            continue
        fs = project_features(p)
        drew = False
        for f in fs:
            is_original = display_text((f.get("properties") or {}).get("role")) == "original_point"
            if f and _add_feature_shape(m, f, color, _project_popup_v35(p, f), tooltip=p.project_name):
                project_shapes += 1
                drew = True
                gt = geometry_type(f)
                type_counts[gt] = type_counts.get(gt, 0) + 1
                if is_original:
                    original_points += 1
        if drew:
            project_count += 1

    if len(bounds) > 1:
        try:
            m.fit_bounds(bounds, padding=(30, 30), max_zoom=16)
        except TypeError:
            m.fit_bounds(bounds, padding=(30, 30))
    elif len(bounds) == 1:
        m.location = bounds[0]
        m.options["zoom"] = 16

    AdaptiveVectorScale().add_to(m)

    OverviewPresentationControl(
        title=export_title,
        title_position=_overview_position_code(title_position),
        title_px=_overview_title_px(title_size),
        display_mode=display_mode,
        export_width=export_width,
    ).add_to(m)

    return m, {
        "projects": project_count,
        "project_shapes": project_shapes,
        "reaches": rch_count,
        "bounds_points": len(bounds),
        "points": type_counts.get("Point", 0),
        "lines": type_counts.get("LineString", 0),
        "polygons": type_counts.get("Polygon", 0),
        "original_points": original_points,
    }

def build_query_results_map(records: Sequence[Dict[str, Any]]) -> Tuple[folium.Map, Dict[str, int]]:
    """工程查詢地圖：沿用工程圖資優先規則，採詳細顯示配色。"""
    try:
        geo = load_query_geo_snapshot()
    except Exception:
        geo = empty_geojson()
    prepared = []
    bounds: List[List[float]] = []
    gis_project_count = 0
    fallback_count = 0
    seen_spatial = set()
    for rec in records:
        pid = display_text(rec.get("project_id"))
        canonical_pid = _spatial_primary_project_id_from_geo_v3651(geo, pid) if pid else ""
        spatial_key = canonical_pid or pid or display_text(rec.get("name"))
        if spatial_key and spatial_key in seen_spatial:
            continue
        if spatial_key:
            seen_spatial.add(spatial_key)
        features: List[Dict[str, Any]] = []
        if pid:
            features = active_manual_features(geo, pid)
            if not features:
                o = original_feature(geo, pid)
                if o:
                    features = [o]
        if features:
            gis_project_count += 1
        else:
            lon, lat = parse_number(rec.get("lon")), parse_number(rec.get("lat"))
            if is_valid_wgs84(lon, lat):
                fallback_count += 1
                features = [{"type": "Feature", "geometry": {"type": "Point", "coordinates": [lon, lat]}, "properties": {"source": "query_fallback"}}]
        if features:
            prepared.append((rec, features))
            for f in features:
                pts = _feature_all_latlon_points(f)
                if pts:
                    bounds.extend(pts)
                else:
                    c = feature_center(f)
                    if c:
                        bounds.append([float(c[0]), float(c[1])])
    center = TAIWAN_CENTER if not bounds else [sum(x[0] for x in bounds)/len(bounds), sum(x[1] for x in bounds)/len(bounds)]
    m = folium.Map(location=center, zoom_start=TAIWAN_ZOOM if not bounds else (15 if len(prepared)==1 else 9), tiles="OpenStreetMap", control_scale=True)
    _add_leaflet_compat_css(m)
    shapes = 0
    for rec, features in prepared:
        color = _detail_project_color(rec.get("status"))
        name = display_text(rec.get("name")) or "未命名工程"
        popup = _query_feature_popup(rec)
        label_done = False
        for f in features:
            if _add_feature_shape(m, f, color, popup, tooltip=name):
                shapes += 1
                if not label_done:
                    c = feature_center(f)
                    if c:
                        _add_query_label(m, float(c[0]), float(c[1]), name)
                        label_done = True
    if len(bounds) > 1:
        m.fit_bounds(bounds, padding=(18,18))
    _add_legend(m, "詳細顯示")
    AdaptiveVectorScale().add_to(m)
    return m, {"projects": len(prepared), "shapes": shapes, "gis_projects": gis_project_count, "fallback_projects": fallback_count}


def _render_overview_tab(rows: Sequence[ProjectRow], geo: Dict[str, Any], reaches: Dict[str, Any]) -> None:
    # V3.6.19：查詢使用 form，避免輸入每一個中文字就整張地圖重建；按查詢後才套用並 fit_bounds。
    if "gis_overview_keyword_applied" not in st.session_state:
        st.session_state["gis_overview_keyword_applied"] = ""

    c1, c2 = st.columns([1.2, 1.3])
    with c1:
        display_mode = st.radio("顯示模式", ["簡易顯示", "詳細顯示"], horizontal=False, key="gis_display_mode")
    with c2:
        counties = sorted({p.county for p in rows if p.county})
        county = st.selectbox("縣市", ["全部"] + counties, key="gis_map_county")

    with st.form("gis_overview_search_form", clear_on_submit=False):
        q1, q2, q3 = st.columns([3.2, 1.05, 0.9])
        with q1:
            keyword_input = st.text_input(
                "關鍵字查詢",
                value=st.session_state.get("gis_overview_keyword_applied", ""),
                placeholder="工程名稱、執行單位、工程地點、河川／排水名稱、圖資關鍵字",
            )
        with q2:
            st.write("")
            search_clicked = st.form_submit_button("🔎 查詢並縮放", use_container_width=True)
        with q3:
            st.write("")
            clear_clicked = st.form_submit_button("清除查詢", use_container_width=True)

    if clear_clicked:
        st.session_state["gis_overview_keyword_applied"] = ""
        st.session_state["gis_overview_query_epoch"] = int(
            st.session_state.get("gis_overview_query_epoch", 0) or 0
        ) + 1
    elif search_clicked:
        st.session_state["gis_overview_keyword_applied"] = display_text(keyword_input).strip()
        st.session_state["gis_overview_query_epoch"] = int(
            st.session_state.get("gis_overview_query_epoch", 0) or 0
        ) + 1

    keyword = st.session_state.get("gis_overview_keyword_applied", "")
    kw = norm_text(keyword)

    # V3.6.58：舊 Session 可能仍保存 15～25；先夾到新範圍，避免 slider 狀態越界。
    try:
        saved_basemap_pct = int(st.session_state.get("gis_overview_basemap_pct", 35))
    except Exception:
        saved_basemap_pct = 35
    saved_basemap_pct = max(30, min(100, saved_basemap_pct))
    st.session_state["gis_overview_basemap_pct"] = saved_basemap_pct
    if "gis_overview_basemap_pct_widget" in st.session_state:
        try:
            widget_basemap_pct = int(st.session_state["gis_overview_basemap_pct_widget"])
        except Exception:
            widget_basemap_pct = saved_basemap_pct
        st.session_state["gis_overview_basemap_pct_widget"] = max(30, min(100, widget_basemap_pct))

    # 顯示設定放在 form 內，使用者調圖名時不會每打一個字就觸發重建地圖。
    with st.expander("🖼️ 圖資顯示與輸出設定", expanded=False):
        with st.form("gis_overview_export_form", clear_on_submit=False):
            e1, e2, e3 = st.columns([2.5, 1.0, 1.0])
            with e1:
                export_title = st.text_input(
                    "輸出圖名（單行）",
                    value=st.session_state.get("gis_export_title", "河道整治總覽圖"),
                    placeholder="例如：六腳排水系統整治圖",
                    key="gis_export_title_widget",
                )
            with e2:
                title_position = st.selectbox(
                    "圖名位置", ["左上", "右上", "左下", "右下"],
                    index=["左上", "右上", "左下", "右下"].index(st.session_state.get("gis_export_title_position", "右上")),
                    key="gis_export_title_position_widget",
                )
            with e3:
                title_size = st.radio(
                    "圖名字級", ["小", "中", "大"], horizontal=True,
                    index=["小", "中", "大"].index(st.session_state.get("gis_export_title_size", "中")),
                    key="gis_export_title_size_widget",
                )
            basemap_pct = st.slider(
                "OpenStreetMap 底圖濃度（數值越低越淡）",
                min_value=30, max_value=100,
                value=saved_basemap_pct,
                step=5,
                key="gis_overview_basemap_pct_widget",
            )
            st.markdown("**圖資類型顯示**")
            d1, d2, d3 = st.columns(3)
            with d1:
                show_points = st.checkbox("顯示點位", value=bool(st.session_state.get("gis_overview_show_points", True)), key="gis_overview_show_points_widget")
            with d2:
                show_lines = st.checkbox("顯示線", value=bool(st.session_state.get("gis_overview_show_lines", True)), key="gis_overview_show_lines_widget")
            with d3:
                show_polygons = st.checkbox("顯示範圍／圖塊", value=bool(st.session_state.get("gis_overview_show_polygons", True)), key="gis_overview_show_polygons_widget")
            hide_original_when_shape = st.checkbox(
                "已有線或範圍圖資的案件，隱藏原始代表點",
                value=bool(st.session_state.get("gis_overview_hide_original_when_shape", True)),
                key="gis_overview_hide_original_when_shape_widget",
            )
            apply_export = st.form_submit_button("套用顯示／輸出設定")

        if apply_export:
            st.session_state["gis_export_title"] = display_text(export_title).strip()
            st.session_state["gis_export_title_position"] = title_position
            st.session_state["gis_export_title_size"] = title_size
            st.session_state["gis_overview_basemap_pct"] = int(basemap_pct)
            st.session_state["gis_overview_show_points"] = bool(show_points)
            st.session_state["gis_overview_show_lines"] = bool(show_lines)
            st.session_state["gis_overview_show_polygons"] = bool(show_polygons)
            st.session_state["gis_overview_hide_original_when_shape"] = bool(hide_original_when_shape)

        st.caption("套用顯示設定後，地圖會保留目前的中心與縮放範圍。")

    # form 尚未按套用時，沿用上一次已套用值；初次則採安全預設。
    export_title_applied = st.session_state.get("gis_export_title", "河道整治總覽圖")
    title_position_applied = st.session_state.get("gis_export_title_position", "右上")
    title_size_applied = st.session_state.get("gis_export_title_size", "中")
    basemap_pct_applied = max(30, min(100, int(st.session_state.get("gis_overview_basemap_pct", 35))))
    show_points_applied = bool(st.session_state.get("gis_overview_show_points", True))
    show_lines_applied = bool(st.session_state.get("gis_overview_show_lines", True))
    show_polygons_applied = bool(st.session_state.get("gis_overview_show_polygons", True))
    hide_original_when_shape_applied = bool(st.session_state.get("gis_overview_hide_original_when_shape", True))

    filtered = list(rows)
    if county != "全部":
        filtered = [p for p in filtered if p.county == county]
    # V3.6.21：工程欄位命中時顯示該工程全部圖資；若只命中某筆 GIS 圖資關鍵字／圖資名稱，
    # 則只顯示命中的那筆人工圖資，避免同一工程其他不相關圖形一併撐大視窗。
    project_feature_filter: Optional[Dict[str, Optional[set]]] = None
    if kw:
        project_feature_filter = {}
        filtered2: List[ProjectRow] = []
        for p in filtered:
            row_match = (
                kw in norm_text(p.project_name)
                or kw in norm_text(p.unit)
                or kw in norm_text(p.address)
            )
            manuals = active_project_edit_features(geo, p.project_id) if p.project_id else []
            matched_ids = {
                display_text((f.get("properties") or {}).get("geo_id"))
                for f in manuals if feature_keyword_match(f, kw)
            }
            matched_ids.discard("")
            if row_match:
                filtered2.append(p)
                if p.project_id:
                    project_feature_filter[p.project_id] = None
            elif matched_ids:
                filtered2.append(p)
                if p.project_id:
                    project_feature_filter[p.project_id] = matched_ids
        filtered = filtered2

    reaches_view = copy.deepcopy(reaches)
    rr = []
    for f in active_reaches(reaches):
        rp = f.get("properties") or {}
        if county != "全部" and str(rp.get("county", "")) != county:
            continue
        if kw:
            hay = "|".join([
                norm_text(rp.get("water_name", "")),
                norm_text(rp.get("reach_name", "")),
                norm_text(rp.get("basis", "")),
                norm_text(rp.get("notes", "")),
                norm_text(rp.get("keywords", "")),
            ])
            if kw not in hay:
                continue
        rr.append(f)
    reaches_view["features"] = rr

    m, counts = build_overview_map_v35(
        filtered,
        geo,
        reaches_view,
        display_mode,
        basemap_opacity=basemap_pct_applied / 100.0,
        export_title=export_title_applied,
        title_position=title_position_applied,
        title_size=title_size_applied,
        show_points=show_points_applied,
        show_lines=show_lines_applied,
        show_polygons=show_polygons_applied,
        hide_original_when_shape=hide_original_when_shape_applied,
        project_feature_filter=project_feature_filter,
    )

    if kw:
        if counts["projects"] == 0 and counts["reaches"] == 0:
            st.warning(f"查無符合「{keyword}」的工程或治理現況圖資。")
        else:
            st.success(
                f"已查詢「{keyword}」，並自動縮放至所有符合圖資的最小可視範圍："
                f"工程 {counts['projects']:,} 件、治理現況 {counts['reaches']:,} 段。"
            )

    st.caption(
        f"工程可顯示 {counts['projects']:,} 件／{counts['project_shapes']:,} 筆圖形；"
        f"治理現況底圖 {counts['reaches']:,} 段；目前顯示：點位 {counts.get('points',0):,}、"
        f"線 {counts.get('lines',0):,}、範圍 {counts.get('polygons',0):,}。"
        f"OSM 底圖濃度 {basemap_pct_applied}%。"
    )

    # 顯示設定改變會使 st_folium 元件重建；記憶鍵只隨查詢條件改變，
    # 因此按「套用顯示／輸出設定」不會跳回自動 fitBounds 的初始位置。
    query_epoch = int(st.session_state.get("gis_overview_query_epoch", 0) or 0)
    overview_view_scope = hashlib.md5(
        f"{display_mode}|{county}|{keyword}|{query_epoch}".encode("utf-8")
    ).hexdigest()[:12]
    BrowserViewportMemory(f"overview::{overview_view_scope}").add_to(m)

    st_folium(
        m,
        width=None,
        height=700,
        returned_objects=[],
        key=f"gis_overview_v3658_{display_mode}_{county}_{hashlib.md5((keyword+'|'+export_title_applied+'|'+title_position_applied+'|'+title_size_applied+'|'+str(basemap_pct_applied)+'|'+str(show_points_applied)+'|'+str(show_lines_applied)+'|'+str(show_polygons_applied)+'|'+str(hide_original_when_shape_applied)).encode('utf-8')).hexdigest()[:10]}",
    )

def _project_selector_point(
    p: ProjectRow,
    geo: Dict[str, Any],
    project_drafts: Dict[str, Any],
    point_overrides: Optional[Dict[str, Dict[str, float]]] = None,
) -> Optional[Tuple[float, float]]:
    """工程圖資編輯的選取點位：優先使用本 Session 待同步座標，再使用人工 Point／Excel 原始點；
    若只有線/面，使用其中心點作為『選取點』，方便從縣市地圖進入編輯。
    回傳 (lat, lon)。
    """
    if p.project_id and _spatial_is_secondary_v3651(geo, p.project_id):
        return None

    if p.project_id and point_overrides and p.project_id in point_overrides:
        try:
            ov = point_overrides[p.project_id]
            lon, lat = float(ov["lon"]), float(ov["lat"])
            if is_valid_wgs84(lon, lat):
                return lat, lon
        except Exception:
            pass

    candidates: List[Dict[str, Any]] = []
    if p.project_id:
        candidates.extend(active_manual_features(geo, p.project_id))
        orig = original_feature(geo, p.project_id)
        if orig:
            candidates.append(orig)
    else:
        candidates.extend(active_project_drafts(project_drafts, p.row_key))

    # 人工校正 Point 優先
    for f in candidates:
        g = f.get("geometry") or {}
        c = g.get("coordinates")
        if g.get("type") == "Point" and isinstance(c, list) and len(c) >= 2:
            try:
                lon, lat = float(c[0]), float(c[1])
                if is_valid_wgs84(lon, lat):
                    return lat, lon
            except Exception:
                pass

    # Excel 目前座標
    if p.lon is not None and p.lat is not None and is_valid_wgs84(p.lon, p.lat):
        return float(p.lat), float(p.lon)

    # 只有 Line/Polygon 時，以圖形中心當作選取點，不改變正式圖資
    for f in candidates:
        c = feature_center(f)
        if c:
            try:
                lat, lon = float(c[0]), float(c[1])
                if is_valid_wgs84(lon, lat):
                    return lat, lon
            except Exception:
                pass
    return None


def _project_edit_progress(
    p: ProjectRow,
    geo: Dict[str, Any],
    project_drafts: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """工程圖資編輯進度判定。

    Excel 自動產生的 original_point 不算『已編輯』；只有人工 GEO、待ID人工草稿，
    或代表點曾由 GIS 校正，才視為已有人工作圖。
    """
    if p.project_id:
        manuals = active_project_edit_features(geo, p.project_id)
    else:
        manuals = active_project_drafts(project_drafts or empty_project_drafts(), p.row_key)
    manual_types = {geometry_type(f) for f in manuals if f.get("geometry")}
    orig = original_feature(geo, p.project_id) if p.project_id else None
    op = (orig or {}).get("properties") or {}
    calibrated = display_text(op.get("coordinate_edit_source")) == "gis_map"
    has_shape = bool({"LineString", "Polygon"} & manual_types)
    has_manual_point = "Point" in manual_types
    edited = bool(manuals or calibrated)
    if has_shape:
        category = "已有線／範圍"
    elif has_manual_point or calibrated:
        category = "僅有人工作點／校正點"
    else:
        category = "尚未有人工作圖"
    return {
        "edited": edited,
        "has_shape": has_shape,
        "has_manual_point": has_manual_point,
        "calibrated": calibrated,
        "category": category,
        "manuals": manuals,
    }


class ProjectProgressLegendControl(MacroElement):
    _template = Template(r"""
    {% macro script(this, kwargs) %}
    (function(){
      var map={{ this._parent.get_name() }};
      var ctl=L.control({position:'bottomright'});
      ctl.onAdd=function(){
        var d=L.DomUtil.create('div','wra-project-progress-legend');
        d.innerHTML='<div style="font-weight:700;margin-bottom:3px">工程圖資進度</div>'+ 
          '<div><span style="display:inline-block;width:11px;height:11px;border-radius:50%;background:#60a5fa;border:2px solid #000;margin-right:6px"></span>彩色點／線／面：已有人工作圖</div>'+ 
          '<div><span style="display:inline-block;width:11px;height:11px;border-radius:50%;background:#d1d5db;border:2px solid #000;margin-right:6px"></span>灰色：尚未有人工作圖</div>'+ 
          '<div><span style="display:inline-block;width:11px;height:11px;border-radius:50%;background:#60a5fa;border:3px solid #fff;box-shadow:0 0 0 1px #111;margin-right:6px"></span>白框：目前選取</div>';
        d.style.background='rgba(255,255,255,.94)';d.style.border='1px solid #9CA3AF';d.style.borderRadius='7px';
        d.style.padding='7px 9px';d.style.fontSize='12px';d.style.color='#111827';d.style.webkitTextFillColor='#111827';
        d.style.lineHeight='1.45';d.style.boxShadow='0 1px 4px rgba(0,0,0,.15)';
        L.DomEvent.disableClickPropagation(d);L.DomEvent.disableScrollPropagation(d);return d;
      };
      ctl.addTo(map);
    })();
    {% endmacro %}
    """)
    def __init__(self):
        super().__init__()
        self._name='ProjectProgressLegendControl'


def _add_project_progress_legend(m: folium.Map) -> None:
    ProjectProgressLegendControl().add_to(m)


def _build_county_project_selector_map(
    county_rows: Sequence[ProjectRow],
    geo: Dict[str, Any],
    project_drafts: Dict[str, Any],
    selected_row_key: str = "",
    point_overrides: Optional[Dict[str, Dict[str, float]]] = None,
) -> Tuple[folium.Map, List[Tuple[ProjectRow, float, float]], int]:
    """建立縣市工程選取地圖；既有工程 Line/Polygon 會直接疊在預覽圖。"""
    display_rows = _spatial_dedupe_rows_v3651(county_rows, geo)
    points: List[Tuple[ProjectRow, float, float]] = []
    for p in display_rows:
        pt = _project_selector_point(p, geo, project_drafts, point_overrides=point_overrides)
        if pt:
            points.append((p, pt[0], pt[1]))

    if points:
        center = [
            sum(x[1] for x in points) / len(points),
            sum(x[2] for x in points) / len(points),
        ]
        zoom = 10
    else:
        center = TAIWAN_CENTER
        zoom = TAIWAN_ZOOM

    # prefer_canvas 對大量 CircleMarker 的瀏覽器繪製較輕量。
    m = folium.Map(location=center, zoom_start=zoom, tiles="OpenStreetMap", control_scale=True, prefer_canvas=True)
    _add_leaflet_compat_css(m)
    bounds: List[List[float]] = []

    # V3.6.25：先預覽已完成的人工作圖（Line/Polygon），再疊上工程選取點。
    # 不能只依 source == manual；active_manual_features 已兼容 manual_draft_promoted 等來源。
    # 待確認ID案件則直接把本 Session/正式 project_drafts 內的 Line/Polygon 一併預覽。
    preview_shape_count = 0
    for p in display_rows:
        if p.project_id:
            preview_features = active_project_edit_features(geo, p.project_id)
        else:
            preview_features = active_project_drafts(project_drafts, p.row_key)
        shape_features = [f for f in preview_features if geometry_type(f) in {"LineString", "Polygon"}]
        if not shape_features:
            continue
        preview_color = _detail_project_color(p.status)
        for f in shape_features:
            if _add_feature_shape(
                m, f, preview_color, "",
                tooltip=f"✅ {display_text(p.project_name) or '未命名工程'}｜已完成編輯圖資",
                opacity=.82, reference=False,
            ):
                preview_shape_count += 1
                bounds.extend(_feature_all_latlon_points(f))

    for p, lat, lon in points:
        is_selected = bool(selected_row_key and p.row_key == selected_row_key)
        prog = _project_edit_progress(p, geo, project_drafts)
        status_color = _detail_project_color(p.status)
        fill_color = status_color if prog["edited"] else "#D1D5DB"
        label_state = "✅ 已有人工作圖" if prog["edited"] else "⬜ 尚未有人工作圖"
        folium.CircleMarker(
            location=[lat, lon],
            radius=9 if is_selected else (7 if prog["edited"] else 5.5),
            color="#FFFFFF" if is_selected else "#000000",
            weight=4 if is_selected else 2.5,
            fill=True,
            fill_color=fill_color,
            fill_opacity=1.0 if is_selected else (0.95 if prog["edited"] else 0.82),
            tooltip=f"{label_state}｜{display_text(p.project_name) or '未命名工程'}｜{prog['category']}",
        ).add_to(m)
        bounds.append([lat, lon])

    if len(bounds) > 1:
        try:
            m.fit_bounds(bounds, padding=(18, 18))
        except Exception:
            pass
    elif len(bounds) == 1:
        m.location = bounds[0]
        m.options["zoom"] = 15
    AdaptiveVectorScale().add_to(m)
    _add_project_progress_legend(m)

    return m, points, preview_shape_count


def _match_clicked_project(
    state: Optional[Dict[str, Any]],
    points: Sequence[Tuple[ProjectRow, float, float]],
) -> Optional[ProjectRow]:
    """把 st_folium 的 marker 點擊座標對回 ProjectRow。"""
    if not state:
        return None
    obj = state.get("last_object_clicked")
    if not isinstance(obj, dict):
        return None
    try:
        lat = float(obj.get("lat"))
        lon = float(obj.get("lng"))
    except Exception:
        return None

    tooltip = display_text(state.get("last_object_clicked_tooltip"))
    best: Optional[ProjectRow] = None
    best_d = 999.0
    for p, plat, plon in points:
        # 若 tooltip 有值，先用名稱縮小候選；座標再解決同名工程。
        if tooltip and display_text(p.project_name) != tooltip:
            continue
        d = (plat - lat) ** 2 + (plon - lon) ** 2
        if d < best_d:
            best_d = d
            best = p

    # tooltip 因元件版本差異未回傳時，單純以座標找最近點。
    if best is None:
        for p, plat, plon in points:
            d = (plat - lat) ** 2 + (plon - lon) ** 2
            if d < best_d:
                best_d = d
                best = p

    # CircleMarker 點擊通常是精確座標；放寬到約 60m 以容忍前端浮點誤差。
    if best is not None and best_d <= 0.0006 ** 2:
        return best
    return None



def representative_point_excel_differences(
    rows: Sequence[ProjectRow],
    geo: Dict[str, Any],
    tolerance_m: float = 1.0,
) -> List[Dict[str, Any]]:
    """比較 GIS 代表點與 current.xlsx 座標，列出尚待同步項目。"""
    result: List[Dict[str, Any]] = []
    row_map = {p.project_id: p for p in rows if p.project_id}
    for pid, p in row_map.items():
        f = original_feature_for_project(geo, pid)
        geom = (f or {}).get("geometry") or {}
        coords = geom.get("coordinates") or []
        if geom.get("type") != "Point" or len(coords) < 2:
            continue
        try:
            glon, glat = float(coords[0]), float(coords[1])
        except Exception:
            continue
        if not is_valid_wgs84(glon, glat):
            continue
        excel_valid = p.lon is not None and p.lat is not None and is_valid_wgs84(p.lon, p.lat)
        distance = None
        if excel_valid:
            distance = _haversine_m(float(p.lat), float(p.lon), glat, glon)
        props = (f or {}).get("properties") or {}
        pending_flag = props.get("excel_sync_status") == "pending"
        if (not excel_valid) or pending_flag or (distance is not None and distance > tolerance_m):
            result.append({
                "project_id": pid,
                "project_name": p.project_name,
                "county": p.county,
                "sheet_name": p.sheet_name,
                "excel_lon": p.lon,
                "excel_lat": p.lat,
                "gis_lon": glon,
                "gis_lat": glat,
                "distance_m": distance,
                "updated_by": display_text(props.get("coordinate_updated_by")),
                "updated_at": display_text(props.get("coordinate_updated_at")),
                "coord_source": display_text(props.get("coord_source")),
                "coord_status": display_text(props.get("coord_status")),
                "coordinate_edit_source": display_text(props.get("coordinate_edit_source")),
                "spatial_project_id": display_text(props.get("spatial_project_id")),
                "spatial_primary_project_id": display_text(props.get("spatial_primary_project_id")),
            })
    return result


def _apply_gis_only_point_to_session_snapshot_v368(result: Dict[str, Any]) -> None:
    """GIS-only 儲存成功後只更新 Session 的 GEO；Excel rows 保持原值以供同步差異比較。"""
    snap = st.session_state.get("_gis_v366_snapshot")
    if isinstance(snap, dict) and result.get("_new_geo") is not None:
        snap["geo"] = result["_new_geo"]
        st.session_state["_gis_v366_snapshot"] = snap
    fn = globals().get("_cached_query_geo_snapshot_v351")
    if fn is not None and hasattr(fn, "clear"):
        try:
            fn.clear()
        except Exception:
            pass


def _render_project_edit_tab(svc: EngineeringGISService, rows: Sequence[ProjectRow], geo: Dict[str, Any], project_drafts: Dict[str, Any], editor: str) -> None:
    if not rows:
        st.info("目前沒有工程資料。")
        return

    # V3.6.8：代表點校正立即儲存到 engineering_geo.geojson，但不碰 current.xlsx。
    # Excel 座標改由系統管理者定期批次同步。
    point_overrides: Dict[str, Dict[str, float]] = {}
    st.caption("📍 代表點校正會立即儲存到 GIS；current.xlsx 不會被修改。Excel 座標由系統管理者日後批次同步。")

    # ------------------------------------------------------------
    # V3.5.2：先選縣市，再直接從地圖點工程進入編輯
    # ------------------------------------------------------------
    counties = sorted({normalize_county_name(p.county) or p.county for p in rows if p.county})
    unresolved = [p for p in rows if not p.county]
    selector_options = ["請選擇縣市"] + counties
    if unresolved:
        selector_options.append(f"未辨識縣市（{len(unresolved)}件）")
    # 最後一道防呆：即使來源 Excel 完全沒有可辨識的縣市欄，也不讓工程清單整頁空白。
    if not counties and rows:
        selector_options.append("全部工程（縣市欄待確認）")

    county = st.selectbox(
        "1. 選擇縣市",
        selector_options,
        key="gis_v352_edit_county",
    )
    _saved_notice = st.session_state.pop("gis_v365_point_saved_notice", None)
    if _saved_notice:
        st.success(_saved_notice)
    if county == "請選擇縣市":
        st.info("請先選擇縣市。選定後，地圖會顯示該縣市全部可定位工程點位；游標移到點位會顯示工程名稱，點一下即可進入編輯。")
        return

    if county.startswith("未辨識縣市"):
        county_rows = unresolved
    elif county == "全部工程（縣市欄待確認）":
        county_rows = list(rows)
    else:
        county_rows = [p for p in rows if (normalize_county_name(p.county) or p.county) == county]

    # V3.6.51：共用圖資的多筆 PRJ 仍保留於管控資料，但工程圖資編輯只顯示一個空間工程。
    spatial_county_rows = _spatial_dedupe_rows_v3651(county_rows, geo)
    shared_hidden_count = max(0, len(spatial_county_rows) - len(spatial_county_rows))
    if shared_hidden_count:
        st.caption(
            f"本縣市另有 {shared_hidden_count:,} 筆管控紀錄已納入共用圖資群組，"
            "不會在工程選取地圖重複畫點。"
        )

    # V3.6.24：工程圖資進度預覽。
    progress_map = {p.row_key: _project_edit_progress(p, geo, project_drafts) for p in spatial_county_rows}
    edited_count = sum(1 for p in spatial_county_rows if progress_map[p.row_key]["edited"])
    shape_count = sum(1 for p in spatial_county_rows if progress_map[p.row_key]["has_shape"])
    point_only_count = sum(1 for p in spatial_county_rows if progress_map[p.row_key]["edited"] and not progress_map[p.row_key]["has_shape"])
    pending_count = max(0, len(spatial_county_rows) - edited_count)
    mc = st.columns(4)
    mc[0].metric("本縣市工程", f"{len(spatial_county_rows):,}")
    mc[1].metric("已有人工作圖", f"{edited_count:,}")
    mc[2].metric("已有線／範圍", f"{shape_count:,}")
    mc[3].metric("尚未有人工作圖", f"{pending_count:,}")
    if point_only_count:
        st.caption(f"另有 {point_only_count:,} 件目前只有人工作點／已校正代表點，尚無人工線或範圍。")

    progress_filter = st.radio(
        "2. 圖資進度預覽與選取",
        ["全部案件", "尚未有人工作圖", "已有人工作圖", "已有線／範圍", "僅有人工作點／校正點"],
        horizontal=True, key=f"gis_v3624_progress_filter_{county}",
    )
    if progress_filter == "尚未有人工作圖":
        preview_rows = [p for p in spatial_county_rows if not progress_map[p.row_key]["edited"]]
    elif progress_filter == "已有人工作圖":
        preview_rows = [p for p in spatial_county_rows if progress_map[p.row_key]["edited"]]
    elif progress_filter == "已有線／範圍":
        preview_rows = [p for p in spatial_county_rows if progress_map[p.row_key]["has_shape"]]
    elif progress_filter == "僅有人工作點／校正點":
        preview_rows = [p for p in spatial_county_rows if progress_map[p.row_key]["edited"] and not progress_map[p.row_key]["has_shape"]]
    else:
        preview_rows = list(spatial_county_rows)

    valid_keys = {p.row_key for p in spatial_county_rows}
    selected_key = display_text(st.session_state.get("gis_v352_selected_project_key"))
    selected_county = display_text(st.session_state.get("gis_v352_selected_county"))
    if selected_county != county or selected_key not in valid_keys:
        st.session_state["gis_v352_selected_county"] = county
        st.session_state.pop("gis_v352_selected_project_key", None)
        selected_key = ""

    smap, selector_points, preview_shape_count = _build_county_project_selector_map(
        preview_rows, geo, project_drafts, selected_row_key=selected_key, point_overrides=point_overrides
    )
    no_point_count = max(0, len(preview_rows) - len(selector_points))
    st.caption(
        f"{county}目前篩選 {len(preview_rows):,} 件；地圖可定位 {len(selector_points):,} 件；"
        f"已直接載入既有人工線／範圍 {preview_shape_count:,} 筆。"
        + (f"另有 {no_point_count:,} 件目前沒有可用座標，可用下方『工程名稱備用選取』進入。" if no_point_count else "")
        + "　將游標移到工程點位可看名稱，點一下會選取該工程，並自動把畫面移到下方正式編輯地圖；上方選取地圖不會消失。"
    )
    if shape_count > 0 and preview_shape_count == 0 and progress_filter in {"全部案件", "已有人工作圖", "已有線／範圍"}:
        st.warning("系統判斷本縣市有已完成線／範圍，但預覽地圖尚未載入任何線／範圍。請按頁面上方『重新載入 GitHub 最新圖資』；若仍為 0，再回報我檢查該批舊 GEO 的 project_id。")
    selector_epoch = int(st.session_state.get("gis_v366_selector_epoch", 0) or 0)
    map_state = st_folium(
        smap,
        width=None,
        height=560,
        returned_objects=["last_object_clicked", "last_object_clicked_tooltip"],
        key=f"gis_v3624_project_selector_{county}_{progress_filter}_{selector_epoch}",
    )
    clicked_p = _match_clicked_project(map_state, selector_points)
    if clicked_p and clicked_p.row_key != selected_key:
        st.session_state["gis_v352_selected_project_key"] = clicked_p.row_key
        st.session_state["gis_v352_selected_county"] = county
        st.session_state["gis_v641_force_editor_scroll"] = True
        action_key0 = clicked_p.project_id or clicked_p.row_key
        st.session_state.setdefault(
            f"gis_project_edit_action::{action_key0}",
            "✏️ 編輯工程線／範圍",
        )
        st.session_state["gis_v366_selector_epoch"] = int(st.session_state.get("gis_v366_selector_epoch", 0) or 0) + 1
        st.rerun()

    # 沒座標或同仁較習慣文字搜尋時，保留備用途徑。
    with st.expander("🔎 工程名稱備用選取", expanded=False):
        name_options = {}
        for q in preview_rows:
            pid_text = q.project_id if q.project_id else "待確認ID"
            label = f"{q.project_name}｜{pid_text}｜{q.sheet_name}"
            name_options[label] = q
        labels = ["請選擇工程"] + list(name_options.keys())
        backup_label = st.selectbox("工程名稱", labels, key=f"gis_v352_name_selector_{county}")
        if backup_label != "請選擇工程":
            backup_p = name_options[backup_label]
            if st.button("選取此工程", use_container_width=True, key="gis_v352_select_by_name"):
                st.session_state["gis_v352_selected_project_key"] = backup_p.row_key
                st.session_state["gis_v352_selected_county"] = county
                st.session_state["gis_v641_force_editor_scroll"] = True
                action_key0 = backup_p.project_id or backup_p.row_key
                st.session_state.setdefault(
                    f"gis_project_edit_action::{action_key0}",
                    "✏️ 編輯工程線／範圍",
                )
                st.session_state["gis_v366_selector_epoch"] = int(st.session_state.get("gis_v366_selector_epoch", 0) or 0) + 1
                st.rerun()

    selected_key = display_text(st.session_state.get("gis_v352_selected_project_key"))
    p = next((x for x in spatial_county_rows if x.row_key == selected_key), None)
    if p is None:
        st.info("請在上方地圖點選一個工程點位；選取後會自動往下移到正式圖資編輯地圖，上方工程選取地圖仍會保留。")
        return

    st.markdown("---")
    st.markdown("### 3. 編輯已選工程")
    st.success(f"✅ 已選取：{p.project_name}　｜　{p.project_id or '待確認ID'}")
    st.caption("上方縣市工程選取地圖會保留；系統已把畫面移到本工程的正式編輯區。要換工程時可直接往上點另一個工程點位。")
    if st.session_state.pop("gis_v641_force_editor_scroll", False):
        _scroll_parent_to_current_component(96)
    if st.button("↩️ 取消目前工程選取", key="gis_v352_clear_selected_project"):
        st.session_state.pop("gis_v352_selected_project_key", None)
        st.session_state["gis_v366_selector_epoch"] = int(st.session_state.get("gis_v366_selector_epoch", 0) or 0) + 1
        st.rerun()

    # ------------------------------------------------------------
    # V3.6.2：選取工程後，先明確選擇「校正代表點」或「編輯工程線／範圍」。
    # 校正代表點不再藏在 toggle 裡；沒有系統工程ID時也會顯示原因。
    # ------------------------------------------------------------
    action_key = p.project_id or p.row_key
    edit_action = st.radio(
        "請選擇要進行的圖資操作",
        ["📍 校正工程代表點", "✏️ 編輯工程線／範圍"],
        index=1,
        horizontal=True,
        key=f"gis_project_edit_action::{action_key}",
    )

    if edit_action == "📍 校正工程代表點":
        st.markdown("#### 📍 校正工程代表點")
        if not p.project_id:
            st.warning(
                "這件工程目前尚未建立『系統工程ID』，因此暫時不能建立正式 GIS 工程代表點。"
                "請先由工程分併標管理建立系統工程ID；完成後回到這裡即可直接校正。"
            )
            return
        st.caption(
            "在下方地圖直接點一下新的正確位置。灰點是目前代表點，紅點是準備儲存的新代表點；"
            "按一次『💾 儲存 GIS 代表點』即正式儲存到 GitHub；current.xlsx 不會在此步驟變更。"
        )
        orig_now = original_feature(geo, p.project_id)
        old_lon = old_lat = None
        if orig_now and (orig_now.get("geometry") or {}).get("type") == "Point":
            cc = (orig_now.get("geometry") or {}).get("coordinates") or []
            if len(cc) >= 2:
                try:
                    old_lon, old_lat = float(cc[0]), float(cc[1])
                except Exception:
                    old_lon = old_lat = None
        if old_lon is None and p.lon is not None and p.lat is not None:
            old_lon, old_lat = float(p.lon), float(p.lat)

        cand_key = f"gis_point_candidate::{p.project_id}"
        candidate = st.session_state.get(cand_key)
        if isinstance(candidate, dict):
            try:
                cand_lat = float(candidate.get("lat"))
                cand_lon = float(candidate.get("lon"))
                if not is_valid_wgs84(cand_lon, cand_lat):
                    candidate = None
            except Exception:
                candidate = None

        # V3.6.33：校正點位一旦已有候選新點，地圖重繪時要以候選點為中心。
        # 舊版每次 st_folium rerun 都重新以 original_point 為中心，若新點距離很遠，
        # 使用者每點一次就會被拉回舊位置。現在只有第一次尚未選新點時才定位原始點。
        if candidate:
            map_center = [float(candidate["lat"]), float(candidate["lon"])]
            map_zoom = 17
        elif old_lat is not None and old_lon is not None:
            map_center = [old_lat, old_lon]
            map_zoom = 17
        elif selector_points:
            map_center = [
                sum(x[1] for x in selector_points) / len(selector_points),
                sum(x[2] for x in selector_points) / len(selector_points),
            ]
            map_zoom = 11
        else:
            map_center, map_zoom = TAIWAN_CENTER, TAIWAN_ZOOM

        pm = folium.Map(location=map_center, zoom_start=map_zoom, tiles="OpenStreetMap", control_scale=True)
        _add_leaflet_compat_css(pm)
        if old_lat is not None and old_lon is not None:
            folium.CircleMarker(
                [old_lat, old_lon], radius=8, color="#FFFFFF", weight=4,
                fill=True, fill_color="#757575", fill_opacity=0.95,
                tooltip="目前工程代表點",
            ).add_to(pm)
        if candidate:
            folium.CircleMarker(
                [float(candidate["lat"]), float(candidate["lon"])], radius=9,
                color="#FFFFFF", weight=4, fill=True, fill_color="#EF4444",
                fill_opacity=0.95, tooltip="準備儲存的新代表點",
            ).add_to(pm)

        PointCalibrationCursor().add_to(pm)
        # V3.6.59：點選校正位置觸發 rerun 時，同時保留整頁捲動位置與地圖視角。
        BrowserViewportMemory(f"point_calibration::{p.project_id}").add_to(pm)
        token = "none" if not candidate else f"{float(candidate['lat']):.6f}_{float(candidate['lon']):.6f}"
        click_state = st_folium(
            pm,
            width=None,
            height=500,
            returned_objects=["last_clicked"],
            key=f"gis_point_cal_map::{p.project_id}::{token}",
        )
        clicked = (click_state or {}).get("last_clicked") or {}
        try:
            click_lat = float(clicked.get("lat"))
            click_lon = float(clicked.get("lng"))
        except Exception:
            click_lat = click_lon = None
        if click_lat is not None and click_lon is not None and is_valid_wgs84(click_lon, click_lat):
            new_candidate = {"lat": click_lat, "lon": click_lon}
            if candidate != new_candidate:
                st.session_state[cand_key] = new_candidate
                st.rerun()

        st.info("操作方式：滑鼠會顯示一般箭頭游標，箭頭尖端會有紅點；在上方地圖直接點一下新位置。灰點是目前位置，紅點是準備儲存的新位置。")

        # 提供精確座標輸入備援，不依賴 Draw / 拖曳元件。
        base_lat = float(candidate["lat"]) if candidate else (float(old_lat) if old_lat is not None else 23.7)
        base_lon = float(candidate["lon"]) if candidate else (float(old_lon) if old_lon is not None else 120.95)
        with st.expander("⌨️ 需要更精確時，可直接輸入 WGS84 座標", expanded=False):
            nc1, nc2, nc3 = st.columns([1, 1, 0.8])
            in_lon = nc1.number_input("經度 E", value=base_lon, format="%.8f", key=f"gis_point_lon::{p.project_id}")
            in_lat = nc2.number_input("緯度 N", value=base_lat, format="%.8f", key=f"gis_point_lat::{p.project_id}")
            with nc3:
                st.markdown("<div style='height:1.68rem'></div>", unsafe_allow_html=True)
                if st.button("套用座標", use_container_width=True, key=f"gis_point_apply_xy::{p.project_id}"):
                    if not is_valid_wgs84(float(in_lon), float(in_lat)):
                        st.error("輸入座標不在臺灣及離島合理範圍。")
                    else:
                        st.session_state[cand_key] = {"lat": float(in_lat), "lon": float(in_lon)}
                        st.rerun()

        candidate = st.session_state.get(cand_key)
        if candidate:
            new_lat, new_lon = float(candidate["lat"]), float(candidate["lon"])
            info_cols = st.columns(3)
            info_cols[0].metric("新經度 E", f"{new_lon:.8f}")
            info_cols[1].metric("新緯度 N", f"{new_lat:.8f}")
            if old_lat is not None and old_lon is not None:
                moved_m = _haversine_m(float(old_lat), float(old_lon), new_lat, new_lon)
                info_cols[2].metric("移動距離", f"{moved_m:,.1f} m")
            else:
                info_cols[2].metric("移動距離", "新增點位")

            # V3.6.27：改用 button on_click callback。
            # callback 會在 Streamlit 正式 rerun 執行頁面內容之前先提交 GitHub，
            # 避免 st_folium 的 component rerun 把第一次按鈕事件吃掉。
            legacy_confirm_key = f"gis_point_confirm::{p.project_id}"
            st.session_state.pop(legacy_confirm_key, None)
            save_error_key = f"gis_point_save_error::{p.project_id}"
            prior_error = st.session_state.pop(save_error_key, None)
            if prior_error:
                st.error(prior_error)

            def _save_point_once_callback():
                try:
                    cand = st.session_state.get(cand_key) or {}
                    lon0 = float(cand.get("lon"))
                    lat0 = float(cand.get("lat"))
                    if not is_valid_wgs84(lon0, lat0):
                        raise ValueError("新代表點座標不在臺灣及離島合理範圍，未執行儲存。")
                    _save_mask = _saving_overlay("正在儲存 GIS 代表點…")
                    try:
                        result0 = svc.update_project_representative_point_gis_only(
                            p.project_id, lon0, lat0, editor,
                            project_name=p.project_name, county=p.county, status=p.status,
                        )
                    finally:
                        _close_saving_overlay(_save_mask)
                    _apply_gis_only_point_to_session_snapshot_v368(result0)
                    st.session_state.pop(cand_key, None)
                    st.session_state.pop("gis_v352_selected_project_key", None)
                    st.session_state["gis_v365_point_saved_notice"] = (
                        f"{p.project_name} 的 GIS 代表點已儲存；current.xlsx 尚未更新。"
                    )
                    st.session_state["gis_v366_selector_epoch"] = int(st.session_state.get("gis_v366_selector_epoch", 0) or 0) + 1
                except Exception as exc:
                    st.session_state[save_error_key] = f"儲存 GIS 代表點失敗：{exc}"

            pc1, pc2 = st.columns(2)
            with pc1:
                if st.button("清除本次新點位", use_container_width=True, key=f"gis_point_clear::{p.project_id}"):
                    st.session_state.pop(cand_key, None)
                    st.rerun()
            with pc2:
                st.button(
                    "💾 儲存 GIS 代表點",
                    type="primary",
                    use_container_width=True,
                    key=f"gis_point_save_once::{p.project_id}",
                    on_click=_save_point_once_callback,
                )
        else:
            st.caption("目前尚未選擇新位置；不會修改任何資料。")


        # 校正代表點模式完成後，不載入下面較重的線／面編輯器。
        return

    clear_notice = st.session_state.pop("gis_project_clear_notice", None)
    if clear_notice:
        st.success(clear_notice)

    st.markdown("#### ✏️ 工程線／範圍圖資")
    st.caption("既有人工圖資會直接載入下方編輯地圖；直接新增點、線或範圍就是新增圖資。部分繪圖動作為了同步草稿仍會觸發一次 Streamlit 更新，但系統會自動回到原本捲動位置與地圖視窗。重疊線可用地圖左上角「🧭 選取重疊線」逐條切換。")

    if p.project_id:
        manuals = active_project_edit_features(geo, p.project_id)
        orig = original_feature(geo, p.project_id)
        mode = "直接編輯"
        baseline = manuals
        draft_key = f"project::{p.project_id}::unified"
    else:
        manuals = active_project_drafts(project_drafts, p.row_key)
        orig = None
        if p.lon is not None and p.lat is not None:
            orig = {"type":"Feature","geometry":{"type":"Point","coordinates":[p.lon,p.lat]},"properties":{"source":"excel_pending"}}
        mode = "待確認ID圖資草稿"
        baseline = manuals
        draft_key = f"project_pending::{p.row_key}"
        st.warning("此工程尚未取得系統工程ID。您可以先畫圖並儲存『草稿』，但不會出現在正式河道整治總覽；分併標管理者確認 ID 後，系統會把草稿轉成正式 GEO。")
    _draft_init(draft_key, baseline)
    current = _draft_current(draft_key)
    _render_latest_edit_notice(current)

    c1, c2 = st.columns([2,1])
    with c1:
        geo_name = st.text_input("圖資名稱（可留空）", key="gis_v35_geo_name", placeholder="例如：左岸護岸、第一工區、抽水站")
    with c2:
        line_weight = st.slider("線條粗細", min_value=2, max_value=12, value=6, key="gis_v35_line_weight")
    keyword_default = feature_keywords_text(baseline)
    geo_keywords = st.text_input(
        "圖資關鍵字（可輸入多個）",
        value=keyword_default,
        placeholder="例如：六腳排水、六腳排水系統、蒜頭地區（可用逗號或頓號分隔）",
        key=f"gis_v3621_geo_keywords::{p.project_id or p.row_key}::{mode}",
    )
    st.caption("關鍵字只存於 GIS 圖資，不寫入 current.xlsx；河道整治總覽可直接用這些關鍵字查詢。")

    # OSM 參考線：安全小範圍搜尋（固定 20 / 50 / 100m，預設 50m）
    osm_radius = 50
    with st.expander("🌐 使用 OpenStreetMap 河道／護岸參考線", expanded=False):
        st.caption(
            "由您的瀏覽器直接向 OSM Overpass 查詢。搜尋範圍固定為 20m、50m、100m，"
            "預設 50m，且程式端強制限制最大 100m，避免誤載過大範圍。"
            "灰色虛線只作參考，點一下才會複製成藍色可編輯草稿。"
        )
        if p.lat is None or p.lon is None:
            st.info("本工程沒有可用的代表點，因此無法以工程位置搜尋 OSM；請先校正代表點或手動畫線。")
        else:
            osm_radius = st.radio(
                "OSM 搜尋範圍", options=[20, 50, 100], index=1, horizontal=True,
                format_func=lambda x: f"{x}m", key=f"gis_osm_browser_radius_{p.project_id or p.row_key}"
            )
            st.info(
                "設定後請到下方地圖右下角使用『🌐 載入工程附近 OSM』，"
                "或『📍 載入點選位置附近 OSM』。找到合適的灰色候選線並複製成藍色草稿後，"
                "可再按『🧭 延伸相連 OSM 河段（1層）』，只從該線兩端各搜尋 80m、"
                "單次最多 20 條，不會自動遞迴；需要更遠時再從新複製的下一段繼續延伸。"
            )

    center_features = current or manuals or ([orig] if orig else [])
    center, zoom = _center_for_features(center_features)
    em = folium.Map(location=center, zoom_start=zoom, tiles="OpenStreetMap", control_scale=True)
    _add_leaflet_compat_css(em)
    # 非編輯參考層
    if orig:
        _add_feature_shape(em, orig, "#666666", "", reference=True)
    fg = _feature_group_from_drawings(current)
    fg.add_to(em)
    _add_edit_visual_overlays(em, current)
    Draw(
        export=False,
        position="topleft",
        feature_group=fg,
        show_geometry_on_click=False,
        draw_options={"polyline": True, "polygon": True, "marker": True, "rectangle": False, "circle": False, "circlemarker": False},
        edit_options={"edit": True, "remove": True},
    ).add_to(em)
    # V3.5.5a：固定顯示中文工具列，不依賴原生 Leaflet Draw 小圖示是否被 streamlit-folium 正確顯示。
    map_memory_key = f"project::{p.project_id or p.row_key}::{mode}::{osm_radius}"
    BrowserViewportMemory(map_memory_key).add_to(em)
    VisibleDrawToolbar(
        fg, line_weight=line_weight, allow_marker=True, allow_polygon=True,
        tool_memory_key=map_memory_key,
    ).add_to(em)
    if p.lat is not None and p.lon is not None:
        BrowserOSMReferenceControl(
            p.lat, p.lon, radius=osm_radius, line_weight=line_weight, feature_group=fg,
            memory_key=map_memory_key,
        ).add_to(em)
    # V3.6.18：on_change 在 rerun 建地圖前先同步草稿。
    # 若每次畫線後 index 改變就換 key，Streamlit 會把 Folium iframe 當成新元件重建，
    # 造成第一次操作跳離地圖／要做第二次才看得到。一般編輯期間 key 必須穩定。
    render_epoch = int(_draft_state(draft_key).get("render_epoch", 0) or 0)
    component_key = f"gis_v3633_draw_{p.project_id or p.row_key}_{mode}_{render_epoch}_{osm_radius}"
    def _project_map_on_change():
        _sync_folium_component_to_draft(component_key, draft_key)
    _map_kwargs = dict(width=None, height=650, returned_objects=["all_drawings"], key=component_key)
    if _ST_FOLIUM_HAS_ON_CHANGE:
        _map_kwargs["on_change"] = _project_map_on_change
    state = st_folium(em, **_map_kwargs)
    returned = (state or {}).get("all_drawings")
    if returned is not None:
        # st_folium 元件變更本身已觸發本輪 rerun；只需同步草稿，不再額外 rerun 第二次。
        _draft_push(draft_key, returned)
    _render_draft_controls(draft_key)

    b1, b2 = st.columns(2)
    with b1:
        if st.button("💾 儲存工程圖資", type="primary", use_container_width=True, key="gis_v35_save_project"):
            try:
                drawings = _draft_current(draft_key)
                valid_drawings = [d for d in drawings if geometry_type(d) in {"Point", "LineString", "Polygon"}]
                if p.project_id:
                    if not valid_drawings:
                        # V3.6.36：草稿已清空時，不再用「沒有可儲存圖資」擋住。
                        # 只有原本確實有人工圖資時才進二次確認；否則仍視為沒有任何異動可儲存。
                        if manuals:
                            _confirm_clear_project_drawings_dialog(
                                svc, p.project_id, p.project_name, geo_name, geo_keywords,
                                line_weight, editor, draft_key,
                            )
                        else:
                            st.warning("目前沒有任何人工工程圖資，也沒有可儲存的新圖形。")
                    else:
                        _save_mask = _saving_overlay("正在儲存工程圖資…")
                        try:
                            res = save_project_drawings_v35(
                                svc, p.project_id, p.project_name, valid_drawings, geo_name, geo_keywords, line_weight, editor,
                                replace_manual=True,
                            )
                        finally:
                            _close_saving_overlay(_save_mask)
                        msg = "已儲存正式圖資：" + "、".join(res["created_geo_ids"])
                        _draft_clear(draft_key)
                        st.success(msg)
                        st.rerun()
                else:
                    if not valid_drawings:
                        st.warning("目前沒有可儲存的待確認ID圖資草稿。")
                    else:
                        _save_mask = _saving_overlay("正在儲存工程圖資草稿…")
                        try:
                            res = save_pending_project_drafts(
                                svc, p, valid_drawings, geo_name, geo_keywords, line_weight, editor
                            )
                        finally:
                            _close_saving_overlay(_save_mask)
                        msg = f"已儲存待確認ID圖資草稿 {len(res['draft_ids'])} 筆。"
                        _draft_clear(draft_key)
                        st.success(msg)
                        st.rerun()
            except Exception as exc:
                st.error(f"儲存失敗：{exc}")
    with b2:
        if not p.project_id:
            st.caption("待確認ID工程目前只儲存草稿；取得 ID 後才會成為正式圖資。")



def _project_locator_search_text(p: ProjectRow, geo: Dict[str, Any]) -> str:
    """治理現況底圖的資料庫工程快速定位：建立一筆工程可搜尋文字。"""
    parts = [
        p.project_name, p.water_system, p.address, p.unit, p.project_content,
        p.project_type, p.status, p.sheet_name,
    ]
    if p.project_id:
        for f in active_project_edit_features(geo, p.project_id):
            fp = (f or {}).get("properties") or {}
            parts.extend([
                fp.get("keywords", ""), fp.get("geo_name", ""), fp.get("project_name", ""),
                fp.get("water_name", ""), fp.get("reach_name", ""), fp.get("notes", ""),
            ])
    return "|".join(norm_text(x) for x in parts if display_text(x))


def _project_locator_features(p: ProjectRow, geo: Dict[str, Any]) -> List[Dict[str, Any]]:
    """取得用來定位工程的圖形。人工 GIS 圖資優先，其次 GIS 代表點，再其次 Excel 座標。"""
    if p.project_id:
        manuals = active_project_edit_features(geo, p.project_id)
        if manuals:
            return [copy.deepcopy(f) for f in manuals if f.get("geometry")]
        orig = original_feature(geo, p.project_id)
        if orig and orig.get("geometry"):
            return [copy.deepcopy(orig)]
    if p.lon is not None and p.lat is not None and is_valid_wgs84(p.lon, p.lat):
        return [{
            "type": "Feature",
            "geometry": {"type": "Point", "coordinates": [float(p.lon), float(p.lat)]},
            "properties": {
                "project_id": p.project_id,
                "project_name": p.project_name,
                "source": "excel_locator_fallback",
                "role": "locator_point",
            },
        }]
    return []


def _project_locator_status(p: ProjectRow, geo: Dict[str, Any]) -> str:
    fs = _project_locator_features(p, geo)
    if any(geometry_type(f) in {"LineString", "Polygon"} for f in fs):
        return "已有線／範圍"
    if any(geometry_type(f) == "Point" for f in fs):
        if p.project_id and active_project_edit_features(geo, p.project_id):
            return "已有人工作點"
        return "有代表點"
    return "無可定位圖資"


def _search_projects_for_reach_locator(
    rows: Sequence[ProjectRow], geo: Dict[str, Any], county: str, query: str
) -> List[ProjectRow]:
    """縣市內工程快速搜尋。空白不查；多個空白分隔詞採 AND。"""
    q = display_text(query)
    if not q or not county:
        return []
    terms = [norm_text(x) for x in re.split(r"\s+", q) if norm_text(x)]
    out: List[ProjectRow] = []
    for p in rows:
        if p.county != county:
            continue
        hay = _project_locator_search_text(p, geo)
        if terms and all(t in hay for t in terms):
            out.append(p)
    return out


def _locator_project_label(p: ProjectRow, geo: Dict[str, Any]) -> str:
    status = _project_locator_status(p, geo)
    extras = [x for x in [p.water_system, p.address] if display_text(x)]
    extra = "｜" + "｜".join(extras[:2]) if extras else ""
    return f"{p.project_name}｜{status}{extra}"


def _add_locator_reference_feature(m: folium.Map, feature: Dict[str, Any], project_name: str, selected: bool = False) -> None:
    """治理現況編輯的工程定位暫時參考層，不會進入可編輯 FeatureGroup 或正式儲存。"""
    geom = (feature or {}).get("geometry") or {}
    gt = geom.get("type")
    c = geom.get("coordinates")
    color = "#FFEB3B" if selected else "#00BCD4"
    try:
        if gt == "Point" and c and len(c) >= 2:
            lon, lat = float(c[0]), float(c[1])
            folium.CircleMarker(
                [lat, lon], radius=10 if selected else 8,
                color="#FFFFFF" if selected else "#000000", weight=4 if selected else 2.5,
                fill=True, fill_color=color, fill_opacity=0.95,
                tooltip=f"資料庫工程定位：{project_name}",
            ).add_to(m)
        elif gt == "LineString" and c:
            locs = [[float(y), float(x)] for x, y, *_ in c]
            folium.PolyLine(locs, color="#FFFFFF" if selected else "#000000", weight=10 if selected else 8, opacity=.95, interactive=False).add_to(m)
            folium.PolyLine(locs, color=color, weight=6 if selected else 4.5, opacity=.95, tooltip=f"資料庫工程定位：{project_name}").add_to(m)
        elif gt == "Polygon" and c and c[0]:
            locs = [[float(y), float(x)] for x, y, *_ in c[0]]
            folium.Polygon(
                locs, color="#FFFFFF" if selected else "#000000", weight=5 if selected else 3.5,
                fill=True, fill_color=color, fill_opacity=.18,
                tooltip=f"資料庫工程定位：{project_name}",
            ).add_to(m)
    except Exception:
        pass

def _reach_bulk_water_name_change_info(
    selected_features: Sequence[Dict[str, Any]],
    new_water_name: str,
    threshold: int = 3,
) -> Optional[Dict[str, Any]]:
    """治理現況多段編輯的河川／排水名稱大量變更防呆。

    只有「實際會被改名」的既有線達 threshold 條以上才要求二次確認，
    避免正常只修 1～2 段時過度打斷操作。
    """
    new_name = display_text(new_water_name)
    changed_counts: Dict[str, int] = {}
    changed_ids: List[str] = []
    selected_ids: List[str] = []
    for f in selected_features or []:
        p = (f or {}).get("properties") or {}
        rid = display_text(p.get("reach_id"))
        if rid:
            selected_ids.append(rid)
        old_name = display_text(p.get("water_name"))
        if norm_text(old_name) == norm_text(new_name):
            continue
        label = old_name or "（原名稱空白）"
        changed_counts[label] = changed_counts.get(label, 0) + 1
        if rid:
            changed_ids.append(rid)

    changed_count = sum(changed_counts.values())
    if changed_count < max(1, int(threshold)):
        return None

    payload = {
        "changed_count": changed_count,
        "selected_count": len(selected_features or []),
        "old_name_counts": changed_counts,
        "new_name": new_name,
        "changed_ids": sorted(changed_ids),
        "selected_ids": sorted(selected_ids),
    }
    payload["signature"] = hashlib.md5(
        json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()
    return payload


def _format_reach_bulk_water_name_warning(info: Dict[str, Any]) -> str:
    old_counts = info.get("old_name_counts") or {}
    new_name = display_text(info.get("new_name")) or "（空白）"
    total = int(info.get("changed_count") or 0)
    if len(old_counts) == 1:
        old_name, count = next(iter(old_counts.items()))
        detail = f'本次儲存有 **{count} 條「{old_name}」河道線**會修正為 **「{new_name}」**。'
    else:
        rows = [f'- 「{name}」：{count} 條' for name, count in old_counts.items()]
        detail = (
            f'本次儲存將修改 **{total} 條既有河道線**的河川／排水名稱：\n'
            + "\n".join(rows)
            + f'\n\n以上將全部統一為 **「{new_name}」**。'
        )
    return (
        "⚠️ **大量河道名稱變更確認**\n\n"
        + detail
        + "\n\n本次操作會將所有圈選之排水路套用目前輸入的同一個「河川／排水名稱」。"
        + "請確認這是您要執行的修改。"
    )


def _render_reach_edit_tab(svc: EngineeringGISService, rows: Sequence[ProjectRow], geo: Dict[str, Any], reaches: Dict[str, Any], editor: str) -> None:
    st.caption("可先把整段河道／排水畫好或從 OSM 複製；尚未確認治理狀態時可先以藍色『未指定現況』儲存，之後再逐段補登『尚待治理』或『已完成治理』。")
    items = active_reaches(reaches)
    action = st.radio("操作", ["新增治理現況河段", "編輯既有治理現況河段"], horizontal=True, key="gis_reach_action")
    selected_features: List[Dict[str, Any]] = []
    if action == "編輯既有治理現況河段":
        if not items:
            st.info("目前尚無治理現況底圖。")
            return
        opts: Dict[str, Dict[str, Any]] = {}
        for f in items:
            p = f.get("properties") or {}
            rid0 = display_text(p.get("reach_id")) or "無RCH-ID"
            status_label = display_text(p.get("status")) or REACH_STATUS_UNASSIGNED
            label = f"{p.get('water_name') or '未填河川/排水'}｜{status_label}｜{p.get('reach_name') or rid0}｜{rid0}"
            # 理論上 label 已含 RCH-ID；仍防止極端重複。
            if label in opts:
                label = f"{label}｜{len(opts)+1}"
            opts[label] = f
        chosen_labels = st.multiselect(
            "選擇要一起編輯／合併的既有河段（可複選）",
            list(opts.keys()),
            key="gis_reach_multi_select",
            placeholder="可選 1 段進行一般編輯，或選 2 段以上後使用『合併相鄰線段』",
        )
        selected_features = [opts[x] for x in chosen_labels]
        if not selected_features:
            st.info("請至少選擇 1 段既有治理線。若要合併，請一次選取 2 段以上。")
            return
        if len(selected_features) > 1:
            st.success(f"已載入 {len(selected_features)} 段既有治理線，可在下方地圖直接使用『🔗 合併相鄰線段』。")

    selected = selected_features[0] if selected_features else None
    sp = (selected or {}).get("properties") or {}
    selected_ids = [display_text((f.get("properties") or {}).get("reach_id")) for f in selected_features]
    selected_ids = [x for x in selected_ids if x]
    if len(selected_features) > 1:
        # 多段資料的欄位以第一段為預設；若來源資料不一致要明確提醒，避免誤改。
        fields = ("county", "water_name", "basis")
        mixed = [name for name in fields if len({display_text((f.get("properties") or {}).get(name)) for f in selected_features}) > 1]
        if mixed:
            st.warning("目前選取的多段治理線在「" + "、".join(mixed) + "」資料不完全一致；儲存時會以本頁輸入值統一更新選取後的線段。")
    st.info("治理狀態在線上逐段指定：藍色＝未指定現況、紅色＝尚待治理、綠色＝已完成治理。藍色線可先儲存，待確認後再回到『編輯既有治理現況河段』補登狀態。")
    c1, c2 = st.columns(2)
    with c1:
        counties = sorted({p.county for p in rows if p.county})
        county_options = [""] + counties
        current_county = str(sp.get("county", ""))
        if current_county and current_county not in county_options:
            county_options.append(current_county)
        county = st.selectbox("縣市", county_options, index=county_options.index(current_county) if current_county in county_options else 0, key="gis_reach_county")
        water_name = st.text_input("河川／排水名稱", value=str(sp.get("water_name", "")), key="gis_reach_water")
        reach_name = st.text_input("河段名稱（可留空）", value=str(sp.get("reach_name", "")), key="gis_reach_name")
    with c2:
        basis = st.selectbox("治理依據", REACH_BASIS_OPTIONS, index=REACH_BASIS_OPTIONS.index(sp.get("basis")) if sp.get("basis") in REACH_BASIS_OPTIONS else 0, key="gis_reach_basis")
        line_weight = st.slider("線條粗細", 2, 12, int(sp.get("line_weight", 6) or 6), key="gis_reach_weight")
    notes = st.text_area("備註", value=str(sp.get("notes", "")), key="gis_reach_notes")
    reach_keywords = st.text_input(
        "圖資關鍵字（可輸入多個）",
        value=display_text(sp.get("keywords")),
        placeholder="例如：六腳排水、六腳排水系統、蒜頭地區（可用逗號或頓號分隔）",
        key=f"gis_v3621_reach_keywords::{str(sp.get('reach_id') or 'NEW')}::{action}",
    )
    st.caption("關鍵字會跟治理現況線一起存入 GIS，河道整治總覽可直接以關鍵字搜尋。")

    # V3.6.29：資料庫工程快速定位。只做查詢/定位，不會把工程圖資加入治理線草稿。
    locator_query_key = f"gis_reach_locator_query_applied::{county}"
    locator_mode_key = f"gis_reach_locator_mode::{county}"
    locator_selected_key = f"gis_reach_locator_selected::{county}"
    locator_epoch_key = f"gis_reach_locator_epoch::{county}"
    with st.expander("📍 資料庫工程快速定位", expanded=True):
        if not county:
            st.info("請先選擇縣市，再搜尋資料庫工程位置。")
        else:
            with st.form(f"gis_reach_locator_form::{county}"):
                locator_input = st.text_input(
                    "工程／水系／地點關鍵字",
                    value=st.session_state.get(locator_query_key, ""),
                    placeholder="例如：六腳排水、朴子溪、蒜頭、抽水站",
                    key=f"gis_reach_locator_input::{county}",
                )
                locator_submit = st.form_submit_button("🔎 搜尋並顯示全部符合案件", use_container_width=True)
            if locator_submit:
                st.session_state[locator_query_key] = display_text(locator_input)
                st.session_state[locator_mode_key] = "all"
                st.session_state[locator_selected_key] = ""
                st.session_state[locator_epoch_key] = int(st.session_state.get(locator_epoch_key, 0) or 0) + 1
                st.rerun()

            locator_query = display_text(st.session_state.get(locator_query_key, ""))
            locator_results = _search_projects_for_reach_locator(rows, geo, county, locator_query) if locator_query else []
            locatable_results = [p0 for p0 in locator_results if _project_locator_features(p0, geo)]
            if locator_query:
                if not locator_results:
                    st.warning(f"在 {county} 查無符合「{locator_query}」的工程。")
                else:
                    st.success(
                        f"找到 {len(locator_results):,} 件符合工程，其中 {len(locatable_results):,} 件有可定位座標／圖資。"
                    )
                    if len(locator_results) > 100:
                        st.warning("符合案件超過 100 件，建議再增加一個關鍵字縮小範圍。")
                    labels = {}
                    for p0 in locator_results[:100]:
                        label0 = _locator_project_label(p0, geo)
                        if label0 in labels:
                            label0 = f"{label0}｜{p0.sheet_name}#{p0.row}"
                        labels[label0] = p0
                    if labels:
                        current_sel = display_text(st.session_state.get(locator_selected_key, ""))
                        default_index = 0
                        label_list = list(labels.keys())
                        if current_sel:
                            for ii, lab in enumerate(label_list):
                                if labels[lab].row_key == current_sel:
                                    default_index = ii
                                    break
                        chosen_label = st.selectbox(
                            "符合工程", label_list, index=default_index,
                            key=f"gis_reach_locator_result_select::{county}::{locator_query}",
                        )
                        chosen_project = labels[chosen_label]
                        b_loc1, b_loc2, b_loc3 = st.columns([1, 1, 1])
                        with b_loc1:
                            if st.button("📍 定位選取工程", use_container_width=True, key=f"gis_reach_locator_one::{county}"):
                                if not _project_locator_features(chosen_project, geo):
                                    st.warning("此工程目前沒有可用座標或 GIS 圖資，無法定位。")
                                else:
                                    st.session_state[locator_mode_key] = "one"
                                    st.session_state[locator_selected_key] = chosen_project.row_key
                                    st.session_state[locator_epoch_key] = int(st.session_state.get(locator_epoch_key, 0) or 0) + 1
                                    st.rerun()
                        with b_loc2:
                            if st.button("🗺️ 顯示全部符合案件", use_container_width=True, key=f"gis_reach_locator_all::{county}"):
                                st.session_state[locator_mode_key] = "all"
                                st.session_state[locator_selected_key] = ""
                                st.session_state[locator_epoch_key] = int(st.session_state.get(locator_epoch_key, 0) or 0) + 1
                                st.rerun()
                        with b_loc3:
                            if st.button("🧹 清除工程定位", use_container_width=True, key=f"gis_reach_locator_clear::{county}"):
                                st.session_state[locator_query_key] = ""
                                st.session_state[locator_mode_key] = ""
                                st.session_state[locator_selected_key] = ""
                                st.session_state.pop(f"gis_reach_locator_input::{county}", None)
                                st.session_state[locator_epoch_key] = int(st.session_state.get(locator_epoch_key, 0) or 0) + 1
                                st.rerun()
                        st.caption(
                            f"選取：{chosen_project.project_name}｜{_project_locator_status(chosen_project, geo)}"
                            + (f"｜水系：{chosen_project.water_system}" if chosen_project.water_system else "")
                            + (f"｜地點：{chosen_project.address}" if chosen_project.address else "")
                        )
            st.caption("定位圖資只作為暫時參考，不會被存成治理現況線；定位後可直接使用下方 OSM 20m／50m／100m 搜尋。")

    reach_osm_radius = 50
    with st.expander("🌐 使用 OpenStreetMap 河道／排水參考線", expanded=True):
        st.caption(
            "治理現況底圖採『地圖點選位置』的小範圍搜尋。先選搜尋半徑，再到下方地圖左下角按"
            "『📍 點選位置搜尋附近 OSM』，接著在地圖上點一下搜尋中心。"
        )
        reach_osm_radius = st.radio(
            "OSM 搜尋範圍", options=[20, 50, 100], index=1, horizontal=True,
            format_func=lambda x: f"{x}m", key=f"gis_reach_osm_radius_{action}"
        )
        st.caption(
            "預設 50m；最大固定 100m。地圖縮放大小不會改變實際 OSM 查詢範圍。"
            "找到合適候選線並複製成藍色草稿後，可在地圖 OSM 工具按「🧭 延伸相連 OSM 河段（1層）」；"
            "只查該線兩端各 80m、單次最多 20 條，不會自動遞迴。"
        )
        if county:
            st.info(
                f"下方地圖左下角會出現黃色『🌐 OSM 參考水路搜尋』框；可點位置搜尋附近 OSM，"
                f"也可直接在「{county}」內輸入水路名稱搜尋。"
            )
        else:
            st.info("請先選擇縣市。下方地圖左下角會出現黃色『🌐 OSM 參考水路搜尋』框。")

    locator_focus_features: List[Tuple[ProjectRow, Dict[str, Any]]] = []
    locator_query_applied = display_text(st.session_state.get(f"gis_reach_locator_query_applied::{county}", "")) if county else ""
    locator_mode = display_text(st.session_state.get(f"gis_reach_locator_mode::{county}", "")) if county else ""
    locator_selected_row_key = display_text(st.session_state.get(f"gis_reach_locator_selected::{county}", "")) if county else ""
    if county and locator_query_applied:
        _loc_results = _search_projects_for_reach_locator(rows, geo, county, locator_query_applied)
        if locator_mode == "one" and locator_selected_row_key:
            _loc_results = [p0 for p0 in _loc_results if p0.row_key == locator_selected_row_key]
        for p0 in _loc_results[:100]:
            for f0 in _project_locator_features(p0, geo):
                locator_focus_features.append((p0, f0))

    baseline = []
    if selected_features:
        for sf in selected_features:
            base_feature = copy.deepcopy(sf)
            bp = base_feature.setdefault("properties", {})
            bp["reach_status_draft"] = str(bp.get("status") or "")
            bp["reach_id_draft"] = str(bp.get("reach_id") or "")
            baseline.append(base_feature)
    rid = "MULTI-" + hashlib.md5("|".join(sorted(selected_ids)).encode("utf-8")).hexdigest()[:10] if len(selected_ids) > 1 else str(sp.get("reach_id", "NEW"))
    draft_key = f"reach::{rid}::{action}"
    _draft_init(draft_key, baseline)
    current = _draft_current(draft_key)
    _render_latest_edit_notice(current)
    center, zoom = _center_for_features(current or ([selected] if selected else []))
    if not current and selected is None and county:
        county_points = [(float(pr.lat), float(pr.lon)) for pr in rows
                         if pr.county == county and pr.lat is not None and pr.lon is not None]
        if county_points:
            center = [
                sum(pt[0] for pt in county_points) / len(county_points),
                sum(pt[1] for pt in county_points) / len(county_points),
            ]
            zoom = 11
    em = folium.Map(location=center, zoom_start=zoom, tiles="OpenStreetMap", control_scale=True)
    _add_leaflet_compat_css(em)
    # 其他治理河段作參考，不進 editable FeatureGroup
    selected_id_set = set(selected_ids)
    for f in items:
        p = f.get("properties") or {}
        if display_text(p.get("reach_id")) in selected_id_set:
            continue
        _add_feature_shape(em, f, "#999999", "", reference=True, tooltip=str(p.get("water_name", "")))

    # V3.6.29：資料庫工程定位參考層。青色＝全部搜尋結果；黃色/白框＝單筆定位。
    locator_bounds: List[List[float]] = []
    locator_one = locator_mode == "one" and bool(locator_selected_row_key)
    for lp, lf in locator_focus_features:
        _add_locator_reference_feature(em, lf, lp.project_name, selected=locator_one)
        locator_bounds.extend(_feature_all_latlon_points(lf))
    if locator_bounds:
        try:
            if len(locator_bounds) > 1:
                em.fit_bounds(locator_bounds, padding=(35, 35), max_zoom=17)
            else:
                em.location = locator_bounds[0]
                em.options["zoom"] = 17
        except TypeError:
            em.fit_bounds(locator_bounds, padding=(35, 35))
        except Exception:
            pass

    fg = _feature_group_from_drawings(current, name="治理現況底圖編輯", reach_status_edit=True)
    fg.add_to(em)
    _add_edit_visual_overlays(em, current)
    Draw(
        export=False, position="topleft", feature_group=fg, show_geometry_on_click=False,
        draw_options={"polyline": True, "polygon": False, "marker": False, "rectangle": False, "circle": False, "circlemarker": False},
        edit_options={"edit": True, "remove": True},
    ).add_to(em)
    locator_epoch = int(st.session_state.get(f"gis_reach_locator_epoch::{county}", 0) or 0) if county else 0
    map_memory_key = f"reach::{rid}::{action}::{county}::{reach_osm_radius}::locator{locator_epoch}"
    BrowserViewportMemory(map_memory_key).add_to(em)
    VisibleDrawToolbar(
        fg, line_weight=line_weight, allow_marker=False, allow_polygon=False, allow_reach_status=True,
        tool_memory_key=map_memory_key,
    ).add_to(em)
    BrowserOSMReferenceControl(
        center[0], center[1], radius=reach_osm_radius, line_weight=line_weight,
        feature_group=fg, click_to_select=True, county=county, memory_key=map_memory_key
    ).add_to(em)
    # V3.6.18：on_change 預同步＋瀏覽器記住 viewport／OSM 候選線／治理狀態模式。
    # OSM 複製、畫線、截斷、治理狀態指定都由瀏覽器事件觸發；事件本身只推進草稿，
    # 不應因 index 改變而重建整個 iframe。Undo/Redo 才透過 render_epoch 強制重畫。
    render_epoch = int(_draft_state(draft_key).get("render_epoch", 0) or 0)
    component_key = f"gis_reach_draw_v3633_{rid}_{render_epoch}_{county}_{reach_osm_radius}_loc{locator_epoch}"
    def _reach_map_on_change():
        _sync_folium_component_to_draft(component_key, draft_key)
    _map_kwargs = dict(width=None, height=650, returned_objects=["all_drawings"], key=component_key)
    if _ST_FOLIUM_HAS_ON_CHANGE:
        _map_kwargs["on_change"] = _reach_map_on_change
    state = st_folium(em, **_map_kwargs)
    returned = (state or {}).get("all_drawings")
    if returned is not None:
        # st_folium 元件變更本身已觸發本輪 rerun；只需同步草稿，不再額外 rerun 第二次。
        _draft_push(draft_key, returned)
    _render_draft_controls(draft_key)

    current_now = _draft_current(draft_key)
    reach_lines_now = [d for d in current_now if geometry_type(d) == "LineString"]
    pending_count = 0
    completed_count = 0
    unassigned_count = 0
    for d in reach_lines_now:
        dp = d.get("properties") or {}
        ds = str(dp.get("reach_status_draft") or dp.get("status") or "").strip()
        if ds == "尚待治理":
            pending_count += 1
        elif ds == "已完成治理":
            completed_count += 1
        else:
            unassigned_count += 1
    m1, m2, m3 = st.columns(3)
    m1.metric("🔴 尚待治理", pending_count)
    m2.metric("🟢 已完成治理", completed_count)
    m3.metric("🔵 未指定現況", unassigned_count)
    if unassigned_count:
        st.warning(
            f"目前有 {unassigned_count} 段藍色線尚未確認治理狀態；"
            "可先儲存為「未指定現況」，之後再編輯補登為尚待治理或已完成治理。"
        )
    elif reach_lines_now:
        st.success("所有線段都已指定治理狀態，可以儲存。")

    # V3.6.43：多段既有治理線若有 3 條以上會被改成不同的河川／排水名稱，
    # 儲存前必須二次確認，避免誤把整條大排水大量覆寫成錯誤名稱。
    rename_info = (
        _reach_bulk_water_name_change_info(selected_features, water_name, threshold=3)
        if action == "編輯既有治理現況河段" and selected_features
        else None
    )
    rename_confirm_key = f"gis_v3643_reach_rename_confirm::{rid}"
    rename_signature = display_text((rename_info or {}).get("signature"))
    stored_signature = display_text(st.session_state.get(rename_confirm_key))
    if stored_signature and stored_signature != rename_signature:
        st.session_state.pop(rename_confirm_key, None)
        stored_signature = ""
    rename_confirm_pending = bool(rename_info and stored_signature == rename_signature)

    def _commit_reach_save_v3643() -> None:
        try:
            _save_mask = _saving_overlay("正在儲存治理現況底圖…")
            try:
                res = save_reach_features_batch(
                    svc, current_now, water_name, reach_name, county, basis, notes, reach_keywords, line_weight, editor,
                    existing_reach_id="" if selected is None else str(sp.get("reach_id", "")),
                    existing_reach_ids=selected_ids,
                )
            finally:
                _close_saving_overlay(_save_mask)
            st.session_state.pop(rename_confirm_key, None)
            _draft_clear(draft_key)
            created = len(res.get("created_reach_ids") or [])
            updated = len(res.get("updated_reach_ids") or [])
            deactivated = len(res.get("deactivated_reach_ids") or [])
            saved_unassigned = int(res.get("unassigned_count") or 0)
            st.success(
                f"已儲存 {res.get('saved_count', 0)} 段治理現況底圖"
                f"（更新 {updated} 段、新增 {created} 段、停用合併前舊段 {deactivated} 段）"
                + (
                    f"；其中 {saved_unassigned} 段為藍色『未指定現況』，可日後再補登。"
                    if saved_unassigned else "。"
                )
            )
            st.rerun()
        except Exception as exc:
            st.error(f"儲存失敗：{exc}")

    if rename_confirm_pending:
        st.warning(_format_reach_bulk_water_name_warning(rename_info))
        c_ok, c_cancel = st.columns(2)
        with c_ok:
            if st.button(
                "✅ 確認名稱修改並儲存",
                type="primary",
                use_container_width=True,
                key=f"gis_v3643_confirm_reach_rename::{rid}",
            ):
                _commit_reach_save_v3643()
        with c_cancel:
            if st.button(
                "↩️ 取消，回去檢查名稱",
                use_container_width=True,
                key=f"gis_v3643_cancel_reach_rename::{rid}",
            ):
                st.session_state.pop(rename_confirm_key, None)
                st.rerun()
    else:
        if rename_info:
            st.warning(
                f"偵測到本次有 {rename_info['changed_count']} 條既有治理線的「河川／排水名稱」"
                "會被大量修改；按儲存後會先顯示完整清單，必須再次確認才會真正寫入。"
            )

        b1, b2 = st.columns(2)
        with b1:
            if st.button("💾 儲存治理現況底圖", type="primary", use_container_width=True, key="gis_save_reach"):
                if rename_info:
                    st.session_state[rename_confirm_key] = rename_signature
                    st.rerun()
                else:
                    _commit_reach_save_v3643()
        with b2:
            if len(selected_features) == 1:
                if st.button("🗑️ 停用此治理現況河段", use_container_width=True, key="gis_delete_reach"):
                    _confirm_delete_reach_dialog(svc, str(sp.get("reach_id")), editor)
            elif len(selected_features) > 1:
                st.caption("多段共同編輯時不提供批次停用；請個別選取河段後再停用，避免誤刪。")


def _render_split_merge_tab(svc: EngineeringGISService, rows: Sequence[ProjectRow], geo: Dict[str, Any], history: Dict[str, Any], registry: Dict[str, Any], editor: str) -> None:
    blank_rows = [p for p in rows if not p.project_id]
    assigned_rows = [p for p in rows if p.project_id]
    c1,c2,c3 = st.columns(3)
    c1.metric("工程資料筆數", f"{len(rows):,}")
    c2.metric("已有系統工程ID", f"{len(assigned_rows):,}")
    c3.metric("待確認／編號", f"{len(blank_rows):,}")
    st.caption("第一次初始化會整批建立既有案件的 PRJ-ID；日後空白 ID 才在這裡判斷新核定、分標或併標。工程沿革由系統自動產生，不提供直接修改。")
    current_ids = {p.project_id for p in rows if p.project_id}
    known_all = all_known_projects(rows, geo, history, registry)
    historical_ids = [
        pid for pid in known_all
        if PROJECT_ID_RE.match(display_text(pid)) and pid not in current_ids
    ]
    consumed_ids = history_consumed_source_ids(history)
    eligible_history = eligible_historical_source_projects(rows, geo, history, registry)
    st.caption(
        f"永久工程識別目錄目前可辨識 {len(known_all):,} 件；"
        f"其中 {len(historical_ids):,} 件已不在目前管控表，"
        f"{len(set(historical_ids) & consumed_ids):,} 件已完成過分併標而不再列入來源；"
        f"目前可選歷史來源 {len(eligible_history):,} 件。"
    )

    with st.expander("📥 找不到舊母工程？從舊版管控表補回歷史工程目錄", expanded=False):
        st.write("只匯入舊表中的『系統工程ID、工程名稱、縣市／工作表』到 id_registry.json；不會修改目前 current.xlsx，也不會把舊工程重新加入管控表。")
        old_catalog_file = st.file_uploader(
            "選擇仍保留分標前／併標前案件的舊版 Excel",
            type=["xlsx", "xlsm"],
            key="gis_history_catalog_upload",
        )
        if old_catalog_file is not None:
            if st.button("匯入歷史工程目錄", type="primary", key="gis_import_history_catalog"):
                try:
                    _save_mask = _saving_overlay("正在匯入歷史工程目錄…")
                    try:
                        res = svc.import_historical_project_catalog(old_catalog_file.getvalue(), old_catalog_file.name)
                    finally:
                        _close_saving_overlay(_save_mask)
                    st.success(
                        f"匯入完成：舊表找到 {res['rows_with_id']:,} 筆有效ID；"
                        f"新增歷史ID {res['new_ids']:,} 筆；目前工程識別目錄共 {res['catalog_total']:,} 件。"
                    )
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

    if not assigned_rows:
        st.warning("目前尚未建立任何系統工程ID。")
        if st.button("🚀 第一次初始化：自動建立全部 PRJ-ID 與原始點位", type="primary"):
            try:
                _save_mask = _saving_overlay("正在初始化工程 PRJ-ID 與原始點位…")
                try:
                    res = svc.initialize_existing_projects()
                finally:
                    _close_saving_overlay(_save_mask)
                st.success(f"初始化完成：{res['projects']:,} 件工程；建立 {res['assigned']:,} 個 PRJ-ID。")
                st.rerun()
            except Exception as exc:
                st.error(str(exc))
        return

    if st.button("🔄 同步工程名稱、執行情形及 Excel 原始座標"):
        try:
            _save_mask = _saving_overlay("正在同步工程名稱、執行情形與原始座標…")
            try:
                res = svc.sync_current_projects()
            finally:
                _close_saving_overlay(_save_mask)
            st.success(f"同步完成：{res['projects']:,} 件工程。")
            st.rerun()
        except Exception as exc:
            st.error(str(exc))

    if blank_rows:
        st.divider()
        row_opts = {_row_option_label(p): p for p in blank_rows}
        action = st.radio(
            "待編號工程屬於",
            ["新核定工程", "既有工程分標（1→多）", "既有工程併標（多→1）", "分併標重組（多→多）"],
            horizontal=True,
            key="gis_v35_id_action",
        )
        if action == "新核定工程":
            labels = st.multiselect("選擇本次新核定工程（可一次多選整批）", list(row_opts.keys()), key="gis_v35_new_ids")
            st.caption("若一整批案件的核定屬性／核定年度明顯是新批次，可一次全選後整批發號；系統不會要求逐件輸入 ID。")
            if st.button("整批建立新 PRJ-ID", disabled=not labels, type="primary"):
                try:
                    selected = [(row_opts[x].row_key, row_opts[x].project_name) for x in labels]
                    _save_mask = _saving_overlay("正在建立新工程 PRJ-ID…")
                    try:
                        res = svc.assign_new_projects(selected)
                    finally:
                        _close_saving_overlay(_save_mask)
                    mapping = {rk: pair for (rk, _), pair in zip(selected, res["assigned"])}
                    promoted = promote_project_drafts(svc, mapping, editor)
                    st.success(f"已建立 {len(res['assigned'])} 個新工程ID；同步正式化 {promoted.get('promoted',0)} 筆圖資草稿。")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
        else:
            # V3.6.39：來源只顯示「歷史上有 PRJ-ID、目前 current.xlsx 已不存在、
            # 且尚未完成過分併標」的案件，不再把目前 Excel 全部工程塞入選單。
            hist_known = eligible_historical_source_projects(rows, geo, history, registry)
            hist_opts = {
                _known_option_with_history(pid, name, current_ids): pid
                for pid, name in sorted(hist_known.items())
            }

            if not hist_opts:
                st.warning(
                    "目前沒有可選的歷史來源工程。只有『已有 PRJ-ID、目前 current.xlsx 已移除、"
                    "而且尚未完成過分併標』的案件會出現在這裡。"
                )
            elif action == "既有工程分標（1→多）":
                parent = st.selectbox("分標前歷史母工程", list(hist_opts.keys()), key="gis_v35_split_parent")
                labels = st.multiselect("分標後的新工程列（至少2筆）", list(row_opts.keys()), key="gis_v35_split_rows")
                st.caption("適用 1 標 → 2 標以上；新工程沿用母工程的子 ID 編號。")
                if st.button("建立分標子ID", disabled=len(labels)<2, type="primary"):
                    try:
                        sel = [(row_opts[x].row_key, row_opts[x].project_name) for x in labels]
                        _save_mask = _saving_overlay("正在建立分標工程 ID…")
                        try:
                            res = svc.split_project(hist_opts[parent], sel)
                        finally:
                            _close_saving_overlay(_save_mask)
                        mapping = {rk: pair for (rk, _), pair in zip(sel, res["children"])}
                        promoted = promote_project_drafts(svc, mapping, editor)
                        st.success("分標完成：" + "、".join(x[0] for x in res["children"]) + f"；正式化草稿 {promoted.get('promoted',0)} 筆")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
            elif action == "既有工程併標（多→1）":
                src = st.multiselect("併標前歷史來源工程（至少2件）", list(hist_opts.keys()), key="gis_v35_merge_sources")
                target = st.selectbox("併標後的新工程列", list(row_opts.keys()), key="gis_v35_merge_target")
                st.caption("適用 2 標以上 → 1 標；併標後工程建立新的根 PRJ-ID。")
                if st.button("建立併標新根ID", disabled=len(src)<2, type="primary"):
                    try:
                        t = row_opts[target]
                        _save_mask = _saving_overlay("正在建立併標工程 ID…")
                        try:
                            res = svc.merge_projects([hist_opts[x] for x in src], (t.row_key, t.project_name))
                        finally:
                            _close_saving_overlay(_save_mask)
                        promoted = promote_project_drafts(svc, {t.row_key: (res["new_id"], t.project_name)}, editor)
                        st.success(f"併標完成，新工程ID：{res['new_id']}；正式化草稿 {promoted.get('promoted',0)} 筆")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
            else:
                src = st.multiselect(
                    "重組前歷史來源工程（至少2件）",
                    list(hist_opts.keys()),
                    key="gis_v639_restructure_sources",
                )
                targets = st.multiselect(
                    "重組後的新工程列（至少2筆）",
                    list(row_opts.keys()),
                    key="gis_v639_restructure_targets",
                )
                st.caption(
                    "適用多對多，例如 2標→3標、5標→3標。因沒有單一母工程，"
                    "每一筆重組後工程都會建立新的根 PRJ-ID，沿革則保存整組來源與整組結果。"
                )
                if src and targets:
                    st.info(f"本次重組：{len(src)} 標 → {len(targets)} 標")
                if st.button(
                    "建立多對多分併標重組",
                    disabled=(len(src)<2 or len(targets)<2),
                    type="primary",
                    key="gis_v639_restructure_submit",
                ):
                    try:
                        selected_sources = [hist_opts[x] for x in src]
                        selected_targets = [(row_opts[x].row_key, row_opts[x].project_name) for x in targets]
                        _save_mask = _saving_overlay("正在建立多對多分併標重組…")
                        try:
                            res = svc.restructure_projects(selected_sources, selected_targets)
                        finally:
                            _close_saving_overlay(_save_mask)
                        mapping = {
                            row_key: (new_id, name)
                            for row_key, new_id, name in res["targets"]
                        }
                        promoted = promote_project_drafts(svc, mapping, editor)
                        new_ids = "、".join(new_id for _rk, new_id, _name in res["targets"])
                        st.success(
                            f"分併標重組完成：{len(res['sources'])} 標 → {len(res['targets'])} 標；"
                            f"新工程ID：{new_ids}；正式化草稿 {promoted.get('promoted',0)} 筆"
                        )
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
    else:
        st.success("目前沒有待編號工程。")

    st.divider()
    st.markdown("#### 工程沿革查詢（唯讀）")
    events = list(reversed(history.get("events", [])))
    if not events:
        st.info("目前尚無分標／併標沿革。")
    else:
        keyword = st.text_input("搜尋 PRJ-ID／工程名稱", key="gis_history_keyword")
        if keyword.strip():
            k = norm_text(keyword)
            events = [e for e in events if k in norm_text(json.dumps(e, ensure_ascii=False))]
        st.dataframe(events, hide_index=True, use_container_width=True)
    st.caption("這裡是『工程分併標業務沿革』，不是 GitHub 歷史版本。GitHub 歷史恢復不在網站提供。")



@st.cache_data(ttl=45, max_entries=4, show_spinner=False)
def _cached_gis_page_snapshot_v351(
    owner: str,
    repo: str,
    branch: str,
    excel_path: str,
    geo_path: str,
    history_path: str,
    registry_path: str,
    reach_path: str,
    draft_path: str,
    _token: str,
):
    """一次讀取 GIS 頁面所需 GitHub 檔案；45 秒內切頁不重複下載。"""
    cfg = GitHubSettings(
        owner=owner,
        repo=repo,
        branch=branch,
        token=_token,
        excel_path=excel_path,
        geo_path=geo_path,
        history_path=history_path,
        registry_path=registry_path,
    )
    store = GitHubRepoStore(cfg)
    commit, _ = store.get_head_commit()
    # V3.5.3：同一 commit 下的 6 個檔案平行讀取，避免逐檔等待網路。
    paths = [excel_path, geo_path, history_path, registry_path, reach_path, draft_path]
    def _read_one(path):
        return store.read_file(path, ref=commit, allow_missing=(path != excel_path))
    with ThreadPoolExecutor(max_workers=6) as pool:
        vals = list(pool.map(_read_one, paths))
    excel, geo_b, history_b, registry_b, reach_b, draft_b = vals
    if excel is None:
        raise GitHubError(f"找不到 {excel_path}")
    geo = json_load_bytes(geo_b, empty_geojson())
    history = json_load_bytes(history_b, empty_history())
    registry = json_load_bytes(registry_b, empty_registry())
    reaches = json_load_bytes(reach_b, empty_reaches())
    drafts = json_load_bytes(draft_b, empty_project_drafts())
    return excel, geo, history, registry, reaches, drafts


@st.cache_data(ttl=45, max_entries=4, show_spinner=False)
def _cached_scan_rows_v351(excel_bytes: bytes, excel_path: str):
    """OpenPyXL 掃描工程列也快取，避免每個 widget rerun 都重讀整本 Excel。"""
    _, rows, _ = scan_workbook(excel_bytes, excel_path, ensure_id_cols=False)
    return rows


@st.cache_data(ttl=45, max_entries=4, show_spinner=False)
def _cached_query_geo_snapshot_v351(
    owner: str,
    repo: str,
    branch: str,
    geo_path: str,
    _token: str,
):
    cfg = GitHubSettings(owner=owner, repo=repo, branch=branch, token=_token, geo_path=geo_path)
    store = GitHubRepoStore(cfg)
    commit, _ = store.get_head_commit()
    return json_load_bytes(store.read_file(geo_path, ref=commit, allow_missing=True), empty_geojson())


def render_engineering_gis_page():
    # V3.6.59：先補回上一輪儲存遮罩，再安裝整頁 scroll／地圖錨點記憶。
    # 點工程後需要主動捲到下方編輯區時，該輪暫停一般位置恢復，避免互相拉扯。
    _resume_busy_overlay_if_needed()
    _force_editor_scroll = bool(st.session_state.get("gis_v641_force_editor_scroll"))
    _install_gis_scroll_memory(restore=not _force_editor_scroll)

    st.title("🗺️ 工程空間圖資")
    st.caption("公開瀏覽｜工程圖資｜座標品質檢核｜縣市座標回填審核｜治理現況底圖｜重複工程共用圖資｜工程分併標｜可切換帳號權限")
    if st.button("🔄 重新載入 GitHub 最新圖資", key="gis_v366_force_refresh", help="只有需要取得其他使用者剛更新的資料時才按；平常切縣市不需要重新載入。"):
        _clear_gis_read_caches()
        st.rerun()

    try:
        settings = GitHubSettings.from_streamlit()
        store = GitHubRepoStore(settings)
        svc = EngineeringGISService(store)

        # V3.5.1：不再每次進頁面都執行 bootstrap/多次 GitHub 檢查；
        # 缺少的 JSON 先以空集合顯示，第一次正式寫入時自然建立。
        # V3.6.6：GIS 主資料在本次瀏覽 Session 常駐。
        # 切縣市、點工程、Undo/Redo 都不再反覆從 GitHub/快取反序列化整套資料。
        snap = _get_gis_session_snapshot_v366(settings)
        excel_bytes = snap["excel_bytes"]
        geo = snap["geo"]
        history = snap["history"]
        registry = snap["registry"]
        reaches = snap["reaches"]
        project_drafts = snap["project_drafts"]
        rows = snap["rows"]
    except Exception as exc:
        st.error(f"工程空間圖資初始化失敗：{exc}")
        _release_busy_overlay_when_stable()
        return

    duplicates = validate_duplicate_ids(rows)
    if duplicates:
        st.error("偵測到重複系統工程ID，為避免圖資錯接，編輯功能已暫停。")

    security_config, security_store = load_security_config()
    identity, perms = identity_and_permissions(security_config)
    if admin_unlocked():
        perms = {"edit_reach": True, "edit_project_geo": True, "manage_split_merge": True}
        identity = {"user_id": "MASTER", "display_name": "系統管理者", "role": "ADMIN", "anonymous": False}

    with st.expander("👤 圖資編輯身分與權限", expanded=False):
        render_login_panel(security_config)
        identity, perms = identity_and_permissions(security_config)
        if admin_unlocked():
            perms = {"edit_reach": True, "edit_project_geo": True, "manage_split_merge": True}
            identity = {"user_id": "MASTER", "display_name": "系統管理者", "role": "ADMIN", "anonymous": False}
        if identity.get("role") in ROLE_LABELS:
            st.caption(f"目前身分：{identity.get('display_name')}｜{ROLE_LABELS.get(identity.get('role'))}")

    # 只重整功能階層；各功能原有名稱、內容、顯示及權限流程均不變。
    main_section = st.radio(
        "圖資功能分類",
        ["🗺️ 圖資總覽", "📍 座標作業", "✏️ 圖資編輯", "🔗 工程關聯", "⚙️ 系統管理"],
        horizontal=True,
        key="gis_main_section_v3653",
        label_visibility="collapsed",
    )

    if main_section == "🗺️ 圖資總覽":
        subpage = "河道整治總覽"
    elif main_section == "📍 座標作業":
        subpage = st.radio(
            "座標作業",
            ["座標品質檢核", "縣市座標回填審核"],
            horizontal=True,
            key="gis_coordinate_subpage_v3653",
            label_visibility="collapsed",
        )
    elif main_section == "✏️ 圖資編輯":
        subpage = st.radio(
            "圖資編輯",
            ["工程圖資編輯", "治理現況底圖編輯"],
            horizontal=True,
            key="gis_edit_subpage_v3653",
            label_visibility="collapsed",
        )
    elif main_section == "🔗 工程關聯":
        subpage = st.radio(
            "工程關聯",
            ["工程分併標管理", "重複工程與共用圖資整理"],
            horizontal=True,
            key="gis_relation_subpage_v3653",
            label_visibility="collapsed",
        )
    else:
        subpage = "系統管理"

    editor_name = str(identity.get("display_name") or identity.get("user_id") or "未知使用者")
    edit_blocked = bool(duplicates)

    if subpage == "河道整治總覽":
        _render_overview_tab(rows, geo, reaches)

    elif subpage == "工程圖資編輯":
        if edit_blocked or not perms.get("edit_project_geo"):
            st.info("此功能需要『編輯者』以上權限。自由編輯模式下可直接使用；帳號模式請以有權限帳號登入。")
        else:
            _render_project_edit_tab(svc, rows, geo, project_drafts, editor_name)

    elif subpage == "座標品質檢核":
        if edit_blocked or not perms.get("edit_project_geo"):
            st.info("此功能需要『編輯者』以上權限。")
        else:
            _render_coordinate_quality_tab(
                svc, rows, geo, registry, excel_bytes, editor_name
            )

    elif subpage == "縣市座標回填審核":
        if edit_blocked or not perms.get("edit_project_geo"):
            st.info("此功能需要『編輯者』以上權限。")
        else:
            _render_county_coordinate_return_review_tab(
                svc, rows, geo, registry, editor_name
            )

    elif subpage == "治理現況底圖編輯":
        if edit_blocked or not perms.get("edit_reach"):
            st.info("此功能需要『初階編輯者』以上權限。")
        else:
            _render_reach_edit_tab(svc, rows, geo, reaches, editor_name)

    elif subpage == "重複工程與共用圖資整理":
        if edit_blocked or not perms.get("manage_split_merge"):
            st.info(
                "重複工程與共用圖資整理會影響多筆工程的空間關聯，"
                "需要『進階編輯者』以上權限。"
            )
        else:
            _render_duplicate_spatial_project_tab_v3651(
                svc, rows, geo, registry, editor_name
            )

    elif subpage == "工程分併標管理":
        if edit_blocked or not perms.get("manage_split_merge"):
            st.info("工程分併標管理屬高風險功能，需要『進階編輯者』以上權限。自由編輯模式不會自動開放此功能。")
        else:
            _render_split_merge_tab(svc, rows, geo, history, registry, editor_name)

    elif subpage == "系統管理":
        render_admin_panel(security_config, security_store)
        if admin_unlocked():
            st.divider()
            st.markdown("#### 📍 GIS代表點同步至 current.xlsx")
            st.caption(
                "平常工程代表點以 GIS 為最新座標。只有在這裡由系統管理者確認後，"
                "才會一次把差異座標寫回 current.xlsx；同步只修改座標儲存格。"
            )
            diffs = representative_point_excel_differences(rows, geo, tolerance_m=1.0)
            total_gis_points = sum(1 for p in rows if p.project_id and original_feature(geo, p.project_id))
            county_return_pending = sum(
                1 for d in diffs
                if d.get("coordinate_edit_source") == "county_return_review"
                or d.get("coord_source") == "縣市回填核准"
            )
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("GIS代表點", f"{total_gis_points:,}")
            mc2.metric("待同步至Excel", f"{len(diffs):,}")
            mc3.metric("其中縣市回填核准", f"{county_return_pending:,}")
            mc4.metric("已無差異", f"{max(total_gis_points-len(diffs),0):,}")

            if diffs:
                preview = []
                for d in diffs:
                    preview.append({
                        "系統工程ID": d["project_id"],
                        "縣市": d["county"],
                        "工程名稱": d["project_name"],
                        "Excel經度E": d["excel_lon"],
                        "Excel緯度N": d["excel_lat"],
                        "GIS經度E": d["gis_lon"],
                        "GIS緯度N": d["gis_lat"],
                        "差異(m)": "無Excel座標" if d["distance_m"] is None else round(float(d["distance_m"]),1),
                        "SPJ-ID": d.get("spatial_project_id"),
                        "共用圖資主PRJ": d.get("spatial_primary_project_id"),
                        "GIS座標來源": d.get("coord_source"),
                        "GIS同步狀態": d.get("coord_status"),
                        "GIS最後編輯者": d["updated_by"],
                        "GIS最後編輯時間": d["updated_at"],
                    })
                with st.expander(f"查看 {len(preview):,} 件待同步差異", expanded=False):
                    st.dataframe(preview, hide_index=True, use_container_width=True, height=420)

                confirm_sync = st.checkbox(
                    f"我已確認，準備把上述 {len(diffs):,} 件 GIS 代表點同步至 current.xlsx",
                    key="gis_v368_admin_confirm_excel_sync",
                )
                if st.button(
                    f"📥 一次同步 {len(diffs):,} 件代表點至 current.xlsx",
                    type="primary", use_container_width=True,
                    disabled=not confirm_sync, key="gis_v368_admin_sync_excel",
                ):
                    try:
                        ids = [d["project_id"] for d in diffs]
                        _save_mask = _saving_overlay(
                            f"正在同步 {len(ids):,} 件 GIS 代表點至 current.xlsx…"
                        )
                        try:
                            result = svc.sync_gis_representative_points_to_excel(ids, editor_name)
                        finally:
                            _close_saving_overlay(_save_mask)
                        st.success(f"已完成 {result.get('count',0):,} 件代表點同步至 current.xlsx。")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"GIS代表點同步至 Excel 失敗：{exc}")
            else:
                st.success("目前 GIS 代表點與 current.xlsx 座標沒有待同步差異。")

    # 若本輪是儲存後的重建，直到所有目前頁面元件都送出後才開始計時移除中央遮罩。
    _release_busy_overlay_when_stable()


# 覆寫 V3.4 的 ID 欄建立函式：系統工程ID固定加在最後一欄並隱藏。
def ensure_id_column(ws, header_row: int) -> int:
    from openpyxl.comments import Comment
    hmap = normalize_header_map(ws, header_row)
    existing = find_col(hmap, [PROJECT_ID_HEADER])
    if existing:
        letter = get_column_letter(existing)
        ws.column_dimensions[letter].hidden = True
        return existing

    c = ws.max_column + 1
    dst = ws.cell(header_row, c)
    dst.value = PROJECT_ID_HEADER
    if c > 1:
        src = ws.cell(header_row, c - 1)
        if src.has_style:
            dst._style = copy.copy(src._style)
        if src.font:
            dst.font = copy.copy(src.font)
        if src.fill:
            dst.fill = copy.copy(src.fill)
        if src.border:
            dst.border = copy.copy(src.border)
        if src.alignment:
            dst.alignment = copy.copy(src.alignment)
    dst.comment = Comment("系統專用欄位，請勿刪除、清空或自行修改。", "WRA GIS")
    letter = get_column_letter(c)
    ws.column_dimensions[letter].width = 18
    ws.column_dimensions[letter].hidden = True
    return c

if __name__ == "__main__":
    st.set_page_config(page_title="工程空間圖資", layout="wide")
    render_engineering_gis_page()
