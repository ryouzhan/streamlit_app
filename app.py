# -*- coding: utf-8 -*-
"""
发货单处理与物流发票批量生成工具 (Modern Clean Business Edition)
放弃脆弱的主题强行覆盖，采用原生高质感现代商务风，彻底告别错位与黑块。
"""
import io
import os
import re
import copy
import zipfile
import requests
from datetime import datetime
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

# ===================== 全版本向下兼容安全垫 =====================
def safe_rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.experimental_rerun()

def safe_data_editor(df, key=None):
    if hasattr(st, "data_editor"):
        try:
            return st.data_editor(df, num_rows="dynamic", key=key)
        except TypeError:
            return st.data_editor(df, key=key)
    elif hasattr(st, "experimental_data_editor"):
        try:
            return st.experimental_data_editor(df, num_rows="dynamic", key=key)
        except TypeError:
            return st.experimental_data_editor(df, key=key)
    else:
        st.dataframe(df)
        return df

# ===================== 页面全局配置 =====================
st.set_page_config(
    page_title="发货单与物流发票协同工具",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ===================== 极简非侵入式现代排版 CSS =====================
st.markdown("""
<style>
    /* 移除所有强行改变系统底色的 !important 恶性覆盖 */
    #MainMenu, footer, header {visibility: hidden;}

    /* 卡片容器 */
    .metric-card-box {
        background-color: rgba(255, 255, 255, 0.85);
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.1rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
        margin-bottom: 0.5rem;
    }
    .metric-card-title {
        font-size: 0.82rem;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        margin-bottom: 0.25rem;
    }
    .metric-card-val {
        font-size: 1.65rem;
        font-weight: 700;
        color: #0F172A;
    }
    .metric-card-unit {
        font-size: 0.85rem;
        color: #94A3B8;
        font-weight: 500;
        margin-left: 0.25rem;
    }

    /* 按钮微调，保持专业利落 */
    div.stButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 0.55rem 1.4rem !important;
    }
    div.stDownloadButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 0.65rem 1.8rem !important;
    }
</style>
""", unsafe_allow_html=True)

# ===================== 发票云端默认配置 =====================
# 敏感信息支持从 Streamlit Cloud 的 Secrets 中读取（推荐）：
#   应用管理页 -> Settings -> Secrets，填：
#   KDOCS_TOKEN = "你的token"
# 不配置 Secrets 也能跑，会退回下面的默认值。
_kdocs_token = st.secrets.get("KDOCS_TOKEN", "RRZwCZarLOHD4gHKtfSVi") if hasattr(st, "secrets") else "RRZwCZarLOHD4gHKtfSVi"

INVOICE_CONFIG = {
    "kdocs_webhook": "https://www.kdocs.cn/api/v3/ide/file/cdH0A450EedY/script/V2-6x7HgWruLz4P74YP1PIsVf/sync_task",
    "kdocs_token": _kdocs_token,
    "template_urls": {
        "鼎邦物流": "https://gitee.com/zhanliuliang/zll/raw/master/dingbang.xlsx",
        "明日之星": "https://gitee.com/zhanliuliang/zll/raw/master/mingrizhixing_new.xlsx",
        "易达物流": ""
    },
    # 明日之星新模板订单区默认值（可按需调整）
    "mrzx_order_defaults": {
        "是否FIST订单": "否",
        "提货方式": "自送",
        "交税方式": "包税",
        "申报币种": "美元",
        "是否带电": "否",
        "是否带磁": "否",
        "是否液体": "否",
        "是否粉末": "否",
        "是否危险品": "否"
    }
}

# ===================== 状态初始化 =====================
DEFAULT_YIDA_PRODUCTS = [
    {
        "中文品名": "塑料装饰盘",
        "国外清关商编": "3924.10.4000",
        "单个重量(kg)": 0.35,
        "国内单价(USD)": 1.0
    },
    {
        "中文品名": "塑料托盘",
        "国外清关商编": "3924.10.4000",
        "单个重量(kg)": 1.3,
        "国内单价(USD)": 0.7
    }
]

if "yida_product_df" not in st.session_state:
    st.session_state.yida_product_df = pd.DataFrame(DEFAULT_YIDA_PRODUCTS)

# =====================================================================
# 模块 1：发货单处理业务逻辑
# =====================================================================
def beautify_excel(file_stream, col_width=15, row_height=20):
    try:
        wb = load_workbook(file_stream)
        font = Font(name="Calibri", size=11, bold=True)
        alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin = Side(border_style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        colors = [
            "FFCCCC", "CCFFCC", "CCCCFF", "FFFFCC", "CCFFFF", "FFCCFF",
            "FF9999", "99FF99", "9999FF", "FFCC99", "99FFCC", "CCCC99"
        ]

        for sheet_name in ['详细数据', '汇总结果']:
            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]
            for col in ws.columns:
                column_letter = col[0].column_letter
                ws.column_dimensions[column_letter].width = col_width

            for row in ws.iter_rows():
                if not row[0].row:
                    continue
                ws.row_dimensions[row[0].row].height = row_height
                for cell in row:
                    if cell.value is None:
                        continue
                    cell.alignment = alignment
                    cell.border = border
                    if cell.row == 1:
                        cell.font = font

            if sheet_name == '详细数据':
                header_row = [cell.value for cell in ws[1]]
                if "货件编号" not in header_row:
                    continue
                item_code_col_idx = header_row.index("货件编号") + 1

                item_codes = {}
                for row in ws.iter_rows(min_row=2, min_col=item_code_col_idx, max_col=item_code_col_idx):
                    val = row[0].value
                    if val and val not in item_codes:
                        item_codes[val] = None

                unique_codes = list(item_codes.keys())
                for idx, code in enumerate(unique_codes):
                    item_codes[code] = colors[idx % len(colors)]

                for row in ws.iter_rows(min_row=2):
                    code_cell = row[item_code_col_idx - 1]
                    code_val = code_cell.value
                    if code_val in item_codes:
                        fill = PatternFill(start_color=item_codes[code_val],
                                           end_color=item_codes[code_val],
                                           fill_type="solid")
                        for cell in row:
                            cell.fill = fill

        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        return output_stream
    except Exception as e:
        raise Exception(f"美化出错: {str(e)}")

def load_and_merge_data(uploaded_file):
    try:
        xl_file = pd.ExcelFile(uploaded_file)
        sheet_names = xl_file.sheet_names

        required_sheets = ['发货单详情', '装箱信息']
        for sheet in required_sheets:
            if sheet not in sheet_names:
                raise KeyError(f"文件缺少必须工作表：{sheet}")

        df_detail = pd.read_excel(uploaded_file, sheet_name='发货单详情')
        df_pack = pd.read_excel(uploaded_file, sheet_name='装箱信息')

        df_detail.rename(columns={
            '货件编号(发货商品)': '货件编号',
            '发货单号(基本信息)': '发货单号',
            'SKU(发货商品)': 'SKU',
            '发货量(发货商品)': '发货量',
            '品名(发货商品)': '品名',
            '商品图片(发货商品)': '商品图片',
            '商品品牌': '供应商'
        }, inplace=True)

        df_pack.rename(columns={
            '货件编号(装箱信息)': '货件编号',
            '发货单号(装箱信息)': '发货单号',
            'SKU(装箱信息)': 'SKU',
            '发货量(装箱信息)': '发货量',
            'ReferenceId(装箱信息)': 'ReferenceId'
        }, inplace=True)

        merged = pd.merge(
            df_detail, df_pack,
            on=['货件编号', '发货单号', 'SKU', '发货量'],
            how='left'
        )

        keep_cols = [
            'SKU', '品名', '商品图片', '发货量', '箱数', '箱规单箱数量', '物流中心编码',
            '供应商', '货件编号', 'ReferenceId', '箱号', '总箱数',
            '外箱重量(kg)', '外箱总重量(kg)', '外箱长(cm)', '外箱宽(cm)', '外箱高(cm)',
            '外箱体积(m³)', '外箱总体积(m³)', '外箱总体积重(kg)', '创建时间', '发货时间', '物流商'
        ]
        exist_cols = [c for c in keep_cols if c in merged.columns]
        merged = merged[exist_cols]

        merged.rename(columns={
            '箱规单箱数量': '单箱数量',
            '总箱数': '总箱数编号'
        }, inplace=True)

        return merged
    except Exception as e:
        raise Exception(f"数据加载合并失败：{str(e)}")

def perform_summary(df):
    df_copy = df.copy()
    df_copy['物流中心编码'] = df_copy['物流中心编码'].fillna('未知地址')
    df_copy['货件编号'] = df_copy['货件编号'].fillna('未知编号')
    df_copy['ReferenceId'] = df_copy['ReferenceId'].fillna('无ReferenceId')

    summary = df_copy.groupby(['物流中心编码', '货件编号', 'ReferenceId'], dropna=False).agg(
        品名=('品名', 'first'),
        供应商=('供应商', 'first'),
        总箱数=('箱数', 'sum'),
        外箱总重量=('外箱总重量(kg)', 'sum'),
        外箱总体积=('外箱总体积(m³)', 'sum')
    ).reset_index()

    summary['外箱总体积'] = summary['外箱总体积'].fillna(0)
    summary['外箱总体积重(kg)'] = (summary['外箱总体积'] * 167).round(2)
    return summary

def save_excel(df, summary_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='详细数据', index=False)
        summary_df.to_excel(writer, sheet_name='汇总结果', index=False)
    output.seek(0)
    return output

# =====================================================================
# 模块 2：物流发票处理业务逻辑
# =====================================================================
def fetch_kdocs_po_data(webhook_url, api_token):
    url = webhook_url.strip()
    token = api_token.strip()
    if not url or not token:
        raise ValueError("未配置有效的金山文档 Webhook 或 Token！")

    headers = {"AirScript-Token": token, "Content-Type": "application/json"}
    resp = requests.post(url, json={"Context": {"argv": {}}}, headers=headers, timeout=25)
    if resp.status_code != 200:
        raise ValueError(f"请求金山文档失败 (HTTP {resp.status_code}): {resp.text}")

    res_json = resp.json()
    data = None
    if isinstance(res_json, dict):
        if "data" in res_json and isinstance(res_json["data"], dict) and "result" in res_json["data"]:
            data = res_json["data"]["result"]
        elif "result" in res_json:
            data = res_json["result"]
        elif "data" in res_json and isinstance(res_json["data"], list):
            data = res_json["data"]
    elif isinstance(res_json, list):
        data = res_json

    if not data or not isinstance(data, list):
        raise ValueError(f"金山文档未返回有效数据列表，返回内容: {res_json}")

    return pd.DataFrame(data)

def fetch_template_bytes(template_url, carrier_name):
    url = str(template_url).strip() if template_url else ""
    if "/blob/" in url:
        url = re.sub(r'/blob/([^/]+)/', r'/raw//', url)

    if url and url.startswith(("http://", "https://")):
        try:
            resp = requests.get(url, timeout=20)
            if resp.status_code == 200 and len(resp.content) > 1000 and not resp.content.startswith(b"<!DOCTYPE"):
                return resp.content
        except Exception:
            pass

    app_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
    local_candidates = [
        os.path.join(app_dir, f"{carrier_name}.xlsx"),
        os.path.join(app_dir, f"{carrier_name}发票.xlsx"),
    ]
    if "鼎邦" in carrier_name:
        local_candidates.append(os.path.join(app_dir, "dingbang.xlsx"))
    elif "明日" in carrier_name:
        local_candidates.extend([
            os.path.join(app_dir, "mingrizhixing_new.xlsx"),
            os.path.join(app_dir, "明日之星 新版.xlsx"),
            os.path.join(app_dir, "mingrizhixing.xlsx")
        ])
    elif "易达" in carrier_name:
        local_candidates.extend([
            os.path.join(app_dir, "yida.xlsx"),
            os.path.join(app_dir, "订单批量预报模版_海运V23.08.14.xlsx"),
            os.path.join(app_dir, "易达物流.xlsx")
        ])

    for loc_file in local_candidates:
        if loc_file and os.path.isfile(loc_file):
            with open(loc_file, 'rb') as f:
                return f.read()

    return None

class InvoiceProcessor:
    @staticmethod
    def clean_sku(val):
        if val is None or pd.isna(val):
            return ""
        s = str(val).strip()
        if s.endswith('.0') and s[:-2].isdigit():
            s = s[:-2]
        chars_to_strip = chr(39) + chr(34) + "\t\r\n "
        return s.strip(chars_to_strip)

    @staticmethod
    def to_num(val):
        """长/宽/高等数值列清洗：可转 float 则转，否则原样（空值返回空串）"""
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ""
        try:
            return float(str(val).strip())
        except (ValueError, TypeError):
            s = str(val).strip()
            return "" if s.lower() in ('', 'nan', 'none') else s

    @staticmethod
    def normalize_carrier_name(raw_name):
        if not raw_name or pd.isna(raw_name):
            return ""
        name = str(raw_name).strip()
        name = re.sub(r'[\r\n\t]+', '', name)
        if '鼎邦' in name:
            return '鼎邦物流'
        if '明日' in name:
            return '明日之星'
        if '易达' in name:
            return '易达物流'
        return name

    @staticmethod
    def load_summary_sheet_data(file_obj):
        ref_mapping = {}
        carrier_mapping = {}
        summary_name = None
        try:
            excel_obj = pd.ExcelFile(file_obj)
            sheet_names = excel_obj.sheet_names
            target_sheet = None
            for s in sheet_names:
                if '汇总结果' in str(s) or '汇总' in str(s):
                    target_sheet = s
                    break
            if target_sheet is None and len(sheet_names) >= 2:
                target_sheet = sheet_names[1]

            if target_sheet is not None:
                summary_name = str(target_sheet)
                df_summary = pd.read_excel(excel_obj, sheet_name=target_sheet)

                fba_col, ref_col, carrier_col = None, None, None
                for c in df_summary.columns:
                    c_clean = str(c).strip().lower().replace(' ', '').replace('_', '').replace('*', '')
                    if not fba_col and any(cand in c_clean for cand in ['货件编号', 'fba编号', 'shipmentid', 'fbano', '货件号']):
                        fba_col = c
                    elif not ref_col and any(cand in c_clean for cand in ['referenceid', 'reference', '货物追踪编号', '参考号', '追踪编号', '追踪号']):
                        ref_col = c

                for cand in ['物流商', '承运商', '物流公司', '渠道', '发货物流', '物流']:
                    for c in df_summary.columns:
                        c_clean = str(c).strip().replace('*', '')
                        if any(ex in c_clean for ex in ['中心', '编码', '仓', '代码', 'code', '单号', '体积', '重', '费']):
                            continue
                        if cand in c_clean:
                            carrier_col = c
                            break
                    if carrier_col:
                        break

                if fba_col:
                    for _, row in df_summary.iterrows():
                        fba_val = str(row[fba_col]).strip() if pd.notna(row[fba_col]) else ""
                        if not fba_val or fba_val.lower() == 'nan':
                            continue
                        if ref_col and pd.notna(row[ref_col]):
                            ref_val = str(row[ref_col]).strip()
                            if ref_val and ref_val.lower() != 'nan':
                                ref_mapping[fba_val] = ref_val
                                ref_mapping[fba_val.upper()] = ref_val
                        if carrier_col and pd.notna(row[carrier_col]):
                            car_val = str(row[carrier_col]).strip()
                            if car_val and car_val.lower() != 'nan':
                                car_clean = InvoiceProcessor.normalize_carrier_name(car_val)
                                carrier_mapping[fba_val] = car_clean
                                carrier_mapping[fba_val.upper()] = car_clean
        except Exception:
            pass
        return ref_mapping, carrier_mapping, summary_name

    @staticmethod
    def clean_warehouse_code(code):
        return re.sub(r'[\(（][^\)）]*[\)）]', '', str(code)).strip() if code else ""

    @staticmethod
    def clean_name(name):
        return re.sub(r'[\\/*?:"<>|]', '_', str(name).strip())

    @staticmethod
    def get_dict_val(data_dict, candidate_keys):
        for cand in candidate_keys:
            cand_clean = str(cand).strip().lower()
            for k, val in data_dict.items():
                if cand_clean == str(k).strip().lower():
                    if pd.notna(val) and str(val).strip() != "":
                        return val
        return ""

    @staticmethod
    def parse_box_ranges(box_str):
        if not box_str:
            return None
        s = str(box_str).strip()
        s = re.sub(r'[,;，；、\s]+', ',', s)
        boxes = set()
        has_number = False
        for part in s.split(','):
            part = part.strip()
            if not part:
                continue
            m_range = re.search(r'(\d+)\s*[-~至到_–—－~～]\s*(\d+)', part)
            if m_range:
                start, end = int(m_range.group(1)), int(m_range.group(2))
                boxes.update(range(min(start, end), max(start, end) + 1))
                has_number = True
            else:
                m_single = re.findall(r'\b\d+\b', part)
                if m_single:
                    for num in m_single:
                        boxes.add(int(num))
                    has_number = True
        return boxes if has_number else None

    @staticmethod
    def extract_row_boxes(s_dict, cur_box):
        start_box_val = InvoiceProcessor.get_dict_val(s_dict, ['起始箱号', '开始箱号', '起始箱'])
        end_box_val = InvoiceProcessor.get_dict_val(s_dict, ['截止箱号', '结束箱号', '终止箱号', '截止箱'])
        if start_box_val and end_box_val:
            try:
                s_int = int(re.search(r'\d+', str(start_box_val)).group(0))
                e_int = int(re.search(r'\d+', str(end_box_val)).group(0))
                return set(range(min(s_int, e_int), max(s_int, e_int) + 1)), max(s_int, e_int) - min(s_int, e_int) + 1
            except Exception:
                pass

        box_no_val = InvoiceProcessor.get_dict_val(s_dict, ['箱号', '外箱箱号', '箱号编号', 'CTN NO', 'Carton No'])
        box_cnt_val = InvoiceProcessor.get_dict_val(s_dict, ['箱数', '外箱数', '总箱数'])

        if box_no_val and re.search(r'(\d+)\s*[-~至到_–—－~～]\s*(\d+)', str(box_no_val)):
            b = InvoiceProcessor.parse_box_ranges(box_no_val)
            if b:
                return b, len(b)

        count_int = 1
        if box_cnt_val:
            try:
                count_int = int(float(str(box_cnt_val).strip()))
            except Exception:
                pass

        if box_no_val and str(box_no_val).strip().isdigit() and (not box_cnt_val or str(box_no_val).strip() != str(box_cnt_val).strip()):
            start_num = int(str(box_no_val).strip())
            return set(range(start_num, start_num + count_int)), count_int

        return set(range(cur_box, cur_box + count_int)), count_int

    @staticmethod
    def format_box_ranges(boxes_set):
        if not boxes_set:
            return ""
        sorted_boxes = sorted(list(boxes_set))
        if len(sorted_boxes) == 1:
            return f"{sorted_boxes[0]}"
        if sorted_boxes[-1] - sorted_boxes[0] + 1 == len(sorted_boxes):
            return f"{sorted_boxes[0]}-{sorted_boxes[-1]}"
        ranges = []
        start = sorted_boxes[0]
        end = sorted_boxes[0]
        for n in sorted_boxes[1:]:
            if n == end + 1:
                end = n
            else:
                ranges.append(f"{start}" if start == end else f"{start}-{end}")
                start = n
                end = n
        ranges.append(f"{start}" if start == end else f"{start}-{end}")
        return ",".join(ranges)

    @staticmethod
    def format_dimensions(length, width, height):
        def clean_val(v):
            if v is None:
                return ""
            s = str(v).strip()
            if not s or s.lower() == "nan":
                return ""
            try:
                fv = float(s)
                return f"{int(fv)}" if fv.is_integer() else f"{fv}"
            except Exception:
                return s
        l_str = clean_val(length)
        w_str = clean_val(width)
        h_str = clean_val(height)
        if l_str and w_str and h_str:
            return f"{l_str}*{w_str}*{h_str}"
        return l_str or w_str or h_str or ""

# =====================================================================
# 易达物流预报生成
# =====================================================================
def generate_yida_summary_excel(template_bytes, shipment_groups, summary_ref_mapping, df_summary=None, product_config_list=None, po_dict=None, missing_records=None):
    if template_bytes:
        wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "预报模版"

    main_headers = [
        "客户编码", "客户订单号", "referenceID", "产品编码", "地址速记码", "开航日期", "卸货港", "唛头",
        "报关方式", "是否包税", "箱数", "单位", "毛重", "体积", "备注", "对账合同号",
        "货件创建日期", "发货地址", "起运港", "配送起始日期", "配送结束日期", "SKU"
    ]
    sub_headers = [
        "国内商编", "中文品名", "英文品名", "品牌", "材质", "用途", "数量", "单个产品重量",
        "币种", "国内单价", "申报总价", "国外清关单价", "国外清关商编", "箱号备注", "商品链接"
    ]

    header_font = Font(name="微软雅黑", size=10, bold=True, color="1E293B")
    header_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    table_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for rng in list(ws.merged_cells.ranges):
        if rng.min_row == 1 and rng.max_row == 1:
            ws.unmerge_cells(str(rng))
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=37)
    title_cell = ws.cell(row=1, column=1, value="订单批量预报模版_海运")
    title_cell.font = Font(name="微软雅黑", size=13, bold=True, color="0F172A")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 25

    for col_idx, h_text in enumerate(main_headers, 1):
        ws.merge_cells(start_row=2, start_column=col_idx, end_row=3, end_column=col_idx)
        cell = ws.cell(row=2, column=col_idx, value=h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = table_border
        ws.cell(row=3, column=col_idx).border = table_border

    ws.merge_cells(start_row=2, start_column=23, end_row=2, end_column=37)
    cat_cell = ws.cell(row=2, column=23, value="申报详情")
    cat_cell.font = header_font
    cat_cell.fill = header_fill
    cat_cell.alignment = header_align
    cat_cell.border = table_border
    for c in range(23, 38):
        ws.cell(row=2, column=c).border = table_border
        sub_text = sub_headers[c - 23]
        scell = ws.cell(row=3, column=c, value=sub_text)
        scell.font = header_font
        scell.fill = header_fill
        scell.alignment = header_align
        scell.border = table_border

    summary_weight_map, summary_vol_map, summary_boxes_map = {}, {}, {}
    if df_summary is not None and not df_summary.empty:
        for _, s_row in df_summary.iterrows():
            fba = str(s_row.get("货件编号", "")).strip()
            if fba:
                try:
                    summary_weight_map[fba] = round(float(s_row.get("外箱总重量", 0) or 0), 2)
                except Exception:
                    summary_weight_map[fba] = 0
                try:
                    summary_vol_map[fba] = round(float(s_row.get("外箱总体积", 0) or 0), 2)
                except Exception:
                    summary_vol_map[fba] = 0
                try:
                    summary_boxes_map[fba] = int(s_row.get("总箱数", 0) or 0)
                except Exception:
                    summary_boxes_map[fba] = 0

    shipment_color_palette = [
        "F8FAFC", "F0F9FF", "FEFCE8", "F0FDF4", "FAF5FF",
        "FFF1F2", "FEF3C7", "ECFDF5", "F5F3FF", "FFF7ED"
    ]

    current_row = 4
    today_str = datetime.now().strftime("%Y-%m-%d")
    data_font = Font(name="微软雅黑", size=9, color="1E293B")
    center_align = Alignment(horizontal="center", vertical="center")

    configs = product_config_list if product_config_list else DEFAULT_YIDA_PRODUCTS
    config_map = {}
    for conf in configs:
        c_name = str(conf.get("中文品名", "")).strip()
        if c_name:
            config_map[c_name] = conf

    for ship_idx, (fba_shipment_id, group_df) in enumerate(shipment_groups):
        first_row_data = group_df.iloc[0].to_dict()

        ship_color = shipment_color_palette[ship_idx % len(shipment_color_palette)]
        ship_fill = PatternFill(start_color=ship_color, end_color=ship_color, fill_type="solid")

        ref_id = summary_ref_mapping.get(fba_shipment_id) or first_row_data.get('ReferenceId', '')
        wh_code = InvoiceProcessor.clean_warehouse_code(
            InvoiceProcessor.get_dict_val(first_row_data, ['物流中心编码', '仓码', '仓库编码', '地址速记码'])
        )

        total_boxes = summary_boxes_map.get(fba_shipment_id)
        if not total_boxes:
            total_boxes = int(group_df['箱数'].sum() if '箱数' in group_df.columns else len(group_df))

        total_weight = summary_weight_map.get(fba_shipment_id)
        if not total_weight:
            try:
                total_weight = round(float(group_df['外箱总重量(kg)'].sum()), 2)
            except Exception:
                total_weight = 0

        total_vol = summary_vol_map.get(fba_shipment_id)
        if not total_vol:
            try:
                total_vol = round(float(group_df['外箱总体积(m³)'].sum()), 2)
            except Exception:
                total_vol = 0

        category_aggregated = {}
        for _, row in group_df.iterrows():
            row_dict = row.to_dict()
            row_sku = InvoiceProcessor.clean_sku(row_dict.get('SKU', ''))

            po_item = po_dict.get(row_sku) or po_dict.get(row_sku.upper()) or po_dict.get(row_sku.lower()) or {}
            if not po_item and missing_records is not None:
                missing_records.append({'货件编号': fba_shipment_id, '未匹配SKU': row_sku, '品名': row_dict.get('品名', '')})

            po_pname = InvoiceProcessor.get_dict_val(po_item, ['中文品名', '品名', '商品品名', '产品名称', '商品名称'])
            cn_name = po_pname or str(row_dict.get('品名', '')).strip()

            en_name = InvoiceProcessor.get_dict_val(po_item, ['英文品名', '英文名'])
            hs_cn = InvoiceProcessor.get_dict_val(po_item, ['HS编码', '海关编码', '国内商编'])
            material = InvoiceProcessor.get_dict_val(po_item, ['材质中/英*', '材质中/英', '材质'])
            usage = InvoiceProcessor.get_dict_val(po_item, ['用途中/英*', '用途中/英', '用途'])
            brand = "无"

            pcs_val = None
            if po_item:
                raw_pcs = InvoiceProcessor.get_dict_val(po_item, ['PCS', '单箱数量', '箱规', '单箱PCS', 'PCS/CTN'])
                if raw_pcs != "":
                    try:
                        pcs_val = float(raw_pcs)
                    except Exception:
                        pcs_val = None

            if pcs_val is None:
                raw_pcs = InvoiceProcessor.get_dict_val(row_dict, ['单箱数量', '箱规单箱数量', 'PCS'])
                if raw_pcs != "":
                    try:
                        pcs_val = float(raw_pcs)
                    except Exception:
                        pcs_val = None

            if pcs_val is None or pcs_val <= 0:
                pcs_val = 1.0

            try:
                delivery_boxes = float(row_dict.get('发货量', 1)) if pd.notna(row_dict.get('发货量')) else 1.0
            except Exception:
                delivery_boxes = 1.0

            row_total_pcs = int(round(delivery_boxes * pcs_val))

            matched_conf = config_map.get(cn_name)

            if matched_conf:
                hs_overseas = matched_conf.get("国外清关商编", "")
                try:
                    unit_weight = float(matched_conf.get("单个重量(kg)", 0))
                except Exception:
                    unit_weight = ""
                try:
                    unit_price = float(matched_conf.get("国内单价(USD)", 0))
                except Exception:
                    unit_price = ""
            else:
                hs_overseas = ""
                unit_weight = ""
                po_price = InvoiceProcessor.get_dict_val(po_item, ['申报单价', '单价', '国内单价'])
                try:
                    unit_price = float(po_price) if po_price != "" else ""
                except Exception:
                    unit_price = ""

            if cn_name not in category_aggregated:
                category_aggregated[cn_name] = {
                    'cn_name': cn_name,
                    'en_name': en_name,
                    'hs_cn': hs_cn,
                    'hs_overseas': hs_overseas,
                    'brand': brand,
                    'material': material,
                    'usage': usage,
                    'unit_weight': unit_weight,
                    'unit_price': unit_price,
                    'qty': 0,
                    'is_matched': bool(matched_conf),
                    'raw_dict': row_dict
                }
            category_aggregated[cn_name]['qty'] += row_total_pcs

        gross_weight = float(total_weight or 0)

        tray_item = category_aggregated.get("塑料托盘")
        tray_total_weight = 0.0
        if tray_item:
            try:
                t_wt = float(tray_item.get('unit_weight', 1.3) or 1.3)
            except Exception:
                t_wt = 1.3
            tray_total_weight = tray_item['qty'] * t_wt

        plate_item = category_aggregated.get("塑料装饰盘")
        if plate_item and plate_item['qty'] > 0 and gross_weight > 0:
            rem_net_weight = gross_weight * 0.95 - tray_total_weight
            if rem_net_weight > 0:
                plate_item['unit_weight'] = round(rem_net_weight / plate_item['qty'], 2)

        shipping_mark = f"US-SEA-{wh_code}" if wh_code else "US-SEA"

        for idx, (c_name, item) in enumerate(category_aggregated.items()):
            qty = item['qty']
            unit_price = item['unit_price']
            try:
                total_price = round(qty * float(unit_price), 2) if (unit_price != "" and unit_price is not None) else ""
            except Exception:
                total_price = ""

            is_first_item = (idx == 0)

            currency_val = "USD" if (unit_price != "" and unit_price is not None) else ""
            unit_weight_val = item['unit_weight'] if (item['unit_weight'] != "" and item['unit_weight'] is not None) else ""
            unit_price_val = unit_price if (unit_price != "" and unit_price is not None) else ""

            row_values = {
                1: "C1709200001" if is_first_item else "",
                2: fba_shipment_id if is_first_item else "",
                3: ref_id if is_first_item else "",
                4: "US-SEA" if is_first_item else "",
                5: wh_code if is_first_item else "",
                8: shipping_mark if is_first_item else "",
                9: "退税" if is_first_item else "",
                10: "是" if is_first_item else "",
                11: total_boxes if is_first_item else "",
                12: "CTNS" if is_first_item else "",
                13: total_weight if is_first_item else "",
                14: total_vol if is_first_item else "",
                17: today_str,
                18: "美国境外发货",
                22: "",
                23: item['hs_cn'],
                24: item['cn_name'],
                25: item['en_name'],
                26: item['brand'],
                27: item['material'],
                28: item['usage'],
                29: qty,
                30: unit_weight_val,
                31: currency_val,
                32: unit_price_val,
                33: total_price,
                34: "",
                35: item['hs_overseas'],
                36: "",
                37: ""
            }

            ws.row_dimensions[current_row].height = 24
            for col_idx in range(1, 38):
                val = row_values.get(col_idx, "")
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.alignment = center_align
                cell.font = data_font
                cell.fill = ship_fill
                cell.border = table_border

            current_row += 1

        ws.row_dimensions[current_row].height = 14
        current_row += 1

    custom_col_widths = {
        1: 15, 2: 18, 3: 15, 4: 12, 5: 14, 6: 14, 7: 14, 8: 20,
        9: 12, 10: 11, 11: 10, 12: 10, 13: 12, 14: 12, 15: 14,
        16: 14, 17: 15, 18: 15, 19: 12, 20: 14, 21: 14, 22: 18,
        23: 16, 24: 18, 25: 20, 26: 10, 27: 12, 28: 12, 29: 12,
        30: 16, 31: 10, 32: 13, 33: 14, 34: 14, 35: 16, 36: 12,
        37: 15
    }
    for col_num in range(1, 38):
        col_letter = get_column_letter(col_num)
        ws.column_dimensions[col_letter].width = custom_col_widths.get(col_num, 15)
        ws.column_dimensions[col_letter].hidden = False

    for cd in ws.column_dimensions.values():
        cd.hidden = False

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.getvalue()

# =====================================================================
# 顶部全局导航栏
# =====================================================================
st.title("📦 发货单与物流发票协同处理工具")
st.caption("Cross-Border Logistics Shipment Summary & Invoice Batch Generator")

# 侧边栏统一放置参数字典维护，杜绝前台布局错位与突兀黑框
with st.sidebar:
    st.subheader("⚙️ 业务字典维护")
    with st.expander("📦 易达参数字典", expanded=True):
        st.caption("维护国外商编、单重与单价：")
        
        # 兼容性表格展示/编辑
        if hasattr(st, "data_editor") or hasattr(st, "experimental_data_editor"):
            edited_df = safe_data_editor(
                st.session_state.yida_product_df,
                key="sidebar_yida_editor"
            )
            st.session_state.yida_product_df = edited_df
        else:
            st.dataframe(st.session_state.yida_product_df)

        st.markdown("---")
        st.markdown("**➕ 快捷新增品名规则**")
        with st.form("add_yida_product_form", clear_on_submit=True):
            new_name = st.text_input("中文品名 (需与采购单一致)")
            new_hs = st.text_input("国外清关商编", value="3924.10.4000")
            c_w, c_p = st.columns(2)
            with c_w:
                new_wt = st.number_input("单重(kg)", value=0.0, step=0.01)
            with c_p:
                new_price = st.number_input("单价(USD)", value=0.0, step=0.01)
            
            submitted = st.form_submit_button("➕ 添加到字典")
            if submitted:
                if new_name.strip():
                    new_item = {
                        "中文品名": new_name.strip(),
                        "国外清关商编": new_hs.strip(),
                        "单个重量(kg)": round(float(new_wt), 3),
                        "国内单价(USD)": round(float(new_price), 2)
                    }
                    st.session_state.yida_product_df = pd.concat(
                        [st.session_state.yida_product_df, pd.DataFrame([new_item])],
                        ignore_index=True
                    )
                    st.success(f"已成功添加: {new_name}")
                    safe_rerun()
                else:
                    st.warning("中文品名不能为空！")

        if len(st.session_state.yida_product_df) > 0:
            st.markdown("---")
            col_del1, col_del2 = st.columns([2, 1])
            with col_del1:
                del_item = st.selectbox(
                    "选择要删除的品名",
                    options=st.session_state.yida_product_df["中文品名"].tolist(),
                    key="del_product_sel"
                )
            with col_del2:
                st.write("")
                if st.button("🗑️ 删除", key="btn_del_product"):
                    st.session_state.yida_product_df = st.session_state.yida_product_df[
                        st.session_state.yida_product_df["中文品名"] != del_item
                    ].reset_index(drop=True)
                    safe_rerun()

        st.markdown("---")
        if st.button("↺ 恢复初始默认参数", key="btn_reset_settings"):
            st.session_state.yida_product_df = pd.DataFrame(DEFAULT_YIDA_PRODUCTS)
            safe_rerun()

# =====================================================================
# 主页面 Tabs（发货单合并 + 发票生成）
# =====================================================================
tab_shipment, tab_invoice = st.tabs([
    "📦 1. 发货单合并汇总工具",
    "📑 2. 物流发票批量生成工具"
])

# ---------- 工具 1：发货单合并汇总 ----------
with tab_shipment:
    uploaded_ship_file = st.file_uploader(
        "上传发货单原始 Excel 文件",
        type=["xlsx", "xls"],
        help="需包含「发货单详情」与「装箱信息」两张工作表",
        key="ship_uploader"
    )

    if uploaded_ship_file is not None:
        try:
            with st.spinner("正在处理发货单数据并套用美化格式..."):
                merged_df = load_and_merge_data(uploaded_ship_file)
                summary_df = perform_summary(merged_df)
                raw_excel = save_excel(merged_df, summary_df)
                excel_bytes = beautify_excel(raw_excel)

            total_boxes = int(summary_df['总箱数'].sum())
            total_weight = f"{summary_df['外箱总重量'].sum():,.2f}"
            total_vol = f"{summary_df['外箱总体积'].sum():,.2f}"
            total_vol_weight = f"{summary_df['外箱总体积重(kg)'].sum():,.2f}"

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f'''<div class="metric-card-box">
                    <div class="metric-card-title">总装箱量</div>
                    <div class="metric-card-val">{total_boxes:,}<span class="metric-card-unit">箱</span></div>
                </div>''', unsafe_allow_html=True)
            with col2:
                st.markdown(f'''<div class="metric-card-box">
                    <div class="metric-card-title">实重合计</div>
                    <div class="metric-card-val">{total_weight}<span class="metric-card-unit">kg</span></div>
                </div>''', unsafe_allow_html=True)
            with col3:
                st.markdown(f'''<div class="metric-card-box">
                    <div class="metric-card-title">外箱总体积</div>
                    <div class="metric-card-val">{total_vol}<span class="metric-card-unit">m³</span></div>
                </div>''', unsafe_allow_html=True)
            with col4:
                st.markdown(f'''<div class="metric-card-box">
                    <div class="metric-card-title">总体积重 (1:167)</div>
                    <div class="metric-card-val">{total_vol_weight}<span class="metric-card-unit">kg</span></div>
                </div>''', unsafe_allow_html=True)

            st.write("")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"发货单处理结果_{timestamp}.xlsx"

            st.download_button(
                label="⬇️ 导出处理后的 Excel 文件",
                data=excel_bytes,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"处理发生异常：{str(e)}")

# ---------- 工具 2：物流发票批量生成 ----------
with tab_invoice:
    col_sel, col_up = st.columns([1, 2])
    with col_sel:
        carrier_option = st.selectbox(
            "选择物流模板模式",
            options=["鼎邦物流", "明日之星", "易达物流", "【自动识别】按表格物流列"],
            index=0,
            key="inv_carrier_sel"
        )
    with col_up:
        uploaded_inv_file = st.file_uploader(
            "上传发货单处理结果文件",
            type=["xlsx", "xls"],
            key="inv_uploader"
        )

    if uploaded_inv_file is not None:
        if st.button("🚀 开始批量生成发票", key="btn_run_inv"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                status_text.text("【1/4】正在连接金山文档云端，同步采购单主数据...")
                df_po = fetch_kdocs_po_data(INVOICE_CONFIG["kdocs_webhook"], INVOICE_CONFIG["kdocs_token"])
                progress_bar.progress(20)

                po_dict = {}
                sku_col = next((c for c in df_po.columns if 'sku' in str(c).strip().lower()), None)
                if not sku_col:
                    raise ValueError("采购单中未找到 SKU 列")
                for _, row in df_po.iterrows():
                    sku_val = InvoiceProcessor.clean_sku(row.get(sku_col))
                    if sku_val:
                        po_dict[sku_val] = row.to_dict()
                        po_dict[sku_val.upper()] = row.to_dict()
                        po_dict[sku_val.lower()] = row.to_dict()

                status_text.text("【2/4】正在解析发货单明细与汇总数据...")
                uploaded_inv_file.seek(0)
                df_ship = pd.read_excel(uploaded_inv_file)

                df_summary_sheet = None
                try:
                    uploaded_inv_file.seek(0)
                    xl_temp = pd.ExcelFile(uploaded_inv_file)
                    for s in xl_temp.sheet_names:
                        if '汇总' in s:
                            df_summary_sheet = pd.read_excel(uploaded_inv_file, sheet_name=s)
                            break
                except Exception:
                    pass

                ship_sku_col = next((c for c in df_ship.columns if 'sku' in str(c).strip().lower()), 'SKU')
                ship_id_col = next((c for c in df_ship.columns if any(k in str(c) for k in ['货件编号', 'FBA货件编号', 'ShipmentID'])), '货件编号')

                if ship_id_col not in df_ship.columns:
                    raise ValueError(f"发货单中未找到【货件编号】列，现有列: {list(df_ship.columns)}")

                uploaded_inv_file.seek(0)
                summary_ref_mapping, summary_carrier_mapping, _ = InvoiceProcessor.load_summary_sheet_data(uploaded_inv_file)

                valid_rows = [row for _, row in df_ship.iterrows() if pd.notna(row.get(ship_id_col)) and str(row.get(ship_id_col)).strip() not in ('', 'nan')]
                df_valid_ship = pd.DataFrame(valid_rows)
                groups = list(df_valid_ship.groupby(ship_id_col))
                total_groups = len(groups)

                progress_bar.progress(35)
                status_text.text(f"【3/4】共识别到 {total_groups} 个独立货件，正在匹配生成发票...")

                template_cache = {}
                generated_zip_files = {}
                missing_records, skipped_records = [], []
                center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                thin_border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))

                yida_groups = []
                standard_groups = []

                for fba_shipment_id, group_df in groups:
                    first_row_data = group_df.iloc[0].to_dict()
                    c_carrier = ""
                    if carrier_option != "【自动识别】按表格物流列":
                        c_carrier = InvoiceProcessor.normalize_carrier_name(carrier_option)
                    else:
                        c_carrier = summary_carrier_mapping.get(fba_shipment_id) or ""
                        if not c_carrier:
                            for col_k in ['物流', '物流商', '承运商', '物流公司', '渠道', '发货物流']:
                                v = first_row_data.get(col_k)
                                if pd.notna(v) and str(v).strip():
                                    c_carrier = InvoiceProcessor.normalize_carrier_name(v)
                                    break
                    if not c_carrier:
                        c_carrier = "未知物流"

                    if c_carrier == "易达物流":
                        yida_groups.append((fba_shipment_id, group_df))
                    else:
                        standard_groups.append((fba_shipment_id, group_df, c_carrier))

                # ---------- 处理 1：易达物流 ----------
                if yida_groups:
                    status_text.text(f"正在汇总美化生成易达物流预报单 ({len(yida_groups)} 个货件)...")
                    yida_tpl_bytes = fetch_template_bytes(INVOICE_CONFIG["template_urls"].get("易达物流", ""), "易达物流")
                    current_yida_cfg = st.session_state.yida_product_df.to_dict('records')
                    yida_excel_bytes = generate_yida_summary_excel(
                        yida_tpl_bytes, yida_groups, summary_ref_mapping, df_summary_sheet, current_yida_cfg,
                        po_dict=po_dict, missing_records=missing_records
                    )
                    yida_fn = f"易达海运批量预报_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    generated_zip_files[yida_fn] = yida_excel_bytes

                # ---------- 处理 2：鼎邦 / 明日之星 ----------
                current_idx = 0
                for fba_shipment_id, group_df, current_carrier in standard_groups:
                    current_idx += 1
                    first_row_data = group_df.iloc[0].to_dict()

                    if current_carrier not in template_cache:
                        tpl_url = INVOICE_CONFIG["template_urls"].get(current_carrier, "")
                        template_cache[current_carrier] = fetch_template_bytes(tpl_url, current_carrier)

                    template_bytes = template_cache.get(current_carrier)
                    if not template_bytes:
                        tpl_url_cfg = INVOICE_CONFIG["template_urls"].get(current_carrier, "")
                        reason = f"模板下载失败且本地无模板文件（URL: {tpl_url_cfg or '未配置'}，已尝试本地: {carrier_name}.xlsx / mingrizhixing_new.xlsx 等）" if current_carrier == "明日之星" else "未配置发票模板"
                        skipped_records.append({'货件编号': fba_shipment_id, '识别物流': current_carrier, '原因': reason})
                        continue

                    account_name = "未知"
                    for _, s_row in group_df.iterrows():
                        s_sku = InvoiceProcessor.clean_sku(s_row.get(ship_sku_col, ''))
                        if len(s_sku) >= 2:
                            account_name = s_sku[:2].upper()
                            break

                    clean_acc = InvoiceProcessor.clean_name(account_name)
                    clean_fba = InvoiceProcessor.clean_name(fba_shipment_id)
                    clean_fba_wh_code = InvoiceProcessor.clean_warehouse_code(InvoiceProcessor.get_dict_val(first_row_data, ['物流中心编码', '仓码', '仓库编码']))

                    val_ref_id = summary_ref_mapping.get(fba_shipment_id) or ""
                    if not val_ref_id:
                        for _, s_row in group_df.iterrows():
                            r_val = InvoiceProcessor.get_dict_val(s_row.to_dict(), ['ReferenceId', 'Reference ID', 'Reference', '参考号'])
                            if r_val:
                                val_ref_id = str(r_val).strip()
                                break

                    raw_items = []
                    all_shipment_boxes = set()
                    cur_box = 1

                    for _, s_row in group_df.iterrows():
                        s_dict = s_row.to_dict()
                        s_sku = InvoiceProcessor.clean_sku(s_dict.get(ship_sku_col, ''))
                        po_item = po_dict.get(s_sku) or {}

                        if not po_item:
                            missing_records.append({'货件编号': fba_shipment_id, '未匹配SKU': s_sku, '品名': s_dict.get('品名', '')})

                        parsed_boxes, row_box_cnt = InvoiceProcessor.extract_row_boxes(s_dict, cur_box)
                        if parsed_boxes:
                            all_shipment_boxes.update(parsed_boxes)
                            cur_box = max(cur_box, max(parsed_boxes) + 1)

                        val_len = InvoiceProcessor.get_dict_val(s_dict, ['外箱长(cm)', '外箱长', '长(cm)', '长'])
                        val_width = InvoiceProcessor.get_dict_val(s_dict, ['外箱宽(cm)', '外箱宽', '宽(cm)', '宽'])
                        val_height = InvoiceProcessor.get_dict_val(s_dict, ['外箱高(cm)', '外箱高', '高(cm)', '高'])
                        val_vol_dim = InvoiceProcessor.format_dimensions(val_len, val_width, val_height)

                        raw_items.append({
                            'sku': s_sku,
                            'box_count': row_box_cnt,
                            'val_box_num': InvoiceProcessor.format_box_ranges(parsed_boxes),
                            'parsed_boxes': parsed_boxes,
                            'val_fba_no': fba_shipment_id,
                            'val_ref_id': val_ref_id,
                            'val_name_cn': InvoiceProcessor.get_dict_val(po_item, ['品名', '中文品名']) or InvoiceProcessor.get_dict_val(s_dict, ['品名']),
                            'val_name_en': InvoiceProcessor.get_dict_val(po_item, ['英文品名', '英文名']),
                            'val_qty_per_box': InvoiceProcessor.get_dict_val(s_dict, ['单箱数量', 'PCS/CTN']),
                            'val_unit_price': InvoiceProcessor.get_dict_val(po_item, ['申报单价', '单价']),
                            'val_material': InvoiceProcessor.get_dict_val(po_item, ['材质中/英*', '材质中/英', '材质']),
                            'val_usage': InvoiceProcessor.get_dict_val(po_item, ['用途中/英*', '用途中/英', '用途']),
                            'val_hs_code': InvoiceProcessor.get_dict_val(po_item, ['HS编码', '海关编码']),
                            'val_weight': InvoiceProcessor.get_dict_val(s_dict, ['外箱重量(kg)', '外箱重量', '单箱重']),
                            'val_len': val_len, 'val_width': val_width, 'val_height': val_height,
                            'val_vol_dim': val_vol_dim,
                            'val_img_url': InvoiceProcessor.get_dict_val(s_dict, ['商品图片', '图片'])
                        })

                    raw_items.sort(key=lambda it: min(it['parsed_boxes']) if it.get('parsed_boxes') else 999999)
                    merged_items = []
                    for item in raw_items:
                        cn, en = str(item['val_name_cn']).strip().lower(), str(item['val_name_en']).strip().lower()
                        curr_start = min(item['parsed_boxes']) if item.get('parsed_boxes') else None

                        can_merge = False
                        if merged_items:
                            prev = merged_items[-1]
                            prev_cn, prev_en = str(prev['val_name_cn']).strip().lower(), str(prev['val_name_en']).strip().lower()
                            if (cn, en) == (prev_cn, prev_en):
                                can_merge = True  # 明日之星：同品名(中英文一致)即合并，不要求箱号连续

                        if can_merge:
                            prev = merged_items[-1]
                            prev['box_count'] += item['box_count']
                            if prev['parsed_boxes'] and item['parsed_boxes']:
                                prev['parsed_boxes'].update(item['parsed_boxes'])
                                prev['val_box_num'] = InvoiceProcessor.format_box_ranges(prev['parsed_boxes'])
                        else:
                            merged_items.append(copy.copy(item))

                    final_items = merged_items
                    total_boxes = sum(it['box_count'] for it in raw_items) or len(all_shipment_boxes) or len(group_df)

                    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws = wb.active

                    header_filled = {}
                    if current_carrier == "明日之星":
                        # ============ 明日之星新版导入模板 ============
                        # 模板结构（客户订单导入模板 sheet）：
                        #   订单区：B2=客户订单号  F2=仓库代码(触发XLOOKUP带出地址)  B9=总件数
                        #   箱子信息：表头固定第21行，A~Q 17列，数据从第22行起
                        ws_order = ws
                        # 订单区三项必填
                        ws_order["B2"] = fba_shipment_id              # 客户订单号 = FBA货件编号
                        ws_order["F2"] = clean_fba_wh_code            # 仓库代码 = 物流中心编码（XLOOKUP 自动带地址）
                        ws_order["B9"] = total_boxes                  # 总件数 = 总箱数
                        for addr in ("B2", "F2", "B9"):
                            ws_order[addr].alignment = center_alignment

                        # 订单区默认值（标签→写入格 映射；提货时间 = 当天日期）
                        mz_default_cells = {
                            "是否FIST订单": "B3",
                            "提货方式": "B4",
                            "报关方式": "B6",
                            "交税方式": "B7",
                            "申报币种": "B8",
                            "是否带电": "F13",
                            "是否带磁": "F14",
                            "是否液体": "F15",
                            "是否粉末": "F16",
                            "是否危险品": "F17",
                        }
                        for label, addr in mz_default_cells.items():
                            default_val = INVOICE_CONFIG["mrzx_order_defaults"].get(label)
                            if default_val is not None:
                                ws_order[addr] = default_val
                        ws_order["B5"] = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)  # 提货时间=当天
                        ws_order["B5"].number_format = "yyyy/mm/dd"
                        for addr in list(mz_default_cells.values()) + ["B5"]:
                            ws_order[addr].alignment = center_alignment

                        # 箱子信息表头定位（优先精确第21行，异常时回退到关键词扫描）
                        box_header_row = None
                        try:
                            if str(ws.cell(row=21, column=1).value or "").strip() == "箱号":
                                box_header_row = 21
                        except Exception:
                            pass
                        if not box_header_row:
                            for r in range(15, min(30, ws.max_row + 1)):
                                row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
                                if any("箱号" in v for v in row_vals) and any("Reference" in v for v in row_vals):
                                    box_header_row = r
                                    break
                        if box_header_row:
                            mz_cols = {}
                            for c in range(1, ws.max_column + 1):
                                cv = str(ws.cell(row=box_header_row, column=c).value or "").strip()
                                if cv:
                                    mz_cols[cv] = c

                            def mz_col(candidates):
                                for cand in candidates:
                                    for h_name, col_i in mz_cols.items():
                                        if cand in h_name:
                                            return col_i
                                return None

                            mz_box = mz_col(['箱号'])
                            mz_ref = mz_col(['Reference ID', 'ReferenceId', '参考号'])
                            mz_cn = mz_col(['中文品名'])
                            mz_en = mz_col(['英文品名'])
                            mz_gw = mz_col(['毛重'])
                            mz_len = mz_col(['长(cm)', '长'])
                            mz_wid = mz_col(['宽(cm)', '宽'])
                            mz_hei = mz_col(['高(cm)', '高'])
                            mz_mat = mz_col(['材质'])
                            mz_brand = mz_col(['品牌'])
                            mz_mark = mz_col(['唛头'])
                            mz_price = mz_col(['申报单价'])
                            mz_qty = mz_col(['申报数量'])
                            mz_hs = mz_col(['编码HSC', 'HS'])
                            mz_usage = mz_col(['产品用途'])
                            mz_link = mz_col(['销售链接'])
                            mz_img = mz_col(['图片'])

                            mz_row = box_header_row + 1
                            for item in final_items:
                                # 箱号格式：FBA货件编号+U+起始箱流水号-末尾箱流水号
                                # 依据 Amazon 箱标：Box N → FBA19T0HSCF3U00000N
                                # 该品名占据第 X~Y 箱时填 FBA19XXXXXU0000X-Y，如 FBA19SX3FH0XU000001-12
                                if item.get('parsed_boxes'):
                                    first_box = min(item['parsed_boxes'])
                                    last_box = max(item['parsed_boxes'])
                                else:
                                    first_box, last_box = 1, item['box_count']
                                box_display = f"{clean_fba}U{first_box:06d}-{last_box:06d}"
                                mz_writes = [
                                    (mz_box, box_display),                                   # 箱号：FBA19XXXXXU000001-000012 形式（起-末流水号）
                                    (mz_ref, item['val_ref_id'] or fba_shipment_id),         # Reference ID
                                    (mz_cn, item['val_name_cn']), (mz_en, item['val_name_en']),
                                    (mz_gw, item['val_weight']),                             # 毛重(KG) = 外箱重量
                                    (mz_len, InvoiceProcessor.to_num(item['val_len'])),
                                    (mz_wid, InvoiceProcessor.to_num(item['val_width'])),
                                    (mz_hei, InvoiceProcessor.to_num(item['val_height'])),
                                    (mz_mat, item['val_material']),
                                    (mz_brand, "/"), (mz_mark, "/"), (mz_link, "/"),         # 占位，按亮哥指示
                                    (mz_price, item['val_unit_price']),
                                    (mz_qty, item['val_qty_per_box']),                       # 申报数量 = 单箱数量
                                    (mz_hs, item['val_hs_code']),
                                    (mz_usage, item['val_usage']),
                                ]
                                ws.row_dimensions[mz_row].height = 65 if mz_img else 28
                                for c_idx, val in mz_writes:
                                    if c_idx:
                                        cell = ws.cell(row=mz_row, column=c_idx, value=val if pd.notna(val) else "")
                                        cell.alignment = center_alignment
                                        cell.font = Font(name="微软雅黑", size=9)
                                        cell.border = thin_border

                                val_img_url = item['val_img_url']
                                if mz_img and pd.notna(val_img_url) and str(val_img_url).startswith(('http://', 'https://')):
                                    try:
                                        resp = requests.get(str(val_img_url).strip(), timeout=6)
                                        if resp.status_code == 200:
                                            pil_img = PILImage.open(io.BytesIO(resp.content))
                                            pil_img.thumbnail((75, 75))
                                            img_buf = io.BytesIO()
                                            pil_img.save(img_buf, format='PNG')
                                            img_buf.seek(0)
                                            xl_img = OpenpyxlImage(img_buf)
                                            xl_img.width, xl_img.height = pil_img.width, pil_img.height
                                            ws.add_image(xl_img, ws.cell(row=mz_row, column=mz_img).coordinate)
                                            ws.column_dimensions[openpyxl.utils.get_column_letter(mz_img)].width = 16
                                    except Exception:
                                        pass

                                mz_row += 1

                            # 明日之星：新版模板已写完，直接保存并跳过下方通用填充逻辑（避免重复写入）
                            timestamp_date = datetime.now().strftime("%Y%m%d")
                            out_filename = f"{clean_acc}_{current_carrier}_{clean_fba}-{clean_fba_wh_code}-{total_boxes}箱-{timestamp_date}.xlsx"
                            out_buf = io.BytesIO()
                            wb.save(out_buf)
                            generated_zip_files[out_filename] = out_buf.getvalue()
                            progress_bar.progress(int(35 + (current_idx / max(len(standard_groups), 1)) * 60))
                            continue
                    else:
                        for r in range(1, min(15, ws.max_row + 1)):
                            for c in range(1, min(20, ws.max_column + 1)):
                                cell_val = str(ws.cell(row=r, column=c).value or "").strip()
                                if '客户单号' in cell_val and '客户单号' not in header_filled:
                                    ws.cell(row=r, column=c + 1, value=fba_shipment_id).alignment = center_alignment
                                    header_filled['客户单号'] = True
                                elif ('FBA仓码' in cell_val or '仓码' in cell_val) and 'FBA仓码' not in header_filled:
                                    ws.cell(row=r, column=c + 1, value=clean_fba_wh_code).alignment = center_alignment
                                    header_filled['FBA仓码'] = True

                    header_row_idx = None
                    col_mapping = {}
                    for r in range(1, min(25, ws.max_row + 1)):
                        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
                        if sum(1 for v in row_vals if any(k in v for k in ['箱数', '中文品名', '英文品名', '单箱数量', '单价', '海关编码'])) >= 3:
                            header_row_idx = r
                            for c in range(1, ws.max_column + 1):
                                cv = str(ws.cell(row=r, column=c).value or "").strip()
                                if cv:
                                    col_mapping[cv] = c
                            break

                    if not header_row_idx:
                        continue

                    def get_col(candidates):
                        for cand in candidates:
                            for h_name, col_i in col_mapping.items():
                                if cand in h_name:
                                    return col_i
                        return None

                    col_fba = get_col(['FBA编号', '货件编号'])
                    col_ref = get_col(['Reference ID', 'ReferenceId', '参考号'])
                    col_vol = get_col(['材积CM', '长*宽*高', '材积'])
                    col_box = get_col(['箱数', '箱号'])
                    col_wt = get_col(['单件毛重', '单箱重', '外箱重量'])
                    col_hs = get_col(['海关编码', 'HSCODE', 'HS编码'])
                    col_cn = get_col(['中文品名', '中文名', '品名'])
                    col_en = get_col(['英文品名', '英文名'])
                    col_qty = get_col(['单箱数量', 'PCS/CTN', '每箱数量'])
                    col_price = get_col(['单价', '单价USD', '申报单价'])
                    col_mat = get_col(['材质中/英*', '材质中/英', '材质'])
                    col_use = get_col(['用途中/英*', '用途中/英', '用途'])
                    col_img = get_col(['产品图片', '原高清图', '商品图片', '图片'])

                    start_row = header_row_idx + 1
                    current_row = start_row

                    for item in final_items:
                        box_display = item['val_box_num']
                        field_writes = [
                            (col_fba, item['val_fba_no']), (col_ref, item['val_ref_id']),
                            (col_vol, item['val_vol_dim']), (col_box, box_display),
                            (col_wt, item['val_weight']), (col_hs, item['val_hs_code']),
                            (col_cn, item['val_name_cn']), (col_en, item['val_name_en']),
                            (col_qty, item['val_qty_per_box']), (col_price, item['val_unit_price']),
                            (col_mat, item['val_material']), (col_use, item['val_usage'])
                        ]

                        ws.row_dimensions[current_row].height = 65 if col_img else 28
                        for c_idx, val in field_writes:
                            if c_idx:
                                cell = ws.cell(row=current_row, column=c_idx, value=val if pd.notna(val) else "")
                                cell.alignment = center_alignment
                                cell.font = Font(name="微软雅黑", size=9)
                                cell.border = thin_border

                        val_img_url = item['val_img_url']
                        if col_img and pd.notna(val_img_url) and str(val_img_url).startswith(('http://', 'https://')):
                            try:
                                resp = requests.get(str(val_img_url).strip(), timeout=6)
                                if resp.status_code == 200:
                                    pil_img = PILImage.open(io.BytesIO(resp.content))
                                    pil_img.thumbnail((75, 75))
                                    img_buf = io.BytesIO()
                                    pil_img.save(img_buf, format='PNG')
                                    img_buf.seek(0)
                                    xl_img = OpenpyxlImage(img_buf)
                                    xl_img.width, xl_img.height = pil_img.width, pil_img.height
                                    ws.add_image(xl_img, ws.cell(row=current_row, column=col_img).coordinate)
                                    ws.column_dimensions[openpyxl.utils.get_column_letter(col_img)].width = 16
                            except Exception:
                                pass

                        current_row += 1

                    timestamp_date = datetime.now().strftime("%Y%m%d")
                    out_filename = f"{clean_acc}_{current_carrier}_{clean_fba}-{clean_fba_wh_code}-{total_boxes}箱-{timestamp_date}.xlsx"
                    out_buf = io.BytesIO()
                    wb.save(out_buf)
                    generated_zip_files[out_filename] = out_buf.getvalue()

                    progress_bar.progress(int(35 + (current_idx / max(len(standard_groups), 1)) * 60))

                if missing_records:
                    miss_df = pd.DataFrame(missing_records)
                    miss_buf = io.BytesIO()
                    miss_df.to_excel(miss_buf, index=False)
                    generated_zip_files[f"未匹配缺失SKU清单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"] = miss_buf.getvalue()

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fn, data in generated_zip_files.items():
                        zf.writestr(fn, data)
                zip_buffer.seek(0)

                progress_bar.progress(100)
                status_text.text("✅ 所有发票与预报单处理完毕！")

                yida_count = len(yida_groups)
                std_count = len(generated_zip_files) - (1 if yida_groups else 0) - (1 if missing_records else 0)

                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.markdown(f'''<div class="metric-card-box">
                        <div class="metric-card-title">总识别货件数</div>
                        <div class="metric-card-val">{total_groups}<span class="metric-card-unit">个</span></div>
                    </div>''', unsafe_allow_html=True)
                with c2:
                    st.markdown(f'''<div class="metric-card-box">
                        <div class="metric-card-title">易达汇总货件</div>
                        <div class="metric-card-val">{yida_count}<span class="metric-card-unit">个</span></div>
                    </div>''', unsafe_allow_html=True)
                with c3:
                    st.markdown(f'''<div class="metric-card-box">
                        <div class="metric-card-title">常规单件发票</div>
                        <div class="metric-card-val">{std_count}<span class="metric-card-unit">份</span></div>
                    </div>''', unsafe_allow_html=True)
                with c4:
                    st.markdown(f'''<div class="metric-card-box">
                        <div class="metric-card-title">未匹配 SKU</div>
                        <div class="metric-card-val">{len(missing_records)}<span class="metric-card-unit">条</span></div>
                    </div>''', unsafe_allow_html=True)

                if skipped_records:
                    st.warning(f"⚠️ 有 {len(skipped_records)} 个货件因模板未配置或无法匹配而跳过。")
                if missing_records:
                    st.info(f"💡 检测到 {len(missing_records)} 条 SKU 未在采购单中匹配，已自动生成明细表放入压缩包中。")

                st.write("")

                if yida_groups and len(generated_zip_files) == 1:
                    yida_fn = list(generated_zip_files.keys())[0]
                    st.download_button(
                        label="⬇️ 一键下载【易达批量预报 Excel】",
                        data=generated_zip_files[yida_fn],
                        file_name=yida_fn,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    zip_filename = f"物流发票打包_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                    st.download_button(
                        label="⬇️ 一键下载全部发票 (ZIP 压缩包)",
                        data=zip_buffer,
                        file_name=zip_filename,
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"处理失败：{str(e)}")
