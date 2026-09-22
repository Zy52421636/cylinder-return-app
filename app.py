import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ================= 页面基础配置 =================
st.set_page_config(page_title="ESM特气处理系统", layout="wide")
st.title("📦 ESM特气回空与入库检查系统")

# 终极防崩溃检测
try:
    import pandas as pd
    from openpyxl import Workbook
except ImportError as e:
    st.error(f"🚨 云端服务器环境尚未就绪，缺少核心组件：**{e.name}**")
    st.info("请确保 GitHub 根目录下包含 `requirements.txt` 文件，并写入 streamlit, pandas, openpyxl, xlrd 四行。然后重启应用。")
    st.stop()

# 只保留两个数据源上传入口
col1, col2 = st.columns(2)
with col1:
    file_scan = st.file_uploader("1. 上传【回空扫描数据】(主数据)", type=["xlsx", "xls", "xlsm"])
with col2:
    file_inventory = st.file_uploader("2. 上传【当日库存数据】(参考数据)", type=["xlsx", "xls", "xlsm"])

def get_val(row, cols):
    """优先从提供的列名列表中获取数据"""
    for c in cols:
        if c in row.index:
            v = str(row[c]).strip()
            if v and v.lower() != "nan":
                return v
    return ""

def load_correct_sheet(file_obj, keyword):
    """智能寻找对应的工作表，防止读错 Tab 导致数据量爆炸"""
    xls = pd.ExcelFile(file_obj)
    for sheet in xls.sheet_names:
        if keyword in sheet:
            return pd.read_excel(xls, sheet_name=sheet)
    # 如果没找到带关键字的，默认读第一个表
    return pd.read_excel(xls, sheet_name=0)

def draw_template_format(ws):
    """100% 还原工作簿2.xlsx的精确模板格式"""
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    bottom_line_border = Border(bottom=Side(style='thin')) 
    
    widths = {'A': 4.5, 'B': 25, 'C': 12, 'D': 20, 'E': 6.5, 'F': 7, 
              'G': 8.5, 'H': 13, 'I': 13, 'J': 15, 'K': 13, 'L': 25}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.row_dimensions[1].height = 21
    ws.row_dimensions[2].height = 38.25
    ws.row_dimensions[3].height = 51
    for r in range(4, 24):
        ws.row_dimensions[r].height = 22.5
    ws.row_dimensions[24].height = 25.5
    ws.row_dimensions[25].height = 23.25
    ws.row_dimensions[26].height = 16
    ws.row_dimensions[27].height = 26.25
    ws.row_dimensions[28].height = 26.25

    ws.merge_cells('A1:C2')
    ws.merge_cells('D1:J2')
    ws.merge_cells('K1:L2')

    ws['D1'] = 'ESM特气仓库空瓶入库检查表'
    ws['D1'].font = Font(name='微软雅黑', size=18, bold=True)
    ws['D1'].alignment = Alignment(horizontal='center', vertical='center')

    ws['K1'] = '文件号：ALCH-SOP-ESM/WH-001-RD02\n修    订：2025-12-31\n版    本：1'
    ws['K1'].font = Font(name='微软雅黑', size=10)
    ws['K1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    headers = ['No', '客户名称', '气体名称', '钢瓶号', '容积', '库位', 
               '外观*', '瓶帽*', '阀门*', '标签*\n（尤其Barcode标签）', '文件核对*', '检查结论+备注']
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.font = Font(name='微软雅黑', size=10, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border

    for r in range(4, 24):
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(name='微软雅黑', size=9)

    ws.merge_cells('A24:B24')
    ws['A24'] = '检查人'
    ws.merge_cells('C24:E24')
    for c in range(3, 6): ws.cell(row=24, column=c).border = bottom_line_border
    
    ws.merge_cells('F24:G24')
    ws['F24'] = '复核人'
    ws.merge_cells('H24:I24')
    for c in range(8, 10): ws.cell(row=24, column=c).border = bottom_line_border
        
    ws['J24'] = '检查日期'
    ws.merge_cells('K24:L24')
    for c in range(11, 13): ws.cell(row=24, column=c).border = bottom_line_border
    
    notes = [
        ('1.', '外观：瓶体清洁无锈迹，油漆完好无损坏，无凹痕、腐蚀等异常。瓶身喷漆字迹完好。保护套完好（钢瓶），容器附件完好（如Ton tank防撞栏）；', '5.', '文件核对：DO单，可能有客户返回空瓶记录表；'),
        ('2.', '瓶帽：瓶帽与瓶体匹配，内外部清洁无锈迹，油漆完好无损坏；', '6.', '如发现异常状况，在备注栏填写相关信息，并立即通报相关人员。'),
        ('3.', '阀门：阀门无锈迹，阀门及底座周围无腐蚀、损坏等异常。阀门出口垂直（Ton tank）；', '7.', '如使用花篮，需检查绑带（5年有效期）、棘轮、花篮框架有无异常；'),
        ('4.', '标签：包含产品合格证、气体性质标签、满瓶标签及barcode标签（三张标签确保内容一致性）。', '', '')
    ]
    
    for idx, (n1, t1, n2, t2) in enumerate(notes, 25):
        ws.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=6)
        ws.merge_cells(start_row=idx, start_column=8, end_row=idx, end_column=12)
        ws.cell(row=idx, column=1, value=n1)
        ws.cell(row=idx, column=2, value=t1)
        ws.cell(row=idx, column=7, value=n2)
        ws.cell(row=idx, column=8, value=t2)

    for r in range(24, 29):
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name='微软雅黑', size=9)
            if c in [1, 6, 7, 10]: 
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

if st.button("🚀 开始提取并生成报表", type="primary"):
    if file_inventory and file_scan:
        st.info("🔄 正在清洗数据，并以【回空扫描数据】为主表提取及生成报表...")
        
        try:
            # ================= 1. 智能读取数据 =================
            df_stock = load_correct_sheet(file_inventory, "库存")
            df_scan = load_correct_sheet(file_scan, "扫描")
            
            df_stock.columns = [str(c).strip().upper() for c in df_stock.columns]
            df_scan.columns = [str(c).strip().upper() for c in df_scan.columns]
            
            # 在库存表中剔除异常预期数据
            if "INVENTORY_ITEM_STATUS" in df_stock.columns:
                mask = df_stock["INVENTORY_ITEM_STATUS"].astype(str).str.strip() != "Not scanned but Expected"
                df_stock = df_stock[mask].copy()

            scan_barcode_col = "BARCODE_NO" if "BARCODE_NO" in df_scan.columns else "ITEM_BARCODE"
            stock_barcode_col = "ITEM_BARCODE" if "ITEM_BARCODE" in df_stock.columns else "BARCODE_NO"
            
            if scan_barcode_col not in df_scan.columns:
                st.error("❌ 在【回空扫描数据】中未找到条码列，无法处理！")
                st.stop()
                
            # 【关键修改】：强制清理扫描表中的空行，防止读取到Excel尾部的空白格式行
            df_scan = df_scan.dropna(subset=[scan_barcode_col])
            df_scan = df_scan[df_scan[scan_barcode_col].astype(str).str.strip() != ""]
            df_scan = df_scan[df_scan[scan_barcode_col].astype(str).str.lower() != "nan"]
                
            # 库存表仅作为参考字典
            df_stock_unique = df_stock.drop_duplicates(subset=[stock_barcode_col]).copy()
            df_stock_unique[stock_barcode_col] = df_stock_unique[stock_barcode_col].astype(str).str.strip().str.upper()
            stock_dict = df_stock_unique.set_index(stock_barcode_col).to_dict('index')
            
            # ================= 2. 绝对以【扫描数据】为主干提取 =================
            records = []
            for _, row in df_scan.iterrows():
                vBN = str(row.get(scan_barcode_col, "")).strip().upper()
                
                vLoc = get_val(row, ["EXPECTED_LOCATION_NAME", "EXPECTED_LOCATION_NO"])
                vProd = get_val(row, ["PROD_CODE", "PROD_NO"])
                vProdDesc = get_val(row, ["PROD_DESCR"])
                vSN = get_val(row, ["SERIAL_NO"])
                vDef = get_val(row, ["DEFECT_DESCR"])
                vRem = get_val(row, ["备注", "REMARK"])

                if vBN in stock_dict:
                    s_row = stock_dict[vBN]
                    if not vSN:
                        vSN = str(s_row.get("SERIAL_NO", "")).strip()
                        if vSN.lower() == "nan": vSN = ""
                    if not vDef:
                        vDef = str(s_row.get("DEFECT_DESCR", "")).strip()
                        if vDef.lower() == "nan": vDef = ""
                    if not vRem:
                        vRem = str(s_row.get("备注", "")).strip()
                        if vRem.lower() == "nan": vRem = ""
                
                records.append({
                    "EXPECTED_LOCATION_NAME": vLoc,
                    "PROD_DESCR": vProdDesc,
                    "PROD_CODE": vProd,
                    "SERIAL_NO": vSN,
                    "BARCODE_NO": vBN,
                    "DEFECT_DESCR": vDef,
                    "备注": vRem
                })
                
            df_kendan = pd.DataFrame(records)
            trade_mask = df_kendan["PROD_CODE"].str.upper().isin(["EC1MQ1", "EJ1CO1"])
            df_kendan.loc[trade_mask, "备注"] = "贸易"
            df_kendan = df_kendan.sort_values(by=["PROD_CODE", "EXPECTED_LOCATION_NAME"], ascending=[True, True])
            
            st.success(f"✅ 数据提取完毕，完全以回空扫描为主干，成功精准提取 {len(df_kendan)} 条有效记录。")

            # ================= 3. 生成入库检查表的数据结构 =================
            check_data = []
            pending_fzx = ""
            
            for _, row in df_kendan.iterrows():
                loc = str(row["EXPECTED_LOCATION_NAME"]).strip()
                desc = str(row["PROD_DESCR"]).strip()
                code = str(row["PROD_CODE"]).strip().upper()
                sn = str(row["SERIAL_NO"]).strip()
                
                if code == "FZX20T":
                    clean_sn = sn.replace(" ", "")
                    if len(check_data) > 0:
                        if not check_data[-1]["检查结论+备注"]:
                            check_data[-1]["检查结论+备注"] = clean_sn
                        else:
                            check_data[-1]["检查结论+备注"] += " " + clean_sn
                    else:
                        pending_fzx += (" " + clean_sn) if pending_fzx else clean_sn
                else:
                    gas_name = ""
                    volume = ""
                    if desc:
                        arr_desc = desc.split()
                        gas_name = arr_desc[0]
                        for part in arr_desc:
                            if "L" in part.upper() and part[0].isdigit():
                                volume = part
                                break
                    
                    check_data.append({
                        "客户名称": loc,
                        "气体名称": gas_name,
                        "钢瓶号": sn,
                        "容积": volume,
                        "库位": "",
                        "检查结论+备注": pending_fzx.strip()
                    })
                    pending_fzx = ""

            # ================= 4. 构建 Excel 并自动化绘制模板 =================
            wb = Workbook()
            default_ws = wb.active
            default_ws.title = "回空啃单数据"
            ws_kendan = default_ws
            
            headers_kendan = list(df_kendan.columns)
            ws_kendan.append(headers_kendan)
            for cell in ws_kendan[1]:
                cell.font = Font(bold=True)
                
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            
            for r_idx, row in enumerate(dataframe_to_rows(df_kendan, index=False, header=False), 2):
                for c_idx, val in enumerate(row, 1):
                    cell = ws_kendan.cell(row=r_idx, column=c_idx, value=val)
                    if c_idx in [4, 5]: 
                        cell.number_format = '@'
                    if c_idx == 7 and val == "贸易":
                        cell.fill = yellow_fill

            chunks = [check_data[i:i + 20] for i in range(0, len(check_data), 20)]
            if not chunks: chunks = [[]]
            
            for idx, chunk in enumerate(chunks):
                sheet_name = "ESM特气仓库空瓶入库检查表" if idx == 0 else f"ESM特气仓库空瓶入库检查表_{idx+1}"
                ws_check = wb.create_sheet(sheet_name)
                
                draw_template_format(ws_check)
                
                for r_idx, item in enumerate(chunk, 4):
                    ws_check.cell(row=r_idx, column=1, value=r_idx - 3)
                    ws_check.cell(row=r_idx, column=2, value=item["客户名称"])
                    ws_check.cell(row=r_idx, column=3, value=item["气体名称"])
                    cell_sn = ws_check.cell(row=r_idx, column=4, value=item["钢瓶号"])
                    cell_sn.number_format = '@' 
                    ws_check.cell(row=r_idx, column=5, value=item["容积"])
                    ws_check.cell(row=r_idx, column=6, value=item["库位"])
                    cell_rem = ws_check.cell(row=r_idx, column=12, value=item["检查结论+备注"])
                    cell_rem.number_format = '@'

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.download_button(
                label="📥 点击下载当日生成的 【回空啃单及检查表.xlsx】",
                data=output,
                file_name="当日生成_回空啃单及检查表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ 处理过程中出现错误：{str(e)}")
    else:
        st.warning("⚠️ 请确保【库存数据】和【扫描数据】均已上传！")
