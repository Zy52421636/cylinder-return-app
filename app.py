import streamlit as st
import pandas as pd
import io
import base64
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image

# ================= 页面基础配置 =================
st.set_page_config(page_title="ESM特气处理系统", layout="wide")
st.title("📦 ESM特气回空与入库检查系统")

# 终极防崩溃检测
try:
    import pandas as pd
    from openpyxl import Workbook
except ImportError as e:
    st.error(f"🚨 云端服务器环境尚未就绪，缺少核心组件：**{e.name}**")
    st.info("请确保 GitHub 根目录下包含 `requirements.txt` 文件，并写入 streamlit, pandas, openpyxl, xlrd, Pillow 五行。然后重启应用。")
    st.stop()

# 内置 Air Liquide Logo 的 Base64 编码，实现无文件依赖自动绘制 Logo
LOGO_BASE64 = "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 "

# 智能寻找工作表，防止读错 Tab
def load_correct_sheet(file_obj, keyword):
    xls = pd.ExcelFile(file_obj)
    for sheet in xls.sheet_names:
        if keyword in sheet:
            return pd.read_excel(xls, sheet_name=sheet)
    return pd.read_excel(xls, sheet_name=0)

def draw_template_format(ws):
    """100% 还原模板格式：包含全线条边框、内置 Logo、下划线及 A4 横向打印设置"""
    
    # 全线框样式
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    bottom_line_border = Border(bottom=Side(style='thin', color='000000')) 
    
    # 1. 列宽设置
    widths = {'A': 4.5, 'B': 25, 'C': 12, 'D': 20, 'E': 6.5, 'F': 7, 
              'G': 8.5, 'H': 13, 'I': 13, 'J': 15, 'K': 13, 'L': 25}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # 2. 行高设置
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

    # 3. 合并单元格
    ws.merge_cells('A1:C2')
    ws.merge_cells('D1:J2')
    ws.merge_cells('K1:L2')

    # 4. 解码并插入内置的 Logo 图片（完美解决云端找不到图片的问题）
    try:
        img_data = base64.b64decode(LOGO_BASE64)
        img = OpenpyxlImage(io.BytesIO(img_data))
        img.width = 135
        img.height = 42
        ws.add_image(img, "A1")
    except Exception:
        pass

    # 5. 第 1-2 行头部区域添加全边框
    for r in range(1, 3):
        for c in range(1, 13):
            ws.cell(row=r, column=c).border = thin_border

    # 6. 写入标题与文件声明信息
    ws['D1'] = 'ESM特气仓库空瓶入库检查表'
    ws['D1'].font = Font(name='微软雅黑', size=18, bold=True)
    ws['D1'].alignment = Alignment(horizontal='center', vertical='center')

    ws['K1'] = '文件号：ALCH-SOP-ESM/WH-001-RD02\n修    订：2025-12-31\n版    本：1'
    ws['K1'].font = Font(name='微软雅黑', size=10)
    ws['K1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 7. 写入第三行表头及全线框
    headers = ['No', '客户名称', '气体名称', '钢瓶号', '容积', '库位', 
               '外观*', '瓶帽*', '阀门*', '标签*\n（尤其Barcode标签）', '文件核对*', '检查结论+备注']
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.font = Font(name='微软雅黑', size=10, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border

    # 8. 数据区(4-23行)全线框及格式预设
    for r in range(4, 24):
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(name='微软雅黑', size=9)

    # 9. 底部签字栏与下划线
    ws.merge_cells('A24:B24')
    ws['A24'] = '检查人'
    ws.merge_cells('C24:E24')
    ws.merge_cells('F24:G24')
    ws['F24'] = '复核人'
    ws.merge_cells('H24:I24')
    ws['J24'] = '检查日期'
    ws.merge_cells('K24:L24')

    for c in range(1, 13):
        ws.cell(row=24, column=c).border = thin_border
    
    notes = [
        ('1.', '外观：瓶体清洁无锈迹，油漆完好无损坏，无凹痕、腐蚀等异常。瓶身喷漆字迹完好。保护套完好（钢瓶），容器附件完好（如Ton tank防撞栏）；', '5.', '文件核对：DO单，可能有客户返回空瓶记录表；'),
        ('2.', '瓶帽：瓶帽与瓶体匹配，内外部清洁无锈迹，油漆完好无损坏；', '6.', '如发现异常状况，在备注栏填写相关信息，并立即通报相关人员。'),
        ('3.', '阀门：阀门无锈迹，阀门及底座周围无腐蚀、损坏等异常。阀门出口垂直（Ton tank）；', '7.', '如使用花篮，需检查绑带（5年有效期），棘轮、花篮框架有无异常；'),
        ('4.', '标签：包含产品合格证、气体性质标签、满瓶标签及barcode标签（三张标签确保内容一致性）。', '', '')
    ]
    
    for idx, (n1, t1, n2, t2) in enumerate(notes, 25):
        ws.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=6)
        ws.merge_cells(start_row=idx, start_column=8, end_row=idx, end_column=12)
        ws.cell(row=idx, column=1, value=n1)
        ws.cell(row=idx, column=2, value=t1)
        ws.cell(row=idx, column=7, value=n2)
        ws.cell(row=idx, column=8, value=t2)

        for c in range(1, 13):
            cell = ws.cell(row=idx, column=c)
            cell.border = thin_border
            cell.font = Font(name='微软雅黑', size=8)
            if c in [1, 7]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 10. A4 纸横向打印自适应配置
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4

if st.button("🚀 开始提取并生成报表", type="primary"):
    if file_inventory and file_scan:
        st.info("🔄 正在提取回空扫描数据，并从当日库存中匹配 DEFECT_DESCR 与 备注...")
        
        try:
            # 1. 读取数据
            df_stock = load_correct_sheet(file_inventory, "库存")
            df_scan = load_correct_sheet(file_scan, "扫描")
            
            df_stock.columns = [str(c).strip().upper() for c in df_stock.columns]
            df_scan.columns = [str(c).strip().upper() for c in df_scan.columns]
            
            # 过滤扫描表及库存表中的 Not scanned 行
            if "INVENTORY_ITEM_STATUS" in df_scan.columns:
                mask_scan = ~df_scan["INVENTORY_ITEM_STATUS"].astype(str).str.lower().str.contains("not scanned", na=False)
                df_scan = df_scan[mask_scan].copy()

            if "INVENTORY_ITEM_STATUS" in df_stock.columns:
                mask_stock = ~df_stock["INVENTORY_ITEM_STATUS"].astype(str).str.lower().str.contains("not scanned", na=False)
                df_stock = df_stock[mask_stock].copy()

            scan_barcode_col = "BARCODE_NO" if "BARCODE_NO" in df_scan.columns else "ITEM_BARCODE"
            stock_barcode_col = "ITEM_BARCODE" if "ITEM_BARCODE" in df_stock.columns else "BARCODE_NO"
            
            if scan_barcode_col not in df_scan.columns:
                st.error("❌ 在【回空扫描数据】中未找到条码列！")
                st.stop()
                
            df_scan = df_scan.dropna(subset=[scan_barcode_col])
            df_scan = df_scan[df_scan[scan_barcode_col].astype(str).str.strip() != ""]
            df_scan = df_scan[df_scan[scan_barcode_col].astype(str).str.lower() != "nan"]
            
            # 建立条码与序列号的双重查表字典
            stock_barcode_dict = {}
            stock_sn_dict = {}
            
            for _, s_row in df_stock.iterrows():
                b_val = str(s_row.get(stock_barcode_col, "")).strip().upper()
                sn_val = str(s_row.get("SERIAL_NO", "")).strip().upper()
                row_dict = s_row.to_dict()
                
                if b_val and b_val != "NAN":
                    stock_barcode_dict[b_val] = row_dict
                if sn_val and sn_val != "NAN":
                    stock_sn_dict[sn_val] = row_dict
            
            # 2. 以扫描数据为主表提取
            records = []
            for _, row in df_scan.iterrows():
                vBN = str(row.get(scan_barcode_col, "")).strip().upper()
                
                vLoc = get_val(row, ["EXPECTED_LOCATION_NAME", "EXPECTED_LOCATION_NO"])
                vProd = get_val(row, ["PROD_CODE", "PROD_NO"])
                vProdDesc = get_val(row, ["PROD_DESCR"])
                vSN = get_val(row, ["SERIAL_NO"])
                vDef = get_val(row, ["DEFECT_DESCR"])
                vRem = get_val(row, ["备注", "REMARK"])

                ref_row = stock_barcode_dict.get(vBN) or stock_sn_dict.get(vSN.upper())
                
                if ref_row:
                    if not vSN:
                        sn = str(ref_row.get("SERIAL_NO", "")).strip()
                        if sn.lower() != "nan": vSN = sn
                    if not vDef:
                        dfc = str(ref_row.get("DEFECT_DESCR", "")).strip()
                        if dfc.lower() != "nan": vDef = dfc
                    if not vRem:
                        rmk = str(ref_row.get("备注", ref_row.get("REMARK", ""))).strip()
                        if rmk.lower() != "nan": vRem = rmk
                
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
            
            st.success(f"✅ 数据提取成功，共提取 {len(df_kendan)} 条回空记录。")

            # 3. 生成入库检查表数据结构
            check_data = []
            pending_fzx = ""
            
            for _, row in df_kendan.iterrows():
                loc = str(row["EXPECTED_LOCATION_NAME"]).strip()
                desc = str(row["PROD_DESCR"]).strip()
                code = str(row["PROD_CODE"]).strip().upper()
                sn = str(row["SERIAL_NO"]).strip()
                def_desc = str(row["DEFECT_DESCR"]).strip()
                rem_desc = str(row["备注"]).strip()
                
                combine_remark = []
                if def_desc: combine_remark.append(def_desc)
                if rem_desc and rem_desc != "贸易": combine_remark.append(rem_desc)
                final_remark = ", ".join(combine_remark)
                
                if code == "FZX20T":
                    clean_sn = sn.replace(" ", "")
                    if len(check_data) > 0:
                        existing = check_data[-1]["检查结论+备注"]
                        check_data[-1]["检查结论+备注"] = (existing + " " + clean_sn).strip()
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
                    
                    if pending_fzx:
                        final_remark = (final_remark + " " + pending_fzx).strip() if final_remark else pending_fzx.strip()
                        pending_fzx = ""
                        
                    check_data.append({
                        "客户名称": loc,
                        "气体名称": gas_name,
                        "钢瓶号": sn,
                        "容积": volume,
                        "库位": "",
                        "检查结论+备注": final_remark
                    })

            # 4. 构建 Excel 并自动化绘制模板
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
                
                # 绘制表头与框架（包含 Base64 解码并绘制内置 Logo）
                draw_template_format(ws_check)
                
                # 填入数据
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
