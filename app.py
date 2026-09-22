import os
import sys
import subprocess

# 终极防报错机制：在代码层面强制检测并安装缺失的库，彻底告别 requirements.txt 报错！
try:
    import openpyxl
    import xlrd
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "xlrd", "pandas", "streamlit"])

import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="ESM特气处理系统", layout="wide")
st.title("📦 ESM特气回空与入库检查系统")

# 恢复三个文件的上传逻辑，确保严格使用原模板格式
col1, col2, col3 = st.columns(3)
with col1:
    file_scan = st.file_uploader("1. 上传【回空扫描数据】(主数据)", type=["xlsx", "xls", "xlsm"])
with col2:
    file_inventory = st.file_uploader("2. 上传【当日库存数据】(参考数据)", type=["xlsx", "xls", "xlsm"])
with col3:
    template_file = st.file_uploader("3. 上传【回空模板】(如回空模板_3.xlsm)", type=["xlsm", "xlsx"])

if st.button("🚀 开始提取并生成报表", type="primary"):
    if file_inventory and file_scan and template_file:
        st.info("🔄 正在以【扫描数据】为主干进行匹配，并严格套用模板格式...")
        
        try:
            # ================= 1. 读取数据 =================
            df_stock = pd.read_excel(file_inventory)
            df_scan = pd.read_excel(file_scan)
            
            df_stock.columns = [str(c).strip().upper() for c in df_stock.columns]
            df_scan.columns = [str(c).strip().upper() for c in df_scan.columns]
            
            # 确定条码列
            scan_barcode_col = "BARCODE_NO" if "BARCODE_NO" in df_scan.columns else "ITEM_BARCODE"
            stock_barcode_col = "ITEM_BARCODE" if "ITEM_BARCODE" in df_stock.columns else "BARCODE_NO"
            
            if scan_barcode_col not in df_scan.columns:
                st.error("❌ 在【回空扫描数据】中未找到条码列！")
                st.stop()
                
            # 将库存表放入字典，仅作为“参考字典”使用
            df_stock_unique = df_stock.drop_duplicates(subset=[stock_barcode_col]).copy()
            df_stock_unique[stock_barcode_col] = df_stock_unique[stock_barcode_col].astype(str).str.strip().str.upper()
            stock_dict = df_stock_unique.set_index(stock_barcode_col).to_dict('index')
            
            # ================= 2. 严格以【扫描数据】为基准提取 =================
            records = []
            for _, row in df_scan.iterrows():
                vBN = str(row.get(scan_barcode_col, "")).strip().upper()
                if vBN == "NAN" or not vBN:
                    continue
                
                # 优先抓取扫描表数据
                vLoc = str(row.get("EXPECTED_LOCATION_NAME", "")).strip()
                if not vLoc or vLoc == "nan":
                    vLoc = str(row.get("EXPECTED_LOCATION_NO", "")).strip()
                    
                vProd = str(row.get("PROD_CODE", "")).strip()
                if not vProd or vProd == "nan":
                    vProd = str(row.get("PROD_NO", "")).strip()
                    
                vProdDesc = str(row.get("PROD_DESCR", "")).strip()
                vDef = str(row.get("DEFECT_DESCR", "")).strip()
                vSN = ""

                # 清理空值标识
                if vLoc == "nan": vLoc = ""
                if vProd == "nan": vProd = ""
                if vProdDesc == "nan": vProdDesc = ""
                if vDef == "nan": vDef = ""
                
                # 扫描表缺失的数据，去库存表（参考数据）中找补充
                if vBN in stock_dict:
                    s_row = stock_dict[vBN]
                    # 序列号一定来自库存
                    if not vSN:
                        sn = str(s_row.get("SERIAL_NO", "")).strip()
                        if sn != "nan": vSN = sn
                    # 补齐其他缺失项
                    if not vLoc:
                        loc = str(s_row.get("CURRENT_LOCATION", "")).strip()
                        if loc != "nan": vLoc = loc
                    if not vProd:
                        pc = str(s_row.get("PROD_CODE", "")).strip()
                        if pc != "nan": vProd = pc
                    if not vProdDesc:
                        pdc = str(s_row.get("PROD_DESCR", "")).strip()
                        if pdc != "nan": vProdDesc = pdc
                    if not vDef:
                        dfc = str(s_row.get("DEFECT_DESCR", "")).strip()
                        if dfc != "nan": vDef = dfc
                
                records.append({
                    "EXPECTED_LOCATION_NAME": vLoc,
                    "PROD_DESCR": vProdDesc,
                    "PROD_CODE": vProd,
                    "SERIAL_NO": vSN,
                    "BARCODE_NO": vBN,
                    "DEFECT_DESCR": vDef,
                    "备注": ""
                })
                
            df_kendan = pd.DataFrame(records)
            
            # 贸易涂黄
            trade_mask = df_kendan["PROD_CODE"].str.upper().isin(["EC1MQ1", "EJ1CO1"])
            df_kendan.loc[trade_mask, "备注"] = "贸易"
            
            # 严格双重排序
            df_kendan = df_kendan.sort_values(by=["PROD_CODE", "EXPECTED_LOCATION_NAME"], ascending=[True, True])
            st.success(f"✅ 数据提取完毕，完全以回空扫描数据为准，共计 {len(df_kendan)} 条记录。")

            # ================= 3. 生成入库检查表的数据结构 =================
            check_data = []
            pending_fzx = ""
            
            for _, row in df_kendan.iterrows():
                loc = str(row["EXPECTED_LOCATION_NAME"]).strip()
                desc = str(row["PROD_DESCR"]).strip()
                code = str(row["PROD_CODE"]).strip().upper()
                sn = str(row["SERIAL_NO"]).strip()
                
                # FZX20T 拼接逻辑
                if code == "FZX20T":
                    clean_sn = sn.replace(" ", "")
                    if len(check_data) > 0:
                        if not check_data[-1]["备注"]:
                            check_data[-1]["备注"] = clean_sn
                        else:
                            check_data[-1]["备注"] += " " + clean_sn
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
                        "气体": gas_name,
                        "钢瓶号": sn,
                        "容积": volume,
                        "库位": "NA",
                        "备注": pending_fzx.strip()
                    })
                    pending_fzx = ""

            # ================= 4. 原汁原味操作您的模板文件 =================
            wb = load_workbook(template_file, keep_vba=True)
            
            # --- 写入【回空啃单数据】 ---
            if "回空啃单数据" not in wb.sheetnames:
                ws_kendan = wb.create_sheet("回空啃单数据")
                ws_kendan.append(["EXPECTED_LOCATION_NAME", "PROD_DESCR", "PROD_CODE", "SERIAL_NO", "BARCODE_NO", "DEFECT_DESCR", "备注"])
            else:
                ws_kendan = wb["回空啃单数据"]
                ws_kendan.delete_rows(2, ws_kendan.max_row)
                
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            
            for r_idx, row in enumerate(df_kendan.itertuples(index=False), 2):
                for c_idx, val in enumerate(row, 1):
                    cell = ws_kendan.cell(row=r_idx, column=c_idx, value=val)
                    if c_idx in [4, 5]: cell.number_format = '@'
                    if c_idx == 7 and val == "贸易": cell.fill = yellow_fill

            # --- 写入【ESM特气仓库空瓶入库检查表】并保留模板格式 ---
            for sheet_name in wb.sheetnames:
                if sheet_name.startswith("ESM特气仓库空瓶入库检查表_"):
                    del wb[sheet_name]
                    
            if "ESM特气仓库空瓶入库检查表" not in wb.sheetnames:
                ws_check_base = wb.create_sheet("ESM特气仓库空瓶入库检查表")
            else:
                ws_check_base = wb["ESM特气仓库空瓶入库检查表"]
                
            # 寻找表头
            header_row = 1
            col_map = {"NO": 1, "客户名称": 2, "气体": 3, "钢瓶号": 4, "容积": 5, "库位": 6, "备注": 7}
            for r in range(1, 15):
                for c in range(1, 15):
                    val = str(ws_check_base.cell(r, c).value or "").strip()
                    if "客户名称" in val:
                        header_row = r
                        col_map["客户名称"] = c
                        for cc in range(1, 15):
                            v = str(ws_check_base.cell(r, cc).value or "").strip()
                            if v == "NO": col_map["NO"] = cc
                            elif "气体" in v: col_map["气体"] = cc
                            elif "钢瓶号" in v: col_map["钢瓶号"] = cc
                            elif "容积" in v: col_map["容积"] = cc
                            elif "库位" in v: col_map["库位"] = cc
                            elif "结论" in v or "备注" in v: col_map["备注"] = cc
                        break
                if header_row != 1: break
            
            # 20行分页写入
            chunks = [check_data[i:i + 20] for i in range(0, len(check_data), 20)]
            if not chunks: chunks = [[]]
            
            for idx, chunk in enumerate(chunks):
                if idx == 0:
                    ws = ws_check_base
                else:
                    ws = wb.copy_worksheet(ws_check_base)
                    ws.title = f"ESM特气仓库空瓶入库检查表_{idx+1}"
                
                # 仅清空数据，完全保留原有的边框、行高、打印格式
                for r in range(1, 21):
                    for c in col_map.values():
                        ws.cell(header_row + r, c).value = ""
                
                for r_idx, item in enumerate(chunk):
                    ws_r = header_row + 1 + r_idx
                    ws.cell(ws_r, col_map["NO"]).value = r_idx + 1
                    ws.cell(ws_r, col_map["客户名称"]).value = item["客户名称"]
                    ws.cell(ws_r, col_map["气体"]).value = item["气体"]
                    cell_sn = ws.cell(ws_r, col_map["钢瓶号"], value=item["钢瓶号"])
                    cell_sn.number_format = '@'
                    ws.cell(ws_r, col_map["容积"]).value = item["容积"]
                    ws.cell(ws_r, col_map["库位"]).value = item["库位"]
                    cell_rem = ws.cell(ws_r, col_map["备注"], value=item["备注"])
                    cell_rem.number_format = '@'

            # ================= 5. 保存并提供下载 =================
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"✅ 入库检查表已基于【回空模板】的原生格式完美生成！共计 {len(chunks)} 页。")
            
            st.download_button(
                label="📥 点击下载生成完毕的报表",
                data=output,
                file_name="已生成_回空处理结果.xlsm",
                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
            )
            
        except Exception as e:
            st.error(f"❌ 处理过程中出现错误：{str(e)}")
    else:
        st.warning("⚠️ 请确保三个文件都已上传完毕后再点击按钮！")
