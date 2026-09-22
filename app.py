import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="ESM特气处理系统", layout="wide")
st.title("📦 ESM特气回空与入库检查系统")

# 只保留两个数据源上传入口，彻底告别模板！
col1, col2 = st.columns(2)
with col1:
    file_scan = st.file_uploader("1. 上传【回空扫描数据】(主数据)", type=["xlsx", "xls", "xlsm"])
with col2:
    file_inventory = st.file_uploader("2. 上传【当日库存数据】(参考数据)", type=["xlsx", "xls", "xlsm"])

def get_column_data(df, possible_cols):
    """辅助函数：按优先级查找列，找不到则返回空字符串"""
    for col in possible_cols:
        if col in df.columns:
            return df[col].fillna("").astype(str).str.strip()
    return pd.Series([""] * len(df), index=df.index)

if st.button("🚀 开始提取并生成报表", type="primary"):
    if file_inventory and file_scan:
        st.info("🔄 正在以【扫描数据】为主干进行匹配，并自动绘制标准报表...")
        
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

                if vLoc == "nan": vLoc = ""
                if vProd == "nan": vProd = ""
                if vProdDesc == "nan": vProdDesc = ""
                if vDef == "nan": vDef = ""
                
                # 扫描表缺失的数据，去库存表（参考数据）中找补充
                if vBN in stock_dict:
                    s_row = stock_dict[vBN]
                    if not vSN:
                        sn = str(s_row.get("SERIAL_NO", "")).strip()
                        if sn != "nan": vSN = sn
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
                        "气体": gas_name,
                        "钢瓶号": sn,
                        "容积": volume,
                        "库位": "NaN",
                        "检查结论+备注": pending_fzx.strip()
                    })
                    pending_fzx = ""

            # ================= 4. 无中生有：从零构建 Excel 文件 =================
            wb = Workbook()
            default_ws = wb.active
            default_ws.title = "回空啃单数据"
            ws_kendan = default_ws
            
            # 写入【啃单数据】表头
            headers_kendan = list(df_kendan.columns)
            ws_kendan.append(headers_kendan)
            for cell in ws_kendan[1]:
                cell.font = Font(bold=True)
                
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            
            # 写入啃单数据
            for r_idx, row in enumerate(dataframe_to_rows(df_kendan, index=False, header=False), 2):
                for c_idx, val in enumerate(row, 1):
                    cell = ws_kendan.cell(row=r_idx, column=c_idx, value=val)
                    if c_idx in [4, 5]:  # 序列号和条码强转纯文本
                        cell.number_format = '@'
                    if c_idx == 7 and val == "贸易":
                        cell.fill = yellow_fill

            # --- 写入【ESM特气仓库空瓶入库检查表】并进行 20 行分页 ---
            check_headers = ["NO", "客户名称", "气体", "钢瓶号", "容积", "库位", "检查结论+备注"]
            
            chunks = [check_data[i:i + 20] for i in range(0, len(check_data), 20)]
            if not chunks: chunks = [[]]
            
            for idx, chunk in enumerate(chunks):
                sheet_name = "ESM特气仓库空瓶入库检查表" if idx == 0 else f"ESM特气仓库空瓶入库检查表_{idx+1}"
                ws_check = wb.create_sheet(sheet_name)
                
                # 写入标准表头
                ws_check.append(check_headers)
                for cell in ws_check[1]:
                    cell.font = Font(bold=True)
                
                # 写入分页数据
                for r_idx, item in enumerate(chunk, 2):
                    ws_check.cell(row=r_idx, column=1, value=r_idx - 1)
                    ws_check.cell(row=r_idx, column=2, value=item["客户名称"])
                    ws_check.cell(row=r_idx, column=3, value=item["气体"])
                    
                    cell_sn = ws_check.cell(row=r_idx, column=4, value=item["钢瓶号"])
                    cell_sn.number_format = '@'
                    
                    ws_check.cell(row=r_idx, column=5, value=item["容积"])
                    ws_check.cell(row=r_idx, column=6, value=item["库位"])
                    
                    cell_rem = ws_check.cell(row=r_idx, column=7, value=item["检查结论+备注"])
                    cell_rem.number_format = '@'

            # ================= 5. 保存并提供下载 =================
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"✅ 生成成功！入库检查表已分为 {len(chunks)} 页。")
            
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
