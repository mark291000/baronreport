# =========================================================
# TASK DASHBOARD WITH IMAGE DISPLAY - FULL VERSION
# =========================================================

import pandas as pd
import openpyxl
import base64
from datetime import datetime
from IPython.display import display, HTML
import plotly.graph_objects as go
import plotly.express as px

# === 1️⃣ Cấu hình ===
file_path = "/content/New Go Plastic Wanek 6.xlsx"
today = pd.Timestamp.now().normalize()

# === 2️⃣ Đọc Excel với openpyxl để lấy hình ảnh ===
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb.active

# Lưu hình ảnh (cột PICTURE)
images = {}
for image in ws._images:
    cell = image.anchor._from
    cell_coord = f"{openpyxl.utils.get_column_letter(cell.col + 1)}{cell.row + 1}"
    img_bytes = image.ref.getvalue() if hasattr(image.ref, 'getvalue') else image.ref
    images[cell_coord] = base64.b64encode(img_bytes).decode('utf-8')

# === 3️⃣ Đọc Excel với pandas (header dòng 3) ===
df = pd.read_excel(file_path, header=2)  # header ở dòng 3
df.columns = df.columns.str.strip()  # trim cột tránh khoảng trắng dư

# === 4️⃣ Xử lý Status với New Task ===
def get_status(row):
    confirm = str(row.get("CONFIRM FROM BARON", "")).strip().lower()
    start_date = row.get("START DATE")
    due = row.get("DUE DATE")
    
    # Kiểm tra Completed
    if "go" in confirm:
        return "Completed"
    
    # Kiểm tra New Task (start date > today)
    if pd.notna(start_date):
        try:
            if pd.to_datetime(start_date).date() > today.date():
                return "New Task"
        except:
            pass
    
    # Kiểm tra Delay và Working
    if pd.notna(due):
        try:
            if pd.to_datetime(due).date() < today.date():
                return "Delay"
            else:
                return "Working"
        except:
            return "Working"
    
    return "Working"

df["STATUS"] = df.apply(get_status, axis=1)

# === 5️⃣ Định dạng ngày và xử lý NaN ===
df["START DATE"] = pd.to_datetime(df["START DATE"], errors="coerce").dt.strftime("%m/%d/%Y")
df["DUE DATE"] = pd.to_datetime(df["DUE DATE"], errors="coerce").dt.strftime("%m/%d/%Y")

# Thay thế tất cả NaN và NaT bằng chuỗi rỗng
df = df.fillna("")
df = df.replace("NaT", "")

# === 6️⃣ Pie chart trạng thái ===
status_counts = df["STATUS"].value_counts()
if not status_counts.empty:
    fig_pie = px.pie(
        df, 
        names="STATUS",
        title="Tỷ lệ STATUS các Task",
        color="STATUS",
        color_discrete_map={
            "Completed": "green",
            "Working": "orange",
            "Delay": "red",
            "New Task": "blue"
        }
    )
    fig_pie.update_traces(textinfo='percent+label', pull=[0.05]*len(df["STATUS"].unique()))
    fig_pie.show()

# === 7️⃣ Bar chart số task theo tháng (grouped by status - 4 cột) ===
df_with_dates = df[df["START DATE"] != ""].copy()
if not df_with_dates.empty:
    df_with_dates["month"] = pd.to_datetime(df_with_dates["START DATE"], errors="coerce").dt.strftime("%Y-%m")
    df_summary = df_with_dates.groupby(["month", "STATUS"]).size().reset_index(name="count")
    df_summary = df_summary[df_summary["month"].notna()]
    
    if not df_summary.empty:
        # Lấy tất cả các tháng unique
        all_months = sorted(df_summary["month"].unique())
        all_statuses = ["Completed", "Working", "New Task", "Delay"]
        
        # Tạo DataFrame đầy đủ với tất cả combinations
        full_data = []
        for month in all_months:
            for status in all_statuses:
                existing = df_summary[(df_summary["month"] == month) & (df_summary["STATUS"] == status)]
                if not existing.empty:
                    full_data.append({"month": month, "STATUS": status, "count": existing["count"].values[0]})
                else:
                    full_data.append({"month": month, "STATUS": status, "count": 0})
        
        df_full = pd.DataFrame(full_data)
        
        # Tạo bar chart với 4 cột cho mỗi tháng
        fig_bar = go.Figure()
        
        colors = {
            "Completed": "green",
            "Working": "orange",
            "New Task": "blue",
            "Delay": "red"
        }
        
        for status in all_statuses:
            df_status = df_full[df_full["STATUS"] == status]
            fig_bar.add_trace(go.Bar(
                x=df_status["month"],
                y=df_status["count"],
                name=status,
                marker_color=colors.get(status, "gray"),
                text=df_status["count"],
                textposition='outside',  # Label nằm ngoài trên cột
                textfont=dict(size=12),
            ))
        
        fig_bar.update_layout(
            title="Số lượng Task theo tháng",
            xaxis_title="Tháng",
            yaxis_title="Số lượng Task",
            barmode='group',  # Grouped bar chart
            xaxis=dict(
                tickformat="%Y-%m",
                type='category'
            ),
            hovermode='x unified',
            height=500,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        fig_bar.show()

# === 8️⃣ Hàm xử lý giá trị rỗng cho HTML ===
def safe_value(val):
    """Chuyển đổi giá trị an toàn cho HTML, tránh hiển thị nan"""
    if pd.isna(val) or val == "" or str(val).lower() == "nan" or str(val) == "NaT":
        return ""
    return str(val).strip()

# === 9️⃣ Lấy header từ dòng 3 của Excel ===
header_row = [cell.value for cell in ws[3]]
required_cols = ["TASK", "Requester", "START DATE", "DUE DATE", "CONFIRM FROM BARON", "STATUS", "PICTURE"]

# Tìm index của các cột cần thiết
col_indices = {}
for col in required_cols:
    try:
        col_indices[col] = header_row.index(col)
    except ValueError:
        col_indices[col] = None

# === 🔟 Tạo bảng HTML với hình ảnh ===
status_options = df["STATUS"].unique().tolist()

table_html = """
<style>
  .task-table {
    border-collapse: collapse;
    width: 100%;
    min-width: 1000px;
  }
  .task-table th {
    background: #4CAF50;
    color: white;
    padding: 10px;
    text-align: center;
    position: sticky;
    top: 0;
    z-index: 10;
  }
  .task-table td {
    padding: 8px;
    text-align: center;
    vertical-align: middle;
    border: 1px solid #ddd;
  }
  .task-table tr:hover {
    background-color: #f5f5f5;
  }
  .image-cell {
    width: 120px;
    text-align: center;
  }
  .image-cell img {
    max-width: 100px;
    max-height: 100px;
    object-fit: contain;
    cursor: pointer;
    border: 1px solid #ddd;
    border-radius: 4px;
    transition: transform 0.2s;
  }
  .image-cell img:hover {
    transform: scale(3.5);
    z-index: 1000;
    position: relative;
    box-shadow: 0 8px 16px rgba(0,0,0,0.3);
  }
  .status-completed {
    background-color: #d4edda;
    font-weight: bold;
    color: #155724;
  }
  .status-working {
    background-color: #fff3cd;
    font-weight: bold;
    color: #856404;
  }
  .status-delay {
    background-color: #f8d7da;
    font-weight: bold;
    color: #721c24;
  }
  .status-newtask {
    background-color: #cce5ff;
    font-weight: bold;
    color: #004085;
  }
</style>

<h3>📋 Task Dashboard</h3>
<label for="statusFilter">Lọc theo STATUS: </label>
<select id="statusFilter" onchange="filterTable()">
  <option value="All">All</option>
""" + "".join([f"<option value='{s}'>{s}</option>" for s in status_options]) + """
</select>

<div style='margin-top:10px; overflow-x:auto; max-height: 600px; overflow-y: auto;'>
<table id="taskTable" class="task-table">
  <thead>
    <tr>
""" + "".join([f"<th>{col}</th>" for col in required_cols]) + """
    </tr>
  </thead>
  <tbody>
"""

# Duyệt qua các dòng từ dòng 4 (row index 3 trong openpyxl, vì bắt đầu từ 0)
for row_idx, row in enumerate(ws.iter_rows(min_row=4), start=4):
    # Lấy giá trị STATUS từ DataFrame
    df_row_idx = row_idx - 4  # Vì DataFrame bắt đầu từ 0
    status_value = ""
    
    if df_row_idx < len(df):
        status_value = safe_value(df.iloc[df_row_idx]["STATUS"])
    
    row_html = f"<tr data-status='{status_value}'>"
    
    for col_name in required_cols:
        col_idx = col_indices[col_name]
        value = ""
        
        if col_idx is not None:
            cell = row[col_idx]
            value = cell.value if cell.value is not None else ""
            
            # Xử lý cột STATUS - lấy từ DataFrame thay vì Excel
            if col_name == "STATUS":
                status_class = ""
                if status_value == "Completed":
                    status_class = "status-completed"
                elif status_value == "Working":
                    status_class = "status-working"
                elif status_value == "Delay":
                    status_class = "status-delay"
                elif status_value == "New Task":
                    status_class = "status-newtask"
                row_html += f"<td class='{status_class}'>{status_value}</td>"
                continue
            
            # Xử lý cột PICTURE
            if col_name == "PICTURE":
                cell_coord = cell.coordinate
                if cell_coord in images:
                    img_tag = f"<img src='data:image/png;base64,{images[cell_coord]}' alt='Task Image'/>"
                    row_html += f"<td class='image-cell'>{img_tag}</td>"
                else:
                    row_html += "<td class='image-cell'></td>"
                continue
            
            # Xử lý ngày tháng
            if col_name in ["START DATE", "DUE DATE"]:
                if value and value != "":
                    try:
                        if isinstance(value, datetime):
                            value = value.strftime("%m/%d/%Y")
                        else:
                            value = pd.to_datetime(value).strftime("%m/%d/%Y")
                    except:
                        pass
        
        row_html += f"<td>{value}</td>"
    
    row_html += "</tr>"
    table_html += row_html

table_html += """
  </tbody>
</table>
</div>

<script>
function filterTable(){
  var select = document.getElementById("statusFilter");
  var filter = select.value;
  var table = document.getElementById("taskTable");
  var tr = table.getElementsByTagName("tr");
  
  for(var i = 1; i < tr.length; i++){
    var status = tr[i].getAttribute("data-status");
    if(status){
      tr[i].style.display = (filter === "All" || status === filter) ? "" : "none";
    }
  }
}
</script>
"""

# === 1️⃣1️⃣ Hiển thị dashboard ===
display(HTML(table_html))
