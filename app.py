import os
import io
import openpyxl
import streamlit as st
from lxml import etree
from shapely.geometry import LineString
from bs4 import BeautifulSoup

# -------------------------------------
# 🌟 ตกแต่ง UI ด้วย CSS & Bootstrap
# -------------------------------------
st.set_page_config(page_title="KML to Excel Converter", page_icon="📄", layout="wide")

st.markdown("""
    <style>
    body {
        background-color: #f8f9fa;
    }
    .main-title {
        color: #007bff;
        text-align: center;
        font-size: 36px;
        font-weight: bold;
        margin-bottom: 10px;
    }
    .sub-title {
        color: #6c757d;
        text-align: center;
        font-size: 18px;
        margin-bottom: 20px;
    }
    .stButton>button {
        background-color: #28a745;
        color: white;
        font-size: 16px;
        border-radius: 10px;
        padding: 12px 20px;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #218838;
        transform: scale(1.05);
    }
    .stDownloadButton>button {
        background-color: #007bff;
        color: white;
        font-size: 16px;
        border-radius: 10px;
        padding: 12px 20px;
        transition: 0.3s;
    }
    .stDownloadButton>button:hover {
        background-color: #0056b3;
        transform: scale(1.05);
    }
    </style>
    """, unsafe_allow_html=True)

# -------------------------------------
# 🚀 ฟังก์ชันประมวลผล
# -------------------------------------
def extract_description_data(description_html):
    """Extract structured data from the HTML description."""
    soup = BeautifulSoup(description_html, 'html.parser')
    extracted_data = {}
    rows = soup.find_all('tr')
    
    for row in rows:
        cells = row.find_all('td')
        if len(cells) >= 2:
            header = cells[0].get_text(strip=True)  
            value = cells[1].get_text(strip=True)  
            extracted_data[header] = value  

    return extracted_data

def load_kml_lines(uploaded_file):
    """Load lines from a KML file and return them as Shapely LineString objects with metadata."""
    kml_data = etree.fromstring(uploaded_file.getvalue())  
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}  
    lines = []
    
    for placemark in kml_data.xpath('.//kml:Placemark', namespaces=ns):
        line_string = placemark.xpath('.//kml:LineString/kml:coordinates', namespaces=ns)
        if line_string:
            coords_text = line_string[0].text.strip()
            coords = [(float(c.split(',')[0]), float(c.split(',')[1])) for c in coords_text.split()]
            line_geom = LineString(coords) if len(coords) > 1 else None
            name = placemark.xpath('./kml:name/text()', namespaces=ns)
            name = name[0] if name else "Unnamed"
            description = placemark.xpath('./kml:description/text()', namespaces=ns)
            description_data = extract_description_data(description[0]) if description else {}

            if line_geom:
                lines.append({
                    "Name": name,
                    "Description": description_data,  
                    "Line": line_geom,
                    "Start_Coordinate": coords[0] if coords else None,
                    "End_Coordinate": coords[-1] if coords else None
                })
            else:
                st.warning(f"Skipped invalid LineString: {name}")

    return lines

def save_to_excel_memory(lines):
    """Save line data to Excel in memory."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "KML Data"
    headers = ['Name', 'Start_Coordinate', 'End_Coordinate']
    
    if lines:
        first_line = lines[0]
        if isinstance(first_line['Description'], dict):
            headers.extend(first_line['Description'].keys())

    sheet.append(headers)

    for line in lines:
        row = [
            line["Name"],
            f"{line['Start_Coordinate'][0]},{line['Start_Coordinate'][1]}" if line["Start_Coordinate"] else "N/A",
            f"{line['End_Coordinate'][0]},{line['End_Coordinate'][1]}" if line["End_Coordinate"] else "N/A"
        ]
        if isinstance(line['Description'], dict):
            for header in headers[3:]:
                row.append(line['Description'].get(header, "N/A"))
        sheet.append(row)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    
    return output

# -------------------------------------
# 🎯 ส่วน UI ของ Streamlit
# -------------------------------------
st.markdown('<p class="main-title">KML to Excel Converter 🚀</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">📄 แปลงไฟล์ KML เป็น Excel พร้อมข้อมูล Metadata</p>', unsafe_allow_html=True)

uploaded_files = st.file_uploader("📂 อัปโหลดไฟล์ KML", type="kml", accept_multiple_files=True)
output_folder = st.text_input("📁 โฟลเดอร์ปลายทาง (เว้นว่างเพื่อดาวน์โหลด)")

if uploaded_files:
    if st.button('⚡ เริ่มประมวลผล'):
        with st.status("⏳ กำลังประมวลผล...", expanded=True) as status:
            progress_bar = st.progress(0)
            total_files = len(uploaded_files)

            for i, uploaded_file in enumerate(uploaded_files):
                lines = load_kml_lines(uploaded_file)
                output_memory = save_to_excel_memory(lines)

                if output_folder:  # หากมีการกำหนดโฟลเดอร์ปลายทาง
                    if not os.path.exists(output_folder):
                        os.makedirs(output_folder)
                    
                    output_path = os.path.join(output_folder, uploaded_file.name.replace(".kml", ".xlsx"))
                    with open(output_path, "wb") as f:
                        f.write(output_memory.read())
                    
                    st.success(f"✅ ไฟล์บันทึกที่ {output_path}")
                else:
                    # ให้ดาวน์โหลดไฟล์จาก Streamlit โดยตรง
                    st.download_button(
                        label=f"📥 ดาวน์โหลด {uploaded_file.name.replace('.kml', '.xlsx')}",
                        data=output_memory,
                        file_name=f"{uploaded_file.name.replace('.kml', '.xlsx')}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                progress_bar.progress((i + 1) / total_files)

            status.update(label="✅ เสร็จสิ้น!", state="complete")

