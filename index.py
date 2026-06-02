import json
from urllib.request import Request, urlopen
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import folium

# Configuration files
API_URL = "https://supercharge.info/service/supercharge/allSites"
EXCEL_FILE = "uk_tesla_superchargers.xlsx"
MAP_FILE = "uk_superchargers_map.html"

def fetch_superchargers():
    print("Fetching live network data from Supercharge.info...")
    try:
        # A basic User-Agent ensures the request isn't blocked by basic API rate filters
        req = Request(API_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(req) as response:
            return json.loads(response.read().decode())
    except Exception as e:
        print(f"Error fetching data from API: {e}")
        return []

def main():
    sites = fetch_superchargers()
    if not sites:
        print("No live dataset was retrieved. Exiting.")
        return

    uk_stations = []
    
    # 1. Broad exclusion list to heavily screen and discard all co-branded/partner third-party locations
    exclusion_keywords = [
        "EVOTM", "EV OTM", "EVPOINT", "EV POINT", 
        "EG ON THE MOVE", "EG GROUP", "EG-GROUP", 
        "BP PULSE", "IONITY", "PARTNER"
    ]

    for site in sites:
        address = site.get("address", {})

        # Filter: Keep only sites within the United Kingdom
        if address.get("country") == "United Kingdom":
            
            # 2. STRICT FILTER: Discard anything that is not fully live/operational
            status = site.get("status", "").upper()
            if status != "OPEN" and status != "OPEN - LIMITED HOURS":
                continue
                
            # 3. STRICT FILTER: Strip out partner stations or white-labeled V4 hardware locations
            name = site.get("name", "")
            if any(kw in name.upper() for kw in exclusion_keywords):
                continue
                
            stalls = site.get("stallCount", 0)
            gps = site.get("gps", {})
            latitude = gps.get("latitude")
            longitude = gps.get("longitude")

            # Missing GPS data safeguards against broken map coordinates or empty links
            if not latitude or not longitude:
                continue

            street = address.get("street", "")
            city = address.get("city", "")
            postcode = address.get("zip", "")
            
            # Construct exact coordinate fallback URL to use in links
            google_maps_link = f"https://www.google.com/maps/search/?api=1&query={latitude},{longitude}"

            uk_stations.append({
                "Station Name": name,
                "Status": status,
                "Stalls": stalls,
                "Street": street,
                "City": city,
                "Postcode": postcode,
                "Latitude": latitude,
                "Longitude": longitude,
                "Google Maps Link": google_maps_link
            })

    # Sort alphabetized neatly by Station Name
    uk_stations.sort(key=lambda x: x["Station Name"])

    if not uk_stations:
        print("No stations found matching your strict criteria.")
        return

    # =========================================================================
    # PART 1: GENERATE STYLED EXCEL FILE (With Executable Formulas)
    # =========================================================================
    print(f"Generating formatted spreadsheet with {len(uk_stations)} sites...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UK Superchargers"
    
    # Force grid lines to render visibly alongside filled rows
    ws.views.sheetView[0].showGridLines = True
    
    # Design Formatting Configurations
    font_name = "Segoe UI"
    header_font = Font(name=font_name, size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # Corporate Steel Navy
    body_font = Font(name=font_name, size=10)
    link_font = Font(name=font_name, size=10, color="0563C1", underline="single") # Hyperlink Royal Blue
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    alt_row_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid") # Zebra light gray
    
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    headers = ["Station Name", "Status", "Stalls", "Street", "City", "Postcode", "Latitude", "Longitude", "Google Maps Link"]
    ws.append(headers)
    
    # Apply styling directly onto header block
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center if header in ["Status", "Stalls", "Postcode"] else align_left
        cell.border = thin_border
        
    # Inject Clean Rows 
    for row_idx, station in enumerate(uk_stations, start=2):
        ws.cell(row=row_idx, column=1, value=station["Station Name"]).alignment = align_left
        ws.cell(row=row_idx, column=2, value=station["Status"]).alignment = align_center
        ws.cell(row=row_idx, column=3, value=station["Stalls"]).alignment = align_right
        ws.cell(row=row_idx, column=4, value=station["Street"]).alignment = align_left
        ws.cell(row=row_idx, column=5, value=station["City"]).alignment = align_left
        ws.cell(row=row_idx, column=6, value=station["Postcode"]).alignment = align_center
        ws.cell(row=row_idx, column=7, value=station["Latitude"]).alignment = align_right
        ws.cell(row=row_idx, column=8, value=station["Longitude"]).alignment = align_right
        
        # NATIVE HYPERLINK FORMULA: Forces Excel to interpret links safely as interactive strings
        link_cell = ws.cell(row=row_idx, column=9)
        link_cell.value = f'=HYPERLINK("{station["Google Maps Link"]}", "Open in Maps")'
        link_cell.font = link_font
        link_cell.alignment = align_center
        
        # Clean down structural borders and set alternating row paints
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=row_idx, column=col_idx)
            if col_idx != 9: 
                c.font = body_font
            c.border = thin_border
            if row_idx % 2 == 0: 
                c.fill = alt_row_fill

    # Set cell spacing variables for row profiles
    ws.row_dimensions[1].height = 26
    for r in range(2, len(uk_stations) + 2):
        ws.row_dimensions[r].height = 20

    # Auto-adjust column widths cleanly based on data contents
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if val_str.startswith("=HYPERLINK"): 
                val_str = "Open in Maps" # Fallback metric calculation length
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 12)
        
    # Append native data filters and freeze headers from viewport scrolling
    ws.auto_filter.ref = f"A1:I{len(uk_stations) + 1}"
    ws.freeze_panes = "A2"
    
    try:
        wb.save(EXCEL_FILE)
        print(f"-> Excel Sheet saved successfully: {EXCEL_FILE}")
    except Exception as e:
        print(f"Error compiling Excel output: {e}")

    # =========================================================================
    # PART 2: GENERATE INTERACTIVE HTML MAP (With Locked UK Viewport Bounds)
    # =========================================================================
    print("Generating restricted boundary HTML map layout...")
    
    # Strict boundary box isolating the UK [South-West, North-East]
    uk_bounds = [[49.8, -10.5], [60.9, 2.0]] 
    
    # Cartodb Positron avoids the strict access blockers imposed by OSM tile servers on local systems
    uk_map = folium.Map(
        location=[54.3, -2.5], 
        zoom_start=6, 
        tiles="Cartodb Positron",
        max_bounds=True,
        min_lat=uk_bounds[0][0], max_lat=uk_bounds[1][0],
        min_lon=uk_bounds[0][1], max_lon=uk_bounds[1][1]
    )
    
    # Automatically fit map to the target box parameters
    uk_map.fit_bounds(uk_bounds)
    
    for station in uk_stations:
        # Custom elegant HTML card architecture inside individual map nodes
        popup_html = f"""
        <div style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; width: 220px; line-height: 1.4;">
            <strong style="font-size: 14px; color: #1F497D;">{station['Station Name']}</strong><br/>
            <span style="color: #2e7d32; font-weight: bold;">● Active Tesla Network</span><br/>
            <b>Stalls available:</b> {station['Stalls']}<br/>
            <b>Location:</b> {station['City']}, {station['Postcode']}<br/><br/>
            <a href="{station['Google Maps Link']}" target="_blank" 
               style="display: inline-block; padding: 6px 10px; background-color: #0563C1; color: white; 
                      text-decoration: none; border-radius: 4px; font-weight: bold; text-align: center; width: 90%;">
               Navigate in Maps
            </a>
        </div>
        """
        
        # Add custom red lightning flash pins over target data coordinate nodes
        folium.Marker(
            location=[station["Latitude"], station["Longitude"]],
            popup=folium.Popup(popup_html, max_width=250),
            tooltip=station["Station Name"],
            icon=folium.Icon(color="red", icon="flash", prefix="fa")
        ).add_to(uk_map)

    try:
        uk_map.save(MAP_FILE)
        print(f"-> Interactive Map saved successfully: {MAP_FILE}")
        print("\nAll tasks completed flawlessly! Open the generated files to see the filtered tracking results.")
    except Exception as e:
        print(f"Error compiling HTML map output: {e}")

if __name__ == "__main__":
    main()