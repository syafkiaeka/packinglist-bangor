# engine.py - UNIFIED PACKING LIST ENGINE (3 FORMAT SUPPORT)
import os
import math
import copy
from collections import OrderedDict, defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side


class PackingListEngine:
    def __init__(self, input_file, output_file, progress_callback=None, master_wb=None):
        self.input_file = input_file
        self.output_file = output_file
        self.progress = progress_callback or (lambda msg: None)
        self.master_wb = master_wb
        self.wb = None
        self.MasterData = {}
        self.DOGroups = OrderedDict()
        self.TransformResult = OrderedDict()
        self.FinalResult = {}
        self.format_type = None
        self.target_sheet = ""
        self.logs = []

    def log(self, msg):
        self.logs.append(msg)
        self.progress(msg)

    def Txt(self, v):
        return "" if v is None else str(v).strip().upper()

    def TxtRaw(self, v):
        return "" if v is None else str(v).strip()

    def Val(self, v):
        try:
            if v is None or v == "": return 0
            return float(v)
        except:
            return 0

    def SafeSheetName(self, name):
        if not name: return "Sheet"
        bad = ['\\', '/', '*', '?', ':', '[', ']']
        for c in bad: name = name.replace(c, " ")
        return name[:31].strip()

    def NormalizeHeader(self, v):
        text = self.Txt(v)
        for old, new in [(" ", ""), (".", ""), ("_", ""), ("-", ""), ("/", ""), ("\\", "")]:
            text = text.replace(old, new)
        return text

    def NormalizeUnit(self, unit):
        unit_upper = self.Txt(unit)
        if unit_upper in ["DOZ", "DOZEN"]:
            return "DUS"
        return unit_upper

    def load_master_data(self):
        self.log("   📚 Membaca Master Data...")
        master_sheet = next((s for s in self.wb.sheetnames if "Master Data" in s or "Master" in s), None)
        if not master_sheet:
            raise Exception(f"Sheet 'Master Data Packinglist' tidak ditemukan. Sheet yang ada: {self.wb.sheetnames}")
        
        ws = self.wb[master_sheet]
        for row in range(6, ws.max_row + 1):
            nama = self.Txt(ws[f"B{row}"].value)
            if not nama or nama in self.MasterData: continue
            self.MasterData[nama] = {
                "Nama": nama,
                "QtyMax": self.Val(ws[f"C{row}"].value),
                "Unit": self.Txt(ws[f"E{row}"].value),
                "Kategori": self.Txt(ws[f"F{row}"].value)
            }
        self.log(f"   ✅ Master Data: {len(self.MasterData)} item terbaca.")

    def detect_format(self):
        self.log("    Mendeteksi format Delivery Order...")
        sheets = self.wb.sheetnames
        
        # Cek Format 1: Rincian Pemindahan Barang (IT lama)
        rincian_sheet = next((s for s in sheets if "Rincian" in s and "Pemindahan" in s), None)
        # Cek Format 2: Delivery Order Detail (Inggris)
        do_sheet = next((s for s in sheets if "Delivery Order" in s or "DO Detail" in s), None)
        # Cek Format 3: Rincian Pengiriman Pesanan (Indonesia baru - ACCURATE)
        pengiriman_sheet = next((s for s in sheets if "Rincian" in s and "Pengiriman" in s), None)

        if rincian_sheet:
            self.format_type = "Rincian"
            self.target_sheet = rincian_sheet
            self.log(f"   📄 Format: Rincian Pemindahan Barang (IT Lama)")
        elif pengiriman_sheet:
            self.format_type = "Pengiriman"
            self.target_sheet = pengiriman_sheet
            self.log(f"    Format: Rincian Pengiriman Pesanan (Indonesia/ACCURATE)")
        elif do_sheet:
            self.format_type = "DeliveryOrder"
            self.target_sheet = do_sheet
            self.log(f"   📄 Format: Delivery Order Detail (Inggris)")
        else:
            raise Exception(f"Format DO tidak dikenali. Sheet yang tersedia: {sheets}")
        self.log(f"   📄 Sheet: '{self.target_sheet}'")

    def parse_rincian(self):
        self.log(f"    Parsing sheet: '{self.target_sheet}'...")
        ws = self.wb[self.target_sheet]
        header_keywords = {
            "DO": ["NOPEMINDAHAN", "NOPEMINDAHAN#", "NOMORPEMINDAHAN", "NOTRANSAKSI", "NOMORTRANSAKSI", "NO", "DO"],
            "TANGGAL": ["TANGGAL", "DATE", "TGL"],
            "GUDANG": ["GUDANGTUJUAN/DARI", "GUDANGTUJUAN", "GUDANGDARI", "GUDANG", "OUTLET"],
            "NAMA": ["NAMABARANG", "NAMAITEM", "ITEM", "BARANG"],
            "UNIT": ["SATUAN", "UNIT"],
            "QTY": ["KUANTITAS", "QTY", "QUANTITY", "JUMLAH"],
            "SHIPVIA": ["SHIPVIA", "SHIP", "VIA", "EKSPEDISI"]
        }
        header_row, header_map = None, {}
        for r in range(1, min(ws.max_row, 30) + 1):
            kandidat = {}
            for c in range(1, ws.max_column + 1):
                norm = self.NormalizeHeader(ws.cell(row=r, column=c).value)
                if not norm: continue
                for field, kws in header_keywords.items():
                    for kw in kws:
                        if norm == self.NormalizeHeader(kw):
                            kandidat[field] = c; break
            if "NAMA" in kandidat and "QTY" in kandidat:
                header_row, header_map = r, kandidat; break
                
        if not header_row:
            raise Exception(f"Header tidak ditemukan.")
        
        last_do, last_tgl, last_gudang, last_ship = None, None, None, None
        for r in range(header_row + 1, ws.max_row + 1):
            def cell(field):
                return ws.cell(row=r, column=header_map.get(field, 0)).value if field in header_map else None
            no_do = cell("DO")
            if no_do is not None:
                no_do = self.TxtRaw(no_do)
                if no_do: last_do = no_do
            else: no_do = last_do
            tgl = cell("TANGGAL")
            if tgl is not None: last_tgl = tgl
            else: tgl = last_tgl
            gudang = cell("GUDANG")
            if gudang is not None:
                g = self.TxtRaw(gudang)
                if g: last_gudang = g
            gudang = last_gudang
            ship = cell("SHIPVIA")
            if ship is not None:
                s = self.TxtRaw(ship)
                if s: last_ship = s
            ship = last_ship
            nama = cell("NAMA")
            if nama is None: continue
            nama = self.TxtRaw(nama)
            if not nama: continue
            qty = cell("QTY")
            try:
                if qty is None: continue
                qty = float(qty)
            except: continue
            if qty <= 0: continue
            unit = self.TxtRaw(cell("UNIT")) if "UNIT" in header_map else ""
            if not no_do: continue
            if no_do not in self.DOGroups:
                self.DOGroups[no_do] = {"Tanggal": tgl, "Outlet": gudang, "ShipVia": ship, "Items": []}
            self.DOGroups[no_do]["Items"].append({"Nama": nama, "Qty": qty, "Unit": unit})
        self.log(f"   ✅ DO terbaca: {len(self.DOGroups)}")

    def parse_pengiriman(self):
        """Parser untuk Rincian Pengiriman Pesanan (format ACCURATE Indonesia)"""
        self.log(f"    Parsing sheet: '{self.target_sheet}' (Format ACCURATE)...")
        ws = self.wb[self.target_sheet]
        
        # Cari semua baris yang mengandung "Nomor #" di kolom B
        do_start_rows = []
        for r in range(1, ws.max_row + 1):
            val_b = self.Txt(ws.cell(row=r, column=2).value)
            if val_b == "NOMOR#":
                do_start_rows.append(r)
        
        if not do_start_rows:
            raise Exception(f"Tidak ditemukan header 'Nomor #' di Kolom B.")
        
        self.log(f"   -> Ditemukan {len(do_start_rows)} blok DO.")
        
        for idx, start_row in enumerate(do_start_rows):
            # Tentukan end_row (baris sebelum DO berikutnya atau end of sheet)
            if idx + 1 < len(do_start_rows):
                end_row = do_start_rows[idx + 1]
            else:
                end_row = ws.max_row + 1
            
            # Baca Nomor DO (kolom E, setelah "Nomor # :")
            nomor_do = ""
            for c in range(3, 10):
                val = ws.cell(row=start_row, column=c).value
                if val:
                    val_str = self.TxtRaw(val)
                    if val_str and not val_str.startswith(":"):
                        nomor_do = val_str
                        break
            
            # Baca Pelanggan/Outlet (kolom M, setelah "Pelanggan :")
            outlet = ""
            for c in range(10, 20):
                val = ws.cell(row=start_row, column=c).value
                if val:
                    val_str = self.TxtRaw(val)
                    if val_str and not val_str.startswith(":"):
                        outlet = val_str
                        break
            
            # Baca Tanggal (baris setelah start_row, kolom E)
            tanggal = ""
            for r in range(start_row + 1, min(start_row + 5, end_row)):
                val_b = self.Txt(ws.cell(row=r, column=2).value)
                if val_b == "TANGGAL":
                    for c in range(3, 10):
                        val = ws.cell(row=r, column=c).value
                        if val:
                            tanggal = val
                            break
                    break
            
            # Baca Ship Via (jika ada)
            ship_via = ""
            
            if not nomor_do: continue
            
            self.DOGroups[nomor_do] = {"Outlet": outlet, "Tanggal": tanggal, "ShipVia": ship_via, "Items": []}
            
            # Cari header item (baris dengan "Kode #" di kolom C)
            item_header_row = None
            for r in range(start_row + 1, min(start_row + 10, end_row)):
                val_c = self.Txt(ws.cell(row=r, column=3).value)
                if val_c == "KODE#":
                    item_header_row = r
                    break
            
            if not item_header_row: continue
            
            # Baca item data
            for r in range(item_header_row + 1, end_row):
                # Skip jika baris kosong atau total
                val_l = ws.cell(row=r, column=12).value  # Kolom L (Kuantitas)
                if val_l is None: continue
                try:
                    qty = float(val_l)
                except:
                    continue
                if qty <= 0: continue
                
                nama = self.TxtRaw(ws.cell(row=r, column=6).value)  # Kolom F (Nama Barang)
                if not nama or nama.upper() == "NAMA BARANG": continue
                if "DUS" in self.Txt(nama): continue
                
                unit = self.TxtRaw(ws.cell(row=r, column=15).value)  # Kolom O (Satuan)
                unit_normalized = self.NormalizeUnit(unit)
                
                self.DOGroups[nomor_do]["Items"].append({
                    "Nama": nama, "Qty": qty, "Unit": unit_normalized
                })
        
        self.log(f"   ✅ DO terbaca: {len(self.DOGroups)}")

    def parse_delivery_order(self):
        self.log(f"    Parsing sheet: '{self.target_sheet}'...")
        ws = self.wb[self.target_sheet]
        header_rows = []
        for r in range(1, ws.max_row + 1):
            if self.Txt(ws.cell(row=r, column=3).value) == "NUMBER":
                header_rows.append(r)
        header_rows.append(ws.max_row + 1)
        if not header_rows or header_rows[0] > ws.max_row:
            raise Exception(f"Tidak ditemukan header 'NUMBER' di Kolom C.")
            
        for i in range(len(header_rows) - 1):
            start, end = header_rows[i], header_rows[i + 1]
            nomor_do = self.TxtRaw(ws.cell(row=start, column=6).value)
            outlet = self.TxtRaw(ws.cell(row=start, column=13).value)
            tanggal = ""
            if start + 1 < end:
                tanggal_val = ws.cell(row=start+1, column=6).value
                if tanggal_val: tanggal = tanggal_val
            ship_via = ""
            for col in range(14, 17):
                val = ws.cell(row=start, column=col).value
                if val:
                    ship_via = self.TxtRaw(val)
                    break
            if not nomor_do: continue
            self.DOGroups[nomor_do] = {"Outlet": outlet, "Tanggal": tanggal, "ShipVia": ship_via, "Items": []}
            for r in range(start, end):
                nama = self.TxtRaw(ws.cell(row=r, column=8).value)
                if "DUS" in self.Txt(nama) or not nama: continue
                if nama.upper() == "ITEM NAME": continue
                qty = self.Val(ws.cell(row=r, column=14).value)
                if qty <= 0: continue
                unit_raw = self.TxtRaw(ws.cell(row=r, column=16).value)
                unit_normalized = self.NormalizeUnit(unit_raw)
                self.DOGroups[nomor_do]["Items"].append({
                    "Nama": nama, "Qty": qty, "Unit": unit_normalized
                })
        self.log(f"   ✅ DO terbaca: {len(self.DOGroups)}")

    def get_category_and_rank(self, kategori_asli, is_receh):
        kat = self.Txt(kategori_asli)
        
        if kat in ["HAMPERS", "GRILL BOX", "INNER", "CRAZY BUCKET", "MINYAK", "PACKAGING"]:
            kat = "LAIN-LAIN"

        if kat == "DAGING":
            return ("Daging Receh", 2) if is_receh else ("Daging", 1)
        elif kat == "FROZEN":
            return ("Frozen Receh", 4) if is_receh else ("Frozen", 3)
        elif kat == "KEJU":
            return ("Keju Receh", 6) if is_receh else ("Keju", 5)
        elif kat == "KENTANG":
            return ("Kentang Receh", 8) if is_receh else ("Kentang", 7)
        elif kat == "SAUS":
            return ("Saus Receh", 10) if is_receh else ("Saus", 9)
        elif kat == "BUTTER":
            return ("Butter", 11)
        elif kat == "DRY":
            return ("Dry Receh", 13) if is_receh else ("Dry", 12)
        elif kat == "LAIN-LAIN":
            return ("Receh Lain-lain", 15) if is_receh else ("Lain-lain", 14)
        elif kat in ["BUN", "ROTI"]:
            return ("Bun", 16)
        else:
            return (f"{kat} Receh" if is_receh else kat, 99)

    def transform(self):
        self.log("   ⚙️ Transforming data + Breakdown Koli...")
        for do_no, data in self.DOGroups.items():
            hasil = []
            for item in data["Items"]:
                nama = self.TxtRaw(item["Nama"])
                key = self.Txt(nama)
                master = self.MasterData.get(key)
                if not master: continue
                if "DUS" in key: continue
                
                qty_max = master["QtyMax"]
                unit_master = master["Unit"]
                kategori_asli = master["Kategori"]
                unit_do = self.Txt(item["Unit"])
                qty_do = item["Qty"]
                if qty_max <= 0 or qty_do <= 0: continue
                nama_upper = key
                
                if unit_master == "DUS":
                    qty_real, qty_koli_max, unit_out = qty_do, qty_max, "DUS"
                elif unit_master == "PACK":
                    unit_out = "PACK"
                    if "DUS" in unit_do:
                        if "BEEF PATTY SMALL" in nama_upper or "BEEF PATTY LARGE" in nama_upper: isi_dus = 18
                        elif "THOUSAND ISLAND" in nama_upper: isi_dus = 20
                        else: isi_dus = qty_max
                        qty_real = qty_do * isi_dus
                    else: qty_real = qty_do
                    qty_koli_max = 20 if "THOUSAND ISLAND" in nama_upper else qty_max
                elif unit_master == "BKS":
                    unit_out = "BKS"
                    qty_real = qty_do * qty_max if "DUS" in unit_do else qty_do
                    qty_koli_max = qty_max
                else:
                    unit_out, qty_real, qty_koli_max = unit_master, qty_do, qty_max
                    
                if qty_koli_max <= 0: qty_koli_max = qty_max
                
                while qty_real > 0.0001:
                    qty_koli = min(qty_real, qty_koli_max)
                    is_receh = qty_koli < (qty_koli_max - 0.0001)
                    out_kat, rank = self.get_category_and_rank(kategori_asli, is_receh)

                    hasil.append({
                        "NamaBarang": nama, "Qty": qty_koli, "Unit": unit_out, 
                        "Kategori": out_kat, "Rank": rank
                    })
                    qty_real -= qty_koli
                    
            hasil.sort(key=lambda x: (x["Rank"], x["NamaBarang"].upper()))
            self.TransformResult[do_no] = {
                "Outlet": data["Outlet"], "Tanggal": data.get("Tanggal", ""),
                "ShipVia": data.get("ShipVia", ""), "Items": hasil
            }
        self.log(f"   ✅ Transform: {len(self.TransformResult)} DO")

    def group_receh(self):
        self.log("   📦 Finalizing (Tanpa Grouping MaxReceh)...")
        for do_no, do_data in self.TransformResult.items():
            self.FinalResult[do_no] = {
                "Outlet": do_data["Outlet"], 
                "Tanggal": do_data.get("Tanggal", ""), 
                "ShipVia": do_data.get("ShipVia", ""), 
                "Items": do_data["Items"]
            }
        self.log(f"   ✅ Final Result: {len(self.FinalResult)} DO")

    def export(self):
        self.log("   📝 Exporting Excel...")
        if self.master_wb is not None:
            out_wb = self.master_wb
        else:
            out_wb = Workbook()
            out_wb.remove(out_wb.active)

        if not self.FinalResult:
            self.log("   ⚠️ Tidak ada data DO yang berhasil diproses.")
            if self.master_wb is None:
                ws = out_wb.active; ws.title = "Info"
                ws["A1"] = "Tidak ada data Packing List yang bisa diproses."
                out_wb.save(self.output_file)
            return out_wb
            
        thin = Side(style="thin"); medium = Side(style="medium")
        border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
        border_medium = Border(left=medium, right=medium, top=medium, bottom=medium)
        
        for do_no, data in self.FinalResult.items():
            sheet_name = self.SafeSheetName(data["Outlet"])
            if not sheet_name: sheet_name = "DO"
            
            base_name = sheet_name; n = 2
            while sheet_name in out_wb.sheetnames:
                sheet_name = f"{base_name[:28]}_{n}"; n += 1
                
            ws = out_wb.create_sheet(title=sheet_name)
            ws.column_dimensions["A"].width = 10; ws.column_dimensions["B"].width = 38
            ws.column_dimensions["C"].width = 12; ws.column_dimensions["D"].width = 14; ws.column_dimensions["E"].width = 25
            
            ws.merge_cells("A2:E2"); ws["A2"] = "PACKING LIST"
            ws["A2"].font = Font(bold=True, size=18); ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[2].height = 30
            
            ws["A4"] = "DELIVERY\nORDER :"; ws["B4"] = do_no
            ws["D4"] = "Ship Via:"; ws["E4"] = data.get("ShipVia", "")
            ws["A5"] = "OUTLET :"; ws["B5"] = data["Outlet"]
            ws["A6"] = "DELIVERY\nDATE :"; ws["B6"] = data.get("Tanggal", "")
            
            for row in range(4, 7):
                ws.cell(row=row, column=1).font = Font(bold=True, size=12)
                ws.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                ws.cell(row=row, column=2).font = Font(size=12)
                ws.cell(row=row, column=2).alignment = Alignment(horizontal="left", vertical="center")
            if data.get("ShipVia"):
                ws.cell(row=4, column=4).font = Font(bold=True, size=12)
                ws.cell(row=4, column=5).font = Font(size=12)
            if data.get("Tanggal"):
                try: ws["B6"].number_format = "dd-mm-yyyy"
                except: pass
                
            ws.row_dimensions[4].height = 28; ws.row_dimensions[5].height = 20; ws.row_dimensions[6].height = 28
            
            hr = 9
            for col, h in enumerate(["NO KOLI", "NAMA BARANG", "QTY", "UNIT", "NOTES"], 1):
                c = ws.cell(row=hr, column=col, value=h)
                c.font = Font(bold=True, size=12); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = border_medium
            ws.row_dimensions[hr].height = 25
            
            r = 10
            for item in data["Items"]:
                ws.cell(row=r, column=2, value=item["NamaBarang"]).font = Font(size=12)
                ws.cell(row=r, column=3, value=item["Qty"]).font = Font(size=12)
                ws.cell(row=r, column=4, value=item["Unit"]).font = Font(size=12)
                for col in range(1, 6):
                    cell = ws.cell(row=r, column=col); cell.border = border_thin; cell.font = Font(size=12)
                ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=4).alignment = Alignment(horizontal="center")
                r += 1
                
            ws.page_setup.orientation = "portrait"; ws.page_setup.fitToWidth = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            
        if self.master_wb is None:
            out_wb.save(self.output_file)
        self.log(f"   ✅ Sheet berhasil ditambahkan.")
        return out_wb

    def run(self):
        self.logs = []
        self.wb = load_workbook(self.input_file, data_only=False)
        self.load_master_data()
        self.detect_format()
        if self.format_type == "Rincian": self.parse_rincian()
        elif self.format_type == "Pengiriman": self.parse_pengiriman()
        else: self.parse_delivery_order()
        self.transform()
        self.group_receh()
