# Import Shipment Report

เว็บแอป Streamlit สำหรับอ่านไฟล์ Excel รายงาน Import Shipment แล้วแสดงผลเป็น KPI / กราฟ / ตาราง พร้อมกรองข้อมูลและ export ออกเป็น Excel

---

## สิ่งที่ต้องมีก่อนเริ่ม

| รายการ | รายละเอียด |
| --- | --- |
| Python | เวอร์ชัน 3.9 ขึ้นไป (แนะนำ 3.11+) |
| ไฟล์ข้อมูล | ไฟล์ Excel นามสกุล `.xlsx` ของรายงาน Import Shipment |
| ระบบปฏิบัติการ | Windows (คำสั่งด้านล่างเป็น PowerShell) |

ตรวจสอบว่าติดตั้ง Python แล้วหรือยัง:

```powershell
python --version
```

ถ้าไม่ขึ้นเวอร์ชัน ให้ติดตั้ง Python จาก https://www.python.org/downloads/ แล้วติ๊ก **Add Python to PATH** ตอนติดตั้ง

---

## ขั้นตอนการรัน

### Step 1 — เปิด PowerShell แล้วเข้าไปที่โฟลเดอร์โปรเจกต์

```powershell
cd C:\...\Test
```

### Step 2 — สร้าง virtual environment (ทำครั้งแรกครั้งเดียว)

```powershell
python -m venv .venv
```

จะได้โฟลเดอร์ `.venv` ขึ้นมาในโปรเจกต์ ถ้ามีอยู่แล้วข้ามขั้นตอนนี้ได้เลย

### Step 3 — เปิดใช้งาน virtual environment

```powershell
.\.venv\Scripts\Activate.ps1
```

เมื่อสำเร็จจะเห็น `(.venv)` นำหน้าบรรทัดคำสั่ง

> **ถ้าเจอ error เรื่อง execution policy** ให้รันคำสั่งนี้ก่อนแล้วลอง activate ใหม่:
>
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

### Step 4 — ติดตั้ง dependencies

```powershell
pip install -r requirements.txt
```

ไลบรารีที่จะถูกติดตั้ง:

- `streamlit` — ตัวเว็บแอป
- `pandas` — อ่านและประมวลผลข้อมูล
- `openpyxl` — engine สำหรับอ่าน/เขียนไฟล์ `.xlsx`

### Step 5 — รันแอป

```powershell
streamlit run Webapp.py
```

เบราว์เซอร์จะเปิดขึ้นมาเองที่ `http://localhost:8501` ถ้าไม่เปิดให้ copy URL จากเทอร์มินัลไปวางในเบราว์เซอร์

### Step 6 — อัปโหลดไฟล์ Excel

กดปุ่ม **Import Excel** บนหน้าเว็บ แล้วเลือกไฟล์ `.xlsx` ที่ต้องการ (แอปจะยังไม่แสดงผลอะไรจนกว่าจะอัปโหลดไฟล์)

### Step 7 — ใช้งานฟิลเตอร์ที่แถบซ้าย

- **ใช้วันที่จาก** — เลือกคอลัมน์วันที่ที่จะใช้กรอง (ETA / ATA / DORelease / DeliveryDate)
- **ช่วงวันที่** — เลือกวันเริ่มต้นและวันสิ้นสุด
- **Job ID / Invoice No / Container No** — พิมพ์คำค้น (ค้นแบบบางส่วน ไม่สนตัวพิมพ์เล็กใหญ่)
- **Origin / Release Port / Carrier / Container Size** — เลือกได้หลายค่า
- **Declaration** — All / Declared (409) / Pending
- **เฉพาะที่มีปัญหาเอกสาร** — ติ๊กเพื่อดูเฉพาะรายการที่มี issue

### Step 8 — ดูผลลัพธ์

- **KPI** — จำนวน Jobs, Invoices, Containers, CBM, รายการที่ยังไม่ declare, ค่าเฉลี่ย ATA→Delivery (วัน)
- **กราฟ** — Containers แยกตาม Origin / Carrier / รายวัน / ขนาดตู้
- **ตาราง Data** — ข้อมูลที่ผ่านฟิลเตอร์ทั้งหมด

### Step 9 — Export รายงาน

กดปุ่ม **Export Report (Excel)** ด้านล่างสุด จะได้ไฟล์ `export_report.xlsx` ที่มี 4 sheet:

| Sheet | เนื้อหา |
| --- | --- |
| `Data` | ข้อมูลทั้งหมดที่ผ่านฟิลเตอร์ |
| `By Origin` | สรุปจำนวน Invoice / CBM / Cartons แยกตาม Origin |
| `By Carrier` | จำนวนตู้แยกตาม Carrier |
| `Pending` | รายการที่ยังไม่ declare |

### Step 10 — ปิดแอป

กลับไปที่เทอร์มินัลแล้วกด `Ctrl + C`

ออกจาก virtual environment ด้วย:

```powershell
deactivate
```

---

## การรันครั้งถัดไป

ครั้งต่อไปไม่ต้องทำ Step 2 และ 4 ซ้ำ ใช้แค่:

```powershell
cd C:\...\Test
.\.venv\Scripts\Activate.ps1
streamlit run Webapp.py
```

---

## รูปแบบไฟล์ Excel ที่รองรับ

แอปจะจับคู่ชื่อคอลัมน์แบบ **contains** (ไม่สนตัวพิมพ์เล็กใหญ่) ชื่อหัวตารางจึงไม่ต้องตรงเป๊ะ แต่ต้องมีคำเหล่านี้อยู่:

`job`, `eta`, `ata`, `invoice`, `origin`, `release port`, `carrier`, `ctnr no. count`, `ctnr no.`, `seal`, `container size`, `cartons`, `cbm`, `filenamedate`, `do release`, `shipping document receive`, `issue reason`, `shipping document complete`, `status declaration`, `(0209)`, `delivery plan`, `(0309)`, `delivery date`, `delivery place`

ข้อกำหนดเพิ่มเติม:

- แถวที่ใช้งานได้ต้องมี `Job ID` ขึ้นต้นด้วยตัว **I** ตามด้วยตัวเลข เช่น `I2600123` (แถวอื่นจะถูกตัดทิ้งอัตโนมัติ)
- คอลัมน์วันที่รองรับรูปแบบ `MM/DD/YYYY` หรือค่า date ของ Excel
- `FileNameDate` รองรับรูปแบบ `YYYYMMDD`
- `Status Declaration` ที่เป็น `409` จะถือว่า **Declared**

---

## แก้ปัญหาที่พบบ่อย

| อาการ | วิธีแก้ |
| --- | --- |
| `streamlit : The term 'streamlit' is not recognized` | ยังไม่ได้ activate venv หรือยังไม่ได้ `pip install -r requirements.txt` |
| `Activate.ps1 cannot be loaded ... execution policy` | รัน `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser` |
| `KeyError: 'JobID'` | ไฟล์ Excel ไม่มีคอลัมน์ที่มีคำว่า `job` ในหัวตาราง ตรวจสอบว่าแถวหัวตารางอยู่บรรทัดแรก |
| `คอลัมน์ ... ไม่มีวันที่ที่อ่านได้` | รูปแบบวันที่ในไฟล์ไม่ตรงกับที่รองรับ ให้ลองเปลี่ยนคอลัมน์วันที่ หรือจัดรูปแบบใน Excel ใหม่ |
| ตารางว่างเปล่า | ฟิลเตอร์แคบเกินไป ลองล้างค่าฟิลเตอร์ที่แถบซ้าย |
| แก้โค้ดแล้วผลไม่เปลี่ยน | ผลลัพธ์ถูก cache ไว้ กด **C** บนหน้าเว็บ หรือเมนูขวาบน > **Clear cache** แล้ว rerun |
| พอร์ต 8501 ถูกใช้งานอยู่ | รันด้วยพอร์ตอื่น: `streamlit run Webapp.py --server.port 8502` |
