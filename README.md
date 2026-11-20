# ใบเช็กลิสต์ยืมของ ห้องผ่าตัด

แอปเช็กเครื่องมือ/เวชภัณฑ์ที่ต้องยืม ใช้ได้หลายเครื่องพร้อมกัน ข้อมูลตรงกันทุกเครื่อง

| ส่วน | อยู่ที่ | หน้าที่ |
|---|---|---|
| หน้าแอป | GitHub Pages | ที่น้องๆ เปิดใช้ในมือถือ/คอม |
| ข้อมูล | Google Sheet + Apps Script | ศูนย์กลางข้อมูลชุดเดียว |
| โค้ด | GitHub (repo นี้) | แก้ที่นี่ที่เดียว ทุกอย่างอัปเดตเอง |
| สำรอง | OneDrive ส่วนตัว | ไฟล์ Excel ทุกวันจันทร์ 01:00 น. |

ตั้งค่าครั้งแรกใช้คอมพิวเตอร์ ประมาณ 30–45 นาที ทำครั้งเดียว

---

## ส่วนที่ 1 — Google Sheet (ศูนย์กลางข้อมูล)

1. sheets.google.com → สร้างชีตใหม่ ตั้งชื่อ `ยืมของ OR`
2. เมนู **ส่วนขยาย → Apps Script**
3. ลบโค้ดเดิม → วางเนื้อหา `apps-script/Code.gs` ทั้งไฟล์
4. แก้บรรทัด `const TEAM_KEY = '';` เป็นรหัสทีม เช่น `const TEAM_KEY = 'or2569';` (**ต้องตั้ง** ไม่งั้นสำรองข้อมูลไม่ได้) → บันทึก
5. เลือกฟังก์ชัน **setup** → **Run** → อนุญาต (Advanced → Go to … → Allow)
6. เลือกฟังก์ชัน **testExport** → **Run** → อนุญาตเพิ่ม → ดูบันทึกว่าขึ้น `OK`
7. **Deploy → New deployment → Web app** · Execute as: **Me** · Who has access: **Anyone** → Deploy
8. จดไว้ 3 อย่าง
   - **Web app URL** (ลงท้าย `/exec`)
   - **Deployment ID** (Deploy → Manage deployments → ตัวอักษรยาวใต้ชื่อ)
   - **Script ID** (ไอคอนเฟือง Project Settings → IDs)

## ส่วนที่ 2 — GitHub (หน้าแอป + อัปเดตอัตโนมัติ)

1. github.com → **New repository** ชื่อ `or-checklist` → **Public** → Create
2. **Add file → Upload files** → ลากไฟล์และโฟลเดอร์ทั้งหมดในชุดนี้ (`index.html`, `apps-script/`, `.github/`, `README.md`, `.gitignore`) → Commit
   - ถ้าโฟลเดอร์ `.github` มองไม่เห็นในเครื่อง ให้เปิดแสดงไฟล์ซ่อน หรือสร้างไฟล์ผ่านเว็บ: Add file → Create new file → พิมพ์ชื่อ `.github/workflows/deploy.yml` แล้ววางเนื้อหา (ทำแบบเดียวกันกับ `backup.yml`)
3. **Settings → Pages → Source: GitHub Actions**
4. **Settings → Secrets and variables → Actions → New repository secret** ใส่ทีละตัว:

| ชื่อ secret | ใส่อะไร | จำเป็นไหม |
|---|---|---|
| `API_URL` | Web app URL จากส่วนที่ 1 | จำเป็น |
| `TEAM_KEY` | รหัสทีม (ตัวเดียวกับใน Code.gs) | จำเป็น (สำรอง) |
| `RCLONE_CONF` | จากส่วนที่ 3 | จำเป็น (สำรอง) |
| `SCRIPT_ID` | Script ID | สำหรับอัปเดต Apps Script อัตโนมัติ |
| `DEPLOYMENT_ID` | Deployment ID | สำหรับอัปเดต Apps Script อัตโนมัติ |
| `CLASPRC_JSON` | จากส่วนที่ 4 | สำหรับอัปเดต Apps Script อัตโนมัติ |
| `GH_PAT` | จากส่วนที่ 5 | ไม่บังคับ (ต่ออายุ OneDrive เอง) |

5. แท็บ **Actions → อัปเดตแอปอัตโนมัติ → Run workflow** → รอขึ้นเครื่องหมายถูกสีเขียว
6. ลิงก์แอป: `https://<ชื่อผู้ใช้>.github.io/or-checklist/` (เชื่อมกับชีตให้เองแล้ว)

> repo เป็น Public ได้ เพราะไม่มีข้อมูลการยืมอยู่ในนี้ ข้อมูลอยู่ใน Google Sheet และ OneDrive ของพี่เท่านั้น อย่าใส่รหัสทีมลงในไฟล์ใน repo (ใส่ใน Secrets เท่านั้น)

## ส่วนที่ 3 — เชื่อม OneDrive ส่วนตัว (rclone)

1. ดาวน์โหลด rclone: rclone.org/downloads → Windows → แตกไฟล์
2. เปิด Command Prompt ในโฟลเดอร์นั้น พิมพ์ `rclone config` แล้วตอบ:
   - `n` (new remote) → name: **onedrive**
   - Storage: พิมพ์ **onedrive**
   - client_id / client_secret: กด Enter ข้าม
   - region: **global**
   - Edit advanced config: **n**
   - Use web browser: **y** → ล็อกอิน Microsoft ส่วนตัว → อนุญาต
   - เลือก **OneDrive Personal or Business** → เลือกไดรฟ์แรก → **y** → **q**
3. พิมพ์ `rclone config file` จะบอกที่อยู่ไฟล์ → เปิดด้วย Notepad → คัดลอก**ทั้งหมด** → ใส่เป็น secret `RCLONE_CONF`
4. ทดสอบ: Actions → **สำรองข้อมูลไป OneDrive → Run workflow** → เช็ก OneDrive โฟลเดอร์ `OR-Checklist-Backup`
   - เปลี่ยนชื่อโฟลเดอร์ได้ที่ Settings → Secrets and variables → Actions → แท็บ Variables → `ONEDRIVE_FOLDER`

## ส่วนที่ 4 — ให้ Apps Script อัปเดตเองจาก GitHub (clasp)

1. เปิด script.google.com/home/usersettings → เปิด **Google Apps Script API** = On
2. ติดตั้ง Node.js (nodejs.org) → เปิด Command Prompt พิมพ์
   `npx @google/clasp@2.4.2 login` → ล็อกอินบัญชี Google เดียวกับที่สร้างชีต
3. เปิดไฟล์ `C:\Users\<ชื่อ>\.clasprc.json` ด้วย Notepad → คัดลอกทั้งหมด → secret `CLASPRC_JSON`
4. ตั้งแต่นี้ แก้ `apps-script/Code.gs` ใน GitHub → Commit → Apps Script อัปเดตลิงก์เดิมเอง

> ถ้ายังไม่ทำส่วนนี้ หน้าแอปยังอัปเดตอัตโนมัติ แต่ถ้าแก้ Code.gs ต้องไปวางใน Apps Script เอง แล้ว Manage deployments → แก้ → New version

## ส่วนที่ 5 (ไม่บังคับ) — ต่ออายุการเชื่อม OneDrive อัตโนมัติ

การเชื่อม OneDrive จะหมดอายุถ้าไม่ต่อ (ประมาณ 3 เดือน) ตั้ง `GH_PAT` แล้วระบบจะต่อให้เองทุกสัปดาห์
1. GitHub → รูปโปรไฟล์ → Settings → Developer settings → **Fine-grained tokens** → Generate
2. Repository access: เฉพาะ `or-checklist` · Permissions → **Secrets: Read and write** · อายุ 1 ปี
3. คัดลอก token → secret `GH_PAT`
4. ถ้าไม่ตั้ง: เมื่อการสำรองขึ้นสีแดง ให้ทำส่วนที่ 3 ใหม่ แล้วอัปเดต `RCLONE_CONF`

---

## การใช้งานหลังติดตั้ง

- **แจกน้องๆ:** ในแอป แท็บ แพทย์/ชื่อ → คัดลอกลิงก์ให้น้องๆ (มีรหัสทีมในลิงก์ ส่งเฉพาะในกลุ่มทีม)
- **แก้แอป / อัปเดตอัตโนมัติ:** แก้ไฟล์ใน GitHub (กดรูปดินสอ ✏️ หรือ Add file → Upload files ทับ) → Commit changes → Actions อัปเดต Apps Script + หน้าแอปเอง ภายใน 2–3 นาที → มือถือทุกเครื่องโหลดแอปใหม่เองเมื่อเปิดแอปหรือภายใน 5 นาที (ยกเว้นเครื่องที่กำลังกรอกใบค้าง จะอัปเดตหลังกรอกเสร็จ หรือกดปุ่ม "อัปเดตแอป")
- **ดูข้อมูล:** เปิด Google Sheet ได้ตลอด แท็บ "บันทึก" คอลัมน์ `borrowIns` / `borrowSup` สรุปของที่ยืม
- **ดูการสำรอง:** แท็บ Actions ใน GitHub (สีเขียว = สำเร็จ) และในโฟลเดอร์ OneDrive
- ไม่ควรแก้คอลัมน์ `ins` / `sup` ในชีตด้วยมือ
- ใบนี้ไม่เก็บ HN/ข้อมูลผู้ป่วย แต่ควรแจ้งหัวหน้า/IT ว่าใช้ Google Sheets และ OneDrive ส่วนตัวเก็บข้อมูลการยืมของ
