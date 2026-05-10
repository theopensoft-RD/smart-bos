# SKILL.md — Smart Plant 1 Comply Spec Workflow

## โครงการ
**โครงการเพิ่มศักยภาพการบริหารการจัดการและควบคุมการบำบัดน้ำเสียด้วยโครงข่ายอัจฉริยะ ระยะที่ 1 เมืองพัทยา**

เป้าหมาย: จัดทำเอกสารประกวดราคา (TOR / BOQ / Comply Spec) สำหรับระบบควบคุมและตรวจวัดโรงบำบัดน้ำเสีย

---

## โครงสร้างโฟลเดอร์

```
co-work/
├── TOR/          # Terms of Reference (.docx + .pdf) — แหล่งข้อมูลต้นทาง
│   └── TOR โครงการเพิ่มประสิทธภาพโรงบำบัดน้ำเสีย Smart Plant P1  09-04-69.{docx,pdf}
├── BOQ/          # Bill of Quantities (มาจาก TOR เช่นกัน)
│   └── 5.ปริมาณงาน Smart Plant P1  23-04-2569.xlsx
├── template/     # ไฟล์ต้นแบบ (reference สำหรับ Comply spec + catalog PDF สำรอง)
│   ├── Comply spec Sensor P2 Trio Smart 06-03-69.xlsx
│   └── 6.2.5-2 ตู้ควบคุมสถานีวัดอัตราการไหลของน้ำอัจฉริยะ.pdf
├── catalog/      # datasheet อุปกรณ์ต่างๆ ต้นฉบับ (vendor source PDFs)
├── output/       # ไฟล์ที่กำลังทำ
│   ├── Comply spec Smart Plant 1.xlsx   ← ไฟล์หลัก
│   ├── Comply spec Smart Plant 1.pdf    ← ไฟล์ output PDF
│   ├── _archive/                         # backup + เวอร์ชันเก่า + รายงาน verification
│   ├── 5.1.1. .../...                   # section folders (work in progress)
│   ├── 5.1.2. .../5.1.2.-1/...pdf      # catalog PDFs แต่ละ section
│   ├── 5.1.3. .../5.1.3.1. .../...
│   ├── 5.1.4. .../5.1.4.1. .../...
│   ├── 5.1.5 .../5.1.5.1. .../...
│   ├── 5.1.6 .../5.1.6.1. .../...
│   └── 5.2 .../5.2.X.-N/...
├── scripts/      # Python utilities
│   ├── pdf_header.py     # เครื่องมือเพิ่ม header/footer PDF
│   ├── fix_uv_headers.py # script แก้ header UV cabinet -1
│   └── version.py        # snapshot-based version control
├── _versions/    # snapshot storage (created by version.py)
│   └── snapshots/<YYYY-MM-DD_HHMMSS_tag>/
│       ├── manifest.json
│       ├── Comply spec Smart Plant 1.xlsx
│       ├── SKILL.md
│       └── output.tar.gz  # full snap only
└── knowledge_base/  # KB for agent development
    ├── KB.md          # human-readable knowledge base
    ├── catalogs.json  # catalog inventory (vendor, model, sections)
    ├── sections.json  # section status + assignments
    ├── rect_coords.json # reusable rect coord templates
    ├── pipelines.md   # workflow recipes
    └── pitfalls.md    # lessons learned + common issues
```

**หมายเหตุ**: ชื่อไฟล์ TOR และ BOQ มี **2 spaces** ระหว่าง "P1" กับวันที่ — ใช้ exact path

---

## โครงสร้าง Comply spec Smart Plant 1.xlsx

| Column | หัวคอลัมน์ | คำอธิบาย | แหล่งที่มา |
|--------|-----------|----------|-----------|
| A | หัวข้อ | รหัส section (5, 5.1, 5.2 ...) | TOR |
| B | คุณลักษณะที่ต้องการ | ข้อกำหนด TOR (มีคำเปรียบเทียบ) | **คัดลอกจาก TOR ตรงตัว — รวม typo** |
| C | คุณลักษณะที่เสนอ | คุณสมบัติที่ผู้เสนอราคายืนยัน (ค่าจริงจาก catalog) | catalog datasheet |
| D | เอกสารอ้างอิง | ชี้ไปยัง catalog PDF ที่รองรับข้อกำหนด | catalog ref / ยี่ห้อ-รุ่น / ยินดีปฏิบัติ |
| E | เสนอโดย | ผู้เสนอราคา (vendor) — ใส่ตามผู้รับผิดชอบจริงในแต่ละกลุ่มอุปกรณ์ | manual |

### กฎสำคัญสำหรับ Col B / C / D

- **Col B preserved จาก TOR ตรงตัว** — แม้ TOR มี typo (เช่น "Elecptro", "ภายนอกอาคาร ทั่วไป" ที่มีช่องว่างเกิน) **ห้ามแก้** เพราะถือเป็นข้อความ TOR ต้นฉบับ
- **Col C** = ตัดคำเปรียบเทียบจาก B + ใช้ค่าจริงจาก catalog (ถ้า catalog ดีกว่า spec → ใช้ค่า catalog) + ซ่อม typo ที่ B มี
- **Col D** มี 6 รูปแบบ:
  1. **catalog ref (เทียบเท่า)**: `เทียบเท่าข้อกำหนด เอกสาร {section}-{N} {ชื่อตู้} หน้า {P} ข้อ {section} ข้อ X) ข้อย่อย N.` (สำหรับ ข้อย่อย ทั่วไป — catalog spec ตรงกับ TOR)
  2. **catalog ref (สูงกว่า)**: `สูงกว่าข้อกำหนด เอกสาร {section}-{N} {ชื่อตู้} หน้า {P} ข้อ {section} ข้อ X) ข้อย่อย N.` (เมื่อ catalog spec **ดีกว่า/มากกว่า** TOR เช่น TOR ขอ 5mm ได้ 3.91mm, TOR ขอ 50,000 hr ได้ 100,000 hr)
  3. **ยี่ห้อ/รุ่น**: `ยี่ห้อ {brand} รุ่น {model}` (สำหรับ parent ข้อ ที่ระบุรุ่นชัดเจน)
  4. **dash brand**: `ยี่ห้อ - รุ่น {model}` (สำหรับงาน fabricate ที่ไม่มี brand แต่มีรุ่น เช่น Vibration sensor)
  5. **commitment**: `ยินดีปฏิบัติตามข้อกำหนด` (Software, งานติดตั้ง, commitment statements)
  6. **empty**: section/sub-section header ที่ไม่ระบุ — เว้นว่าง

**เพิ่มเติม** สำหรับ ข้อย่อย ที่อ้างอิง separate catalog (single-row) — ใช้ filename format:
  - `5.X.Y.Z {Col B description minus จำนวน} {model}` (ดู กฎข้อ 8)
  - หรือ shortform: `5.X.Y.Z-N keyword model` / `5.X.Y.Z. name`
  - หรือ model name อย่างเดียว (สำหรับ ข้อย่อย under multi-row parent ที่ใช้ separate catalog)

### รูปแบบ Column C: คำเปรียบเทียบที่ต้องตัดจาก B

| คำใน B | C ทำอย่างไร |
|--------|------------|
| `หรือดีกว่า` | ลบออก |
| `ไม่น้อยกว่า` | ลบออก (เหลือแค่ค่า) |
| `ไม่น้อยไปกว่า` | ลบออก |
| `ไม่มากกว่า` | ลบออก |
| `ต้องสามารถ` | เปลี่ยนเป็น `สามารถ` |
| `จะต้อง` | ลบออก |

### Typo จาก TOR ที่ต้องแก้ใน Col C เท่านั้น

- "Elec**p**tro" → "Electro" (R140 BOD UV cabinet)
- "ภายนอกอาคาร ทั่วไป" (ช่องว่างเกิน) → "ภายนอกอาคารทั่วไป" (R211 LED)
- "จะต้อง" → "" (R129/R167/R239 เสา ข้อย่อย 3)

### Vendor Distribution (Col E "เสนอโดย")

โครงการนี้มี vendors 3 ราย แบ่งตามกลุ่มอุปกรณ์:

| Vendor | ขอบเขต | ตัวอย่างอุปกรณ์ |
|---|---|---|
| **SMART** | ตู้ + sensor + cable + civil work + ส่วนใหญ่ของ section 5.1.2-5.1.6 | UV cabinet (LINK), RCBO (Schneider), Controller (Cytron IRIV), Vibration (VTall), BOD/DO/LED sensors, Fiber optic patch panel (UF-2010A/UF-4112A), Twisted Pair (US-9106LSZH), VCT (Thai union), EMT (union emt 3/4"), Outdoor cabinet 5.2.5.2/5.2.6.2 (LINK UV-9012H-IP55), งานติดตั้ง+งานขุดฝัง |
| **TRIO** | Network/IT equipment ทั้งหมด | Server, PC, NGFW, NAS, Tablet, UPS, Core Switch, L2/L3 Switch, PoE Switch, NVR, CCTV camera, Access Point |
| **SR** | ท่อ HDPE เท่านั้น | ท่อ HDPE 32 mm. PE100 PN10 (5 entries ใน 5.2.x) |

**Pattern การใส่ Col E:**
- Section header rows (Col D ว่าง) → ใส่ vendor ที่รับผิดชอบทั้ง section
- Item rows (Col D มี ยี่ห้อ/รุ่น/catalog ref/ยินดีปฏิบัติ) → ใส่ vendor ที่เสนอ item นั้น
- Sub-item rows (ข้อย่อย ของ ข้อ X)) → vendor เดียวกับ parent

### Col D ของ Sub-item ที่ใช้ Brand เดียวกับ Parent

เมื่อ section header (parent ข้อ) มี ยี่ห้อ X รุ่น Y ใน Col D, sub-item (ข้อย่อย N) สามารถใช้รูปแบบย่อใน Col D:

| Type | Format | ตัวอย่าง |
|---|---|---|
| **Parent row** | `ยี่ห้อ X รุ่น Y` (ระบุยี่ห้อ + รุ่นเต็ม) | `ยี่ห้อ 19" German รุ่น G3-61142` |
| **Sub-item row** | ระบุเพียง model code (brand implied) | `G7-00012`, `G7-05002` |
| **Sub-item อื่น** | ระบุเพียง model spec | `UF-2010A`, `UFC9312A`, `union emt 3/4"`, `US-9106LSZH` |

**ใช้กับ:** ตู้ Rack ที่มีพัดลม/ปลั๊กไฟภายใน, ตู้ outdoor + ส่วนประกอบ, สาย+อุปกรณ์ที่อยู่ในกลุ่มเดียวกัน

---

## Comply spec — แผนผัง Section 5.1

```
5.1   ระบบบริหารจัดการน้ำเสียอัจฉริยะ (row 6)
│
├── 5.1.1  ระบบควบคุมการบำบัดน้ำเสียอัจฉริยะ (rows 7–60)        ← ❌ ยังไม่ได้ทำ
│   ├── 5.1.1.1  ตู้ Rack 42U
│   ├── 5.1.1.2  เครื่องแม่ข่าย
│   ├── 5.1.1.3  Next Generation Firewall
│   ├── 5.1.1.4  L2 Switch 24 ช่อง
│   ├── 5.1.1.5  NAS
│   ├── 5.1.1.6  Tablet
│   └── 5.1.1.7  UPS 10kVA
│
├── 5.1.2  ตู้ควบคุมเครื่องจักรภายในโรงบำบัด (rows 61–95)        ← ✅ เสร็จ
│   ├── ข้อ 1) ตู้ → 5.1.2-1 (LINK / UV-9012H-SUS Stainless)
│   ├── ข้อ 2) RCBO → 5.1.2-2 (Schneider Electric / QO116C06RCBO30)
│   ├── ข้อ 3) Micro Controller → 5.1.2-3 (Cytron Technologies / IRIV PiControl CM5)
│   ├── ข้อ 4) Vibration sensor → 5.1.2-4 (- / VTall-S203L-2)
│   └── ข้อ 5) งานติดตั้ง → ยินดีปฏิบัติฯ
│
├── 5.1.3  ตู้ควบคุมเครื่องสูบน้ำเสียภายนอก (rows 96–137)        ← ✅ เสร็จ
│   ├── 5.1.3.1  ตู้ควบคุม → 5.1.3.1-1~4 (เหมือน 5.1.2)
│   ├── 5.1.3.2  เสา (rows 126–130) → 5.1.3.2 (P1+P2 ฐานราก)
│   └── 5.1.3.3  งานติดตั้ง (rows 131–137) → ยินดีปฏิบัติฯ
│
├── 5.1.4  ตู้ BOD (rows 138–175)                              ← ✅ เสร็จ
│   ├── 5.1.4.1  ตู้ประมวลผล BOD → 5.1.4.1-1~3
│   ├── 5.1.4.2  BOD Sensor (rows 159–163) → Proteus Instruments / Water Quality Probe
│   ├── 5.1.4.3  เสา (rows 164–168) → ดู เสา convention
│   └── 5.1.4.4  งานติดตั้ง (rows 169–175) → ยินดีปฏิบัติฯ
│
├── 5.1.5  ตู้ DO (rows 176–208)                               ← ✅ เสร็จ
│   ├── 5.1.5.1  ตู้ประมวลผล DO → 5.1.5.1-1~3
│   ├── 5.1.5.2  DO Sensor (rows 197–202) → JIANT / JG-LDO-N01
│   └── 5.1.5.3  งานติดตั้ง (rows 203–208) → ยินดีปฏิบัติฯ
│
├── 5.1.6  ตู้ควบคุมแสดงผล (rows 209–247)                       ← ✅ เสร็จ
│   ├── 5.1.6.1  ตู้ควบคุม → 5.1.6.1-1~3
│   ├── 5.1.6.2  LED Display (rows 230–235) → Fahchy / LED Display P3.91 SMD OUTDOOR
│   ├── 5.1.6.3  เสา (rows 236–240)
│   └── 5.1.6.4  งานติดตั้ง (rows 241–247) → ยินดีปฏิบัติฯ
│
├── 5.1.7  ระบบ Software SCADA (rows 248–273)                  ← ❌ ยังไม่ได้ทำ
└── 5.1.8  งานเดินสาย (rows 274–277)                            ← ❌ ยังไม่ได้ทำ

5.2   งานระบบเครือข่ายคอมพิวเตอร์ (rows 278–663)                ← ❌ ยังไม่ได้ทำ
```

---

## Catalog PDF — Annotation Standard

### Header (ทุกหน้า)
- **สี**: แดง (`1 0 0 rg`, `color:#FF0000`)
- **font**: HeBo (Helvetica-Bold) ขนาด 14-18pt — เลือกใหญ่สุดที่ fit ในหน้า ถ้ายาวเกินให้ wrap 2 บรรทัด
- **alignment**: center (`Q=1`, `text-align:center`)
- **rect**: เต็มความกว้างหน้า margin 15pt ซ้าย/ขวา, สูง ~30pt (1 บรรทัด) หรือ ~53pt (2 บรรทัด wrap)
- **content**: `{section}-{N} {ชื่อตู้} หน้า {P}`

### Square + FreeText label — กฎทั่วไป

- **Square** ครอบเฉพาะแถวของ datasheet ที่ตรงกับ spec — ไม่ครอบ row อื่น
- **FreeText label** วางใน**พื้นที่ขาว** ขวาของ rect (rightmost catalog text + 5pt) — คงขนานกับ y-range ของ rect
- เนื้อหาใต้ rect **ต้องตรงกับ Col C** (ถ้าไม่ตรง → ถาม user ก่อน อย่าแก้ rect ใหม่เอง)

### 🔑 Brand/Model Annotation Convention (สำคัญที่สุด!)

ใน catalog แต่ละ section มี **2 รูปแบบ label พิเศษ** ที่ใช้แทน "5.1.X ข้อ N)":

| Rect ที่ครอบ | Label ที่ใช้ | ตัวอย่าง |
|---|---|---|
| โลโก้ยี่ห้อ/บริษัท ใน catalog | **`ยี่ห้อ`** | LINK, Schneider, Cytron, PROTEUS, Fahchy, Jiant logos |
| ชื่อรุ่น/Model number | **`รุ่น`** | UV-9012H-SUS, QO116C06RCBO30, VTall-S203L-2 |

**rect ที่เหลือ (sub-phrase ของ ข้อย่อย)** ใช้ label ตามโครงสร้างเอกสาร:
- Section ที่มี nested ข้อ: `{section} ข้อ X) ข้อย่อย N.` (เช่น `5.1.2 ข้อ 2) ข้อย่อย 1.`)
- Section ที่ ข้อ X) ไม่มี nested: `{section} ข้อ X)`
- Section flat (sensor/LED/เสา): `{section} ข้อย่อย N.`

### Label DA / DS standard (ยี่ห้อ/รุ่น)

```
DA: 1 0 0 rg /HeBo 9 Tf
DS: font: bold Helvetica,sans-serif 9.0pt; text-align:left; color:#FF0000
```

ขนาด font 9pt เป็น standard — แต่ template เก่าบางไฟล์ใช้ 8pt หรือ 14pt (ปล่อยตามต้นฉบับได้ไม่ต้องแก้ ถ้า user ไม่ขอ)

---

## Annotation Pattern แต่ละ -X PDF (Master Tables)

### -1 PDFs (ตู้ UV cabinet) — page 595×842 (A4 portrait)

แตก Col B "1) เป็นตู้สำหรับติดตั้ง..." เป็น 6 sub-phrase rects + 1 model row rect + 1 brand logo rect:

| Rect | Coords | Label | Catalog text |
|---|---|---|---|
| **brand-link** | `[24.3, 757.1, 121.2, 823.4]` | `ยี่ห้อ` (label `[61.1, 732.7, 80.5, 752.7]` ใต้ rect) | LINK + American Standard logo |
| 1-hanging | `[47.2, 497.2, 82.8, 508.8]` | `{section} ข้อ 1)` | "Hanging" (ชนิดแขวน) |
| 1-outdoor | `[136.2, 497.2, 210.8, 508.8]` | `{section} ข้อ 1)` | "outdoor installation," |
| 1-two_layers | `[47.2, 445.2, 113.8, 456.8]` | `{section} ข้อ 1)` | "Two Layers Door" |
| 1-general_outdoor | `[141.2, 445.2, 311.8, 456.8]` | `{section} ข้อ 1)` | "for harsh environment with outdoor installation." |
| 1-material | `[101.2, 419.2, 281.8, 430.8]` | `{section} ข้อ 1)` | "Electro-Galvanized Sheet steel or Stainless steel," |
| 1-ip54 | `[120.2, 432.2, 240.8, 443.8]` | `{section} ข้อ 1)` | "Index of protection IP54 or IP55." |
| 1-uv9012hsus | `[55.2, 50.2, 525.8, 67.8]` | `รุ่น` (label `[528, 51, 590, 67]` ขวา rect) | UV-9012H-SUS row ในตาราง Order Information |

**Col D ของ parent** (R61, R97, R139, R177, R210): `ยี่ห้อ LINK รุ่น UV-9012H-SUS (Stainless)`

### -2 PDFs (RCBO) — page 502×843 (custom)

5 sub-phrase rects (ข้อย่อย 1-5) + 1 model rect + 1 brand rect:

| Rect | Coords | Label | Catalog text |
|---|---|---|---|
| **2-brand-schneider** | `[190.7, 17.0, 258.9, 42.5]` (P1 bottom) | `ยี่ห้อ` (label `[261.1, 23.4, 279.3, 33.4]` ขวา rect) | Schneider Electric logo |
| 2) ข้อย่อย 1 (P1) | `[33.1, 551.5, 439.6, 566.7]` | `{section} ข้อ 2) ข้อย่อย 1.` (label `[342.8, 552.1, 472.8, 566.1]`) | RCBO Product type |
| 2) ข้อย่อย 2 (P1) | `[33.1, 515.7, 439.6, 548.8]` | `{section} ข้อ 2) ข้อย่อย 2.` (label `[201.9, 516.3, 331.9, 548.2]`) | 1P+Ns 16A |
| 2) ข้อย่อย 3 (P1) | `[33.1, 462.0, 439.6, 477.2]` | `{section} ข้อ 2) ข้อย่อย 3.` (label `[181.2, 462.6, 311.2, 476.6]`) | Earth-leakage 30 mA |
| 2) ข้อย่อย 4 (P1) | `[33.1, 203.8, 439.6, 225.5]` | `{section} ข้อ 2) ข้อย่อย 4.` (label `[263.4, 204.4, 393.4, 224.9]`) | [Ics] 6000 A |
| 2) ข้อย่อย 5 (**P2**) | `[33.1, 425.1, 439.6, 458.3]` | `{section} ข้อ 2) ข้อย่อย 5.` (label `[205.6, 425.7, 335.6, 457.7]`) | Standards IEC 61009 |
| **2-model** | `[172.4, 673.4, 240.6, 684.6]` | `รุ่น` (label `[241.6, 674.0, 371.6, 684.0]` ขวา rect) | QO116C06RCBO30 |

**Col D ของ parent** (R63, R99, R141, R179, R212): `ยี่ห้อ Schneider Electric รุ่น QO116C06RCBO30`

> ⚠️ **อย่าใส่ spec ใน Col D ของ "รุ่น"** — รุ่น = model number เท่านั้น (เช่น QO116C06RCBO30) ไม่ใส่ "1P+NS 16A 30mA 6000A" เพราะนั้นเป็น spec ไม่ใช่รุ่น จะสร้างความสับสนให้คนตรวจ

### -3 PDFs (Micro Controller) — page 596×842 (A4 portrait)

10 sub-rects (ข้อย่อย 1-10) บน P3 + 1 brand rect + 2 model rects บน P1:

| Rect | Page | Coords | Label | Catalog text |
|---|---|---|---|---|
| **3-iriv-brand** | P1 | `[139.7, 704.4, 439.3, 803.6]` | `ยี่ห้อ` (label `[445.2, 745.4, 575.2, 761.6]`) | Cytron Technologies brand logo |
| **3-iriv** (model name) | P1 | `[207.2, 648.2, 386.8, 680.8]` | `รุ่น` (label `[389.4, 657.5, 519.4, 671.5]`) | IRIV PiControl |
| **3-cm5** | P1 | `[229.2, 619.2, 275.8, 645.8]` | (no label — share `รุ่น` กับ iriv) | CM5 |
| **shared topic (CPU)** | P3 | `[43.5, 636.5, 153.9, 673.5]` | (ไม่มี label — empty contents) | "CPU" topic label คอลัมน์ซ้ายของตาราง |
| 3) ข้อย่อย 1 (CPU) | P3 | `[354.2, 638.5, 545.8, 670.5]` | `{section} ข้อ 3) ข้อย่อย 1.` | Broadcom BCM2712 + Cortex-A76 SoC @ 2.4GHz **(คอลัมน์ขวาเท่านั้น ไม่ครอบ BCM2711)** |
| 3) ข้อย่อย 2 | P3 | `[42.2, 555.2, 552.8, 573.7]` | `{section} ข้อ 3) ข้อย่อย 2.` | WiFi 2.4/5GHz + BLE |
| 3) ข้อย่อย 3 | P3 | `[42.2, 415.7, 552.8, 434.2]` | `{section} ข้อ 3) ข้อย่อย 3.` | 4x Isolated digital input |
| 3) ข้อย่อย 4 | P3 | `[42.2, 392.5, 552.8, 411.0]` | `{section} ข้อ 3) ข้อย่อย 4.` | 4x Isolated digital output |
| 3) ข้อย่อย 5 | P3 | `[42.2, 369.2, 552.8, 387.7]` | `{section} ข้อ 3) ข้อย่อย 5.` | 4x Isolated analog input |
| 3) ข้อย่อย 6 | P3 | `[42.2, 346.0, 552.8, 364.5]` | `{section} ข้อ 3) ข้อย่อย 6.` | 1x Isolated RS232 |
| 3) ข้อย่อย 7 | P3 | `[42.2, 322.7, 552.8, 341.2]` | `{section} ข้อ 3) ข้อย่อย 7.` | 1x Isolated RS485 |
| 3) ข้อย่อย 8 | P3 | `[42.2, 276.2, 552.8, 294.7]` | `{section} ข้อ 3) ข้อย่อย 8.` | 1x mini PCIe socket |
| 3) ข้อย่อย 9 | P3 | `[42.2, 239.5, 552.8, 258.0]` | `{section} ข้อ 3) ข้อย่อย 9.` | DC 10-30V surge-protected |
| 3) ข้อย่อย 10 | P3 | `[42.2, 100.0, 552.8, 118.5]` | `{section} ข้อ 3) ข้อย่อย 10.` | Metal enclosure DIN rail |
| ข้อย่อย 11 (Software) | — | **ไม่มี annotation** | — | Col D = "ยินดีปฏิบัติตามข้อกำหนด" |

**Col D ของ parent** (R69, R105, R147, R185, R218): `ยี่ห้อ Cytron Technologies รุ่น IRIV PiControl CM5`

### -4 PDFs (Vibration Sensor) — page 595×842 (A4 portrait)

8 sub-rects (ข้อย่อย 1-8) + 1 model rect + **1 shared topic rect บน P7**. **ไม่มี brand logo** ใน catalog → Col D ใช้ `-` แทนยี่ห้อ:

| Rect | Page | Coords | Label | Catalog text |
|---|---|---|---|---|
| **4-model** | P7 | `[436.2, 166.2, 525.8, 180.8]` | `รุ่น` (label `[527.6, 167, 590, 180]` ขวา rect) | VTall-S203L-2 (ในตาราง spec) |
| **shared topic (Sensor range)** | P7 | `[53.2, 64.6, 167.4, 124.5]` | (ไม่มี label — empty contents) | Topic column "Sensor range" ครอบ row ของ ข้อย่อย 2+3 |
| 4) ข้อย่อย 2 | P7 | `[168.6, 101.1, 542.8, 121.1]` (narrow — value column only) | `{section} ข้อ 4) ข้อย่อย 2.` | Vibration acceleration ±16g |
| 4) ข้อย่อย 3 | P7 | `[168.1, 69.9, 542.8, 89.9]` (narrow — value column only) | `{section} ข้อ 4) ข้อย่อย 3.` | Velocity 0-200mm/s |
| 4) ข้อย่อย 4 | P8 | `[169.4, 757.8, 542.8, 777.8]` | `{section} ข้อ 4) ข้อย่อย 4.` | Vibration displacement |
| 4) ข้อย่อย 5 | P8 | `[169.4, 726.6, 542.8, 746.6]` | `{section} ข้อ 4) ข้อย่อย 5.` | Temperature range |
| 4) ข้อย่อย 7 | P8 | `[53.6, 658.7, 542.8, 718.4]` (tall — multi-line) | `{section} ข้อ 4) ข้อย่อย 7.` | Operating temperature |
| 4) ข้อย่อย 1 | P8 | `[54.2, 563.8, 542.8, 656.2]` (tall — multi-line) | `{section} ข้อ 4) ข้อย่อย 1.` | measurement direction X/Y/Z axis |
| 4) ข้อย่อย 8 | P9 | `[52.2, 349.6, 542.8, 400.3]` | `{section} ข้อ 4) ข้อย่อย 8.` | Communication RS485 |
| 4) ข้อย่อย 6 | P9 | `[52.2, 128.7, 542.8, 148.2]` | `{section} ข้อ 4) ข้อย่อย 6.` | IP67 protection |

**Col D ของ parent** (R81, R117): `ยี่ห้อ - รุ่น VTall-S203L-2`

### Sensor PDFs (5.1.4.2 BOD, 5.1.5.2 DO, 5.1.6.2 LED) — flat structure

**Sub-phrase rects** ครอบแถว datasheet ที่ตรงกับ ข้อย่อย แต่ละข้อ — label `{section} ข้อย่อย N.`

**Brand + Model rects (บน P1):**

| PDF | Brand rect (`ยี่ห้อ`) | Model rect (`รุ่น`) | Col D parent |
|---|---|---|---|
| **5.1.4.2 BOD** | `[40, 800, 270, 840]` (PROTEUS INSTRUMENTS logo top) | `[40, 627, 322, 653]` (Proteus Water Quality Probe title) | R159: `ยี่ห้อ Proteus Instruments รุ่น Water Quality Probe` |
| **5.1.5.2 DO** | `[181, 93, 245, 106]` (Shanghai Jiant text bottom) | `[179, 635, 418, 680]` (JG-LDO-N01 large title) | R197: `ยี่ห้อ JIANT รุ่น JG-LDO-N01` |
| **5.1.6.2 LED** | `[461, 419, 563, 449]` (Fahchy text top-right) | `[15, 55, 300, 175]` (LED Display P3.91 SMD OUTDOOR yellow box bottom-left) | R230: `ยี่ห้อ Fahchy รุ่น LED Display P3.91 SMD OUTDOOR` |

**กรณีพิเศษของ sensor PDFs:**
- BOD ข้อย่อย 4 (Output): rect ครอบ "Modbus® RTU" บน **P1** (ไม่ใช่ description ยาวบน P2)
- LED brightness (ข้อย่อย 3): rect บน **P6** ที่ "Brightness 4300CD/m" (ตามที่ Col D ระบุ "หน้า 6")
- LED life time (ข้อย่อย 4): rect บน **P10** ที่ "มีอายุการใช้งานนานถึง 100,000 ชั่วโมง"

### เสา PDFs (5.1.3.2, 5.1.4.3, 5.1.6.3) — งาน fabricate ไม่มียี่ห้อ

**Page sizes:** P1 = 1191×842, P2 = 859×531

| รูป | ตำแหน่ง | Label |
|---|---|---|
| ข้อย่อย 1 (ความสูง+เส้นผ่านศูนย์กลาง) | rect ที่ `Steel Pipe Ø4" OD ±10%Thickness 2.5 mm.` บน P1 | `{section} ข้อย่อย 1.` |
| ข้อย่อย 2 (Galvanize) | rect ที่ `Hot-Dip Galvanize` บน P1 | `{section} ข้อย่อย 2.` |
| ข้อย่อย 3 (Service Door) | rect ที่ Service Door dimension บน P1 | `{section} ข้อย่อย 3.` |
| ข้อย่อย 4 (ฐานราก) | **2 rects บน P2** — PLAN view + SECTION A view | `{section} ข้อย่อย 4.` (each rect) |

**P2 Foundation rects (exact template):**

| Annotation | Rect | Notes |
|---|---|---|
| **PLAN Square** | `[19.2, 103.5, 387.5, 481.8]` | ครอบทั้งภาพ PLAN view (left half) |
| **PLAN Label** `{section} ข้อย่อย 4.` | `[75.3, 73.1, 205.3, 98.1]` | ใต้ rect ที่ "PLAN" caption |
| **SECTION A Square** | `[394.2, 107.1, 824.0, 481.8]` | ครอบทั้งภาพ SECTION A view (right half) |
| **SECTION A Label** `{section} ข้อย่อย 4.` | `[446.8, 77.6, 576.8, 102.6]` | ใต้ rect ที่ "SECTION A" caption |

**Square contents tags:** `{section}(4-plan)` และ `{section}(4-sectionA)`

**Col D ของ ข้อย่อย 4**: `เทียบเท่าข้อกำหนด เอกสาร {section} เสาสำหรับติดตั้งตู้เก็บอุปกรณ์ หน้า 2 ข้อ {section} ข้อย่อย 4.`

**Col D ของ section parent** (R126, R164, R236): **เว้นว่าง** (เสา fabricate ไม่มี brand/model)

---

## กฎพิเศษที่ต้องจำเสมอ

### กฎข้อ 1: "ยินดีปฏิบัติตามข้อกำหนด" vs "ไม่พบใน catalog" — ใช้แยกชัด

**"ยินดีปฏิบัติตามข้อกำหนด"** ใช้กับเฉพาะ:
1. **งานติดตั้ง / install** (parent ข้อ + ทุก ข้อย่อย ภายใน) เช่น "เดินสาย", "ติดตั้ง", "ทดสอบระบบ"
2. **Software/Firmware ที่ทีมเขียนเอง** เช่น SCADA Application, Custom Dashboard, Mobile App ที่ develop เอง
3. Commitment statements ที่ไม่ใช่ spec hardware (เช่น "รับประกัน 1 ปี", "training")

**"ไม่พบใน catalog"** ใช้กับ:
- ข้อย่อยที่เป็น **hardware spec** หรือ **product feature** แต่หา callout/page ใน catalog ไม่เจอ
- ใช้เป็น **flag เตือน user** ให้ตรวจสอบอีกครั้ง — อาจต้อง:
  - หาหน้า catalog ที่ระบุ spec นั้น แล้วเพิ่ม annotation
  - เปลี่ยนรุ่นเป็นรุ่นที่ comply ได้
  - หรือยืนยันว่าเป็น commitment จริงๆ (แล้วเปลี่ยนเป็น "ยินดีปฏิบัติฯ")

**กฎตัดสินใจ:**
```
TOR ข้อย่อยพูดถึง...
├── installation/wiring/setup     → "ยินดีปฏิบัติตามข้อกำหนด"
├── software/firmware เขียนเอง    → "ยินดีปฏิบัติตามข้อกำหนด"
├── warranty/training/commitment  → "ยินดีปฏิบัติตามข้อกำหนด"
└── product spec/feature/standard → ต้องมี catalog ref
                                    ├── หาเจอ → "เทียบเท่าข้อกำหนด เอกสาร ..."
                                    └── หาไม่เจอ → "ไม่พบใน catalog" ⚠ FLAG
```

วิธีใช้:
- Col D เขียน "ยินดีปฏิบัติตามข้อกำหนด" หรือ "ไม่พบใน catalog" ตามกฎข้างต้น
- **ไม่ตี Square หรือ FreeText ใน PDF** สำหรับข้อย่อยทั้ง 2 ประเภทนี้

### กฎข้อ 2: Label positioning ในพื้นที่ขาว

1. หาเนื้อหาที่ขวาสุดของ catalog ในแถวเดียวกับ rect
2. Label `x_left = rightmost_text + 5pt`
3. Label `x_right = x_left + ~130pt` (ไม่เกิน `page_width - 10pt`)
4. Label y-range = same as rect (parallel)
5. ถ้า rect สูงเกิน 20pt → ย่อ label height เป็น 14pt อยู่กึ่งกลาง y

### กฎข้อ 3: Brand/Model rect — ใช้ "ยี่ห้อ" / "รุ่น" ห้ามใช้ "ข้อ N)"

**สำคัญ:** rect ที่ครอบโลโก้ยี่ห้อ หรือ ครอบชื่อรุ่น ห้ามใช้ label "5.1.X ข้อ N)" — ต้องใช้ `ยี่ห้อ` หรือ `รุ่น` ตามหน้าที่ของ rect

### กฎข้อ 4: User template = source of truth

เมื่อ user upload ไฟล์ template ที่แก้แล้ว:
1. **Replace ไฟล์เดิมด้วย upload** (ใช้ `os.open(path, O_WRONLY | O_TRUNC)` + `os.write()` เพื่อหลีกเลี่ยง permission issue)
2. **อ่าน annotations จาก upload** เพื่อหา coords ที่แม่นยำ
3. **Apply same coords + label content** ไปยังไฟล์อื่นใน group เดียวกัน
4. **อย่าเปลี่ยน convention** — ทำตามที่ user ตั้งไว้ในต้นแบบ

### กฎข้อ 5: Image-based brand logos

โลโก้ใน catalog หลายตัวเป็น raster image (LINK, Schneider, PROTEUS, Cytron, Fahchy, LED yellow box) — text extraction returns empty แต่ rect ครอบตำแหน่งภาพถูกต้อง — verify ด้วยการ render ภาพแล้วดู visually

### 🔑 กฎข้อ 6: Shared Topic Rect (สำหรับ catalog เป็นตาราง)

เมื่อ catalog เป็น **ตาราง** ที่หลาย ข้อย่อย share "topic"/row label เดียวกัน:

**ตัวอย่าง:** datasheet ของ Vibration sensor มีตาราง:
| Sensor range | acceleration ±16g  |  ← ข้อย่อย 2
|              | velocity 0-200mm/s  |  ← ข้อย่อย 3

ก่อนหน้านี้: rect ของ ข้อย่อย 2 และ 3 ครอบทั้งแถว (รวม "Sensor range" topic + value) — ทำให้ rect 2 อันแยกกันแต่ครอบ topic ซ้ำ

**Pattern ใหม่ (ถูกต้องกว่า):**
1. **1 shared topic rect** — ครอบเฉพาะ topic column (left), spans หลาย rows ของ ข้อย่อย ที่ share topic เดียวกัน
   - Contents = empty (ไม่มี label เพราะ support หลาย ข้อย่อย พร้อมกัน)
2. **N rects แยก** สำหรับแต่ละ ข้อย่อย — ครอบเฉพาะ value column (right)
   - แต่ละ rect มี label ของตัวเอง (`{section} ข้อ X) ข้อย่อย N.`)

**ข้อดี:** topic ไม่ซ้ำ + label "ข้อย่อย N" ชี้ไปที่ value ของ ข้อย่อย นั้นโดยตรง

**ตัวอย่างที่ใช้ pattern นี้:**

1. **5.1.X.-4 P7 — Sensor range (covers ข้อย่อย 2+3):**
   - Shared topic rect `[53.2, 64.6, 167.4, 124.5]` (covers "Sensor range" cell — y range covers 2 rows)
   - ข้อย่อย 2 rect `[168.6, 101.1, 542.8, 121.1]` (value only, narrow x)
   - ข้อย่อย 3 rect `[168.1, 69.9, 542.8, 89.9]` (value only, narrow x)

2. **5.1.X.-3 P3 — CPU topic (supports ข้อย่อย 1):**
   - Shared topic rect `[43.5, 636.5, 153.9, 673.5]` (covers "CPU" topic cell ในคอลัมน์ซ้ายของตาราง Features)
   - ข้อย่อย 1 rect `[354.2, 638.5, 545.8, 670.5]` (BCM2712 spec — คอลัมน์ขวาเท่านั้น)
   - แม้จะ support แค่ ข้อย่อย 1 แต่ยังตี shared topic rect แยก เพื่อให้ชัดเจนว่า "CPU" คือ topic ของแถวนี้

### กฎข้อ 7: รุ่น vs Spec — ห้ามปน

ใน Col D parent row ของ "ยี่ห้อ X รุ่น Y":
- **Y = Model number / part number เท่านั้น** (เช่น QO116C06RCBO30, UV-9012H-SUS, JG-LDO-N01)
- **ห้ามใส่ spec values** ต่อท้าย (เช่น "1P+NS 16A 30mA 6000A") → จะสร้างความสับสนให้คนตรวจ — spec values อยู่ใน Col B/C อยู่แล้ว
- ถ้าต้องการระบุ variant ใส่เป็นวงเล็บได้ (เช่น `UV-9012H-SUS (Stainless)`)

### 🔑 กฎข้อ 9: Hybrid Annotation Pattern — Highlight ก่อน Rect (NEW 2026-05-10)

**ค้นพบจากศึกษา SR's annotation convention** — pattern ที่เร็วและ print-friendly

> **สำคัญ:** ใช้กับ catalog ที่ **เราทำเอง** — SR's catalog มี annotation ของเค้าอยู่แล้ว **ห้ามแก้**

**กฎ:** เลือก annotation type ตาม PDF type:

| PDF type | Detection | Annotation type | Speed |
|---|---|---|---|
| **Text-based** (text_len > 500 chars/page) | `page.get_text()` length | **Highlight สีเหลือง + callout `N)`** | < 1 วิ/spec |
| **Image-based** (text_len < 100 chars/page) | scanned PDF, image-only | **Rect + callout `N)`** | 5-30 วิ/spec |
| **Mixed** (100-500 chars/page) | partial text layer | ลอง highlight ก่อน → fallback rect | varies |

**SR's callout convention (study reference, see `knowledge_base/sr_pattern.md`):**
- Format: `N)` — สั้น แค่เลข ข้อย่อย + closing paren
- Section number **ไม่ใส่** — implicit จาก folder structure (ไฟล์อยู่ใน folder section นั้น)
- White background (พิมพ์เห็นชัด)
- Position: ใน white space ใกล้ highlight (right margin / table empty area)

**Sample code (Text-based highlight + callout `N)`):**

```python
def hl_callout(page, search_text, sub_n, page_w):
    """Highlight text + add SHORT callout 'N)' at right margin (SR-style)."""
    rects = page.search_for(search_text)
    if not rects: return False
    for r in rects:
        h = page.add_highlight_annot(r)
        h.set_colors(stroke=(1, 1, 0))  # yellow
        h.update()
    first = rects[0]
    # SR convention: short callout 'N)' with white bg
    lbl_rect = fitz.Rect(page_w - 25, first.y0, page_w - 5, first.y0 + 12)
    ft = page.add_freetext_annot(lbl_rect, f"{sub_n})", fontsize=9, fontname="hebo",
                                  text_color=(1, 0, 0),
                                  fill_color=(1, 1, 1),  # WHITE BG (print-clear)
                                  align=fitz.TEXT_ALIGN_CENTER)
    ft.set_border(width=0); ft.update()
    return True

# Usage example (H3C R4900 G7 if WE do annotation):
# hl_callout(p3, "Intel® Xeon® 6 Processors", 1, 596)
# hl_callout(p3, "DDR5 RDIMM Slots", 3, 596)
```

**Brand + Model บน cover** — ใช้ rect (cover มัก image-based + logo เป็นรูป) + label `ยี่ห้อ` / `รุ่น`

**ข้อดี vs rect-only approach:**
1. **เร็วกว่า 10-30x** — ไม่ต้องหา coords manually
2. **ชัดกว่า** — highlight ทับข้อความตรงๆ
3. **Print-friendly** — short callout `N)` พื้นหลังขาว เห็นชัดเวลาพิมพ์
4. **ไม่บดบังเนื้อหา** — callout อยู่ margin

**Cross-ref กับ xlsx Col D:** ตอนนี้ Col D ใช้ format `เทียบเท่าข้อกำหนด เอกสาร {section} {name} หน้า {P} ข้อ {section} ข้อย่อย {N}.` — verify โดยเปิด PDF + ดู callout `N)` ที่ตรงกับ ข้อย่อย N

---

### 🔑 กฎข้อ 8: Single-row Comply Item — Extract spec จาก main item

เมื่อแถว comply เป็น **single-row** ที่ Col B/C **ไม่มี ข้อย่อย 1) 2) 3)...** (เช่น R349 5.2.1.9 "แผงพักสายไฟเบอร์ออฟติคขนาด 24 Core พร้อมอุปกรณ์ จำนวน 2 ชุด") ยังต้องตี rect ใน catalog เหมือน multi-row item โดย:

**Step 1: extract sub-phrases จาก main item Col B**
- หาคำที่เป็น **spec ที่ตรวจสอบได้** ใน catalog (ไม่ใช่จำนวน "ชุด")
- เช่น "24 Core" → bullet ใน catalog ที่ระบุ "24 fiber cores"
- เช่น "พร้อมอุปกรณ์" → "Accessories provided" + รายการ

**Step 2: ตี rects (4 รูปแบบ + custom Thai labels)**

| rect type | label | ตัวอย่าง |
|---|---|---|
| brand logo | `ยี่ห้อ` | LINK logo, UNION logo, Thai union logo |
| model number | `รุ่น` | UFC9312A cell, UF-2010A row, model code in spec table |
| spec sub-phrase | `{section}` หรือ **custom Thai phrase** | `5.2.2.8` หรือ `ใช้ภายนอกอาคาร` |
| context-specific | **custom Thai phrase ตรงกับ Col B** | `Twisted Pair Shield`, `ใช้ภายในอาคาร` |

**🔑 User pattern (จาก 5.2.1.13, 5.2.1.14, 5.2.2.8, 5.2.2.10 references):**

1. **Custom Thai labels แทน `{section}`** — เมื่อ rect ครอบ phrase ที่มีคำไทยเฉพาะ ให้ใช้คำไทยนั้นเป็น label แทน section number:
   - rect ครอบ "OUTDOOR INSTALLATION" body → label `ใช้ภายนอกอาคาร` (จาก Col B)
   - rect ครอบ "Indoor Installation" body → label `ใช้ภายในอาคาร`
   - rect ครอบ "Twisted Pair" image → label `Twisted Pair Shield`
   - ใช้ font Bold red 9pt (สำหรับ A4) / 36pt (สำหรับ 2480×3507)

2. **Multi-page annotations** — single-row item ครอบหลายหน้าได้:
   - US-9106LSZH (3-page catalog): brand+model+spec บน P1, indoor-installation rect บน P3 ด้วย (ที่ Order Information row + Applications text)
   - VCT (12-page): brand บน P1 (cover), model+spec บน P5 (VCT spec page)

3. **Same /Contents tag, different rects** — สามารถใช้ tag เดียวกันสำหรับหลาย rect ที่ครอบ spec เดียวกันคนละตำแหน่ง (เช่น `(indoor-installation)` ใช้กับ rect บน P1 + 2 rects บน P3)

4. **Label position flexibility** — label วางในพื้นที่ขาวที่ใกล้สุด:
   - **ขวา** ของ rect (default — เช่น brand → ยี่ห้อ)
   - **ซ้าย** ของ rect (เช่น EMT brand label, UFC9312A รุ่น label)
   - **ด้านบน** rect (เช่น VCT brand label)
   - **ด้านล่าง** rect

**Step 3: Col D format** — ใช้ Col B description (ตัด section + จำนวน ออก) ต่อด้วย model:
```
{section} {Col B description minus "จำนวน N <unit>"} {model}
```

**กฎเขียน Col D:**
1. ตัด `5.X.Y.Z. ` (section + dot + space) ออกจากต้น Col B
2. ตัด ` จำนวน <number> <unit>` ออกจากท้าย (เช่น `จำนวน 500 เมตร`, `จำนวน 2 ชุด`)
3. คงข้อความที่เหลือ ทั้งหมด (รวม "ขนาด...", "แบบ...", "ใช้ภายในอาคาร" ฯลฯ)
4. ต่อท้ายด้วย model name

**ตัวอย่าง:**
- B: `5.2.1.9. แผงพักสายไฟเบอร์ออฟติคขนาด 24 Core พร้อมอุปกรณ์ จำนวน 2 ชุด`
- D: `5.2.1.9 แผงพักสายไฟเบอร์ออฟติคขนาด 24 Core พร้อมอุปกรณ์ UF-2010A`
- B: `5.2.2.10. สายไฟ VCT 3*2.5 SQ.mm จำนวน 200 เมตร`
- D: `5.2.2.10 สายไฟ VCT 3*2.5 SQ.mm Thai union 300/500 V 70 °C 60227 IEC 53 (VCT)`

**ตัวอย่าง Sister rows (R349 / R408 / R471 / R534):** ทั้งหมดใช้ catalog UF-2010A เหมือนกัน → coords identical, ต่างแค่ section number ใน label

**ค่า rect coords UF-2010A (page 2480x3507):**
- brand-link: `[80, 60, 530, 320]` (LINK logo)
- 24cores: `[30, 2493, 1700, 2552]` (bullets 7-8: "support 24 fiber cores" + "UF-2010A supports 6-24 fiber ports")
- accessories: `[30, 2570, 1500, 2755]` (bullet 9 + 3 sub-items)
- model: `[289, 3304, 461, 3341]` (UF-2010A row in spec table)

**Reference patterns ที่ user ตี (ใช้เป็น template):**

| Catalog | Page(s) | Reference file | Annotations |
|---|---|---|---|
| `US-9106LSZH.pdf` | P1 + P3 | 5.2.1.13 | brand-link, model (US-9106LSZH text), indoor-installation × 3, custom labels: `Twisted Pair Shield`, `ใช้ภายในอาคาร` |
| `แคตตาล็อค Union emt.pdf` | P3 | 5.2.1.14 | brand-union (label LEFT), ul797-spec, size-3-4-row (รุ่น label) |
| `UFC9312A.pdf` | P1 | 5.2.2.8 | brand-link, armored-title × 2 (title bar + body), singlemode-row, model-cell (รุ่น label LEFT), custom label `ใช้ภายนอกอาคาร` |
| `สายไฟ thai union vct.pdf` | P1 + P5 | 5.2.2.10 | P1: brand-thai-union (label ABOVE); P5: vct-title, model (cable text), tis-standard |
| `UF-2010.pdf` | P1 (2480×3507) | 5.2.1.9 | brand-link, 24cores, accessories, model |

**วิธี clone reference ไป sister files:**
1. ใช้ `extract_annots()` อ่าน annotations ทุกหน้า (subtype, rect, contents)
2. Rebuild sister: `shutil.copy2(source, sister) + add_freetext + add_rect_annot` ตาม spec ของ reference
3. เปลี่ยน label content จาก `ref_section` → `sister_section` (เช่น `5.2.1.13` → `5.2.3.9`)
4. คง custom Thai labels (`ใช้ภายในอาคาร`, ฯลฯ) ตามเดิม
5. รักษา /Contents tag เดิม (`(brand-link)`, `(indoor-installation)`, ฯลฯ)
6. Header rebuild จาก sister filename + page number

ใช้ `garbage=4, clean=True` ตอน save เพื่อ purge phantom annotations

---

## Workflow สำหรับ section ใหม่

เมื่อทำ section ใหม่ (เช่น 5.1.1, 5.1.7, 5.1.8, 5.2) ทำตามลำดับ:

1. **อ่าน TOR** เพื่อ verify Col B ตรงตัว (รวม typo)
2. **หา catalog PDFs** ของอุปกรณ์
3. **ใส่ PDFs ใน output/{section}/{section}-N/** พร้อมโฟลเดอร์สั้นๆ
4. **เพิ่ม header** ด้วย `scripts/pdf_header.py` + ปรับเป็นแดง 14-18pt center, span page width
5. **ตี rects** สำหรับแต่ละ ข้อย่อย ที่มีใน datasheet (ครอบเฉพาะแถว/พื้นที่ที่ตรง)
6. **เพิ่ม brand rect** ที่โลโก้ยี่ห้อ (label = `ยี่ห้อ`)
7. **เพิ่ม model rect** ที่ชื่อรุ่น (label = `รุ่น`)
8. **วาง labels ของ ข้อย่อย** ในพื้นที่ขาว ขนาน y-range
9. **อัพเดต Col C** — ตัดคำเปรียบเทียบจาก B + ใช้ค่าจริงจาก catalog
10. **อัพเดต Col D**:
    - parent ข้อ ที่มีรุ่น → `ยี่ห้อ {brand} รุ่น {model}` (หรือ `ยี่ห้อ - รุ่น {model}` ถ้าไม่มี brand)
    - ข้อย่อย hardware spec ที่หา catalog เจอ → `เทียบเท่าข้อกำหนด เอกสาร {section}-{N} ... หน้า {P} ข้อ {section} ข้อ X) ข้อย่อย N.`
    - งานติดตั้ง / software-เขียนเอง / commitment → `ยินดีปฏิบัติตามข้อกำหนด`
    - hardware spec ที่หา catalog ไม่เจอ → `ไม่พบใน catalog` ⚠ (flag ให้ user ตรวจ)
    - section header → ปล่อยว่าง
11. **verify** เนื้อหาใต้ rect ตรงกับ Col C — ถ้าไม่ตรง ถาม user ก่อนแก้
12. **ขอ confirm brand/model values** กับ user ก่อนเขียน Col D

---

## Verification Checklist (ตอนเสร็จ section)

- [ ] Col D ทุกแถวตรง 1 ใน 7 รูปแบบที่ระบุ (เทียบเท่า / สูงกว่า / brand-model / dash-brand / commitment "ยินดีปฏิบัติฯ" / flag "ไม่พบใน catalog" / empty) หรือ filename-format (single-row) / model-only (nested ข้อย่อย)
- [ ] Catalog references ใน Col D resolve ไป rect+label จริงใน PDF
- [ ] Brand rect → label = `ยี่ห้อ`, Model rect → label = `รุ่น`
- [ ] Sub-phrase rects → label `{section} ข้อ X) ข้อย่อย N.` หรือ `{section} ข้อย่อย N.`
- [ ] Headers ทุกหน้า: red, center (Q=1), 14-18pt, span page width
- [ ] Software item / งานติดตั้ง: NO annotation (memory rule)
- [ ] Col B preserve ตาม TOR (รวม typo)
- [ ] Col C ปลอดคำเปรียบเทียบ (หรือดีกว่า/ไม่น้อยกว่า/จะต้อง/ต้องสามารถ)
- [ ] Labels positioned ในพื้นที่ขาว — ไม่ทับเนื้อหา catalog
- [ ] เนื้อหาใต้ rect (visual check ผ่าน pdftoppm) ตรงกับ Col C

---

## เครื่องมือ

### scripts/pdf_header.py
```bash
# เพิ่ม header อัตโนมัติ
python3 scripts/pdf_header.py --overwrite "output/**/*-1.pdf"

# กำหนด text เอง
python3 scripts/pdf_header.py --text "{name} หน้า {page}" --overwrite file.pdf
```
Template variables: `{name}`, `{page}`, `{pages}`, `{dir}`

### scripts/fix_uv_headers.py
Script เฉพาะสำหรับแก้ header UV cabinet -1 ทั้ง 5 section ด้วย PyMuPDF

### scripts/version.py — Snapshot-based version control

ระบบ version control แบบ snapshot ที่ทำงานกับ Google Drive ได้ดี (ไม่ใช่ git — หลีกเลี่ยง sync conflicts)

**Snapshots เก็บใน:** `_versions/snapshots/<ID>/`
- `<ID>` = `YYYY-MM-DD_HHMMSS[_tag]`
- แต่ละ snapshot มี: `manifest.json` + ไฟล์ที่ track + (optional) `output.tar.gz`

**Commands:**
```bash
# Quick snapshot (xlsx + SKILL.md เท่านั้น — ~150KB)
python3 scripts/version.py snap "before-emt-fix"

# Full snapshot (+ output/ tar.gz — ~115MB)
python3 scripts/version.py snap-full "release-v1"

# ดูทั้งหมด
python3 scripts/version.py list

# ดูรายละเอียด snapshot (รับ ID prefix ได้ เช่น 2026-05-09)
python3 scripts/version.py show 2026-05-09_1130

# เทียบ 2 snapshots — แสดง file changes + output tree diff
python3 scripts/version.py diff <id1> <id2>

# Restore xlsx + SKILL.md (ถามก่อน overwrite)
python3 scripts/version.py restore <id>

# Restore ทั้ง output/ จาก tarball
python3 scripts/version.py restore-full <id>

# ลบ snapshots เก่า เก็บแค่ N ล่าสุด
python3 scripts/version.py prune --keep 10

# Auto-snap เฉพาะถ้า xlsx เปลี่ยน (ใส่ใน cron / หรือเรียกก่อนเริ่มงาน)
python3 scripts/version.py auto-snap
```

**Workflow แนะนำ:**
1. ก่อนเริ่มงานใหญ่ → `snap "before-<work>"`
2. หลังเสร็จงานหลัก → `snap "after-<work>"` หรือ `snap-full "milestone-X"`
3. ทุกสิ้นวัน → `snap "eod-YYYY-MM-DD"` หรือ `auto-snap`
4. เดือนละครั้ง → `prune --keep 30`

**Tracked files (quick mode):**
- `output/Comply spec Smart Plant 1.xlsx` (ไฟล์หลัก)
- `SKILL.md` (documentation)
- `output_tree`: รายการไฟล์ใน `output/` พร้อม size + mtime (ไม่เก็บ content)

**Excluded จาก full snapshot:**
- `_archive/` (เก่าแล้ว)
- `~$*` (Excel lock files)
- `.DS_Store`

### Library ที่ใช้
- `pypdf` — read/write PDF annotations (ใน sandbox)
- `pdfplumber` — extract text + word positions ภายใน rect bounds
- `openpyxl` — read/write Comply spec xlsx
- `PIL` — measure text width (สำหรับ header font sizing)
- `fitz` (PyMuPDF) — บน macOS (sandbox มี proxy block)

### Script template ที่ใช้บ่อย

```python
# Load PDF + write back (workaround สำหรับ permission issue)
import os, tempfile
from pypdf import PdfReader, PdfWriter

reader = PdfReader(path)
writer = PdfWriter()
for p in reader.pages:
    writer.add_page(p)
# ... modify ...
tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
tmp.close()
with open(tmp.name, 'wb') as f:
    writer.write(f)
with open(tmp.name, 'rb') as src:
    data = src.read()
fd = os.open(path, os.O_WRONLY | os.O_TRUNC)
os.write(fd, data)
os.close(fd)
os.unlink(tmp.name)
```

---

## หมายเหตุทางเทคนิค

- **PDF coordinate system:** y=0 ที่ขอบล่าง, y เพิ่มขึ้นไปบน (ตรงข้ามกับ pdfplumber's `top` ที่นับจากบน)
- **A4 size:** 595.2 × 841.7 pt → Header rect: `[15, ph-35, pw-15, ph-8]`
- **Header DA**: `1 0 0 rg /HeBo {N} Tf` (N=14-18)
- **Header DS**: `font: bold Helvetica,sans-serif {N}.0pt; text-align:center; color:#FF0000`
- **ยี่ห้อ/รุ่น Label DA**: `1 0 0 rg /HeBo 9 Tf` (standard)
- **ยี่ห้อ/รุ่น Label DS**: `font: bold Helvetica,sans-serif 9.0pt; text-align:left; color:#FF0000`
- **Square C (border color)**: `[1 0 0]` (red RGB)
- **Annotation Thai text**: viewer ต้องมี Thai font fallback (Adobe Reader OK, Poppler limited)
- **ลบ /AP** เมื่อแก้ DA/DS/Q/Contents เพื่อให้ viewer re-render
- **ชื่อโฟลเดอร์ยาว:** Linux NAME_MAX=255 bytes → ตั้งชื่อ sub-folder สั้น
- **/Contents tag** ของ Square — ใช้ระบุประเภท rect (`(brand-link)`, `(2-model)`, `(4-plan)`) สำหรับ programmatic identification

### File operations ใน sandbox

- ไฟล์ที่ user upload หลัง initial state อาจแสดง 0 bytes ใน workspace mount แม้ pypdf อ่านได้ — ต้องระวังเวลา `cp` หรือ `shutil.copy` (อาจ overwrite ด้วยข้อมูล 0 bytes)
- ใช้ Python `open(upload, 'rb').read()` + check `len(data) > 0` ก่อนเขียน
- การลบไฟล์ใน sandbox อาจติด permission error — ใช้ `os.open` with `O_TRUNC` แทน

---

## สถานะปัจจุบัน (อัพเดต 2026-05-07)

### ✅ เสร็จแล้ว — UV Cabinet + Sensors + LED + เสา (Section 5.1.2 — 5.1.6)

- **PDF annotations** ทุกไฟล์ใน 5.1.2 - 5.1.6 (23 PDFs, 212 หน้า)
- **Headers**: แดง, center, 14-18pt — ทุกหน้า
- **Brand annotations** (label `ยี่ห้อ`):
  - 5 -1 PDFs → LINK
  - 5 -2 PDFs → Schneider Electric  
  - 5 -3 PDFs → Cytron Technologies
  - BOD → Proteus Instruments
  - DO → Shanghai Jiant
  - LED → Fahchy
  - -4 PDFs (Vibration), เสา PDFs → no brand
- **Model annotations** (label `รุ่น`):
  - UV-9012H-SUS, QO116C06RCBO30, IRIV PiControl + CM5, VTall-S203L-2
  - Water Quality Probe, JG-LDO-N01, LED Display P3.91 SMD OUTDOOR
- **Sub-phrase rects** ครบทุก ข้อย่อย พร้อม labels
- **เสา P2**: PLAN + SECTION A foundation rects + Col D updated
- **CPU rect** ปรับให้ครอบเฉพาะ BCM2712 column (-3 PDFs)
- **Software item 11** ใน -3 PDFs: ลบ annotation ตามกฎ
- **Comply spec Col C** rows 62-247: ค่าจริงจาก catalog, ซ่อม typo
- **Comply spec Col D** rows 62-247: catalog ref + ยี่ห้อ/รุ่น

### ❌ ยังค้างอยู่

- **Section 5.1.1** (rows 7-60) — Server Rack, NGFW, L2 Switch, NAS, Tablet, UPS 10kVA — ยังไม่มี catalog PDFs
- **Section 5.1.7** SCADA Software (rows 248-273) — น่าจะเป็น "ยินดีปฏิบัติฯ" ทั้งหมด
- **Section 5.1.8** เดินสาย (rows 274-277) — งานติดตั้ง น่าจะเป็น "ยินดีปฏิบัติฯ"
- **Section 5.2** Network Computer (rows 278-663) — section ใหญ่ ครอบคลุม 6 sub-sections
- **BOQ** `5.ปริมาณงาน Smart Plant P1  23-04-2569.xlsx` — ยังไม่ได้ verify ตรงกับ TOR
- **Polish (optional)**:
  - 5.1.2.-3 ยี่ห้อ/รุ่น labels เป็น 14pt (อื่น 9pt) — visual ไม่เป็นเอกภาพ
  - 5.1.2.-2 Schneider rect tag ซ้ำกับ QO model tag (`5.1.2(2-model)` ทั้งคู่)

---

## อัพเดต 2026-05-08 — Section 5.1.1.1 / 5.2.1.1 (Rack G3N + G7-00012/G7-05002)

### Catalog ที่ใช้ (uploaded May 8)
- `catalog/G3N-61142.pdf` (3 หน้า A4) — 19" GERMANY G3 series 42U rack
- `catalog/G7-00012.pdf` (1 หน้า A4) — 19" GERMANY AC Power Distribution 12 outlets, 16A
- `catalog/G7-05002.pdf` (1 หน้า A4) — 19" GERMANY Heavy Duty Fan Set (2x Ø4")
- `catalog/สายไฟ thai union.pdf` (12 หน้า landscape) — VCT cable
- `catalog/แคตตาล็อค Union.pdf` (5 หน้า A4) — EMT/IMC conduit (Arrow Syndicate UNION)
- *ยังไม่ได้รับ:* G1-60609 (9U rack สำหรับ 5.2.1.2/5.2.3.1/5.2.4.1), LINK UV-9012H-IP55,
  UF/US series switches, FortiGate, MikroTik, etc.

### Output folder convention (สำคัญ)

ตามที่ user ทำใน 5.1.1.-1:
- Parent rack catalog (ข้อ 1 ของ section 5.X.Y.Z): file ชื่อ `5.X.Y. [parent_section_name]-Z.pdf` ใน folder `5.X.Y.-Z/`
- Sub-item ที่มี catalog แยก (ข้อ N ของ 5.X.Y.Z): file ชื่อ `5.X.Y.Z-N [keyword] [model].pdf` วางใน folder เดียวกัน

ตัวอย่าง (`output/5.1.1. .../5.1.1.-1/`):
```
5.1.1. ระบบควบคุมการบำบัดน้ำเสียอัจฉริยะ-1.pdf       # parent rack (ข้อ 1)
5.1.1.1-3 ช่องเสียบไฟฟ้า G7-00012.pdf                  # ข้อ 3
5.1.1.1-4 พัดลมระบายความร้อน G7-05002.pdf            # ข้อ 4
```

### Annotation pattern — Multi-page parent catalog (rack)

จากตัวอย่าง user ใน `5.1.1. ระบบควบคุมการบำบัดน้ำเสียอัจฉริยะ-1.pdf`:

**P1 (cover, image-only):**
- Header `[15, 784, 580.22, 837]` 14pt red center: `5.1.1-1 [section_name] หน้า 1`
- Brand rect `[431.9, 724.5, 568.2, 796.0]` (19"GERMANY logo top-right)
- "ยี่ห้อ" label `[378, 720, 428, 758]` 11pt red right
- ข้อ 2 rect `[195, 542.1, 405.9, 560]` (Electro-galvanized text in FEATURES bullet list)
- "5.1.1.1 ข้อ 2)" label `[342.1, 551.7, 482.1, 571.7]` 10pt **BLACK** left

**P2:** header only

**P3 (ORDERING table, image-only):**
- Header (ดังที่ P1)
- Model rect (G3N-61142 cell at 42U row D=1100 col) `[441.1, 230.6, 490.4, 249.4]`
- "รุ่น" label `[481.7, 235.2, 525.6, 246.4]` 11pt red **center**
- 1U+Overall cell rect (42U size + 2050mm height) `[115.8, 229.9, 208.2, 248.7]`
- "5.1.1.1 ข้อ 1)" label `[12.2, 224.5, 152.2, 244.5]` 10pt **RED** center (left margin)
- Cabinet diagram rect (อีกตำแหน่งสำหรับ ข้อ 1) `[439.1, 309.3, 488.4, 328.1]`
- "5.1.1.1 ข้อ 1)" label `[450.7, 304.1, 590.7, 324.1]` 10pt RED center

### Annotation pattern — Single-page catalog (G7 outlet/fan)

**Header at BOTTOM** (เพราะ top มี title stripe + logo): `[15, 5, 580, 30]` 12pt red center

**Brand rect**: ครอบ logo + EXPORT RACK text เต็มความสูง
- G7-00012: `[454.9, 738.2, 579.4, 821.8]`
- G7-05002: `[456.2, 734.9, 580.8, 821.5]`
- "ยี่ห้อ" label วาง overlay ภายในกรอบเลย (ไม่ออกนอก)

**Model rect แบ่งเป็น 2 กรอบ** (Order No column + Description column):
- Order No cell: narrow rect (~55×30pt)
- Description: wider rect (~185×30pt) แยกออกมาขวา
- "รุ่น" label red 11pt center วาง**ใกล้** Order No rect (อาจ above/below)
- "5.X.Y.Z ข้อ N)" label 10pt **BLACK** left วาง**ใกล้** Description rect

ตัวอย่าง G7-00012 (ข้อ 3):
- Order No rect: `[249.3, 133.9, 303.8, 163.6]`  
- รุ่น label: `[247.2, 157.2, 296.5, 174.2]`
- Description rect: `[306.2, 133.4, 490.4, 163.0]`
- ข้อ label (BLACK): `[401.7, 154.4, 541.7, 174.4]`

### สีของ label "5.X.Y.Z ข้อ N)"
- **BLACK (#000000)** สำหรับ rect บนข้อมูลทั่วไป (FEATURES list, Description column)
- **RED (#FF0000)** สำหรับ rect บนตาราง ORDERING (1U cell, cabinet diagram บน rack P3)
- หลักการ: Red ถ้าวางในส่วนที่มีกรอบ/ตารางเด่น ๆ, Black ถ้าวางบน body text

### Comply.xlsx Col D convention (ใหม่)

จาก user แก้ R10-R13:

| ประเภท ข้อ | format Col D |
|------------|---------------|
| Parent (5.X.Y.Z มี catalog) | `ยี่ห้อ [brand] รุ่น [model]` |
| ข้อ N ของ parent + rect ใน parent catalog | `เทียบเท่าข้อกำหนด เอกสาร 5.X.Y-Z [section_name] หน้า {N} ข้อ 5.X.Y.Z ข้อ N)` |
| ข้อ N ของ parent + มี catalog แยก | `5.X.Y.Z-N [keyword] [model]` (เช่น `5.1.1.1-3 ช่องเสียบไฟฟ้า G7-00012`) |
| ข้อย่อยซอฟต์แวร์/ติดตั้ง | `ยินดีปฏิบัติตามข้อกำหนด` |

**สำคัญ:** สำหรับ ข้อ N ที่ต้องการ rect ใน parent catalog (เช่น "Electro-galvanized" บน P1 ของ rack) — ให้ใช้ format "เทียบเท่าข้อกำหนด..." แทน "ยินดีปฏิบัติ" เพราะมี rect แล้ว

### Status (พฤษภาคม 8)

#### ✅ เสร็จแล้ว
- 5.1.1.1 (R9-R13): rack G3N-61142 + G7-00012 + G7-05002 — annotated ตาม template user
- 5.2.1.1 (R281-R285): mirror ของ 5.1.1.1 — annotated โดย agent ตาม template เดียวกัน
- xlsx Col D updated ทั้ง 5.1.1.1 และ 5.2.1.1

#### ⏳ คงเหลือ catalog SMART
- Thai union VCT (12 หน้า landscape) → R443, R506, R569
- Union EMT 3/4" (5 หน้า A4) → R277, R383, R444, R507, R570
- G1-60609 9U rack → R286 (5.2.1.2), R349 (5.2.3.1?), etc. — ยังไม่ได้รับไฟล์
- G7-05002 ต้องทำสำเนาไปที่ folder 5.2.1.-2 ด้วย (ใช้ในตู้ 9U) เมื่อมี G1-60609 catalog

---

## อัพเดต 2026-05-09 — All Comply Done (637/660 rows ครบ)

### Pipeline Strategy (Opus + Sonnet Agent)

ค้นพบว่าสำหรับงานที่มี **sister sections** (catalog เดียวใช้ใน sub-sections หลายอัน) — pipeline ผสม model ประหยัดได้ 40-50%:

```
[Opus 4.7]   inspect catalog + design rect coords + annotate REFERENCE file (1 ครั้ง/catalog)
[Sonnet]     spawn Agent → clone reference annotations → all sister files + bulk xlsx update
```

**ใช้ pipeline นี้สำเร็จกับ:**
| Catalog | Sister sections | Rows |
|---|---|---|
| L3 Switch RG-CS85 | 4 | 40 |
| PoE L2 Switch CS4220 | 6 | 48 |
| TP-Link AP EAP660 | 4 | 36 |
| Dahua CCTV DH-IPC-HFW4231T | 6 | 132 |

**Sonnet Agent prompt template:** ใน `knowledge_base/pipelines.md`

### Single-section catalogs (Opus only)
NVR, Server (re-use), PC+Monitor, Tablet, NAS, Core Switch, UPS, NGFW, L2 Switch — ทำใน Opus เลยเพราะไม่มี sister files

### HTML→PDF (Tablet — Apple iPad spec page)

User upload spec เป็น HTML จาก apple.com — convert ด้วย `weasyprint`:
```bash
weasyprint "iPad spec.html" /tmp/ipad.pdf
```
**Caveats:** image references จาก local cache อาจ broken (warnings, ignored), text + key spec จะ render สมบูรณ์

### "สูงกว่าข้อกำหนด" pattern (Col D format ใหม่ #2)

เมื่อ catalog spec เกินกว่า TOR ต้องการ:
- TOR: "MAC Address ≥ 32,000" → catalog: 64,000 → `สูงกว่าข้อกำหนด`
- TOR: "Image Sensor ≥ 1/3"" → catalog: 1/2.8" → `สูงกว่าข้อกำหนด`
- TOR: "IP66" → catalog: IP67 → `สูงกว่าข้อกำหนด`
- TOR: "Switching Capacity ≥ 1 Tbps" → catalog: 1.44 Tbps → `สูงกว่าข้อกำหนด`

Format เหมือน "เทียบเท่า" ทุกอย่างเปลี่ยนแค่คำหัว

### กฎใหม่: Vendor Inheritance (Col E)

**Sub-item rows (ข้อย่อย) ของ multi-row parent → vendor เดียวกับ parent**

ใน xlsx ตอน update:
- Set Col E parent row first
- Then for child rows, **ใส่ Col E ด้วย vendor เดียวกัน** (อย่าเว้นว่าง)
- ถ้าเว้นว่าง → fix ภายหลังด้วย script "inherit from parent" (วน scan + carry forward)

### CCTV catalog gotcha (ICT TOR pre-annotated)

`catalog/กล้องโทรทัศน์วงจรปิด...DH-IPC-HFW4231T-ZAS_S0(ICT).pdf` มี **(ICT)** suffix — เป็น catalog ที่ผ่านการ annotate ตาม ICT TOR แล้ว มี:
- **Highlight annotations** สีเหลือง (pre-existing)
- **FreeText labels** เป็นตัวเลข "4.1, 4.2, 4.3..." (pre-existing — TOR reference numbers)

Sonnet Agent ที่ clone ไป sister files จะ inherit pre-existing annotations ทั้งหมด — **ไม่ต้องลบ** เพราะ user ICT auditor อาจอ้างอิงเลขเหล่านี้ แค่ตี rect + label เพิ่มของเรา (ใช้ "5.X.Y.Z ข้อย่อย N." แทนเลข ICT) ไปด้วย

### Sonnet Agent miss issue (เคยพบ)

Sonnet Agent ที่ spawn สำหรับ CCTV clone **ไม่ได้ใส่ header** ทุกหน้าตามที่ prompt ระบุ → ต้อง verify หลัง spawn + เพิ่ม header เอง:

```python
# Detection: header = FreeText with "หน้า" + y < 60
# Re-add header for missing pages
for pi in needs_header:
    page = doc[pi]
    page.add_freetext_annot(
        fitz.Rect(15, 10, pw-15, 50),
        f"{filename} หน้า {pi+1}",
        fontsize=10, fontname="hebo", text_color=(1, 0, 0),
        align=fitz.TEXT_ALIGN_CENTER)
```

### Final Status (2026-05-09)

| Status | Count |
|---|---|
| **รอ user ตรวจสอบ** | **637** (97%) |
| Section header (เว้นว่าง) | 23 |
| **รอ catalog** | **0** ✅ |

**101 PDFs annotated** ใน output/, ครอบคลุมทุก section (5.1.1 - 5.1.8, 5.2.1 - 5.2.6)

**ดูรายละเอียด status + cross-reference:** `knowledge_base/sections.json` + `knowledge_base/KB.md`

---

## อัพเดต 2026-05-10 — TRIO_SR_Solution + SR Pattern Adoption

### บริบทใหม่: 2 Consortium proposals
หลัง user แจ้ง Trio Sr Solution + Take IT เป็น 2 consortia แยกข้อเสนอ:
- `output/TRIO_SR_Solution/` (ใช้ catalog จาก SR เป็นหลัก)
- `output/Take_IT/` (ใช้ catalog ที่เราทำเองทั้งหมด)
- `output/Comply spec Smart Plant 1.xlsx` master = 660 rows
- xlsx ของแต่ละ consortium มี sheet tab ของตัวเอง

### SR Pattern (กฎข้อ 9) — ใช้กับ catalog ที่เราทำเอง
**Hybrid annotation:** highlight (text-based PDF) + rect (image-based PDF) + **short callout `N)` พื้นหลังขาว**
- text density >500 chars/page → use `page.search_for(keyword)` + `add_highlight_annot` + callout right margin
- text density <100 chars/page → use rect + callout
- **callout format ใหม่:** `N)` หรือ `N.M)` (ไม่มี section prefix แล้ว) — สั้น พื้นหลังขาว ตัวอักษรแดง 9pt
- **Section prefix implicit จาก folder location** (ไม่ต้องเขียน `5.1.1.2 ข้อย่อย 1.` แบบเดิม)
- **Print clarity:** white fill + red text + no border = ชัดเจนเวลาพิมพ์

```python
ft = page.add_freetext_annot(small_rect, f"{n})", fontsize=9, fontname="hebo",
    text_color=(1, 0, 0), fill_color=(1, 1, 1),
    align=fitz.TEXT_ALIGN_CENTER)
ft.set_border(width=0)
ft.update()
```

### กฎข้อ 10: ห้ามแก้ catalog ของ SR
**Rule:** SR catalogs ใน `catalog/SR/extracted/` = read-only reference เท่านั้น
- **อ่าน** เพื่อศึกษา pattern (highlight + callout placement)
- **Copy** ไปที่ `output/TRIO_SR_Solution/` ตามต้นฉบับ
- **เพิ่ม header เท่านั้น** บนสุด ไม่ทับ annotations ของ SR
- **ห้าม** ลบ/แก้ highlights/rects/callouts ของ SR

### กฎข้อ 11 (NEW): "ไม่พบใน catalog" vs "ยินดีปฏิบัติฯ"
**ใช้ "ยินดีปฏิบัติตามข้อกำหนด"** เฉพาะ:
1. งานติดตั้ง / install (เดินสาย, ติดตั้ง, ตั้งค่า)
2. Software/Firmware ที่ทีมเขียนเอง (SCADA, dashboard, custom app)
3. Commitment statements (warranty, training, document delivery)

**ใช้ "ไม่พบใน catalog"** เมื่อ:
- ข้อย่อยเป็น hardware spec / product feature
- หา callout/page ใน catalog ไม่เจอ → flag ให้ user ตรวจ

User ทำต่อได้ 3 ทาง:
1. หา catalog page → annotate + เปลี่ยนเป็น "เทียบเท่าฯ"
2. เปลี่ยนรุ่นเป็นรุ่นที่ comply ได้
3. ยืนยันเป็น commitment จริง → revert "ยินดีปฏิบัติฯ"

### กฎข้อ 12 (CRITICAL): Continuity Document — เมื่อ context ใกล้เต็ม

**กฎ:** ถ้า context window ใกล้เต็ม (estimate >70% of max), **เสมอ** ทำ continuity document ก่อน auto-compaction

**ทำที่ไหน:** สร้างหรืออัพเดต `_continuity/STATE_<YYYYMMDD>.md` ครอบคลุม:

```markdown
# Continuity State — <date>

## Last completed task
- [task name + result]
- Files modified: [list with line counts]
- Snapshot ID: 2026-MM-DD_HHMMSS_...

## Open in-progress
- TodoWrite todos (copy-paste)
- Pending decisions waiting for user

## Critical context for next session
- Recent user corrections/preferences (verbatim quotes)
- Discovered bugs/workarounds (e.g., PyMuPDF xref_set_key trick)
- Path conventions in use
- Vendor sections currently being worked on

## Next planned action
- Single concrete next step (e.g., "Run /tmp/foo.py on 5 remaining files")
- Why this is the right next step

## Files to read first on resume
1. SKILL.md (always)
2. knowledge_base/KB.md (always)
3. _continuity/STATE_<YYYYMMDD>.md (this file)
4. <other context-specific files>
```

**Trigger ตอนทำ:**
- ก่อน auto-compact (sync hint)
- Long-running batch task ที่อาจใช้ context >5 turns
- User กล่าวว่า "หยุด" / "พอแค่นี้ก่อน" / "บันทึกสถานะ"
- หลัง snapshot สำเร็จที่เป็น milestone ใหญ่

**Don't:**
- ห้ามรอจน context overflow แล้วค่อยทำ
- ห้ามเขียนยาวเกิน — ต้องอ่านได้เร็วใน 30 วินาที
- ห้ามใส่ raw output ของ commands (ใส่แค่ summary)

### Bug fix ใหม่ที่ค้นพบ (PyMuPDF)

**1. delete_annot ไม่ persist กับ stored Annot ref**
```python
# ❌ ไม่ทำงาน — ลบไม่หาย หลัง save
a_list = []
a = page.first_annot
while a: a_list.append(a); a = a.next
for a in a_list: page.delete_annot(a)

# ✅ ทำงาน — clear /Annots array โดยตรง
for pi in range(doc.page_count):
    doc.xref_set_key(doc[pi].xref, "Annots", "[]")
# จากนั้น re-add annotations ที่ต้องการ
```

**2. "annotation not bound to any page" — ต้อง keep page ref alive**
```python
# ❌ page ถูก garbage-collect ระหว่าง iter annotations
a = doc[pi].first_annot

# ✅ keep page ref
page = doc[pi]
a = page.first_annot
```

### TRIO_SR_Solution Final State (2026-05-10)

```
PDFs (104 ทั้งหมด):
├── 75 SR pattern (highlight/rect + short callout N))
├── 27 brand-marker only (ยี่ห้อ/รุ่น single-row)
├── 2 placeholder (BOD/DO Sensor SR proposal 1-page summary)
└── 0 empty / 0 duplicate / 0 long label

xlsx Col D distribution (660 rows):
├── 309 (46.8%) เทียบเท่าข้อกำหนด ✅
├── 115 (17.4%) ไม่พบใน catalog ⚠ (flag for user)
├── 77 (11.7%) ยินดีปฏิบัติฯ (real install/software)
├── 58 (8.8%) ยี่ห้อ-รุ่น
├── 45 (6.8%) filename ref (single-row)
├── 30 (4.5%) สูงกว่าข้อกำหนด
└── 23 (3.5%) section header (empty)

Cross-ref check: 339/340 verified (100% data rows)
```

**Snapshot:** `2026-05-10_111357_Add--ไม-พบใน-catalog--rule...`
