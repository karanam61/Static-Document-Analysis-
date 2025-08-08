## 📂 Digital Forensics – Static Document Analysis Challenge

### 🎯 Purpose of This Analysis
The goal of this challenge was to statically analyze suspicious Microsoft Office documents without executing them, in order to:

- Detect the presence of VBA macros that could execute malicious code.
- Identify auto-execution triggers (e.g., Document_Open) used to launch attacks.
- Extract and review metadata (author, timestamps) that could assist in attribution or timeline building.
- Find Indicators of Compromise (IOCs) such as URLs, domains, and file hashes.
- Practice using forensic tools (exiftool, olevba, oleid, olemeta) commonly applied in malware triage and incident response.

This process simulates real-world document malware investigations, where opening a malicious file in a live environment could infect the system.  
By using static analysis tools, we can safely gather evidence without executing any harmful code.

---

### 1️⃣ Does the file `/root/Desktop/QuestionFiles/PO-465514-180820.doc` contain a VBA macro?

**Method Used:**
```bash
exiftool PO-465514-180820.doc
```

**Observation:**
- The Template field shows `Normal.dotm`, which is macro-enabled.

**Screenshot:**  
![Q1](images/Q1.png)
![q1.1](images/q1.1.png)

**Answer:** ✅ Yes, the file contains a VBA macro.

---

### 2️⃣ Macro keyword enabling malicious activity

**Method Used:**
```bash
olevba PO-465514-180820.doc
```
Extracted macro code and looked for auto-execution keywords.

**Observation:**
- Found `Document_Open` in the first line — runs automatically when the document is opened.

**Answer:** 🔑 Macro Keyword: `Document_Open`

**Screenshot Proof:**  
![Q2](images/Q2.png)
![Q2.1](images/Q2.1.png)

---

### 3️⃣ Who is the author of the file?

**Method Used:**
```bash
oleid PO-465514-180820.doc
```
Retrieved metadata containing author information.


**Screenshot Proof:**  
![Q3](images/Q3.png)


---

### 4️⃣ Last saved time of the file

**Method Used:**
```bash
olemeta PO-465514-180820.doc
```
Displayed metadata including LastSavedTime.

**Answer:** ⏳ Last Saved Time: **(Timestamp from output)**

**Screenshot Proof:**  
![Q4](images/Q4.png)
![Q4.1](images/Q4.1.png)

---

### 5️⃣ Domain from which the malicious XLS file tries to download  
File: `/root/Desktop/QuestionFiles/Siparis_17.xls`

**Method Used:**
```bash
olevba --d Siparis_17.xls
```
Decoded macros to reveal any embedded URLs.

Verified findings by checking the file hash on VirusTotal.

**Observation:**
- Found IOC: `http://hocosco.mobi`
- VirusTotal confirmed similar domain activity.

**Answer:** 🌐 Domain: `hocosco.mobi`

**Screenshot Proof:**  
![Q5](images/Q5.png)
![Q5.1](images/Q5.1.png)
![Q5.2](images/Q5.2.png)



---

### 6️⃣ Number of IOCs in the XLS file according to olevba

**Method Used:**
```bash
olevba Siparis_17.xls
```
Checked the IOC extraction table in the output.

**Observation:**
- Detected **2 Indicators of Compromise**.

**Answer:** 📌 IOCs Count: **2**

**Screenshot Proof:**  
![Q6](images/Q6.png)

---

### 🛠 Tools Used
- **ExifTool** – Extracts file metadata.
- **Oletools Suite:**
  - **oleid** – High-level OLE document analysis.
  - **olemeta** – Extracts document metadata.
  - **olevba** – Extracts and analyzes VBA macros, IOCs, and auto-execution triggers.

---

### ⚠️ Disclaimer
All analysis was performed in a controlled, isolated lab environment.  
Never open suspicious files on your main system.
