This guide summarizes essential steps and tools for performing static analysis on documents, especially those suspected to contain malicious content. Each step is explained to help you understand its purpose and how to execute it.
  ```

- **olevba:**  
  Extracts and analyzes VBA macros.  
  **Important Note:** `olevba` will output any embedded VBA code found in the document.  
  - **Always carefully review the extracted VBA code for suspicious functions, obfuscated strings, auto-execution macros (e.g., `AutoOpen`, `Document_Open`), and calls to external processes (like `Shell`, `CreateObject`, `WScript.Shell`).**  
  - Malicious VBA macros often use obfuscation and may attempt to download and execute payloads or manipulate system files.
  
  ```sh
  olevba filename
  ```

---

## Additional Analysis Points

In addition to the above steps, consider these best practices:

- **Check for embedded objects** (e.g., images, executables, scripts) which may be hidden or disguised.
- **Look for suspicious file extensions or double extensions** (e.g., `.doc.exe`, `.pdf.scr`).
- **Review document properties** for inconsistencies (e.g., odd author names, strange template references).
- **Analyze file structure** using hex editors or forensic tools to spot anomalies.
- **Be cautious of password-protected or encrypted documents**, which may attempt to bypass detection tools.
- **Search for URLs and domains** and check their reputation with threat intelligence services.
- **Check for signs of social engineering** in document content, such as instructions to enable macros or click suspicious links.

---

## Explanations for Each Tool

- **unzip:** Extracts files for further analysis.
- **VirusTotal:** Online service for scanning files/hashes/IPs for malware.
- **exiftool:** Extracts metadata from various file formats.
- **strings:** Finds readable text in binary files.
- **xorsearch:** Detects obfuscated (XORed) content, often used by malware.
- **olemeta, oleid, olevba:** Analyze Microsoft Office document features and macros for signs of exploitation.  
  - **olevba:** Extracts VBA macro code that must be thoroughly analyzed for suspicious or malicious behavior.

---

## Portfolio Message

These steps and tools provide a solid foundation for static document analysis, allowing you to spot potential malware and security risks before executing or sharing unknown files. Integrate these into your workflow for a safer and more thorough investigation process.
