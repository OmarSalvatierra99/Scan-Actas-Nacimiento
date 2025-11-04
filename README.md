# Actas QR — Mass Scanner for Mexican Birth Certificates

**Actas QR** is a specialized web application developed for the **Superior Audit Office of Tlaxcala (OFS)** to streamline the verification and processing of Mexican birth certificates (*actas de nacimiento*).  
It supports scanning via **web camera**, **connected scanners**, or **bulk PDF uploads**, automatically reading and validating official QR codes from the Civil Registry.

🔗 **Live:** [https://actas.omar-xyz.shop](https://actas.omar-xyz.shop)

---

## Features

- **QR-based validation** — Extracts and verifies official data directly from acta QR codes.  
- **Multiple input options** —  
  - Scan using a connected **webcam**  
  - Use an external **scanner**  
  - Upload **multiple PDFs** at once  
- **Automatic data extraction** — Reads and parses certificate data fields (name, CURP, date, registration number).  
- **Mass processing** — Handles hundreds of documents in a single batch.  
- **Institutional report generation** — Exports results in standardized Excel or CSV format.  
- **Responsive web interface** — Lightweight and functional, built for office environments.

---

## Tech Stack

- **Frontend:** HTML, CSS, JavaScript  
- **Backend:** Flask (Python)  
- **PDF & QR Processing:** PyMuPDF, qrcode, pdf2image  
- **Batch Handling:** Pandas, OpenPyXL  
- **Deployment:** Gunicorn + Nginx on Linux

---

## Usage

1. Open [https://actas.omar-xyz.shop](https://actas.omar-xyz.shop).  
2. Choose an input method:  
   - “**Camera Scan**” to use your webcam  
   - “**Connect Scanner**” for physical document scanning  
   - “**Upload PDFs**” to process digital files in bulk  
3. The system reads the embedded QR codes and validates each certificate.  
4. Download the Excel report with the parsed data and verification results.

---

## Institutional Context

This tool is part of the **OFS Tlaxcala** automation suite, designed to simplify and standardize document validation workflows across government offices.  
Developed by **Omar Gabriel Salvatierra García** — 2025.  

© 2025 OFS Tlaxcala — Institutional Software  
