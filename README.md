# DYMO Label Web Editor
A web-based editor for DYMO LabelWriter printers. Supports both legacy `.label` files and modern `.dymo` (DYMO Connect) files.

## Hosting
This project acts as a static website and is designed to be hosted on **GitHub Pages**.

## Local Development

A lightweight web application for uploading DYMO `.label` files, editing text fields, previewing labels, and printing to a DYMO LabelWriter.

## Prerequisites
- **DYMO Connect Software** (recommended) or **DYMO Label Software** installed
- **DYMO Web Service** running (check system tray/menu bar icon)
- A DYMO LabelWriter printer

## Installation & Usage

Since this is a static website, you can run it directly:

1.  **Clone or Download** this repository.
2.  **Open** `index.html` file directly in your browser.
    - *Note: Some browsers block local file access to the DYMO service. If so, use a simple static server.*

### Running a Local Server (Optional)
If you have Python installed:
```bash
python3 -m http.server 8080
```
Then open `http://localhost:8080` in your browser.

## Usage

1. **Check Status** - The app will show a green indicator if the DYMO Web Service is connected
// ... steps continue ...

## Creating Label Templates

1. Open DYMO Label Software or DYMO Connect
2. Create a new label and add text objects
3. Name your text objects (right-click → Properties → Name)
4. Save as `.label` file

## Troubleshooting

### DYMO Web Service Not Detected
- Ensure DYMO Connect or DYMO Label Software is installed
- Check if the DYMO Web Service is running (look for icon in system tray/menu bar)
- Try restarting the DYMO Web Service
- Visit `https://127.0.0.1:41951/DYMO/DLS/Printing/Check` to test directly

### Certificate Errors
- The DYMO Web Service uses a self-signed HTTPS certificate
- This should work automatically if DYMO software installed correctly
- If issues persist, try reinstalling DYMO Connect Software

## License

ISC
