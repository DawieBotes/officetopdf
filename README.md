# 📝 Office to PDF & Excel Recalculator API

A simple Flask-based API in Docker that uses LibreOffice UNO to:

- 📄 Convert `.doc`, `.docx`, `.xls`, `.xlsx` files to PDF
- 🔁 Recalculate all formulas in `.xls`, `.xlsx` Excel files and return the updated Excel file

---

## 📦 Docker Usage

### 1. Build the Docker Image

```bash
docker build -t officeconverttopdf .
```

### 2. Run the Container

```bash
docker run -p 5000:5000 -it officeconverttopdf
```

> If you're using Docker Compose, just run:

```bash
docker-compose up
```

---

## 🚀 API Endpoints

### 🧾 `/convert` — Convert Office file to PDF

**Method:** `POST`  
**Content-Type:** `multipart/form-data`  
**Form field:** `file`

**Supported formats:** `.doc`, `.docx`, `.xls`, `.xlsx`

**Example with curl:**

```bash
curl -X POST http://localhost:5000/convert   -F "file=@/path/to/your/file.xlsx"   --output output.pdf
```

**Returns:** PDF version of the uploaded file

---

### 🔁 `/recalc` — Recalculate and Return Excel

**Method:** `POST`  
**Content-Type:** `multipart/form-data`  
**Form field:** `file`

**Supported formats:** `.xls`, `.xlsx`

**Example with curl:**

```bash
curl -X POST http://localhost:5000/recalc   -F "file=@/path/to/your/formulas.xlsx"   --output recalculated.xlsx
```

**Returns:** A new `.xls` or `.xlsx` file with formulas recalculated

---

## 💡 Health Check

**Endpoint:** `GET /health`  
**Example:**

```bash
curl http://localhost:5000/health
```

**Response:**
```json
{"status": "healthy"}
```

---

## 📂 Project Structure

```
.
├── Dockerfile
├── docker-compose.yml
├── app.py               # Flask API
├── convert.py           # LibreOffice UNO conversion/recalculation logic
├── /files               # Temporary input/output files (inside container)
```

---

## 🛠 Requirements (handled in Dockerfile)

- Python 3
- Flask
- LibreOffice (headless mode)
- PyUNO (LibreOffice UNO bindings)

---

## 🔐 Security & Notes

- All uploaded files are deleted after processing
- LibreOffice runs in headless mode via UNO socket on port 2002
- You can extend support for additional formats like `.odt`, `.ppt`, `.csv`, etc.

---

## ✨ Future Ideas

- Support async processing or queuing with Celery/RQ
- Add size limits or virus scanning (e.g., ClamAV)
- Health monitoring for LibreOffice socket stability

---

## 🧑‍💻 Author

Built with ❤️ for automated document processing.