import io
from docx import Document


def generate_excel_detail(df_detail, title_text):
  """Membuat file Excel dalam bentuk BytesIO untuk detail aset terpilih."""
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_detail.to_excel(writer, index=False, sheet_name="Detail Aset")
  output.seek(0)
  return output


def generate_word_detail(df_detail, title_text):
  """Membuat file Word (.docx) untuk detail aset terpilih."""
  doc = Document()
  doc.add_heading(f"DETAIL ASET: {title_text}", level=1)
  doc.add_paragraph(
      "Dokumen ini dicetak otomatis dari sistem manajemen aset SIMANTAP."
  )

  table = doc.add_table(rows=1, cols=2)
  table.style = "Table Grid"
  hdr_cells = table.rows[0].cells
  hdr_cells[0].text = "Atribut / Kolom"
  hdr_cells[1].text = "Keterangan"

  for _, row in df_detail.iterrows():
    row_cells = table.add_row().cells
    row_cells[0].text = str(row["Atribut / Kolom Data"])
    row_cells[1].text = str(row["Keterangan / Isi Data"])

  output = io.BytesIO()
  doc.save(output)
  output.seek(0)
  return output


def generate_pdf_detail(df_detail, title_text):
  """Membuat file PDF sederhana menggunakan FPDF."""
  pdf = FPDF()
  pdf.add_page()
  pdf.set_font("Arial", "B", 14)
  pdf.cell(0, 10, "DETAIL INFORMASI ASET DAERAH", 0, 1, "C")
  pdf.set_font("Arial", "", 10)
  pdf.cell(0, 8, f"Modul: {title_text}", 0, 1, "C")
  pdf.ln(5)

  pdf.set_font("Arial", "B", 10)
  pdf.set_fill_color(200, 220, 240)
  pdf.cell(95, 8, "Atribut / Kolom Data", 1, 0, "C", True)
  pdf.cell(95, 8, "Keterangan / Isi Data", 1, 1, "C", True)

  pdf.set_font("Arial", "", 9)
  for _, row in df_detail.iterrows():
    col_name = str(row["Atribut / Kolom Data"])[:50]
    col_val = str(row["Keterangan / Isi Data"])[:60]
    pdf.cell(95, 6, col_name, 1, 0, "L")
    pdf.cell(95, 6, col_val, 1, 1, "L")

  pdf_output = pdf.output(dest="S")
  if isinstance(pdf_output, str):
    pdf_output = pdf_output.encode("latin1", "replace")
  return io.BytesIO(pdf_output)