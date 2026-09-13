import io
import glob
import os
from fpdf import FPDF
import openpyxl
import pandas as pd
import streamlit as st


def generate_excel_detail(df_detail, title_text):
  """Membuat file Excel dalam bentuk BytesIO untuk detail aset terpilih."""
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_detail.to_excel(writer, index=False, sheet_name="Detail Aset")
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