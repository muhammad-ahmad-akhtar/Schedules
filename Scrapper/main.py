from processor import PDFScheduleProcessor

def main():
    pdf_path = "faculty.pdf"
    output_path = "faculty.json"

    processor = PDFScheduleProcessor(pdf_path)
    schedules = processor.process_pdf()
    #processor.export_to_json(output_path)
    processor.export_to_excel("CS_Schedules.xlsx")

    print(f"Schedules extracted and saved to {output_path}")

if __name__ == "__main__":
    main()
