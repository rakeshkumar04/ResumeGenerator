import openpyxl

class SummaryGenerator:
    def __init__(self, excel_file, output_tex):
        self.excel_file = excel_file
        self.output_tex = output_tex
        self.data = ""
        self.load_data()

    def load_data(self):
        wb = openpyxl.load_workbook(self.excel_file)
        ws = wb['summary']  # Make sure the sheet is named 'summary'
        summary_data = [cell.value for row in ws.iter_rows(values_only=True) for cell in row]
        self.data = " ".join(summary_data)

    def generate_tex(self):
        with open(self.output_tex, 'w') as f:
            f.write(r"""
            %-------------------------------------------------------------------------------
            % SECTION TITLE
            %-------------------------------------------------------------------------------
            \cvsection{Summary}
            %-------------------------------------------------------------------------------
            % CONTENT
            %-------------------------------------------------------------------------------
            \begin{cvparagraph}
            %---------------------------------------------------------
            """)
            f.write(self.data + "\n")
            f.write(r"""
            \end{cvparagraph}
            """)

if __name__ == "__main__":
    excel_file = "data.xlsx"  # Replace with your Excel file path
    output_tex = "resume/summary.tex"
    
    generator = SummaryGenerator(excel_file, output_tex)
    generator.generate_tex()
