import openpyxl

class CompetenciesGenerator:
    def __init__(self, excel_file, output_tex):
        self.excel_file = excel_file
        self.output_tex = output_tex
        self.data = []
        self.load_data()

    def load_data(self):
        wb = openpyxl.load_workbook(self.excel_file)
        ws = wb['competencies']  # Make sure the sheet is named 'competencies'
        self.data = [cell.value for row in ws.iter_rows(values_only=True) for cell in row]

    def generate_tex(self):
        with open(self.output_tex, 'w') as f:
            f.write(r"""
            % -------------------------------------------------------------------------------
            % SECTION TITLE
            % -------------------------------------------------------------------------------
            \cvsection{Key Competencies}
            % -------------------------------------------------------------------------------
            % CONTENT
            % -------------------------------------------------------------------------------
            % \begin{cvskills}
            % ---------------------------------------------------------
            % \cvskill
            \cventry
            {}
            {}
            {}
            {}
            {
            \begin{cvitems}
            """)
            for item in self.data:
                f.write(f"\t\item{{{item}}}\n")
            f.write(r"""
            \end{cvitems}
            }
            %---------------------------------------------------------
            %---------------------------------------------------------
            % \end{cvskills}
            """)

if __name__ == "__main__":
    excel_file = "data.xlsx"  # Replace with your Excel file path
    output_tex = "resume/competencies.tex"
    
    generator = CompetenciesGenerator(excel_file, output_tex)
    generator.generate_tex()
