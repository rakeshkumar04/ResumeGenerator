import openpyxl

class EducationGenerator:
    def __init__(self, excel_file, output_tex):
        self.excel_file = excel_file
        self.output_tex = output_tex
        self.data = []
        self.load_data()

    def load_data(self):
        wb = openpyxl.load_workbook(self.excel_file)
        ws = wb['education']  # Make sure the sheet is named 'education'
        for row in ws.iter_rows(min_row=2, values_only=True):
            entry = {
                "college": row[0],
                "education": row[1],
                "location": row[2],
                "start_date": row[3],
                "end_date": row[4],
                "description": row[5:]
            }
            self.data.append(entry)

    def generate_tex(self):
        with open(self.output_tex, 'w') as f:
            f.write(r"""
            %-------------------------------------------------------------------------------
            % SECTION TITLE
            %-------------------------------------------------------------------------------
            \cvsection{Education}
            %-------------------------------------------------------------------------------
            % CONTENT
            %-------------------------------------------------------------------------------
            \begin{cventries}
            """)
            
            for entry in self.data:
                f.write(r"""
                %---------------------------------------------------------
                \cventry
                {""" + entry["college"] + r"""} % College
                {""" + entry["education"] + r"""} % Education
                {""" + entry["location"] + r"""} % Location
                {""" + entry["start_date"] + r"""} % Start Date(s)
                {""" + entry["end_date"] + r"""} % End Date(s)
                { \begin{cvitems} % Description(s) of tasks/responsibilities
                """)
                for desc in entry["description"]:
                    if desc:  # Check if the description is not None
                        f.write(f"\t\item {{{desc}}}\n")
                f.write(r"""
                \end{cvitems}
                }
                %---------------------------------------------------------
                """)

            f.write(r"""
            \end{cventries}
            """)

if __name__ == "__main__":
    excel_file = "data.xlsx"  # Replace with your Excel file path
    output_tex = "resume/education.tex"
    
    generator = EducationGenerator(excel_file, output_tex)
    generator.generate_tex()
