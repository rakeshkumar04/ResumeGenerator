import openpyxl

class ExperienceGenerator:
    def __init__(self, excel_file, output_tex):
        self.excel_file = excel_file
        self.output_tex = output_tex
        self.data = []
        self.load_data()

    def load_data(self):
        wb = openpyxl.load_workbook(self.excel_file)
        ws = wb['experience']  # Make sure the sheet is named 'experience'
        for row in ws.iter_rows(min_row=2, values_only=True):
            entry = {
                "organization": row[0],
                "job": row[1],
                "location": row[2],
                "start_date": row[3],
                "end_date": row[4],
                "responsibilities": row[5:]
            }
            self.data.append(entry)

    def generate_tex(self):
        with open(self.output_tex, 'w') as f:
            f.write(r"""
            %-------------------------------------------------------------------------------
            % SECTION TITLE
            %-------------------------------------------------------------------------------
            \cvsection{Work Experience}
            %-------------------------------------------------------------------------------
            % CONTENT
            %-------------------------------------------------------------------------------
            \begin{cventries}
            """)
            
            for entry in self.data:
                f.write(r"""
                %---------------------------------------------------------
                \cventry
                {""" + entry["organization"] + r"""} % Organization
                {""" + entry["job"] + r"""} % Job
                {""" + entry["location"] + r"""} % Location
                {""" + entry["start_date"] + r"""} % Start Date(s)
                {""" + entry["end_date"] + r"""} % End Date(s)
                { \begin{cvitems} % Description(s) of tasks/responsibilities
                """)
                for responsibility in entry["responsibilities"]:
                    if responsibility:  # Check if the responsibility is not None
                        f.write(f"\t\item {{{responsibility}}}\n")
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
    output_tex = "resume/experience.tex"
    
    generator = ExperienceGenerator(excel_file, output_tex)
    generator.generate_tex()
